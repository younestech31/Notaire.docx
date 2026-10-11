/**
 * Desktop Office Bridge (`DesktopOfficeBridge`)
 *
 * Connects the Algerian Notary Office workspace (`OfficeStore`) to:
 * 1. Native desktop shells (`window.__NOTAIRE_DESKTOP__` injected by Tauri / Electron preload)
 * 2. Local File System Access API (`window.showDirectoryPicker`) when a local office folder
 *    is mounted by the notary in Chromium desktop browsers, storing collections as JSON
 *    and sealed originals (`originals/{year}/index_{N}.docx` + `.sha256`) directly on disk.
 */
export interface DesktopOfficeBridge {
  getDataDirectory(): Promise<string>;
  invoke<T>(command: string, payload: unknown): Promise<T>;
}

interface FileSystemDirectoryHandleLike {
  name: string;
  getDirectoryHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<FileSystemDirectoryHandleLike>;
  getFileHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<{
    getFile(): Promise<File>;
    createWritable(): Promise<{
      write(data: string | BufferSource | Blob): Promise<void>;
      close(): Promise<void>;
    }>;
  }>;
}

let mountedLocalFolderBridge: DesktopOfficeBridge | null = null;

function createLocalFileSystemBridge(
  rootHandle: FileSystemDirectoryHandleLike
): DesktopOfficeBridge {
  const readCollectionJson = async <T extends { id?: string }>(
    collection: string
  ): Promise<T[]> => {
    try {
      const dbDir = await rootHandle.getDirectoryHandle('db', { create: true });
      const fileHandle = await dbDir.getFileHandle(`${collection}.json`, {
        create: false,
      });
      const file = await fileHandle.getFile();
      const text = await file.text();
      const parsed = JSON.parse(text);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  };

  const writeCollectionJson = async <T>(
    collection: string,
    items: T[]
  ): Promise<void> => {
    const dbDir = await rootHandle.getDirectoryHandle('db', { create: true });
    const fileHandle = await dbDir.getFileHandle(`${collection}.json`, {
      create: true,
    });
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(items, null, 2));
    await writable.close();
  };

  return {
    async getDataDirectory(): Promise<string> {
      return `local-fs://${rootHandle.name}`;
    },
    async invoke<T>(command: string, payload: unknown): Promise<T> {
      const args = (payload || {}) as Record<string, unknown>;

      if (command === 'storage_load') {
        const collection = String(args.collection || '');
        const items = await readCollectionJson(collection);
        return items as unknown as T;
      }

      if (command === 'storage_save') {
        const collection = String(args.collection || '');
        const item = args.item as { id?: string };
        const items = await readCollectionJson<{ id?: string }>(collection);
        const idx = item?.id ? items.findIndex((x) => x.id === item.id) : -1;
        if (idx >= 0) {
          items[idx] = item;
        } else {
          items.push(item);
        }
        await writeCollectionJson(collection, items);
        return undefined as unknown as T;
      }

      if (command === 'storage_delete') {
        const collection = String(args.collection || '');
        const id = String(args.id || '');
        const items = await readCollectionJson<{ id?: string }>(collection);
        const filtered = items.filter((x) => x.id !== id);
        await writeCollectionJson(collection, filtered);
        return undefined as unknown as T;
      }

      if (command === 'storage_clear') {
        const collection = String(args.collection || '');
        await writeCollectionJson(collection, []);
        return undefined as unknown as T;
      }

      if (command === 'seal_original') {
        const documentId = String(args.documentId || '');
        const year = Number(args.year) || new Date().getFullYear();
        const relativePath = String(
          args.relativePath || `originals/${year}/${documentId}.docx`
        );
        const sha256 = String(args.sha256 || '');
        const docxBytes = args.docxBytes as Uint8Array | undefined;

        const originalsDir = await rootHandle.getDirectoryHandle('originals', {
          create: true,
        });
        const yearDir = await originalsDir.getDirectoryHandle(String(year), {
          create: true,
        });
        const fileName =
          relativePath.split('/').pop() || `${documentId}.docx`;

        if (docxBytes) {
          const docxHandle = await yearDir.getFileHandle(fileName, {
            create: true,
          });
          const writable = await docxHandle.createWritable();
          await writable.write(new Uint8Array(docxBytes));
          await writable.close();
        }

        // Record index metadata for read_original lookup
        const metaItems = await readCollectionJson<{
          id: string;
          year: number;
          fileName: string;
          relativePath: string;
          sha256: string;
        }>('sealed_originals_index');
        const nextMeta = {
          id: documentId,
          year,
          fileName,
          relativePath,
          sha256,
        };
        const existingIdx = metaItems.findIndex((m) => m.id === documentId);
        if (existingIdx >= 0) metaItems[existingIdx] = nextMeta;
        else metaItems.push(nextMeta);
        await writeCollectionJson('sealed_originals_index', metaItems);

        return { relativePath, sha256 } as unknown as T;
      }

      if (command === 'read_original') {
        const documentId = String(args.documentId || '');
        const metaItems = await readCollectionJson<{
          id: string;
          year: number;
          fileName: string;
          relativePath: string;
          sha256: string;
        }>('sealed_originals_index');
        const meta = metaItems.find((m) => m.id === documentId);
        if (!meta) return null as unknown as T;

        try {
          const originalsDir = await rootHandle.getDirectoryHandle(
            'originals',
            { create: false }
          );
          const yearDir = await originalsDir.getDirectoryHandle(
            String(meta.year),
            { create: false }
          );
          const fileHandle = await yearDir.getFileHandle(meta.fileName, {
            create: false,
          });
          const file = await fileHandle.getFile();
          const buf = await file.arrayBuffer();
          return {
            relativePath: meta.relativePath,
            sha256: meta.sha256,
            bytes: new Uint8Array(buf),
          } as unknown as T;
        } catch {
          return null as unknown as T;
        }
      }

      return null as unknown as T;
    },
  };
}

/**
 * Prompts the user to mount a local office directory on disk using the File System Access API
 * (`window.showDirectoryPicker`) so contracts and sealed `.docx` originals are written directly to disk.
 */
export async function mountLocalOfficeDirectory(): Promise<string | null> {
  if (typeof window === 'undefined') return null;
  const win = window as unknown as {
    showDirectoryPicker?: (options?: {
      mode?: 'read' | 'readwrite';
    }) => Promise<FileSystemDirectoryHandleLike>;
  };
  if (typeof win.showDirectoryPicker !== 'function') {
    return null;
  }
  const handle = await win.showDirectoryPicker({ mode: 'readwrite' });
  mountedLocalFolderBridge = createLocalFileSystemBridge(handle);
  return `local-fs://${handle.name}`;
}

export function unmountLocalOfficeDirectory(): void {
  mountedLocalFolderBridge = null;
}

export function getDesktopBridge(): DesktopOfficeBridge | null {
  if (typeof window === 'undefined') return null;
  const nativeBridge = (
    window as unknown as { __NOTAIRE_DESKTOP__?: DesktopOfficeBridge }
  ).__NOTAIRE_DESKTOP__;
  if (nativeBridge) return nativeBridge;
  return mountedLocalFolderBridge;
}
