/**
 * Desktop Office Bridge (`DesktopOfficeBridge`)
 *
 * Connects the Algerian Notary Office workspace (`OfficeStore`) to:
 * 1. Native desktop shells (`window.__NOTAIRE_DESKTOP__` injected by Tauri / Electron preload)
 * 2. Local File System Access API (`window.showDirectoryPicker`) when a local office folder
 *    is mounted by the notary in Chromium desktop browsers, storing collections as JSON
 *    and sealed originals (`originals/{year}/index_{N}.docx` + `.sha256`) directly on disk.
 * 3. Persists the `FileSystemDirectoryHandle` across page reloads in a dedicated handle store
 *    (`NotaryDesktopBridgeMetaDB`) so the notary can reconnect or stay connected across reloads.
 */
export interface DesktopOfficeBridge {
  getDataDirectory(): Promise<string>;
  invoke<T>(command: string, payload: unknown): Promise<T>;
}

export interface FileSystemDirectoryHandleLike {
  name: string;
  queryPermission?: (descriptor?: {
    mode?: 'read' | 'readwrite';
  }) => Promise<'granted' | 'denied' | 'prompt'>;
  requestPermission?: (descriptor?: {
    mode?: 'read' | 'readwrite';
  }) => Promise<'granted' | 'denied' | 'prompt'>;
  getDirectoryHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<FileSystemDirectoryHandleLike>;
  getFileHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<{
    getFile(): Promise<{
      text(): Promise<string>;
      arrayBuffer(): Promise<ArrayBuffer>;
    }>;
    createWritable(): Promise<{
      write(data: string | BufferSource | Blob): Promise<void>;
      close(): Promise<void>;
    }>;
  }>;
}

export interface LocalOfficeDirectoryState {
  status: 'disconnected' | 'needs_permission' | 'connected' | 'native_desktop';
  folderName: string | null;
}

let mountedLocalFolderBridge: DesktopOfficeBridge | null = null;
let mountedFolderName: string | null = null;
let pendingRestoredHandle: FileSystemDirectoryHandleLike | null = null;

const HANDLE_DB_NAME = 'NotaryDesktopBridgeMetaDB';
const HANDLE_STORE_NAME = 'handles';
const HANDLE_KEY = 'office_root_dir';

function openHandleMetaDB(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === 'undefined') return Promise.resolve(null);
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(HANDLE_DB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(HANDLE_STORE_NAME)) {
          db.createObjectStore(HANDLE_STORE_NAME);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function savePersistedDirectoryHandle(
  handle: FileSystemDirectoryHandleLike
): Promise<void> {
  const db = await openHandleMetaDB();
  if (!db) return;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(HANDLE_STORE_NAME, 'readwrite');
      tx.objectStore(HANDLE_STORE_NAME).put(handle, HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch {
      resolve();
    }
  });
}

async function loadPersistedDirectoryHandle(): Promise<FileSystemDirectoryHandleLike | null> {
  const db = await openHandleMetaDB();
  if (!db) return null;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(HANDLE_STORE_NAME, 'readonly');
      const req = tx.objectStore(HANDLE_STORE_NAME).get(HANDLE_KEY);
      req.onsuccess = () =>
        resolve((req.result as FileSystemDirectoryHandleLike) || null);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function clearPersistedDirectoryHandle(): Promise<void> {
  const db = await openHandleMetaDB();
  if (!db) return;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(HANDLE_STORE_NAME, 'readwrite');
      tx.objectStore(HANDLE_STORE_NAME).delete(HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch {
      resolve();
    }
  });
}

export function createLocalFileSystemBridge(
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
          args.relativePath || `archive/${year}/${documentId}.docx`
        );
        const sha256 = String(args.sha256 || '');
        const byteLength = Number(args.byteLength) || 0;
        const base64Data = String(args.base64Data || '');
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

        // Write companion .sha256 file on disk for external verification
        if (sha256) {
          const shaHandle = await yearDir.getFileHandle(`${fileName}.sha256`, {
            create: true,
          });
          const shaWritable = await shaHandle.createWritable();
          await shaWritable.write(`${sha256}  ${fileName}\n`);
          await shaWritable.close();
        }

        // Record in `sealed_originals` collection (compatible with exportFullBackupBundle & read_original)
        const sealedRecords = await readCollectionJson<{
          id: string;
          documentId: string;
          year: number;
          fileName: string;
          relativePath: string;
          sha256: string;
          byteLength: number;
          base64Data: string;
          sealedAt: string;
        }>('sealed_originals');

        const nextRecord = {
          id: documentId,
          documentId,
          year,
          fileName,
          relativePath,
          sha256,
          byteLength: docxBytes ? docxBytes.byteLength : byteLength,
          base64Data,
          sealedAt: new Date().toISOString(),
        };
        const existingIdx = sealedRecords.findIndex(
          (m) => m.id === documentId || m.documentId === documentId
        );
        if (existingIdx >= 0) sealedRecords[existingIdx] = nextRecord;
        else sealedRecords.push(nextRecord);
        await writeCollectionJson('sealed_originals', sealedRecords);

        return { relativePath, sha256 } as unknown as T;
      }

      if (command === 'read_original') {
        const documentId = String(args.documentId || '');
        const sealedRecords = await readCollectionJson<{
          id: string;
          documentId: string;
          year: number;
          fileName: string;
          relativePath: string;
          sha256: string;
          base64Data?: string;
        }>('sealed_originals');
        const meta = sealedRecords.find(
          (m) => m.id === documentId || m.documentId === documentId
        );
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
          if (meta.base64Data) {
            return {
              relativePath: meta.relativePath,
              sha256: meta.sha256,
              base64Data: meta.base64Data,
            } as unknown as T;
          }
          return null as unknown as T;
        }
      }

      if (command === 'backup_database') {
        return { ok: true } as unknown as T;
      }

      return null as unknown as T;
    },
  };
}

/**
 * Attempts to restore a previously saved `FileSystemDirectoryHandle` on startup.
 * If permission is already `'granted'`, activates the local folder bridge immediately.
 * If permission is `'prompt'`, caches the handle in `pendingRestoredHandle` so a single user click
 * on `reconnectPersistedOfficeDirectory()` re-authorizes it without reopening the directory picker.
 */
export async function tryRestorePersistedOfficeDirectory(): Promise<LocalOfficeDirectoryState> {
  if (typeof window === 'undefined') {
    return { status: 'disconnected', folderName: null };
  }
  const nativeBridge = (
    window as unknown as { __NOTAIRE_DESKTOP__?: DesktopOfficeBridge }
  ).__NOTAIRE_DESKTOP__;
  if (nativeBridge) {
    return { status: 'native_desktop', folderName: 'سطح المكتب' };
  }

  const savedHandle = await loadPersistedDirectoryHandle();
  if (!savedHandle) {
    return { status: 'disconnected', folderName: null };
  }

  try {
    const perm = savedHandle.queryPermission
      ? await savedHandle.queryPermission({ mode: 'readwrite' })
      : 'granted';
    if (perm === 'granted') {
      mountedLocalFolderBridge = createLocalFileSystemBridge(savedHandle);
      mountedFolderName = savedHandle.name;
      pendingRestoredHandle = null;
      return { status: 'connected', folderName: savedHandle.name };
    }
    pendingRestoredHandle = savedHandle;
    mountedFolderName = savedHandle.name;
    return { status: 'needs_permission', folderName: savedHandle.name };
  } catch {
    return { status: 'disconnected', folderName: null };
  }
}

/**
 * Re-authorizes a previously persisted directory handle (`pendingRestoredHandle`) on user click,
 * or opens the directory picker if none is cached.
 */
export async function reconnectPersistedOfficeDirectory(): Promise<LocalOfficeDirectoryState> {
  if (pendingRestoredHandle) {
    try {
      const perm = pendingRestoredHandle.requestPermission
        ? await pendingRestoredHandle.requestPermission({ mode: 'readwrite' })
        : 'granted';
      if (perm === 'granted') {
        mountedLocalFolderBridge = createLocalFileSystemBridge(
          pendingRestoredHandle
        );
        mountedFolderName = pendingRestoredHandle.name;
        pendingRestoredHandle = null;
        return { status: 'connected', folderName: mountedFolderName };
      }
    } catch {
      // Fall through to picker
    }
  }
  const mounted = await mountLocalOfficeDirectory();
  if (mounted) {
    return { status: 'connected', folderName: mountedFolderName };
  }
  return getLocalOfficeDirectoryState();
}

/**
 * Prompts the user to mount a local office directory on disk using the File System Access API
 * (`window.showDirectoryPicker`) so contracts and sealed `.docx` originals are written directly to disk,
 * and persists the handle in `NotaryDesktopBridgeMetaDB` across reloads.
 */
export async function mountLocalOfficeDirectory(): Promise<string | null> {
  if (typeof window === 'undefined') return null;
  const win = window as unknown as {
    showDirectoryPicker?: (options?: {
      mode?: 'read' | 'readwrite';
    }) => Promise<FileSystemDirectoryHandleLike>;
  };
  if (typeof win.showDirectoryPicker !== 'function') {
    throw new Error(
      'متصفحك الحالي لا يدعم ربط مجلد محلي مباشر (File System Access API). استخدم متصفح Chromium (Chrome / Edge) أو النسخ الاحتياطي JSON.'
    );
  }
  const handle = await win.showDirectoryPicker({ mode: 'readwrite' });
  mountedLocalFolderBridge = createLocalFileSystemBridge(handle);
  mountedFolderName = handle.name;
  pendingRestoredHandle = null;
  await savePersistedDirectoryHandle(handle);
  return `local-fs://${handle.name}`;
}

export async function unmountLocalOfficeDirectory(): Promise<void> {
  mountedLocalFolderBridge = null;
  mountedFolderName = null;
  pendingRestoredHandle = null;
  await clearPersistedDirectoryHandle();
}

export function getLocalOfficeDirectoryState(): LocalOfficeDirectoryState {
  if (typeof window !== 'undefined') {
    const nativeBridge = (
      window as unknown as { __NOTAIRE_DESKTOP__?: DesktopOfficeBridge }
    ).__NOTAIRE_DESKTOP__;
    if (nativeBridge) {
      return { status: 'native_desktop', folderName: 'سطح المكتب' };
    }
  }
  if (mountedLocalFolderBridge) {
    return { status: 'connected', folderName: mountedFolderName };
  }
  if (pendingRestoredHandle) {
    return { status: 'needs_permission', folderName: mountedFolderName };
  }
  return { status: 'disconnected', folderName: null };
}

export function getDesktopBridge(): DesktopOfficeBridge | null {
  if (typeof window === 'undefined') return null;
  const nativeBridge = (
    window as unknown as { __NOTAIRE_DESKTOP__?: DesktopOfficeBridge }
  ).__NOTAIRE_DESKTOP__;
  if (nativeBridge) return nativeBridge;
  return mountedLocalFolderBridge;
}
