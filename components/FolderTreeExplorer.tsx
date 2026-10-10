'use client';

import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Folder,
  FolderOpen,
  FolderPlus,
  FileText,
  FilePlus2,
  ChevronRight,
  ChevronDown,
  Plus,
  Trash2,
  Edit2,
  Move,
  Search,
  FileCheck,
  Layers,
  GripVertical,
  Save,
  Download,
  ChevronsDownUp,
  ChevronsUpDown,
  CornerRightUp,
} from 'lucide-react';
import {
  ContractFolder,
  SavedDocument,
  NotaryClerk,
  SavedContractDerivedDoc,
  DerivedDocTemplate,
} from '@/lib/types';

interface FolderTreeExplorerProps {
  folders: ContractFolder[];
  documents: SavedDocument[];
  clerks: NotaryClerk[];
  activeClerk: NotaryClerk;
  activeDocument: SavedDocument | null;
  activeDocumentId?: string;
  currentEditorTitle?: string;
  onSelectDocument: (doc: SavedDocument) => void;
  onSaveCurrentToFolder?: (targetFolderId: string | null) => Promise<void> | void;
  onExportDocumentDocx?: (doc: SavedDocument) => Promise<void> | void;
  onCreateFolder: (parentId: string | null, name: string) => void;
  onRenameFolder: (folderId: string, newName: string) => void;
  onDeleteFolder: (folderId: string) => void;
  onMoveDocument: (docId: string, targetFolderId: string | null) => void;
  onMoveFolder: (folderId: string, targetParentId: string | null) => void;
  onCreateContractInFolder: (folderId: string | null) => void;
  onDeleteDocument: (docId: string) => void;
  onUpdateDocumentMeta: (updated: SavedDocument) => void;
  derivedTemplates: DerivedDocTemplate[];
  onOpenDerivedModal: (templateId?: string) => void;
  onSelectDerivedDoc: (derivedDoc: SavedContractDerivedDoc) => void;
  activeDerivedDocId?: string | null;
}

type DragItemPayload = {
  type: 'document' | 'folder' | 'active-editor-doc';
  id: string;
  name: string;
  sourceFolderId: string | null;
};

const STATUS_LABELS: Record<string, string> = {
  draft: 'مسودة',
  editing: 'قيد التحرير',
  ready_to_sign: 'جاهز للتوقيع',
  signed: 'موقع',
  registered: 'مشهر',
};

export default function FolderTreeExplorer({
  folders,
  documents,
  clerks,
  activeClerk,
  activeDocument,
  activeDocumentId,
  currentEditorTitle = 'عقد توثيقي جديد',
  onSelectDocument,
  onSaveCurrentToFolder,
  onExportDocumentDocx,
  onCreateFolder,
  onRenameFolder,
  onDeleteFolder,
  onMoveDocument,
  onMoveFolder,
  onCreateContractInFolder,
  onDeleteDocument,
  onUpdateDocumentMeta,
  onOpenDerivedModal,
  onSelectDerivedDoc,
  activeDerivedDocId,
}: FolderTreeExplorerProps) {
  const [expandedFolders, setExpandedFolders] = useState<Record<string, boolean>>(() => {
    const initial: Record<string, boolean> = {};
    folders.forEach((f) => {
      if (f.parentId === null) initial[f.id] = true;
    });
    return initial;
  });
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [clerkFilter, setClerkFilter] = useState<string>('all');

  // Selected target folder in the "Save Current Editor Contract" bar
  const currentSavedDoc = useMemo(
    () =>
      documents.find((d) => d.id === activeDocumentId) ||
      (activeDocument?.id === activeDocumentId ? activeDocument : null),
    [documents, activeDocument, activeDocumentId]
  );

  const [selectedSaveFolderId, setSelectedSaveFolderId] = useState<string | null>(
    currentSavedDoc?.folderId ?? null
  );

  useEffect(() => {
    setSelectedSaveFolderId(currentSavedDoc?.folderId ?? null);
  }, [currentSavedDoc?.folderId, activeDocumentId]);

  // Auto-expand initial root folders when folders load for the first time
  useEffect(() => {
    if (folders.length > 0 && Object.keys(expandedFolders).length === 0) {
      const next: Record<string, boolean> = {};
      folders.forEach((f) => {
        if (f.parentId === null) next[f.id] = true;
      });
      setExpandedFolders(next);
    }
  }, [folders, expandedFolders]);

  // Modal / Prompt states
  const [showNewFolderModal, setShowNewFolderModal] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  const [targetParentForNewFolder, setTargetParentForNewFolder] = useState<string | null>(null);

  const [renamingFolderId, setRenamingFolderId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState('');

  const [movingItem, setMovingItem] = useState<{
    id: string;
    type: 'folder' | 'document';
    name: string;
  } | null>(null);
  const [targetMoveFolderId, setTargetMoveFolderId] = useState<string | null>(null);

  // Drag & Drop states
  const [draggedItem, setDraggedItem] = useState<DragItemPayload | null>(null);
  const [dragOverTargetId, setDragOverTargetId] = useState<string | 'ROOT' | null>(null);
  const hoverExpandTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Toggle folder expansion
  const toggleFolder = (folderId: string) => {
    setExpandedFolders((prev) => ({ ...prev, [folderId]: !prev[folderId] }));
  };

  const expandFolderAndAncestors = (folderId: string | null) => {
    if (!folderId) return;
    setExpandedFolders((prev) => {
      const next = { ...prev };
      let cursor: string | null = folderId;
      while (cursor) {
        next[cursor] = true;
        const parentObj = folders.find((f) => f.id === cursor);
        cursor = parentObj?.parentId ?? null;
      }
      return next;
    });
  };

  const handleExpandAll = () => {
    const all: Record<string, boolean> = {};
    folders.forEach((f) => {
      all[f.id] = true;
    });
    setExpandedFolders(all);
  };

  const handleCollapseAll = () => {
    setExpandedFolders({});
  };

  // Build hierarchical folder options for <select> menus
  const hierarchicalFolderOptions = useMemo(() => {
    const result: { id: string; name: string; pathLabel: string; depth: number }[] = [];
    const walk = (parentId: string | null, depth: number, prefix: string) => {
      const children = folders.filter((f) => f.parentId === parentId);
      for (const child of children) {
        const pathLabel = prefix ? `${prefix} / ${child.name}` : child.name;
        result.push({ id: child.id, name: child.name, pathLabel, depth });
        walk(child.id, depth + 1, pathLabel);
      }
    };
    walk(null, 0, '');
    return result;
  }, [folders]);

  const getFolderBreadcrumb = (folderId: string | null | undefined): string => {
    if (!folderId) return 'الجذر الرئيسي';
    const found = hierarchicalFolderOptions.find((o) => o.id === folderId);
    return found ? found.pathLabel : 'الجذر الرئيسي';
  };

  // Check if candidateTargetId is the folder itself or a descendant of folderId
  const isInvalidFolderMove = (folderId: string, candidateTargetId: string | null): boolean => {
    if (!candidateTargetId) return false;
    if (folderId === candidateTargetId) return true;
    let cursor: string | null = candidateTargetId;
    while (cursor) {
      if (cursor === folderId) return true;
      const parentObj = folders.find((f) => f.id === cursor);
      cursor = parentObj?.parentId ?? null;
    }
    return false;
  };

  // Count total documents & subfolders inside a folder recursively
  const getFolderCounts = (folderId: string): { docsCount: number; subFoldersCount: number } => {
    const directSubFolders = folders.filter((f) => f.parentId === folderId);
    const directDocs = documents.filter((d) => d.folderId === folderId);
    let totalDocs = directDocs.length;
    let totalSubFolders = directSubFolders.length;
    for (const sf of directSubFolders) {
      const subCounts = getFolderCounts(sf.id);
      totalDocs += subCounts.docsCount;
      totalSubFolders += subCounts.subFoldersCount;
    }
    return { docsCount: totalDocs, subFoldersCount: totalSubFolders };
  };

  // Drag & Drop handlers
  const clearHoverTimer = () => {
    if (hoverExpandTimerRef.current) {
      clearTimeout(hoverExpandTimerRef.current);
      hoverExpandTimerRef.current = null;
    }
  };

  const handleDragStartItem = (e: React.DragEvent, payload: DragItemPayload) => {
    e.stopPropagation();
    setDraggedItem(payload);
    e.dataTransfer.effectAllowed = 'move';
    try {
      e.dataTransfer.setData('application/json', JSON.stringify(payload));
    } catch {
      e.dataTransfer.setData('text/plain', payload.id);
    }
  };

  const handleDragEndItem = () => {
    clearHoverTimer();
    setDraggedItem(null);
    setDragOverTargetId(null);
  };

  const handleDragOverTarget = (
    e: React.DragEvent,
    targetFolderId: string | 'ROOT'
  ) => {
    e.preventDefault();
    e.stopPropagation();
    if (!draggedItem) return;

    const normalizedTarget = targetFolderId === 'ROOT' ? null : targetFolderId;

    // Prevent dropping folder into itself or its descendants
    if (
      draggedItem.type === 'folder' &&
      isInvalidFolderMove(draggedItem.id, normalizedTarget)
    ) {
      e.dataTransfer.dropEffect = 'none';
      return;
    }

    e.dataTransfer.dropEffect = 'move';
    if (dragOverTargetId !== targetFolderId) {
      setDragOverTargetId(targetFolderId);
      clearHoverTimer();

      // Auto-expand collapsed folder after hovering for 500ms
      if (targetFolderId !== 'ROOT' && !expandedFolders[targetFolderId]) {
        hoverExpandTimerRef.current = setTimeout(() => {
          setExpandedFolders((prev) => ({ ...prev, [targetFolderId]: true }));
        }, 500);
      }
    }
  };

  const handleDropOnTarget = async (
    e: React.DragEvent,
    targetFolderId: string | 'ROOT'
  ) => {
    e.preventDefault();
    e.stopPropagation();
    clearHoverTimer();

    let payload = draggedItem;
    if (!payload) {
      try {
        const raw = e.dataTransfer.getData('application/json');
        if (raw) payload = JSON.parse(raw) as DragItemPayload;
      } catch {
        // Fallback
      }
    }

    setDragOverTargetId(null);
    setDraggedItem(null);

    if (!payload) return;
    const normalizedTarget = targetFolderId === 'ROOT' ? null : targetFolderId;

    if (payload.type === 'active-editor-doc') {
      if (onSaveCurrentToFolder) {
        await onSaveCurrentToFolder(normalizedTarget);
        setSelectedSaveFolderId(normalizedTarget);
        expandFolderAndAncestors(normalizedTarget);
      }
      return;
    }

    if (payload.type === 'document') {
      if ((payload.sourceFolderId ?? null) !== normalizedTarget) {
        onMoveDocument(payload.id, normalizedTarget);
        expandFolderAndAncestors(normalizedTarget);
      }
      return;
    }

    if (payload.type === 'folder') {
      if (
        !isInvalidFolderMove(payload.id, normalizedTarget) &&
        (payload.sourceFolderId ?? null) !== normalizedTarget
      ) {
        onMoveFolder(payload.id, normalizedTarget);
        expandFolderAndAncestors(normalizedTarget);
      }
    }
  };

  // Filter documents and folders based on search & clerk role
  const filteredDocs = documents.filter((doc) => {
    if (activeClerk.role !== 'notary' && doc.clerkId && doc.clerkId !== activeClerk.id) {
      return false;
    }
    if (clerkFilter !== 'all' && doc.clerkId !== clerkFilter) return false;
    if (statusFilter !== 'all' && doc.status !== statusFilter) return false;

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const matchTitle = doc.title?.toLowerCase().includes(q);
      const matchClient = doc.clientName?.toLowerCase().includes(q);
      const matchNum = doc.contractNumber?.toLowerCase().includes(q);
      return matchTitle || matchClient || matchNum;
    }
    return true;
  });

  // Check if a folder or any of its descendants match the search/filter
  const folderHasMatchingContent = (folderId: string): boolean => {
    const folder = folders.find((f) => f.id === folderId);
    if (!folder) return false;
    if (activeClerk.role !== 'notary' && folder.clerkId && folder.clerkId !== activeClerk.id) {
      return false;
    }
    if (!searchQuery.trim() && statusFilter === 'all' && clerkFilter === 'all') {
      return true;
    }
    const q = searchQuery.trim().toLowerCase();
    if (q && folder.name.toLowerCase().includes(q)) {
      return true;
    }
    if (filteredDocs.some((d) => d.folderId === folderId)) {
      return true;
    }
    const childFolders = folders.filter((f) => f.parentId === folderId);
    return childFolders.some((cf) => folderHasMatchingContent(cf.id));
  };

  const filteredFolders = folders.filter((f) => folderHasMatchingContent(f.id));

  // Recursive tree renderer
  const renderFolderNode = (parentId: string | null, depth = 0) => {
    const childFolders = filteredFolders.filter((f) => f.parentId === parentId);
    const childDocs = filteredDocs.filter((d) =>
      parentId === null ? !d.folderId : d.folderId === parentId
    );

    if (childFolders.length === 0 && childDocs.length === 0 && searchQuery.trim()) {
      return null;
    }

    return (
      <div className="space-y-1" style={{ paddingRight: depth > 0 ? '10px' : '0' }}>
        {/* Render Folders */}
        {childFolders.map((folder) => {
          const isExpanded = searchQuery.trim() ? true : !!expandedFolders[folder.id];
          const hasChildren =
            filteredFolders.some((f) => f.parentId === folder.id) ||
            filteredDocs.some((d) => d.folderId === folder.id);
          const { docsCount, subFoldersCount } = getFolderCounts(folder.id);
          const isDropTarget = dragOverTargetId === folder.id;
          const isBeingDragged =
            draggedItem?.type === 'folder' && draggedItem.id === folder.id;
          const isInvalidDrop =
            draggedItem?.type === 'folder' &&
            isInvalidFolderMove(draggedItem.id, folder.id);

          return (
            <div
              key={folder.id}
              className={`group/folder rounded-lg transition-all ${
                isBeingDragged ? 'opacity-45' : ''
              }`}
            >
              <div
                draggable={renamingFolderId !== folder.id}
                onDragStart={(e) =>
                  handleDragStartItem(e, {
                    type: 'folder',
                    id: folder.id,
                    name: folder.name,
                    sourceFolderId: folder.parentId,
                  })
                }
                onDragEnd={handleDragEndItem}
                onDragOver={(e) => handleDragOverTarget(e, folder.id)}
                onDrop={(e) => handleDropOnTarget(e, folder.id)}
                className={`flex items-center justify-between px-2 py-1.5 rounded-lg text-xs font-medium cursor-pointer transition-all border ${
                  isDropTarget && !isInvalidDrop
                    ? 'bg-amber-50 border-amber-500 ring-2 ring-amber-400/50 text-amber-950 shadow-xs'
                    : targetParentForNewFolder === folder.id
                    ? 'bg-amber-100/70 border-amber-300 text-amber-900'
                    : 'bg-white/80 border-stone-200/80 hover:border-stone-300 hover:bg-stone-100/80 text-stone-800'
                }`}
                onClick={() => toggleFolder(folder.id)}
              >
                <div className="flex items-center gap-1.5 min-w-0 flex-1">
                  <GripVertical className="w-3 h-3 text-stone-300 group-hover/folder:text-stone-500 shrink-0 cursor-grab active:cursor-grabbing" />
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleFolder(folder.id);
                    }}
                    className="p-0.5 text-stone-400 hover:text-stone-700 rounded shrink-0"
                  >
                    {hasChildren ? (
                      isExpanded ? (
                        <ChevronDown className="w-3.5 h-3.5" />
                      ) : (
                        <ChevronRight className="w-3.5 h-3.5" />
                      )
                    ) : (
                      <span className="w-3.5 h-3.5 inline-block" />
                    )}
                  </button>
                  {isExpanded ? (
                    <FolderOpen className="w-4 h-4 text-amber-600 shrink-0" />
                  ) : (
                    <Folder className="w-4 h-4 text-amber-500 shrink-0" />
                  )}

                  {renamingFolderId === folder.id ? (
                    <input
                      type="text"
                      value={renameValue}
                      onChange={(e) => setRenameValue(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          if (renameValue.trim())
                            onRenameFolder(folder.id, renameValue.trim());
                          setRenamingFolderId(null);
                        } else if (e.key === 'Escape') {
                          setRenamingFolderId(null);
                        }
                      }}
                      onBlur={() => {
                        if (renameValue.trim())
                          onRenameFolder(folder.id, renameValue.trim());
                        setRenamingFolderId(null);
                      }}
                      autoFocus
                      className="px-1.5 py-0.5 border border-amber-500 rounded text-xs bg-white text-stone-900 outline-none flex-1 min-w-0"
                      onClick={(e) => e.stopPropagation()}
                    />
                  ) : (
                    <div className="flex items-center gap-1.5 min-w-0 flex-1">
                      <span className="truncate font-semibold">{folder.name}</span>
                      <span className="text-[10px] text-stone-400 tabular-nums font-mono shrink-0">
                        ({docsCount})
                        {subFoldersCount > 0 ? ` · ${subFoldersCount}م` : ''}
                      </span>
                    </div>
                  )}
                </div>

                {/* Drop indicator or Folder Quick Actions */}
                {isDropTarget && !isInvalidDrop ? (
                  <span className="text-[10px] font-bold text-amber-800 px-1.5 py-0.5 bg-amber-100 rounded shrink-0">
                    أفلت هنا
                  </span>
                ) : (
                  <div
                    className="opacity-0 group-hover/folder:opacity-100 flex items-center gap-0.5 transition-opacity shrink-0"
                    onClick={(e) => e.stopPropagation()}
                  >
                    {onSaveCurrentToFolder && (
                      <button
                        type="button"
                        title={`حفظ العقد المفتوح حالياً (${currentEditorTitle}) داخل مجلد "${folder.name}"`}
                        onClick={async () => {
                          await onSaveCurrentToFolder(folder.id);
                          setSelectedSaveFolderId(folder.id);
                          expandFolderAndAncestors(folder.id);
                        }}
                        className="p-1 hover:bg-amber-100 text-amber-800 rounded transition-colors"
                      >
                        <Save className="w-3 h-3" />
                      </button>
                    )}
                    <button
                      type="button"
                      title="إنشاء عقد فارغ جديد داخل هذا المجلد"
                      onClick={() => {
                        expandFolderAndAncestors(folder.id);
                        onCreateContractInFolder(folder.id);
                      }}
                      className="p-1 hover:bg-stone-200 text-stone-700 rounded transition-colors"
                    >
                      <FilePlus2 className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      title="إضافة مجلد فرعي"
                      onClick={() => {
                        setTargetParentForNewFolder(folder.id);
                        setNewFolderName('');
                        setShowNewFolderModal(true);
                      }}
                      className="p-1 hover:bg-stone-200 text-stone-600 rounded transition-colors"
                    >
                      <FolderPlus className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      title="إعادة تسمية المجلد"
                      onClick={() => {
                        setRenamingFolderId(folder.id);
                        setRenameValue(folder.name);
                      }}
                      className="p-1 hover:bg-stone-200 text-stone-600 rounded transition-colors"
                    >
                      <Edit2 className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      title="نقل المجلد"
                      onClick={() => {
                        setMovingItem({
                          id: folder.id,
                          type: 'folder',
                          name: folder.name,
                        });
                        setTargetMoveFolderId(folder.parentId || null);
                      }}
                      className="p-1 hover:bg-stone-200 text-stone-600 rounded transition-colors"
                    >
                      <Move className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      title="حذف المجلد"
                      onClick={() => onDeleteFolder(folder.id)}
                      className="p-1 hover:bg-red-100 text-red-600 rounded transition-colors"
                    >
                      <Trash2 className="w-3 h-3" />
                    </button>
                  </div>
                )}
              </div>

              {/* Sub-tree */}
              {isExpanded && (
                <div className="pr-2.5 mt-1 border-r border-stone-200/90 space-y-1">
                  {renderFolderNode(folder.id, depth + 1)}
                  {childFolders.length === 0 &&
                    filteredFolders.filter((f) => f.parentId === folder.id).length ===
                      0 &&
                    filteredDocs.filter((d) => d.folderId === folder.id).length ===
                      0 && (
                      <div
                        onDragOver={(e) => handleDragOverTarget(e, folder.id)}
                        onDrop={(e) => handleDropOnTarget(e, folder.id)}
                        className="py-1.5 px-2 text-[10px] text-stone-400 flex items-center justify-between rounded border border-dashed border-stone-200 bg-stone-50/50"
                      >
                        <span>مجلد فارغ — اسحب عقداً إلى هنا أو</span>
                        {onSaveCurrentToFolder && (
                          <button
                            type="button"
                            onClick={async (e) => {
                              e.stopPropagation();
                              await onSaveCurrentToFolder(folder.id);
                              setSelectedSaveFolderId(folder.id);
                            }}
                            className="text-amber-800 font-semibold hover:underline"
                          >
                            احفظ العقد الحالي هنا
                          </button>
                        )}
                      </div>
                    )}
                </div>
              )}
            </div>
          );
        })}

        {/* Render Documents in this folder level */}
        {childDocs.map((doc) => {
          const isActive =
            activeDocument?.id === doc.id || activeDocumentId === doc.id;
          const clerk = clerks.find((c) => c.id === doc.clerkId);
          const isBeingDragged =
            draggedItem?.type === 'document' && draggedItem.id === doc.id;
          const derivedCount = doc.derivedDocuments?.length || 0;
          const statusLabel = STATUS_LABELS[doc.status || 'draft'] || 'مسودة';

          return (
            <div
              key={doc.id}
              draggable
              onDragStart={(e) =>
                handleDragStartItem(e, {
                  type: 'document',
                  id: doc.id,
                  name: doc.title || 'عقد توثيقي',
                  sourceFolderId: doc.folderId || null,
                })
              }
              onDragEnd={handleDragEndItem}
              onClick={() => onSelectDocument(doc)}
              className={`group/doc flex items-center justify-between px-2 py-2 rounded-lg text-xs cursor-pointer transition-all border ${
                isBeingDragged ? 'opacity-45 scale-[0.99]' : ''
              } ${
                isActive
                  ? 'bg-amber-50/90 border-amber-400 text-amber-950 shadow-2xs'
                  : 'bg-white border-stone-200/80 hover:border-stone-300 text-stone-700 hover:bg-stone-50'
              }`}
            >
              <div className="flex items-center gap-1.5 min-w-0 flex-1">
                <GripVertical className="w-3 h-3 text-stone-300 group-hover/doc:text-stone-500 shrink-0 cursor-grab active:cursor-grabbing" />
                <FileText
                  className={`w-4 h-4 shrink-0 ${
                    isActive ? 'text-amber-700' : 'text-stone-400'
                  }`}
                />
                <div className="min-w-0 flex-1">
                  <div className="font-semibold truncate flex items-center gap-1.5">
                    <span className="truncate">{doc.title || 'عقد بدون عنوان'}</span>
                    {isActive && (
                      <span className="text-[9px] font-normal text-amber-800 shrink-0">
                        (مفتوح)
                      </span>
                    )}
                  </div>
                  {/* Unboxed clean metadata with typographic separators */}
                  <div className="flex items-center gap-1.5 text-[10px] text-stone-500 mt-0.5 truncate">
                    {doc.contractNumber && (
                      <span className="font-mono tabular-nums text-stone-600 shrink-0">
                        {doc.contractNumber}
                      </span>
                    )}
                    {doc.contractNumber && <span aria-hidden="true">·</span>}
                    <span className="text-stone-600 shrink-0">{statusLabel}</span>
                    {doc.clientName && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="truncate text-stone-600">{doc.clientName}</span>
                      </>
                    )}
                    {derivedCount > 0 && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="tabular-nums text-amber-800 shrink-0">
                          {derivedCount} مشتق
                        </span>
                      </>
                    )}
                    {clerk && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="text-stone-500 shrink-0">
                          {clerk.name.split(' ')[0]}
                        </span>
                      </>
                    )}
                  </div>
                </div>
              </div>

              {/* Document Actions */}
              <div
                className="opacity-0 group-hover/doc:opacity-100 flex items-center gap-0.5 transition-opacity shrink-0"
                onClick={(e) => e.stopPropagation()}
              >
                {onExportDocumentDocx && (
                  <button
                    type="button"
                    title="تصدير العقد بصيغة Word (.docx)"
                    onClick={() => onExportDocumentDocx(doc)}
                    className="p-1 hover:bg-stone-200 text-stone-600 rounded transition-colors"
                  >
                    <Download className="w-3 h-3" />
                  </button>
                )}
                <button
                  type="button"
                  title="نقل العقد إلى مجلد آخر"
                  onClick={() => {
                    setMovingItem({
                      id: doc.id,
                      type: 'document',
                      name: doc.title || 'عقد',
                    });
                    setTargetMoveFolderId(doc.folderId || null);
                  }}
                  className="p-1 hover:bg-stone-200 text-stone-600 rounded transition-colors"
                >
                  <Move className="w-3 h-3" />
                </button>
                <button
                  type="button"
                  title="حذف العقد"
                  onClick={() => onDeleteDocument(doc.id)}
                  className="p-1 hover:bg-red-100 text-red-600 rounded transition-colors"
                >
                  <Trash2 className="w-3 h-3" />
                </button>
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="flex flex-col h-full bg-stone-50 text-stone-800 text-xs select-none -m-3">
      {/* 1. ACTIVE EDITOR CONTRACT FILING BAR (حفظ العقد المفتوح حالياً في مجلد أو سحبه) */}
      <div className="p-3 bg-white border-b border-stone-200 space-y-2">
        <div
          draggable
          onDragStart={(e) =>
            handleDragStartItem(e, {
              type: 'active-editor-doc',
              id: activeDocumentId || 'current_editor',
              name: currentEditorTitle,
              sourceFolderId: currentSavedDoc?.folderId ?? null,
            })
          }
          onDragEnd={handleDragEndItem}
          className="p-2 rounded-lg border border-amber-300/90 bg-amber-50/60 cursor-grab active:cursor-grabbing transition-all hover:border-amber-400"
          title="اسحب هذه البطاقة وأفلتها فوق أي مجلد في الشجرة أدناه لحفظ العقد المفتوح بداخله فوراً"
        >
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-1.5 min-w-0">
              <GripVertical className="w-3.5 h-3.5 text-amber-700 shrink-0" />
              <FileText className="w-3.5 h-3.5 text-amber-800 shrink-0" />
              <span className="font-bold text-stone-900 truncate">
                {currentEditorTitle || 'عقد توثيقي جديد'}
              </span>
            </div>
            <span className="text-[10px] text-amber-900 shrink-0">
              {currentSavedDoc ? 'محفوظ في الشجرة' : 'مسودة غير مؤرشفة'}
            </span>
          </div>
          <div className="text-[10px] text-stone-600 mt-1 flex items-center justify-between gap-2">
            <span className="truncate">
              الموقع: {getFolderBreadcrumb(currentSavedDoc?.folderId)}
            </span>
            <span className="text-amber-800 font-medium shrink-0">
              اسحب للمجلد أو اختر أدناه
            </span>
          </div>
        </div>

        {/* Direct Folder Selector + Save Button */}
        {onSaveCurrentToFolder && (
          <div className="flex items-center gap-1.5">
            <select
              value={selectedSaveFolderId || ''}
              onChange={(e) =>
                setSelectedSaveFolderId(e.target.value ? e.target.value : null)
              }
              className="flex-1 min-w-0 py-1.5 px-2 bg-stone-50 border border-stone-300 rounded-md text-xs text-stone-800 focus:outline-none focus:border-amber-600 focus:bg-white truncate"
              title="اختر المجلد لحفظ العقد المفتوح حالياً بداخله"
            >
              <option value="">(الجذر الرئيسي للمكتب)</option>
              {hierarchicalFolderOptions.map((opt) => (
                <option key={opt.id} value={opt.id}>
                  {'— '.repeat(opt.depth)}
                  {opt.name}
                </option>
              ))}
            </select>
            <button
              type="button"
              onClick={async () => {
                await onSaveCurrentToFolder(selectedSaveFolderId);
                expandFolderAndAncestors(selectedSaveFolderId);
              }}
              className="px-2.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-md font-semibold inline-flex items-center gap-1 shrink-0 transition-colors shadow-2xs whitespace-nowrap"
              title="حفظ العقد المفتوح حالياً في المحرر داخل المجلد المختار"
            >
              <Save className="w-3.5 h-3.5" />
              <span>حفظ في المجلد</span>
            </button>
          </div>
        )}
      </div>

      {/* 2. TREE TOOLBAR, SEARCH & FILTERS */}
      <div className="p-2.5 border-b border-stone-200 bg-white space-y-2">
        <div className="flex items-center justify-between gap-1">
          <div className="font-bold text-stone-900 flex items-center gap-1.5 min-w-0">
            <Layers className="w-4 h-4 text-amber-700 shrink-0" />
            <span className="truncate">شجرة المجلدات والعقود</span>
            <span className="text-[10px] text-stone-400 font-mono tabular-nums shrink-0">
              ({documents.length})
            </span>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <button
              type="button"
              onClick={handleExpandAll}
              className="p-1 text-stone-500 hover:text-stone-800 hover:bg-stone-100 rounded transition-colors"
              title="توسيع كل المجلدات"
            >
              <ChevronsUpDown className="w-3.5 h-3.5" />
            </button>
            <button
              type="button"
              onClick={handleCollapseAll}
              className="p-1 text-stone-500 hover:text-stone-800 hover:bg-stone-100 rounded transition-colors"
              title="طي كل المجلدات"
            >
              <ChevronsDownUp className="w-3.5 h-3.5" />
            </button>
            <button
              type="button"
              onClick={() => {
                setTargetParentForNewFolder(null);
                setNewFolderName('');
                setShowNewFolderModal(true);
              }}
              className="flex items-center gap-1 px-2 py-1 bg-stone-100 hover:bg-stone-200 text-stone-700 rounded-md font-medium transition-colors whitespace-nowrap"
              title="إنشاء مجلد رئيسي جديد"
            >
              <FolderPlus className="w-3.5 h-3.5 text-amber-700" />
              <span>مجلد</span>
            </button>
            <button
              type="button"
              onClick={() => onCreateContractInFolder(selectedSaveFolderId)}
              className="flex items-center gap-1 px-2 py-1 bg-stone-900 hover:bg-stone-800 text-white rounded-md font-medium transition-colors whitespace-nowrap"
              title="إنشاء عقد فارغ جديد"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>عقد جديد</span>
            </button>
          </div>
        </div>

        {/* Search Bar */}
        <div className="relative">
          <Search className="absolute right-2.5 top-2 w-3.5 h-3.5 text-stone-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="بحث في المجلدات والعقود والزبائن ورقم الفهرس..."
            className="w-full pl-3 pr-8 py-1.5 bg-stone-100 border border-stone-200 rounded-md text-xs text-stone-800 placeholder-stone-400 focus:outline-none focus:border-amber-500 focus:bg-white"
          />
        </div>

        {/* Filters */}
        <div className="flex items-center gap-2">
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="w-1/2 py-1 px-1.5 bg-stone-100 border border-stone-200 rounded text-stone-700 focus:outline-none focus:border-amber-500"
          >
            <option value="all">كل الحالات</option>
            <option value="draft">مسودة</option>
            <option value="editing">قيد التحرير</option>
            <option value="ready_to_sign">جاهز للتوقيع</option>
            <option value="signed">موقع</option>
            <option value="registered">مشهر</option>
          </select>

          {activeClerk.role === 'notary' ? (
            <select
              value={clerkFilter}
              onChange={(e) => setClerkFilter(e.target.value)}
              className="w-1/2 py-1 px-1.5 bg-stone-100 border border-stone-200 rounded text-stone-700 focus:outline-none focus:border-amber-500"
            >
              <option value="all">كل الكتّاب</option>
              {clerks.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.name}
                </option>
              ))}
            </select>
          ) : (
            <div className="w-1/2 py-1 px-2 bg-stone-100 border border-stone-200 rounded text-stone-600 truncate">
              {activeClerk.name}
            </div>
          )}
        </div>
      </div>

      {/* 3. TREE VIEW BODY WITH ROOT DROP ZONE */}
      <div className="flex-1 overflow-y-auto p-2.5 space-y-1.5">
        {/* Root Drop Zone Banner (Visible when dragging an item that belongs to a subfolder) */}
        {draggedItem && draggedItem.sourceFolderId !== null && (
          <div
            onDragOver={(e) => handleDragOverTarget(e, 'ROOT')}
            onDrop={(e) => handleDropOnTarget(e, 'ROOT')}
            className={`p-2 rounded-lg border border-dashed text-center text-[11px] font-medium flex items-center justify-center gap-1.5 transition-all ${
              dragOverTargetId === 'ROOT'
                ? 'border-amber-500 bg-amber-100/80 text-amber-950 ring-2 ring-amber-400/50'
                : 'border-stone-300 bg-white text-stone-600'
            }`}
          >
            <CornerRightUp className="w-3.5 h-3.5 text-amber-700" />
            <span>أفلت هنا لنقل «{draggedItem.name}» إلى الجذر الرئيسي للمكتب</span>
          </div>
        )}

        {renderFolderNode(null, 0)}

        {filteredFolders.length === 0 && filteredDocs.length === 0 && (
          <div className="text-center py-8 px-4 text-stone-400 space-y-2">
            <Folder className="w-8 h-8 mx-auto opacity-40" />
            <p className="text-xs text-stone-500 font-medium">
              لا توجد مجلدات أو عقود مطابقة للبحث.
            </p>
            <p className="text-[11px] text-stone-400">
              يمكنك إنشاء مجلد جديد أو الضغط على «حفظ في المجلد» لحفظ العقد المفتوح حالياً.
            </p>
          </div>
        )}
      </div>

      {/* 4. ACTIVE DOCUMENT DETAILS & DERIVED DOCS DRAWER */}
      {activeDocument && (
        <div className="border-t border-stone-200 bg-white p-3 space-y-2.5 shadow-lg">
          <div className="flex items-center justify-between border-b border-stone-100 pb-1.5">
            <div className="font-bold text-amber-950 truncate flex items-center gap-1.5">
              <FileCheck className="w-4 h-4 text-amber-600 shrink-0" />
              <span className="truncate">بطاقة العقد: {activeDocument.title}</span>
            </div>
          </div>

          {/* Direct Folder Assignment + Status Row */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">
                المجلد المحفوظ فيه
              </label>
              <select
                value={activeDocument.folderId || ''}
                onChange={(e) => {
                  const nextFolderId = e.target.value ? e.target.value : null;
                  onMoveDocument(activeDocument.id, nextFolderId);
                  expandFolderAndAncestors(nextFolderId);
                }}
                className="w-full px-1.5 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500 text-stone-800 truncate"
              >
                <option value="">(الجذر الرئيسي)</option>
                {hierarchicalFolderOptions.map((opt) => (
                  <option key={opt.id} value={opt.id}>
                    {'— '.repeat(opt.depth)}
                    {opt.name}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">
                حالة العقد
              </label>
              <select
                value={activeDocument.status || 'draft'}
                onChange={(e) =>
                  onUpdateDocumentMeta({
                    ...activeDocument,
                    status: e.target.value as SavedDocument['status'],
                    updatedAt: new Date().toISOString(),
                  })
                }
                className="w-full px-1.5 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500 text-stone-800"
              >
                <option value="draft">مسودة</option>
                <option value="editing">قيد التحرير</option>
                <option value="ready_to_sign">جاهز للتوقيع</option>
                <option value="signed">موقع</option>
                <option value="registered">مشهر</option>
              </select>
            </div>
          </div>

          {/* Quick Contract Meta Form */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">
                رقم الفهرس / العقد
              </label>
              <input
                type="text"
                value={activeDocument.contractNumber || ''}
                onChange={(e) =>
                  onUpdateDocumentMeta({
                    ...activeDocument,
                    contractNumber: e.target.value,
                    updatedAt: new Date().toISOString(),
                  })
                }
                placeholder="2026/..."
                className="w-full px-2 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500 font-mono tabular-nums"
              />
            </div>
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">
                اسم الزبون / الطرف
              </label>
              <input
                type="text"
                value={activeDocument.clientName || ''}
                onChange={(e) =>
                  onUpdateDocumentMeta({
                    ...activeDocument,
                    clientName: e.target.value,
                    updatedAt: new Date().toISOString(),
                  })
                }
                placeholder="اسم الزبون"
                className="w-full px-2 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500"
              />
            </div>
          </div>

          {/* Derived Documents List Inside Active Contract */}
          <div className="space-y-1.5 pt-1">
            <div className="flex items-center justify-between">
              <span className="font-semibold text-stone-700 text-[11px]">
                المشتقات المحفوظة تحت العقد (
                <span className="tabular-nums">
                  {activeDocument.derivedDocuments?.length || 0}
                </span>
                ):
              </span>
              <button
                type="button"
                onClick={() => onOpenDerivedModal()}
                className="flex items-center gap-1 px-2 py-0.5 bg-amber-100 hover:bg-amber-200 text-amber-900 rounded font-medium text-[10px] transition-colors whitespace-nowrap"
              >
                <Plus className="w-3 h-3" />
                <span>توليد مشتق</span>
              </button>
            </div>

            <div className="max-h-24 overflow-y-auto space-y-1">
              {!activeDocument.derivedDocuments ||
              activeDocument.derivedDocuments.length === 0 ? (
                <p className="text-stone-400 text-[10px] py-1 text-center">
                  لا توجد وثائق مشتقة محفوظة تحت هذا العقد بعد.
                </p>
              ) : (
                activeDocument.derivedDocuments.map((dd) => {
                  const isDerivedActive = activeDerivedDocId === dd.id;
                  return (
                    <div
                      key={dd.id}
                      onClick={() => onSelectDerivedDoc(dd)}
                      className={`flex items-center justify-between p-1.5 rounded text-xs cursor-pointer border transition-colors ${
                        isDerivedActive
                          ? 'bg-amber-100 border-amber-400 text-amber-950 font-bold'
                          : 'bg-stone-50 border-stone-200 hover:bg-stone-100 text-stone-700'
                      }`}
                    >
                      <div className="flex items-center gap-1.5 truncate">
                        <FileText className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                        <span className="truncate">{dd.title}</span>
                      </div>
                      <span className="text-[9px] text-stone-400 tabular-nums font-mono shrink-0">
                        {new Date(dd.updatedAt).toLocaleDateString('ar-DZ')}
                      </span>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      )}

      {/* New Folder Modal */}
      {showNewFolderModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4">
          <div
            className="bg-white rounded-xl shadow-xl max-w-sm w-full p-5 space-y-4 text-stone-900"
            dir="rtl"
          >
            <h3 className="text-base font-bold text-stone-900 flex items-center gap-2">
              <Folder className="w-5 h-5 text-amber-600" />
              <span>
                {targetParentForNewFolder
                  ? `إنشاء مجلد فرعي داخل «${getFolderBreadcrumb(
                      targetParentForNewFolder
                    )}»`
                  : 'إنشاء مجلد رئيسي جديد'}
              </span>
            </h3>
            <input
              type="text"
              value={newFolderName}
              onChange={(e) => setNewFolderName(e.target.value)}
              placeholder="اسم المجلد (مثل: الإيجارات السكنية، ترقية النور…)"
              className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:border-amber-500"
              autoFocus
              onKeyDown={(e) => {
                if (e.key === 'Enter' && newFolderName.trim()) {
                  onCreateFolder(targetParentForNewFolder, newFolderName.trim());
                  expandFolderAndAncestors(targetParentForNewFolder);
                  setShowNewFolderModal(false);
                }
              }}
            />
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowNewFolderModal(false)}
                className="px-4 py-2 bg-stone-100 hover:bg-stone-200 text-stone-700 rounded-lg text-sm font-medium"
              >
                إلغاء
              </button>
              <button
                type="button"
                onClick={() => {
                  if (newFolderName.trim()) {
                    onCreateFolder(targetParentForNewFolder, newFolderName.trim());
                    expandFolderAndAncestors(targetParentForNewFolder);
                    setShowNewFolderModal(false);
                  }
                }}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-xs"
              >
                إنشاء المجلد
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Move Item Modal */}
      {movingItem && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4">
          <div
            className="bg-white rounded-xl shadow-xl max-w-sm w-full p-5 space-y-4 text-stone-900"
            dir="rtl"
          >
            <h3 className="text-base font-bold text-stone-900 flex items-center gap-2">
              <Move className="w-5 h-5 text-amber-600" />
              <span>نقل: {movingItem.name}</span>
            </h3>
            <p className="text-xs text-stone-600">
              اختر المجلد المستهدف لنقل العنصر إليه:
            </p>

            <select
              value={targetMoveFolderId || ''}
              onChange={(e) =>
                setTargetMoveFolderId(e.target.value ? e.target.value : null)
              }
              className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:border-amber-500 bg-white"
            >
              <option value="">(المجلد الرئيسي - الجذر)</option>
              {hierarchicalFolderOptions
                .filter(
                  (opt) =>
                    movingItem.type !== 'folder' ||
                    !isInvalidFolderMove(movingItem.id, opt.id)
                )
                .map((opt) => (
                  <option key={opt.id} value={opt.id}>
                    {'— '.repeat(opt.depth)}
                    {opt.name}
                  </option>
                ))}
            </select>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setMovingItem(null)}
                className="px-4 py-2 bg-stone-100 hover:bg-stone-200 text-stone-700 rounded-lg text-sm font-medium"
              >
                إلغاء
              </button>
              <button
                type="button"
                onClick={() => {
                  if (movingItem.type === 'document') {
                    onMoveDocument(movingItem.id, targetMoveFolderId);
                    expandFolderAndAncestors(targetMoveFolderId);
                  } else {
                    onMoveFolder(movingItem.id, targetMoveFolderId);
                    expandFolderAndAncestors(targetMoveFolderId);
                  }
                  setMovingItem(null);
                }}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-xs"
              >
                تأكيد النقل
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
