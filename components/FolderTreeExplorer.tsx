'use client';

import React, { useState } from 'react';
import {
  Folder,
  FolderOpen,
  FileText,
  ChevronRight,
  ChevronDown,
  Plus,
  MoreVertical,
  Trash2,
  Edit2,
  Move,
  Search,
  CheckCircle,
  FileCheck,
  Building,
  User,
  Calendar,
  Layers
} from 'lucide-react';
import { ContractFolder, SavedDocument, NotaryClerk, SavedContractDerivedDoc, DerivedDocTemplate } from '@/lib/types';

interface FolderTreeExplorerProps {
  folders: ContractFolder[];
  documents: SavedDocument[];
  clerks: NotaryClerk[];
  activeClerk: NotaryClerk;
  activeDocument: SavedDocument | null;
  onSelectDocument: (doc: SavedDocument) => void;
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

export default function FolderTreeExplorer({
  folders,
  documents,
  clerks,
  activeClerk,
  activeDocument,
  onSelectDocument,
  onCreateFolder,
  onRenameFolder,
  onDeleteFolder,
  onMoveDocument,
  onMoveFolder,
  onCreateContractInFolder,
  onDeleteDocument,
  onUpdateDocumentMeta,
  derivedTemplates,
  onOpenDerivedModal,
  onSelectDerivedDoc,
  activeDerivedDocId,
}: FolderTreeExplorerProps) {
  const [expandedFolders, setExpandedFolders] = useState<Record<string, boolean>>({});
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [clerkFilter, setClerkFilter] = useState<string>('all');
  
  // Modal / Prompt states
  const [showNewFolderModal, setShowNewFolderModal] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  const [targetParentForNewFolder, setTargetParentForNewFolder] = useState<string | null>(null);

  const [renamingFolderId, setRenamingFolderId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState('');

  const [movingItem, setMovingItem] = useState<{ id: string; type: 'folder' | 'document'; name: string } | null>(null);
  const [targetMoveFolderId, setTargetMoveFolderId] = useState<string | null>(null);

  // Toggle folder expansion
  const toggleFolder = (folderId: string) => {
    setExpandedFolders(prev => ({ ...prev, [folderId]: !prev[folderId] }));
  };

  // Filter documents and folders based on search & clerk role
  const filteredDocs = documents.filter(doc => {
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

  const filteredFolders = folders.filter(f => {
    if (activeClerk.role !== 'notary' && f.clerkId && f.clerkId !== activeClerk.id) {
      return false;
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      return f.name.toLowerCase().includes(q);
    }
    return true;
  });

  // Recursive tree renderer
  const renderFolderNode = (parentId: string | null, depth = 0) => {
    const childFolders = filteredFolders.filter(f => f.parentId === parentId);
    const childDocs = filteredDocs.filter(d => (parentId === null ? !d.folderId : d.folderId === parentId));

    if (childFolders.length === 0 && childDocs.length === 0 && searchQuery.trim()) {
      return null;
    }

    return (
      <div className="space-y-1" style={{ paddingRight: depth > 0 ? '12px' : '0' }}>
        {/* Render Folders */}
        {childFolders.map(folder => {
          const isExpanded = !!expandedFolders[folder.id];
          const hasChildren = filteredFolders.some(f => f.parentId === folder.id) || filteredDocs.some(d => d.folderId === folder.id);

          return (
            <div key={folder.id} className="group/folder">
              <div
                className={`flex items-center justify-between px-2 py-1.5 rounded-lg text-xs font-medium cursor-pointer transition-colors ${
                  targetParentForNewFolder === folder.id ? 'bg-amber-100 text-amber-900' : 'hover:bg-stone-100 text-stone-800'
                }`}
                onClick={() => toggleFolder(folder.id)}
              >
                <div className="flex items-center gap-1.5 truncate">
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleFolder(folder.id);
                    }}
                    className="p-0.5 text-stone-400 hover:text-stone-600 rounded"
                  >
                    {hasChildren ? (
                      isExpanded ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />
                    ) : (
                      <span className="w-3.5 h-3.5 inline-block" />
                    )}
                  </button>
                  {isExpanded ? <FolderOpen className="w-4 h-4 text-amber-600 shrink-0" /> : <Folder className="w-4 h-4 text-amber-500 shrink-0" />}
                  {renamingFolderId === folder.id ? (
                    <input
                      type="text"
                      value={renameValue}
                      onChange={(e) => setRenameValue(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          if (renameValue.trim()) onRenameFolder(folder.id, renameValue.trim());
                          setRenamingFolderId(null);
                        } else if (e.key === 'Escape') {
                          setRenamingFolderId(null);
                        }
                      }}
                      onBlur={() => {
                        if (renameValue.trim()) onRenameFolder(folder.id, renameValue.trim());
                        setRenamingFolderId(null);
                      }}
                      autoFocus
                      className="px-1.5 py-0.5 border border-amber-400 rounded text-xs bg-white text-stone-900 outline-none"
                      onClick={(e) => e.stopPropagation()}
                    />
                  ) : (
                    <span className="truncate">{folder.name}</span>
                  )}
                </div>

                {/* Folder Actions */}
                <div className="opacity-0 group-hover/folder:opacity-100 flex items-center gap-1 transition-opacity" onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    title="إضافة مجلد فرعي"
                    onClick={() => {
                      setTargetParentForNewFolder(folder.id);
                      setNewFolderName('');
                      setShowNewFolderModal(true);
                    }}
                    className="p-1 hover:bg-stone-200 text-stone-600 rounded"
                  >
                    <Plus className="w-3 h-3" />
                  </button>
                  <button
                    type="button"
                    title="إضافة عقد داخل المجلد"
                    onClick={() => onCreateContractInFolder(folder.id)}
                    className="p-1 hover:bg-amber-200 text-amber-800 rounded"
                  >
                    <FileText className="w-3 h-3" />
                  </button>
                  <button
                    type="button"
                    title="إعادة تسمية"
                    onClick={() => {
                      setRenamingFolderId(folder.id);
                      setRenameValue(folder.name);
                    }}
                    className="p-1 hover:bg-stone-200 text-stone-600 rounded"
                  >
                    <Edit2 className="w-3 h-3" />
                  </button>
                  <button
                    type="button"
                    title="نقل المجلد"
                    onClick={() => {
                      setMovingItem({ id: folder.id, type: 'folder', name: folder.name });
                      setTargetMoveFolderId(null);
                    }}
                    className="p-1 hover:bg-stone-200 text-stone-600 rounded"
                  >
                    <Move className="w-3 h-3" />
                  </button>
                  <button
                    type="button"
                    title="حذف المجلد"
                    onClick={() => onDeleteFolder(folder.id)}
                    className="p-1 hover:bg-red-100 text-red-600 rounded"
                  >
                    <Trash2 className="w-3 h-3" />
                  </button>
                </div>
              </div>

              {/* Sub-tree */}
              {isExpanded && (
                <div className="pr-3 mt-1 border-r border-stone-200 space-y-1">
                  {renderFolderNode(folder.id, depth + 1)}
                </div>
              )}
            </div>
          );
        })}

        {/* Render Documents in this folder level */}
        {childDocs.map(doc => {
          const isActive = activeDocument?.id === doc.id;
          const clerk = clerks.find(c => c.id === doc.clerkId);

          return (
            <div
              key={doc.id}
              onClick={() => onSelectDocument(doc)}
              className={`group/doc flex items-center justify-between px-2.5 py-2 rounded-lg text-xs cursor-pointer transition-all border ${
                isActive
                  ? 'bg-amber-50 border-amber-300 text-amber-900 shadow-xs'
                  : 'bg-white border-stone-100 hover:border-stone-300 text-stone-700 hover:bg-stone-50'
              }`}
            >
              <div className="flex items-center gap-2 truncate">
                <FileText className={`w-4 h-4 shrink-0 ${isActive ? 'text-amber-600' : 'text-stone-400'}`} />
                <div className="truncate">
                  <div className="font-semibold truncate">{doc.title || 'عقد بدون عنوان'}</div>
                  <div className="flex items-center gap-2 text-[10px] text-stone-500 mt-0.5">
                    {doc.contractNumber && <span className="bg-stone-100 px-1.5 py-0.2 rounded font-mono">{doc.contractNumber}</span>}
                    {doc.clientName && <span className="truncate">{doc.clientName}</span>}
                    {clerk && (
                      <span className="px-1.5 py-0.2 rounded text-white" style={{ backgroundColor: clerk.color || '#666' }}>
                        {clerk.name.split(' ')[0]}
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {/* Document Actions */}
              <div className="opacity-0 group-hover/doc:opacity-100 flex items-center gap-1 transition-opacity shrink-0" onClick={(e) => e.stopPropagation()}>
                <button
                  type="button"
                  title="نقل العقد إلى مجلد آخر"
                  onClick={() => {
                    setMovingItem({ id: doc.id, type: 'document', name: doc.title || 'عقد' });
                    setTargetMoveFolderId(doc.folderId || null);
                  }}
                  className="p-1 hover:bg-stone-200 text-stone-600 rounded"
                >
                  <Move className="w-3 h-3" />
                </button>
                <button
                  type="button"
                  title="حذف العقد"
                  onClick={() => onDeleteDocument(doc.id)}
                  className="p-1 hover:bg-red-100 text-red-600 rounded"
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
    <div className="flex flex-col h-full bg-stone-50 text-stone-800 text-xs select-none">
      {/* Top Toolbar of Tree Explorer */}
      <div className="p-3 border-b border-stone-200 bg-white space-y-2">
        <div className="flex items-center justify-between">
          <div className="font-bold text-stone-900 flex items-center gap-1.5">
            <Layers className="w-4 h-4 text-amber-700" />
            <span>شجرة عقود المكتب</span>
          </div>
          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => {
                setTargetParentForNewFolder(null);
                setNewFolderName('');
                setShowNewFolderModal(true);
              }}
              className="flex items-center gap-1 px-2 py-1 bg-stone-100 hover:bg-stone-200 text-stone-700 rounded-md font-medium transition-colors"
              title="مجلد رئيسي جديد"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>مجلد</span>
            </button>
            <button
              type="button"
              onClick={() => onCreateContractInFolder(null)}
              className="flex items-center gap-1 px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded-md font-medium transition-colors shadow-xs"
              title="عقد جديد في الجذر"
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
            placeholder="بحث في المجلدات والعقود والزبائن..."
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

          {activeClerk.role === 'notary' && (
            <select
              value={clerkFilter}
              onChange={(e) => setClerkFilter(e.target.value)}
              className="w-1/2 py-1 px-1.5 bg-stone-100 border border-stone-200 rounded text-stone-700 focus:outline-none focus:border-amber-500"
            >
              <option value="all">كل الكتّاب</option>
              {clerks.map(c => (
                <option key={c.id} value={c.id}>{c.name}</option>
              ))}
            </select>
          )}
        </div>
      </div>

      {/* Tree View Body */}
      <div className="flex-1 overflow-y-auto p-2 space-y-1">
        {renderFolderNode(null, 0)}

        {filteredFolders.length === 0 && filteredDocs.length === 0 && (
          <div className="text-center py-10 text-stone-400">
            <Folder className="w-8 h-8 mx-auto mb-2 opacity-40" />
            <p>لا توجد مجلدات أو عقود مطابقة للبحث.</p>
          </div>
        )}
      </div>

      {/* Active Document Details & Derived Docs Drawer (when active document is selected) */}
      {activeDocument && (
        <div className="border-t border-stone-200 bg-white p-3 space-y-2.5 shadow-lg">
          <div className="flex items-center justify-between border-b border-stone-100 pb-1.5">
            <div className="font-bold text-amber-900 truncate flex items-center gap-1.5">
              <FileCheck className="w-4 h-4 text-amber-600" />
              <span className="truncate">وثائق هذا العقد: {activeDocument.title}</span>
            </div>
          </div>

          {/* Quick Contract Meta Form */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">رقم العقد</label>
              <input
                type="text"
                value={activeDocument.contractNumber || ''}
                onChange={(e) => onUpdateDocumentMeta({ ...activeDocument, contractNumber: e.target.value, updatedAt: new Date().toISOString() })}
                placeholder="2026/..."
                className="w-full px-2 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500 font-mono"
              />
            </div>
            <div>
              <label className="block text-[10px] text-stone-500 mb-0.5">اسم الزبون</label>
              <input
                type="text"
                value={activeDocument.clientName || ''}
                onChange={(e) => onUpdateDocumentMeta({ ...activeDocument, clientName: e.target.value, updatedAt: new Date().toISOString() })}
                placeholder="اسم الزبون"
                className="w-full px-2 py-1 border border-stone-200 rounded bg-stone-50 focus:bg-white focus:border-amber-500"
              />
            </div>
          </div>

          <div className="flex items-center gap-2">
            <div className="flex-1">
              <label className="block text-[10px] text-stone-500 mb-0.5">الحالة</label>
              <select
                value={activeDocument.status || 'draft'}
                onChange={(e) => onUpdateDocumentMeta({ ...activeDocument, status: e.target.value as any, updatedAt: new Date().toISOString() })}
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

          {/* Derived Documents List Inside Active Contract */}
          <div className="space-y-1.5 pt-1">
            <div className="flex items-center justify-between">
              <span className="font-semibold text-stone-700 text-[11px]">المشتقات المرتبطة (مستخرج، شهر…):</span>
              <button
                type="button"
                onClick={() => onOpenDerivedModal()}
                className="flex items-center gap-1 px-2 py-0.5 bg-amber-100 hover:bg-amber-200 text-amber-900 rounded font-medium text-[10px] transition-colors"
              >
                <Plus className="w-3 h-3" />
                <span>توليد مشتق جديد</span>
              </button>
            </div>

            <div className="max-h-28 overflow-y-auto space-y-1">
              {(!activeDocument.derivedDocuments || activeDocument.derivedDocuments.length === 0) ? (
                <p className="text-stone-400 text-[10px] italic py-1 text-center">لا توجد وثائق مشتقة لهذا العقد بعد.</p>
              ) : (
                activeDocument.derivedDocuments.map(dd => {
                  const isDerivedActive = activeDerivedDocId === dd.id;
                  return (
                    <div
                      key={dd.id}
                      onClick={() => onSelectDerivedDoc(dd)}
                      className={`flex items-center justify-between p-1.5 rounded text-xs cursor-pointer border transition-colors ${
                        isDerivedActive
                          ? 'bg-amber-100 border-amber-400 text-amber-900 font-bold'
                          : 'bg-stone-50 border-stone-200 hover:bg-stone-100 text-stone-700'
                      }`}
                    >
                      <div className="flex items-center gap-1.5 truncate">
                        <FileText className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                        <span className="truncate">{dd.title}</span>
                      </div>
                      <span className="text-[9px] text-stone-400">{new Date(dd.updatedAt).toLocaleDateString('ar-DZ')}</span>
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
          <div className="bg-white rounded-xl shadow-xl max-w-sm w-full p-5 space-y-4 text-stone-900" dir="rtl">
            <h3 className="text-base font-bold text-stone-900 flex items-center gap-2">
              <Folder className="w-5 h-5 text-amber-600" />
              <span>{targetParentForNewFolder ? 'إنشاء مجلد فرعي جديد' : 'إنشاء مجلد رئيسي جديد'}</span>
            </h3>
            <input
              type="text"
              value={newFolderName}
              onChange={(e) => setNewFolderName(e.target.value)}
              placeholder="اسم المجلد (مثل: الإيجارات السكنية، مشاريع الترقية…)"
              className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:border-amber-500"
              autoFocus
              onKeyDown={(e) => {
                if (e.key === 'Enter' && newFolderName.trim()) {
                  onCreateFolder(targetParentForNewFolder, newFolderName.trim());
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
                    setShowNewFolderModal(false);
                  }
                }}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-sm"
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
          <div className="bg-white rounded-xl shadow-xl max-w-sm w-full p-5 space-y-4 text-stone-900" dir="rtl">
            <h3 className="text-base font-bold text-stone-900 flex items-center gap-2">
              <Move className="w-5 h-5 text-amber-600" />
              <span>نقل: {movingItem.name}</span>
            </h3>
            <p className="text-xs text-stone-600">اختر المجلد المستهدف لنقل العنصر إليه:</p>

            <select
              value={targetMoveFolderId || ''}
              onChange={(e) => setTargetMoveFolderId(e.target.value ? e.target.value : null)}
              className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:border-amber-500 bg-white"
            >
              <option value="">(المجلد الرئيسي - الجذر)</option>
              {folders.filter(f => f.id !== movingItem.id).map(f => (
                <option key={f.id} value={f.id}>{f.name}</option>
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
                  } else {
                    onMoveFolder(movingItem.id, targetMoveFolderId);
                  }
                  setMovingItem(null);
                }}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-sm"
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
