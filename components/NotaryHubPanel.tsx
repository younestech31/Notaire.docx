'use client';

import React, { useState, useEffect, useMemo } from 'react';
import {
  Award,
  Building,
  CheckCircle2,
  Clock,
  Download,
  FileCheck,
  FileSpreadsheet,
  FolderArchive,
  Hash,
  Lock,
  Plus,
  Printer,
  RefreshCw,
  Search,
  ShieldAlert,
  ShieldCheck,
  Trash2,
} from 'lucide-react';
import {
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
} from '@/lib/types';
import {
  deletePublicationTrack,
  deleteRegistrationBatch,
  loadPublicationTracks,
  loadRegistrationBatches,
  readSealedOriginal,
  savePublicationTrack,
  saveRegistrationBatch,
} from '@/lib/storage';
import {
  isContractOfficiallyIndexed,
  officiallyIndexAndSealContract,
  parseIndexNumberAndYear,
  suggestNextAnnualIndexNumber,
  suggestNextDepositNumber,
  syncPublicationTrackToContract,
  syncRegistrationBatchToContracts,
} from '@/lib/indexing-engine';

interface NotaryHubPanelProps {
  savedDocuments: SavedDocument[];
  activeDocumentId: string | null;
  documentTitle: string;
  fieldValues: Record<string, string>;
  onUpdateFieldValue: (key: string, val: string) => void;
  onSaveCurrentDocument: (overrideTitle?: string) => Promise<void>;
  onContractIndexedOrSynced?: (updatedDoc?: SavedDocument) => Promise<void> | void;
  showToast: (msg: string) => void;
}

export default function NotaryHubPanel({
  savedDocuments,
  activeDocumentId,
  documentTitle,
  fieldValues,
  onUpdateFieldValue,
  onSaveCurrentDocument,
  onContractIndexedOrSynced,
  showToast,
}: NotaryHubPanelProps) {
  const [hubSubTab, setHubSubTab] = useState<
    'indexing' | 'registration' | 'publication' | 'archive'
  >('indexing');

  const currentYear = new Date().getFullYear();

  const activeSavedDoc = useMemo(
    () => savedDocuments.find((d) => d.id === activeDocumentId) || null,
    [savedDocuments, activeDocumentId]
  );

  const suggestedNextIndex = useMemo(
    () => suggestNextAnnualIndexNumber(savedDocuments, activeSavedDoc?.year || currentYear),
    [savedDocuments, activeSavedDoc?.year, currentYear]
  );

  const suggestedNextDeposit = useMemo(
    () => suggestNextDepositNumber(savedDocuments, activeSavedDoc?.year || currentYear),
    [savedDocuments, activeSavedDoc?.year, currentYear]
  );

  // Indexing state
  const [depositNo, setDepositNo] = useState('');
  const [indexNo, setIndexNo] = useState('');
  const [contractDate, setContractDate] = useState(
    () => fieldValues['تاريخ_العقد'] || new Date().toISOString().slice(0, 10)
  );
  const [isSealing, setIsSealing] = useState(false);

  // Sync suggested index/deposit numbers when active document or savedDocuments change
  useEffect(() => {
    const existingIdx =
      activeSavedDoc?.indexNumber ||
      fieldValues['رقم_الفهرس'] ||
      suggestedNextIndex;
    const existingDep =
      activeSavedDoc?.depositNumber ||
      fieldValues['رقم_الإيداع'] ||
      suggestedNextDeposit;
    const existingDate =
      fieldValues['تاريخ_العقد'] ||
      (activeSavedDoc?.indexedAt ? activeSavedDoc.indexedAt.slice(0, 10) : '') ||
      new Date().toISOString().slice(0, 10);

    setIndexNo(existingIdx);
    setDepositNo(existingDep);
    setContractDate(existingDate);
  }, [
    activeDocumentId,
    activeSavedDoc?.indexNumber,
    activeSavedDoc?.depositNumber,
    activeSavedDoc?.indexedAt,
    fieldValues,
    suggestedNextIndex,
    suggestedNextDeposit,
  ]);

  // Registration Batches state
  const [batches, setBatches] = useState<RegistrationBatch[]>([]);
  const [showNewBatchModal, setShowNewBatchModal] = useState(false);
  const [editingBatchId, setEditingBatchId] = useState<string | null>(null);
  const [newBatchInspectorate, setNewBatchInspectorate] = useState(
    'مفتشية التسجيل والطابع - وهران'
  );
  const [selectedBatchDocIds, setSelectedBatchDocIds] = useState<string[]>([]);
  const [newBatchIndices, setNewBatchIndices] = useState('');
  const [newBatchFees, setNewBatchFees] = useState('');
  const [newBatchReceipt, setNewBatchReceipt] = useState('');
  const [newBatchRegDate, setNewBatchRegDate] = useState(
    new Date().toISOString().slice(0, 10)
  );

  // Publication Tracks state
  const [publications, setPublications] = useState<PublicationTrack[]>([]);
  const [showNewPubModal, setShowNewPubModal] = useState(false);
  const [editingPubId, setEditingPubId] = useState<string | null>(null);
  const [selectedPubDocId, setSelectedPubDocId] = useState<string>('');
  const [newPubOffice, setNewPubOffice] = useState('المحافظة العقارية - وهران');
  const [newPubWilaya, setNewPubWilaya] = useState('وهران');
  const [newPubVolume, setNewPubVolume] = useState('');
  const [newPubNumber, setNewPubNumber] = useState('');
  const [newPubDepositDate, setNewPubDepositDate] = useState(
    new Date().toISOString().slice(0, 10)
  );
  const [newPubPublicationDate, setNewPubPublicationDate] = useState(
    new Date().toISOString().slice(0, 10)
  );

  // Archive search state
  const [archiveYearFilter, setArchiveYearFilter] = useState<string>('all');
  const [archiveQuery, setArchiveQuery] = useState('');

  useEffect(() => {
    loadRegistrationBatches().then(setBatches).catch(() => {});
    loadPublicationTracks().then(setPublications).catch(() => {});
  }, []);

  // Indexed contracts ready for registration/publication
  const indexedContracts = useMemo(() => {
    return savedDocuments.filter(
      (d) =>
        Boolean(d.indexNumber || d.fieldValues?.['رقم_الفهرس']) ||
        isContractOfficiallyIndexed(d)
    );
  }, [savedDocuments]);

  const handleConfirmIndexing = async () => {
    if (!indexNo.trim()) {
      showToast('يرجى إدخال رقم الفهرس الرسمي أولاً');
      return;
    }

    setIsSealing(true);
    try {
      // Ensure active changes in the editor are saved first if needed
      let targetDoc = activeSavedDoc;
      if (!targetDoc) {
        await onSaveCurrentDocument();
      }
      // Reload latest saved documents via activeSavedDoc or fallback temporary record
      const baseDoc: SavedDocument = targetDoc || {
        id: activeDocumentId || `doc_${Date.now()}`,
        title: documentTitle || 'عقد توثيقي',
        bodyHtml: '',
        headerHtml: '',
        footerHtml: '',
        showHeaderFooter: false,
        pageNumberingEnabled: true,
        fieldValues: { ...fieldValues },
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };

      const result = await officiallyIndexAndSealContract({
        document: {
          ...baseDoc,
          title: documentTitle || baseDoc.title,
          fieldValues: {
            ...(baseDoc.fieldValues || {}),
            ...fieldValues,
          },
        },
        allDocuments: savedDocuments,
        indexNumberInput: indexNo.trim(),
        depositNumberInput: depositNo.trim(),
        contractDateInput: contractDate,
      });

      onUpdateFieldValue('رقم_الفهرس', result.updatedDocument.indexNumber || indexNo.trim());
      onUpdateFieldValue('رقم_الإيداع', result.updatedDocument.depositNumber || depositNo.trim());
      onUpdateFieldValue('تاريخ_العقد', contractDate);

      if (onContractIndexedOrSynced) {
        await onContractIndexedOrSynced(result.updatedDocument);
      }

      showToast(
        `تم اعتماد وفهرسة العقد رسمياً برقم (${result.updatedDocument.indexNumber}) وختم الأصل ببصمة SHA-256`
      );
    } catch (err) {
      showToast(
        err instanceof Error ? err.message : 'تعذر إتمام الفهرسة الرسمية وختم الأصل'
      );
    } finally {
      setIsSealing(false);
    }
  };

  const handleDownloadSealedOriginal = async (doc: SavedDocument) => {
    try {
      const sealed = await readSealedOriginal(doc.id);
      if (!sealed) {
        showToast('لم يتم العثور على ملف الأصل المختوم لهذا العقد');
        return;
      }
      const blob = new Blob([new Uint8Array(sealed.bytes)], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safeIdx = (doc.indexNumber || 'أصل').replace(/[\/\\]/g, '-');
      a.download = `الأصل_المختوم_${doc.title}_${safeIdx}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      showToast('تم تحميل النسخة الأصلية المختومة (.docx) مع التحقق من بصمة SHA-256');
    } catch (err) {
      showToast(err instanceof Error ? err.message : 'تعذر قراءة الأصل المختوم');
    }
  };

  const openNewBatchDialog = () => {
    setEditingBatchId(null);
    const defaultIds =
      activeSavedDoc &&
      (activeSavedDoc.indexNumber || activeSavedDoc.fieldValues?.['رقم_الفهرس'])
        ? [activeSavedDoc.id]
        : [];
    setSelectedBatchDocIds(defaultIds);
    const defaultIdx =
      activeSavedDoc?.indexNumber || fieldValues['رقم_الفهرس'] || '';
    setNewBatchIndices(defaultIdx);
    setNewBatchFees(fieldValues['حقوق_التسجيل'] || '');
    setNewBatchReceipt('');
    setNewBatchRegDate(new Date().toISOString().slice(0, 10));
    setShowNewBatchModal(true);
  };

  const openEditBatchDialog = (batch: RegistrationBatch) => {
    setEditingBatchId(batch.id);
    setNewBatchInspectorate(batch.inspectorateName);
    setSelectedBatchDocIds(batch.contractIds || []);
    setNewBatchIndices((batch.indexNumbers || []).join('، '));
    setNewBatchFees(batch.totalFees ? String(batch.totalFees) : '');
    setNewBatchReceipt(batch.receiptNumber || '');
    setNewBatchRegDate(
      batch.registrationDate || batch.statusDate || new Date().toISOString().slice(0, 10)
    );
    setShowNewBatchModal(true);
  };

  const toggleSelectBatchDoc = (doc: SavedDocument) => {
    const exists = selectedBatchDocIds.includes(doc.id);
    const nextIds = exists
      ? selectedBatchDocIds.filter((id) => id !== doc.id)
      : [...selectedBatchDocIds, doc.id];
    setSelectedBatchDocIds(nextIds);

    // Sync index numbers text automatically from selected documents
    const selectedIndices = savedDocuments
      .filter((d) => nextIds.includes(d.id))
      .map((d) => d.indexNumber || d.fieldValues?.['رقم_الفهرس'] || '')
      .filter(Boolean);
    setNewBatchIndices(selectedIndices.join('، '));
  };

  const handleSaveBatch = async () => {
    const manualIndices = newBatchIndices
      .split(/[,،]+/)
      .map((s) => s.trim())
      .filter(Boolean);

    // Resolve contracts matching either selectedBatchDocIds or manualIndices
    const normalizedManual = new Set(
      manualIndices.map((idx) => parseIndexNumberAndYear(idx, currentYear).formatted)
    );

    const matchedDocs = savedDocuments.filter((d) => {
      if (selectedBatchDocIds.includes(d.id)) return true;
      const dIdx = d.indexNumber || d.fieldValues?.['رقم_الفهرس'];
      if (!dIdx) return false;
      return normalizedManual.has(parseIndexNumberAndYear(dIdx, d.year || currentYear).formatted);
    });

    const allIndexNumbers = Array.from(
      new Set([
        ...matchedDocs
          .map((d) => d.indexNumber || d.fieldValues?.['رقم_الفهرس'] || '')
          .filter(Boolean),
        ...manualIndices,
      ])
    );

    if (allIndexNumbers.length === 0) {
      showToast('يرجى اختيار عقد مفهرس واحد على الأقل أو إدخال رقم الفهرس');
      return;
    }

    const existingBatch = editingBatchId
      ? batches.find((b) => b.id === editingBatchId)
      : null;

    const hasReceipt = newBatchReceipt.trim().length > 0;
    const batch: RegistrationBatch = {
      id: existingBatch?.id || `batch_${Date.now()}`,
      batchNumber: existingBatch?.batchNumber || batches.length + 1,
      year: existingBatch?.year || currentYear,
      inspectorateName: newBatchInspectorate.trim() || 'مفتشية التسجيل والطابع',
      contractIds: matchedDocs.map((d) => d.id),
      indexNumbers: allIndexNumbers,
      totalFees: Number(newBatchFees.replace(/[^\d.]/g, '')) || 0,
      status: hasReceipt ? 'registered' : 'deposited',
      statusDate: newBatchRegDate || new Date().toISOString().slice(0, 10),
      receiptNumber: newBatchReceipt.trim() || undefined,
      registrationDate: hasReceipt ? newBatchRegDate : undefined,
      createdAt: existingBatch?.createdAt || new Date().toISOString(),
    };

    await saveRegistrationBatch(batch);
    const updatedBatches = await loadRegistrationBatches();
    setBatches(updatedBatches);

    // Synchronize batch status and [مراجع_التسجيل] back to all linked contracts
    const syncedDocs = await syncRegistrationBatchToContracts(batch);
    if (activeDocumentId) {
      const activeSynced = syncedDocs.find((d) => d.id === activeDocumentId);
      if (activeSynced) {
        if (activeSynced.fieldValues?.['مراجع_التسجيل']) {
          onUpdateFieldValue('مراجع_التسجيل', activeSynced.fieldValues['مراجع_التسجيل']);
        }
        if (activeSynced.fieldValues?.['وصل_التسجيل']) {
          onUpdateFieldValue('وصل_التسجيل', activeSynced.fieldValues['وصل_التسجيل']);
        }
      }
    }

    if (onContractIndexedOrSynced) {
      await onContractIndexedOrSynced();
    }

    setShowNewBatchModal(false);
    setEditingBatchId(null);
    showToast(
      hasReceipt
        ? `تم تسجيل الدفعة رقم ${batch.batchNumber}/${batch.year} وتحديث وسم [مراجع_التسجيل] في (${syncedDocs.length}) عقد`
        : `تم حفظ دفعة الإيداع للتسجيل رقم ${batch.batchNumber}/${batch.year} وربطها بـ (${syncedDocs.length}) عقد`
    );
  };

  const handleDeleteBatch = async (id: string) => {
    await deleteRegistrationBatch(id);
    setBatches((prev) => prev.filter((b) => b.id !== id));
    showToast('تم حذف دفعة التسجيل');
  };

  const handlePrintBordereau = (batch: RegistrationBatch) => {
    const printWin = window.open('', '_blank', 'width=850,height=900');
    if (!printWin) {
      window.print();
      return;
    }
    const rowsHtml = (batch.indexNumbers || [])
      .map((idx, i) => {
        const matchedDoc = savedDocuments.find(
          (d) =>
            d.indexNumber === idx ||
            d.fieldValues?.['رقم_الفهرس'] === idx ||
            batch.contractIds?.includes(d.id)
        );
        return `<tr>
          <td style="border:1px solid #000;padding:6px;text-align:center;">${i + 1}</td>
          <td style="border:1px solid #000;padding:6px;text-align:center;font-weight:bold;">${idx}</td>
          <td style="border:1px solid #000;padding:6px;">${matchedDoc?.title || 'عقد توثيقي رسمي'}</td>
          <td style="border:1px solid #000;padding:6px;text-align:center;">${
            matchedDoc?.fieldValues?.['تاريخ_العقد'] || batch.statusDate
          }</td>
          <td style="border:1px solid #000;padding:6px;">${
            batch.receiptNumber ? `وصل رقم ${batch.receiptNumber}` : 'للإيداع والتسجيل'
          }</td>
        </tr>`;
      })
      .join('');

    printWin.document.write(`<!DOCTYPE html>
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="utf-8"/>
        <title>جدول إرسال للتسجيل رقم ${batch.batchNumber}/${batch.year}</title>
        <style>
          body { font-family: Arial, sans-serif; font-size: 13pt; padding: 20mm; direction: rtl; }
          h2, h3 { margin: 4px 0; text-align: center; }
          table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12pt; }
        </style>
      </head>
      <body>
        <h2>الجمهورية الجزائرية الديمقراطية الشعبية</h2>
        <h3>جدول إرسال العقود للتسجيل الجبائي رقم ${batch.batchNumber}/${batch.year}</h3>
        <p><strong>الجهة المرسل إليها:</strong> ${batch.inspectorateName}</p>
        <p><strong>تاريخ الإرسال / التسجيل:</strong> ${batch.statusDate} — <strong>إجمالي الحقوق:</strong> ${batch.totalFees.toLocaleString('ar-DZ')} دج</p>
        <table>
          <thead>
            <tr style="background:#f1f5f9;">
              <th style="border:1px solid #000;padding:6px;">الرقم</th>
              <th style="border:1px solid #000;padding:6px;">رقم الفهرس</th>
              <th style="border:1px solid #000;padding:6px;">موضوع العقد</th>
              <th style="border:1px solid #000;padding:6px;">تاريخ العقد</th>
              <th style="border:1px solid #000;padding:6px;">ملاحظات / رقم الوصل</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </body>
      </html>`);
    printWin.document.close();
    printWin.focus();
    printWin.print();
  };

  const openNewPublicationDialog = () => {
    setEditingPubId(null);
    const defaultDoc =
      activeSavedDoc || indexedContracts[0] || savedDocuments[0] || null;
    setSelectedPubDocId(defaultDoc?.id || '');
    setNewPubOffice(
      defaultDoc?.conservationOffice || 'المحافظة العقارية - وهران'
    );
    setNewPubWilaya(defaultDoc?.wilaya || 'وهران');
    setNewPubVolume(defaultDoc?.fieldValues?.['مجلد_الشهر'] || '');
    setNewPubNumber(defaultDoc?.fieldValues?.['رقم_الشهر'] || '');
    setNewPubDepositDate(new Date().toISOString().slice(0, 10));
    setNewPubPublicationDate(new Date().toISOString().slice(0, 10));
    setShowNewPubModal(true);
  };

  const openEditPublicationDialog = (track: PublicationTrack) => {
    setEditingPubId(track.id);
    setSelectedPubDocId(track.documentId);
    setNewPubOffice(track.conservationOffice);
    setNewPubWilaya(track.wilaya || 'وهران');
    setNewPubVolume(track.volume || '');
    setNewPubNumber(track.number || '');
    setNewPubDepositDate(track.depositDate || new Date().toISOString().slice(0, 10));
    setNewPubPublicationDate(
      track.publicationDate || new Date().toISOString().slice(0, 10)
    );
    setShowNewPubModal(true);
  };

  const handleSavePublication = async () => {
    const targetDoc = savedDocuments.find((d) => d.id === selectedPubDocId);
    if (!targetDoc) {
      showToast(
        'يرجى اختيار عقد محفوظ ومفهرس من القائمة لربط مسار الشهر العقاري به (يمنع إنشاء سجل يتيم بدون عقد)'
      );
      return;
    }

    const existingPub = editingPubId
      ? publications.find((p) => p.id === editingPubId)
      : null;

    const isPublished = Boolean(newPubVolume.trim() && newPubNumber.trim());
    const docIdx =
      targetDoc.indexNumber ||
      targetDoc.fieldValues?.['رقم_الفهرس'] ||
      indexNo ||
      suggestedNextIndex;

    const track: PublicationTrack = {
      id: existingPub?.id || `pub_${Date.now()}`,
      documentId: targetDoc.id,
      indexNumber: docIdx,
      conservationOffice: newPubOffice.trim() || 'المحافظة العقارية',
      wilaya: newPubWilaya.trim() || 'وهران',
      depositDate: newPubDepositDate || new Date().toISOString().slice(0, 10),
      status: isPublished ? 'published' : 'deposited',
      volume: newPubVolume.trim() || undefined,
      number: newPubNumber.trim() || undefined,
      publicationDate: isPublished ? newPubPublicationDate : undefined,
      updatedAt: new Date().toISOString(),
    };

    await savePublicationTrack(track);
    const updatedTracks = await loadPublicationTracks();
    setPublications(updatedTracks);

    // Synchronize publication status and [مراجع_الشهر] back to the linked contract
    const syncedDoc = await syncPublicationTrackToContract(track);
    if (syncedDoc && syncedDoc.id === activeDocumentId) {
      if (syncedDoc.fieldValues?.['مراجع_الشهر']) {
        onUpdateFieldValue('مراجع_الشهر', syncedDoc.fieldValues['مراجع_الشهر']);
      }
      if (syncedDoc.fieldValues?.['مجلد_الشهر']) {
        onUpdateFieldValue('مجلد_الشهر', syncedDoc.fieldValues['مجلد_الشهر']);
      }
      if (syncedDoc.fieldValues?.['رقم_الشهر']) {
        onUpdateFieldValue('رقم_الشهر', syncedDoc.fieldValues['رقم_الشهر']);
      }
    }

    if (onContractIndexedOrSynced) {
      await onContractIndexedOrSynced(syncedDoc || undefined);
    }

    setShowNewPubModal(false);
    setEditingPubId(null);
    showToast(
      isPublished
        ? `تم تأكيد الشهر العقاري للعقد (${targetDoc.title}) وتحديث وسم [مراجع_الشهر] آلياً`
        : `تم تسجيل إيداع الشهر العقاري للعقد (${targetDoc.title})`
    );
  };

  const handleDeletePub = async (id: string) => {
    await deletePublicationTrack(id);
    setPublications((prev) => prev.filter((p) => p.id !== id));
    showToast('تم حذف سجل متابعة الشهر العقاري');
  };

  // Filtered archive documents
  const filteredArchiveDocs = useMemo(() => {
    return savedDocuments.filter((doc) => {
      if (archiveYearFilter !== 'all') {
        const docYear = String(doc.year || doc.createdAt.slice(0, 4));
        if (docYear !== archiveYearFilter) return false;
      }
      if (archiveQuery.trim()) {
        const q = archiveQuery.toLowerCase();
        const inTitle = doc.title.toLowerCase().includes(q);
        const inCat = (doc.modelTitle || doc.clientName || '').toLowerCase().includes(q);
        const inIndex = (doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'] || '')
          .toLowerCase()
          .includes(q);
        const inParties = Object.values(doc.fieldValues || {}).some((v) =>
          String(v).toLowerCase().includes(q)
        );
        return inTitle || inCat || inIndex || inParties;
      }
      return true;
    });
  }, [savedDocuments, archiveYearFilter, archiveQuery]);

  const archiveYears = useMemo(() => {
    const years = new Set<string>();
    years.add(String(currentYear));
    years.add(String(currentYear - 1));
    savedDocuments.forEach((d) => {
      if (d.year) years.add(String(d.year));
      else if (d.createdAt) years.add(d.createdAt.slice(0, 4));
    });
    return Array.from(years).sort().reverse();
  }, [savedDocuments, currentYear]);

  const isCurrentIndexed = isContractOfficiallyIndexed(activeSavedDoc);

  return (
    <div className="space-y-3 text-right" dir="rtl">
      {/* Sub-navigation bar */}
      <div className="grid grid-cols-4 gap-1 bg-[#0F2744] p-1 rounded-lg text-white">
        <button
          type="button"
          onClick={() => setHubSubTab('indexing')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[10px] font-bold transition ${
            hubSubTab === 'indexing'
              ? 'bg-[#C4A46A] text-[#0F2744]'
              : 'text-slate-200 hover:bg-white/10'
          }`}
        >
          <Hash className="w-3.5 h-3.5 mb-0.5" />
          الفهرسة
        </button>
        <button
          type="button"
          onClick={() => setHubSubTab('registration')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[10px] font-bold transition ${
            hubSubTab === 'registration'
              ? 'bg-[#C4A46A] text-[#0F2744]'
              : 'text-slate-200 hover:bg-white/10'
          }`}
        >
          <FileSpreadsheet className="w-3.5 h-3.5 mb-0.5" />
          التسجيل
        </button>
        <button
          type="button"
          onClick={() => setHubSubTab('publication')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[10px] font-bold transition ${
            hubSubTab === 'publication'
              ? 'bg-[#C4A46A] text-[#0F2744]'
              : 'text-slate-200 hover:bg-white/10'
          }`}
        >
          <Building className="w-3.5 h-3.5 mb-0.5" />
          الشهر العقاري
        </button>
        <button
          type="button"
          onClick={() => setHubSubTab('archive')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[10px] font-bold transition ${
            hubSubTab === 'archive'
              ? 'bg-[#C4A46A] text-[#0F2744]'
              : 'text-slate-200 hover:bg-white/10'
          }`}
        >
          <FolderArchive className="w-3.5 h-3.5 mb-0.5" />
          الأرشيف
        </button>
      </div>

      {/* 1. INDEXING SUB-TAB */}
      {hubSubTab === 'indexing' && (
        <div className="space-y-2.5">
          <div className="rounded-lg border border-[#C4A46A]/40 bg-[#F7F5F0] p-3 space-y-2.5">
            <div className="flex items-center justify-between border-b border-slate-200 pb-2">
              <div className="flex items-center gap-1.5">
                <Award className="w-4 h-4 text-[#0F2744]" />
                <span className="text-xs font-bold text-[#0F2744]">
                  الفهرسة الرسمية وختم الأصل (مرسوم 08-244)
                </span>
              </div>
              <span
                className={`text-[10px] px-2 py-0.5 rounded font-bold tabular-nums ${
                  isCurrentIndexed
                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                    : 'bg-amber-100 text-amber-800 border border-amber-300'
                }`}
              >
                {isCurrentIndexed ? 'مفهرس ومختوم' : 'مسودة قيد الفهرسة'}
              </span>
            </div>

            <p className="text-[11px] text-slate-600 leading-relaxed">
              يقترح النظام آلياً رقم الفهرس السنوي التالي (<strong className="tabular-nums">{suggestedNextIndex}</strong>) مع منع تكرار أي رقم فهرس في نفس السنة، وعند الاعتماد يتم توليد ملف الأصل وختمه ببصمة <strong>SHA-256</strong> غير قابلة للاستبدال.
            </p>

            <div className="grid grid-cols-2 gap-2">
              <div className="bg-white p-2 rounded border border-slate-200">
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-[10px] font-bold text-slate-500">
                    رقم الإيداع المؤقت
                  </label>
                  <button
                    type="button"
                    onClick={() => setDepositNo(suggestedNextDeposit)}
                    title="اقتراح رقم الإيداع التالي"
                    className="text-[9px] text-blue-700 hover:underline flex items-center gap-0.5"
                  >
                    <RefreshCw className="w-2.5 h-2.5" />
                    تلقائي
                  </button>
                </div>
                <input
                  type="text"
                  value={depositNo}
                  onChange={(e) => setDepositNo(e.target.value)}
                  className="w-full text-xs font-mono font-bold text-slate-800 bg-slate-50 border border-slate-200 rounded px-2 py-1 focus:outline-none focus:border-blue-600 tabular-nums"
                  placeholder={suggestedNextDeposit}
                />
                <span className="text-[9px] text-slate-400 mt-0.5 block">
                  للملفات المودعة قيد التحضير
                </span>
              </div>

              <div className="bg-white p-2 rounded border border-emerald-300">
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-[10px] font-bold text-emerald-800">
                    رقم الفهرس الرسمي السنوي
                  </label>
                  <button
                    type="button"
                    onClick={() => setIndexNo(suggestedNextIndex)}
                    title="اقتراح آخر فهرس + 1"
                    className="text-[9px] text-emerald-700 hover:underline flex items-center gap-0.5"
                  >
                    <RefreshCw className="w-2.5 h-2.5" />
                    التالي ({suggestedNextIndex})
                  </button>
                </div>
                <input
                  type="text"
                  value={indexNo}
                  onChange={(e) => setIndexNo(e.target.value)}
                  className="w-full text-xs font-mono font-bold text-emerald-900 bg-emerald-50/50 border border-emerald-300 rounded px-2 py-1 focus:outline-none focus:border-emerald-600 tabular-nums"
                  placeholder={suggestedNextIndex}
                />
                <span className="text-[9px] text-emerald-700 mt-0.5 block">
                  متسلسل غير قابل للتكرار
                </span>
              </div>
            </div>

            <div className="bg-white p-2 rounded border border-slate-200">
              <label className="block text-[10px] font-bold text-slate-600 mb-1">
                تاريخ توقيع وفهرسة العقد
              </label>
              <input
                type="date"
                value={contractDate}
                onChange={(e) => setContractDate(e.target.value)}
                className="w-full text-xs font-mono text-slate-800 bg-slate-50 border border-slate-200 rounded px-2 py-1 focus:outline-none focus:border-blue-600 tabular-nums"
              />
            </div>

            {activeSavedDoc?.sealedOriginalSha256 && (
              <div className="rounded border border-emerald-300 bg-emerald-50/70 p-2 text-[10px] text-emerald-900 space-y-1">
                <div className="flex items-center justify-between">
                  <span className="font-bold flex items-center gap-1">
                    <Lock className="w-3 h-3 text-emerald-700" />
                    الأصل المختوم محفوظ ومؤمّن:
                  </span>
                  <button
                    type="button"
                    onClick={() => handleDownloadSealedOriginal(activeSavedDoc)}
                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-emerald-700 text-white font-bold hover:bg-emerald-800"
                  >
                    <Download className="w-2.5 h-2.5" />
                    تنزيل الأصل (.docx)
                  </button>
                </div>
                <div className="font-mono text-[9px] text-slate-600 truncate tabular-nums">
                  المسار: {activeSavedDoc.sealedOriginalPath}
                </div>
                <div className="font-mono text-[9px] text-slate-600 truncate tabular-nums">
                  SHA-256: {activeSavedDoc.sealedOriginalSha256}
                </div>
              </div>
            )}

            <button
              type="button"
              disabled={isSealing}
              onClick={handleConfirmIndexing}
              className="w-full flex items-center justify-center gap-1.5 py-2 px-3 rounded-lg bg-[#15803D] hover:bg-emerald-800 disabled:opacity-60 text-white text-xs font-bold shadow-sm transition"
            >
              <ShieldCheck className="w-4 h-4" />
              {isSealing
                ? 'جاري توليد الأصل وختمه ببصمة SHA-256...'
                : 'اعتماد وفهرسة العقد رسمياً (ختم الأصل)'}
            </button>
          </div>

          {/* Recent Indexed Contracts list */}
          <div className="rounded-lg border border-slate-200 bg-white p-2.5 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold text-slate-700">
                آخر العقود المفهرسة في السجل السنوي
              </span>
              <span className="text-[10px] text-slate-500 tabular-nums">
                الإجمالي: {indexedContracts.length}
              </span>
            </div>
            {indexedContracts.length === 0 ? (
              <p className="text-[11px] text-slate-400 text-center py-2">
                لا توجد عقود مفهرسة بعد في هذه السنة
              </p>
            ) : (
              <div className="space-y-1.5 max-h-40 overflow-y-auto">
                {indexedContracts.slice(0, 6).map((doc) => (
                  <div
                    key={doc.id}
                    className="flex items-center justify-between text-[11px] p-1.5 rounded bg-slate-50 border border-slate-100"
                  >
                    <div className="truncate max-w-[175px]">
                      <span className="font-bold text-slate-800 block truncate">
                        {doc.title}
                      </span>
                      <span className="text-[9px] text-slate-500 tabular-nums">
                        إيداع: {doc.depositNumber || '—'} · الحالة:{' '}
                        {doc.status === 'published'
                          ? 'مشهر'
                          : doc.status === 'registered'
                          ? 'مسجل'
                          : 'مفهرس'}
                      </span>
                    </div>
                    <div className="flex items-center gap-1">
                      <span className="font-mono text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 tabular-nums">
                        {doc.indexNumber || doc.fieldValues?.['رقم_الفهرس']}
                      </span>
                      {doc.sealedOriginalSha256 && (
                        <button
                          type="button"
                          onClick={() => handleDownloadSealedOriginal(doc)}
                          title="تنزيل الأصل المختوم"
                          className="p-1 rounded text-emerald-700 hover:bg-emerald-100"
                        >
                          <Download className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* 2. REGISTRATION SUB-TAB */}
      {hubSubTab === 'registration' && (
        <div className="space-y-2.5">
          <div className="flex items-center justify-between bg-[#F7F5F0] p-2.5 rounded-lg border border-slate-200">
            <div>
              <h4 className="text-xs font-bold text-[#0F2744]">
                دفعات وجداول التسجيل الجبائي
              </h4>
              <p className="text-[10px] text-slate-500">
                ربط العقود المفهرسة بالدفعة وحقن [مراجع_التسجيل] تلقائياً
              </p>
            </div>
            <button
              type="button"
              onClick={openNewBatchDialog}
              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-[#0F2744] text-white text-[11px] font-bold hover:bg-slate-800 transition"
            >
              <Plus className="w-3.5 h-3.5" />
              دفعة جديدة
            </button>
          </div>

          {showNewBatchModal && (
            <div className="rounded-lg border border-blue-200 bg-blue-50/50 p-3 space-y-2">
              <div className="text-xs font-bold text-blue-950">
                {editingBatchId
                  ? 'تحديث دفعة التسجيل وإدراج رقم الوصل'
                  : 'إنشاء جدول إرسال للتسجيل الجبائي'}
              </div>
              <input
                type="text"
                value={newBatchInspectorate}
                onChange={(e) => setNewBatchInspectorate(e.target.value)}
                placeholder="مفتشية التسجيل والطابع..."
                className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5"
              />

              {/* Interactive checklist of indexed contracts */}
              {indexedContracts.length > 0 && (
                <div className="bg-white border border-slate-200 rounded p-2 space-y-1 max-h-32 overflow-y-auto">
                  <div className="text-[10px] font-bold text-slate-600 mb-1">
                    اختر العقود المفهرسة المشمولة في الدفعة:
                  </div>
                  {indexedContracts.map((doc) => {
                    const idxLabel =
                      doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'] || 'بدون فهرس';
                    const checked = selectedBatchDocIds.includes(doc.id);
                    return (
                      <label
                        key={doc.id}
                        className="flex items-center justify-between gap-2 text-[11px] py-1 px-1.5 rounded hover:bg-slate-50 cursor-pointer"
                      >
                        <div className="flex items-center gap-1.5 truncate">
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleSelectBatchDoc(doc)}
                            className="rounded border-slate-300"
                          />
                          <span className="truncate text-slate-800 font-medium">
                            {doc.title}
                          </span>
                        </div>
                        <span className="font-mono text-[10px] font-bold text-emerald-800 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 tabular-nums shrink-0">
                          {idxLabel}
                        </span>
                      </label>
                    );
                  })}
                </div>
              )}

              <input
                type="text"
                value={newBatchIndices}
                onChange={(e) => setNewBatchIndices(e.target.value)}
                placeholder="أرقام الفهارس (مفصولة بفاصلة: 1/2026، 2/2026)"
                className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5 font-mono tabular-nums"
              />

              <div className="grid grid-cols-2 gap-1.5">
                <input
                  type="text"
                  value={newBatchFees}
                  onChange={(e) => setNewBatchFees(e.target.value)}
                  placeholder="إجمالي الحقوق والرسوم (دج)"
                  className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5 font-mono tabular-nums"
                />
                <input
                  type="date"
                  value={newBatchRegDate}
                  onChange={(e) => setNewBatchRegDate(e.target.value)}
                  className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5 font-mono tabular-nums"
                />
              </div>

              <input
                type="text"
                value={newBatchReceipt}
                onChange={(e) => setNewBatchReceipt(e.target.value)}
                placeholder="رقم وصل التسجيل (عند استلامه لحقن [مراجع_التسجيل] آلياً)"
                className="w-full text-xs bg-white border border-emerald-300 rounded px-2 py-1.5 font-mono tabular-nums"
              />

              <div className="flex items-center gap-1.5 pt-1">
                <button
                  type="button"
                  onClick={handleSaveBatch}
                  className="flex-1 py-1.5 rounded bg-blue-700 text-white text-xs font-bold hover:bg-blue-800"
                >
                  حفظ ومزامنة العقود
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowNewBatchModal(false);
                    setEditingBatchId(null);
                  }}
                  className="px-3 py-1.5 rounded bg-slate-200 text-slate-700 text-xs font-bold"
                >
                  إلغاء
                </button>
              </div>
            </div>
          )}

          <div className="space-y-2 max-h-64 overflow-y-auto">
            {batches.map((batch) => (
              <div
                key={batch.id}
                className="rounded-lg border border-slate-200 bg-white p-2.5 space-y-1.5 shadow-xs"
              >
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-[#0F2744] tabular-nums">
                    دفعة رقم {batch.batchNumber}/{batch.year}
                  </span>
                  <div className="flex items-center gap-1">
                    <span
                      className={`text-[10px] px-2 py-0.5 rounded font-bold ${
                        batch.status === 'registered'
                          ? 'bg-emerald-100 text-emerald-800'
                          : 'bg-amber-100 text-amber-800'
                      }`}
                    >
                      {batch.status === 'registered'
                        ? `مسجلة (وصل ${batch.receiptNumber || ''})`
                        : 'مودعة للتسجيل'}
                    </span>
                    <button
                      type="button"
                      onClick={() => openEditBatchDialog(batch)}
                      title="تحديث رقم الوصل أو العقود"
                      className="text-[10px] px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 font-bold hover:bg-blue-100"
                    >
                      تحديث الوصل
                    </button>
                    <button
                      type="button"
                      onClick={() => handlePrintBordereau(batch)}
                      title="طباعة جدول الإرسال"
                      className="text-slate-500 hover:text-[#0F2744] p-0.5"
                    >
                      <Printer className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDeleteBatch(batch.id)}
                      className="text-slate-400 hover:text-red-600 p-0.5"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
                <div className="text-[11px] text-slate-600">
                  {batch.inspectorateName}
                </div>
                <div className="flex items-center justify-between text-[10px] text-slate-500 pt-1 border-t border-slate-100 tabular-nums">
                  <span>الفهارس: {(batch.indexNumbers || []).join(' ، ')}</span>
                  <span className="font-bold text-slate-800">
                    {batch.totalFees.toLocaleString('ar-DZ')} دج
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 3. LAND PUBLICATION SUB-TAB */}
      {hubSubTab === 'publication' && (
        <div className="space-y-2.5">
          <div className="flex items-center justify-between bg-[#F7F5F0] p-2.5 rounded-lg border border-slate-200">
            <div>
              <h4 className="text-xs font-bold text-[#0F2744]">
                متابعة الشهر العقاري
              </h4>
              <p className="text-[10px] text-slate-500">
                ربط العقد بالمحافظة العقارية وحقن [مراجع_الشهر] آلياً
              </p>
            </div>
            <button
              type="button"
              onClick={openNewPublicationDialog}
              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-[#0F2744] text-white text-[11px] font-bold hover:bg-slate-800 transition"
            >
              <Plus className="w-3.5 h-3.5" />
              إيداع شهر جديد
            </button>
          </div>

          {showNewPubModal && (
            <div className="rounded-lg border border-blue-200 bg-blue-50/50 p-3 space-y-2">
              <div className="text-xs font-bold text-blue-950">
                {editingPubId
                  ? 'تحديث بيانات الشهر العقاري (المجلد والرقم)'
                  : 'تسجيل إيداع بالمحافظة العقارية'}
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-600 mb-1">
                  العقد المرتبط بمسار الشهر:
                </label>
                <select
                  value={selectedPubDocId}
                  onChange={(e) => setSelectedPubDocId(e.target.value)}
                  className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5 font-bold text-slate-800"
                >
                  <option value="">-- اختر عقداً محفوظاً من المكتب --</option>
                  {savedDocuments.map((doc) => {
                    const idx = doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'] || 'مسودة';
                    return (
                      <option key={doc.id} value={doc.id}>
                        [{idx}] {doc.title}
                      </option>
                    );
                  })}
                </select>
              </div>

              <input
                type="text"
                value={newPubOffice}
                onChange={(e) => setNewPubOffice(e.target.value)}
                placeholder="المحافظة العقارية..."
                className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1.5"
              />
              <div className="grid grid-cols-2 gap-1.5">
                <input
                  type="text"
                  value={newPubVolume}
                  onChange={(e) => setNewPubVolume(e.target.value)}
                  placeholder="رقم المجلد (إن شُهر)"
                  className="w-full text-xs bg-white border border-emerald-300 rounded px-2 py-1.5 font-mono tabular-nums"
                />
                <input
                  type="text"
                  value={newPubNumber}
                  onChange={(e) => setNewPubNumber(e.target.value)}
                  placeholder="رقم الشهر (إن شُهر)"
                  className="w-full text-xs bg-white border border-emerald-300 rounded px-2 py-1.5 font-mono tabular-nums"
                />
              </div>
              <div className="grid grid-cols-2 gap-1.5">
                <div>
                  <label className="block text-[9px] text-slate-500">تاريخ الإيداع</label>
                  <input
                    type="date"
                    value={newPubDepositDate}
                    onChange={(e) => setNewPubDepositDate(e.target.value)}
                    className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1 font-mono tabular-nums"
                  />
                </div>
                <div>
                  <label className="block text-[9px] text-slate-500">تاريخ الشهر</label>
                  <input
                    type="date"
                    value={newPubPublicationDate}
                    onChange={(e) => setNewPubPublicationDate(e.target.value)}
                    className="w-full text-xs bg-white border border-slate-200 rounded px-2 py-1 font-mono tabular-nums"
                  />
                </div>
              </div>
              <div className="flex items-center gap-1.5 pt-1">
                <button
                  type="button"
                  onClick={handleSavePublication}
                  className="flex-1 py-1.5 rounded bg-blue-700 text-white text-xs font-bold hover:bg-blue-800"
                >
                  حفظ ومزامنة العقد
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowNewPubModal(false);
                    setEditingPubId(null);
                  }}
                  className="px-3 py-1.5 rounded bg-slate-200 text-slate-700 text-xs font-bold"
                >
                  إلغاء
                </button>
              </div>
            </div>
          )}

          <div className="space-y-2 max-h-64 overflow-y-auto">
            {publications.map((pub) => {
              const linkedDoc = savedDocuments.find((d) => d.id === pub.documentId);
              return (
                <div
                  key={pub.id}
                  className="rounded-lg border border-slate-200 bg-white p-2.5 space-y-1.5 shadow-xs"
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-[#0F2744] truncate max-w-[140px] tabular-nums">
                      {linkedDoc ? linkedDoc.title : `فهرس: ${pub.indexNumber || '—'}`}
                    </span>
                    <div className="flex items-center gap-1">
                      <span
                        className={`text-[10px] px-2 py-0.5 rounded font-bold flex items-center gap-1 ${
                          pub.status === 'published'
                            ? 'bg-emerald-100 text-emerald-800'
                            : pub.status === 'rejected'
                            ? 'bg-red-100 text-red-800'
                            : 'bg-amber-100 text-amber-800'
                        }`}
                      >
                        {pub.status === 'published' ? (
                          <>
                            <CheckCircle2 className="w-3 h-3" /> مشهر
                          </>
                        ) : pub.status === 'rejected' ? (
                          <>
                            <ShieldAlert className="w-3 h-3" /> رفض/تحفظ
                          </>
                        ) : (
                          <>
                            <Clock className="w-3 h-3" /> قيد الإيداع
                          </>
                        )}
                      </span>
                      <button
                        type="button"
                        onClick={() => openEditPublicationDialog(pub)}
                        className="text-[10px] px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 font-bold hover:bg-blue-100"
                      >
                        تحديث الشهر
                      </button>
                      <button
                        type="button"
                        onClick={() => handleDeletePub(pub.id)}
                        className="text-slate-400 hover:text-red-600 p-0.5"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                  <div className="text-[11px] text-slate-600">
                    {pub.conservationOffice} — إيداع: <span className="tabular-nums">{pub.depositDate}</span>
                  </div>
                  {pub.volume && pub.number && (
                    <div className="text-[10px] font-bold text-emerald-800 bg-emerald-50 px-2 py-1 rounded border border-emerald-200 tabular-nums">
                      مشهر بتاريخ {pub.publicationDate} — مجلد: {pub.volume} — رقم:{' '}
                      {pub.number}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* 4. DIGITAL ARCHIVE SUB-TAB */}
      {hubSubTab === 'archive' && (
        <div className="space-y-2.5">
          <div className="rounded-lg border border-slate-200 bg-[#F7F5F0] p-2.5 space-y-2">
            <div className="flex items-center gap-1.5">
              <div className="relative flex-1">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute right-2.5 top-2" />
                <input
                  type="text"
                  value={archiveQuery}
                  onChange={(e) => setArchiveQuery(e.target.value)}
                  placeholder="بحث برقم الفهرس، الطرف، الموضوع..."
                  className="w-full text-xs bg-white border border-slate-200 rounded-lg pr-8 pl-2 py-1.5 focus:outline-none focus:border-blue-600"
                />
              </div>
              <select
                value={archiveYearFilter}
                onChange={(e) => setArchiveYearFilter(e.target.value)}
                className="text-xs bg-white border border-slate-200 rounded-lg px-2 py-1.5 font-bold text-slate-700 tabular-nums"
              >
                <option value="all">كل السنوات</option>
                {archiveYears.map((yr) => (
                  <option key={yr} value={yr}>
                    سنة {yr}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center justify-between text-[10px] text-slate-500 px-1 tabular-nums">
              <span>إجمالي العقود المؤرشفة: {filteredArchiveDocs.length}</span>
              <span>محمية ببصمة SHA-256</span>
            </div>
          </div>

          <div className="space-y-1.5 max-h-64 overflow-y-auto">
            {filteredArchiveDocs.length === 0 ? (
              <div className="text-center py-6 text-xs text-slate-400">
                لا توجد عقود مطابقة للبحث في الأرشيف
              </div>
            ) : (
              filteredArchiveDocs.map((doc) => {
                const idx = doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'];
                return (
                  <div
                    key={doc.id}
                    className="rounded-lg border border-slate-200 bg-white p-2 flex items-center justify-between hover:border-blue-300 transition"
                  >
                    <div className="space-y-0.5 truncate max-w-[185px]">
                      <div className="text-xs font-bold text-slate-800 truncate">
                        {doc.title}
                      </div>
                      <div className="text-[10px] text-slate-500 flex items-center gap-1.5 tabular-nums">
                        <span>{doc.modelTitle || doc.clientName || 'عقد توثيقي'}</span>
                        <span>·</span>
                        <span>{doc.createdAt.slice(0, 10)}</span>
                      </div>
                    </div>
                    <div className="flex items-center gap-1">
                      {idx ? (
                        <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 flex items-center gap-1 tabular-nums">
                          <FileCheck className="w-3 h-3" />
                          {idx}
                        </span>
                      ) : (
                        <span className="text-[10px] px-2 py-0.5 rounded bg-slate-100 text-slate-600">
                          مسودة
                        </span>
                      )}
                      {doc.sealedOriginalSha256 && (
                        <button
                          type="button"
                          onClick={() => handleDownloadSealedOriginal(doc)}
                          title="تنزيل الأصل المختوم"
                          className="p-1 rounded bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200"
                        >
                          <Download className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
}
