'use client';

import React, { useState } from 'react';
import {
  BookOpen,
  Building2,
  Calendar,
  FileCheck,
  Hash,
  Layers,
  Plus,
  ShieldCheck,
  Trash2,
} from 'lucide-react';
import {
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
} from '@/lib/types';
import {
  loadRegistrationBatches,
  saveRegistrationBatch,
  deleteRegistrationBatch,
  loadPublicationTracks,
  savePublicationTrack,
  deletePublicationTrack,
  saveDocumentRecord,
} from '@/lib/storage';

interface NotaryHubPanelProps {
  documents: SavedDocument[];
  activeDocument: SavedDocument | null;
  onOpenDocument: (doc: SavedDocument) => void;
  showToast: (msg: string) => void;
}

export default function NotaryHubPanel({
  documents,
  activeDocument,
  onOpenDocument,
  showToast,
}: NotaryHubPanelProps) {
  const [hubTab, setHubTab] = useState<'indexing' | 'registration' | 'conservation'>('indexing');

  // 1. Indexing state
  const [searchDocTerm, setSearchDocTerm] = useState('');
  const [depositNoInput, setDepositNoInput] = useState('');
  const [indexNoInput, setIndexNoInput] = useState('');

  // 2. Registration Batches state
  const [batches, setBatches] = useState<RegistrationBatch[]>([]);
  const [isNewBatchModalOpen, setIsNewBatchModalOpen] = useState(false);
  const [batchDate, setBatchDate] = useState(new Date().toISOString().slice(0, 10));
  const [batchIndexes, setBatchIndexes] = useState('');
  const [batchValue, setBatchValue] = useState('');
  const [batchCount, setBatchCount] = useState('1');
  const [batchRegRef, setBatchRegRef] = useState('');
  const [batchCheckNo, setBatchCheckNo] = useState('');
  const [batchStatus, setBatchStatus] = useState<'pending' | 'registered' | 'rejected'>('pending');
  const [batchInspectorate, setBatchInspectorate] = useState('مفتشية التسجيل والطابع');

  // 3. Conservation Publication state
  const [tracks, setTracks] = useState<PublicationTrack[]>([]);
  const [isNewTrackModalOpen, setIsNewTrackModalOpen] = useState(false);
  const [trackIndexNo, setTrackIndexNo] = useState('');
  const [trackDepositDate, setTrackDepositDate] = useState(new Date().toISOString().slice(0, 10));
  const [trackPubDate, setTrackPubDate] = useState('');
  const [trackFees, setTrackFees] = useState('');
  const [trackFeesPaid, setTrackFeesPaid] = useState(false);
  const [trackVolume, setTrackVolume] = useState('');
  const [trackNumber, setTrackNumber] = useState('');
  const [trackWilaya, setTrackWilaya] = useState('الجزائر');
  const [trackOffice, setTrackOffice] = useState('المحافظة العقارية الأولى');

  React.useEffect(() => {
    loadRegistrationBatches().then(setBatches).catch(() => {});
    loadPublicationTracks().then(setTracks).catch(() => {});
  }, []);

  const handleSaveBatch = async (e: React.FormEvent) => {
    e.preventDefault();
    const newBatch: RegistrationBatch = {
      id: `batch_${Date.now()}`,
      year: new Date().getFullYear(),
      statusDate: batchDate,
      contractIds: activeDocument ? [activeDocument.id] : [],
      indexNumbers: batchIndexes ? batchIndexes.split(',').map(s => s.trim()) : [],
      totalFees: parseFloat(batchValue || '0'),
      paymentMethod: batchCheckNo ? 'صك بنكي' : 'نقد',
      chequeRef: batchCheckNo,
      receiptNumber: batchRegRef,
      inspectorateName: batchInspectorate,
      status: batchStatus,
      updatedAt: new Date().toISOString(),
    };
    await saveRegistrationBatch(newBatch);
    setBatches(await loadRegistrationBatches());
    setIsNewBatchModalOpen(false);
    setBatchIndexes('');
    setBatchValue('');
    setBatchRegRef('');
    setBatchCheckNo('');
    showToast('تم حفظ دفعة التسجيل لدى المفتشية بنجاح');
  };

  const handleSaveTrack = async (e: React.FormEvent) => {
    e.preventDefault();
    const newTrk: PublicationTrack = {
      id: `trk_${Date.now()}`,
      documentId: activeDocument ? activeDocument.id : 'doc_manual',
      indexNumber: trackIndexNo || '142/2026',
      wilaya: trackWilaya,
      conservationOffice: trackOffice,
      depositDate: trackDepositDate,
      publicationDate: trackPubDate,
      volume: trackVolume,
      number: trackNumber,
      fees: parseFloat(trackFees || '0'),
      feesPaid: trackFeesPaid,
      status: trackPubDate ? 'published' : 'deposited',
      updatedAt: new Date().toISOString(),
    };
    await savePublicationTrack(newTrk);
    setTracks(await loadPublicationTracks());
    setIsNewTrackModalOpen(false);
    setTrackIndexNo('');
    setTrackFees('');
    setTrackVolume('');
    setTrackNumber('');
    showToast('تم حفظ سجل الشهر العقاري بنجاح');
  };

  const filteredDocs = documents.filter(d =>
    d.title.toLowerCase().includes(searchDocTerm.toLowerCase())
  );

  return (
    <div className="flex flex-col h-full bg-slate-50 text-slate-900 overflow-hidden">
      {/* Hub Header & Sub-tabs */}
      <div className="px-4 py-3 bg-white border-b border-slate-200 shrink-0">
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-5 h-5 text-[#0F2744]" />
            <h2 className="text-sm font-bold text-[#0F2744]">منظومة العمل التوثيقي والرقابي</h2>
          </div>
          <span className="text-xs px-2 py-0.5 rounded bg-blue-50 text-blue-800 font-medium">مرسوم 08-244</span>
        </div>
        <div className="grid grid-cols-3 gap-1 bg-slate-100 p-1 rounded-lg">
          <button
            onClick={() => setHubTab('indexing')}
            className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded text-xs font-medium transition-all ${
              hubTab === 'indexing'
                ? 'bg-white text-[#0F2744] shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <BookOpen className="w-3.5 h-3.5" />
            <span>الفهرسة والإيداع</span>
          </button>
          <button
            onClick={() => setHubTab('registration')}
            className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded text-xs font-medium transition-all ${
              hubTab === 'registration'
                ? 'bg-white text-[#0F2744] shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>دفعات المفتشية</span>
          </button>
          <button
            onClick={() => setHubTab('conservation')}
            className={`flex items-center justify-center gap-1.5 py-1.5 px-2 rounded text-xs font-medium transition-all ${
              hubTab === 'conservation'
                ? 'bg-white text-[#0F2744] shadow-2xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Building2 className="w-3.5 h-3.5" />
            <span>الشهر العقاري</span>
          </button>
        </div>
      </div>

      {/* Hub Content Body */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {/* TAB 1: INDEXING & DEPOSIT (فهرس رسمي + إيداع) */}
        {hubTab === 'indexing' && (
          <div className="space-y-4">
            <div className="bg-white rounded-xl border border-slate-200 p-3 shadow-2xs">
              <h3 className="text-xs font-bold text-slate-800 mb-2 flex items-center gap-1.5">
                <Hash className="w-4 h-4 text-emerald-600" />
                <span>إدارة رقم الإيداع وفهرسة العقد</span>
              </h3>
              <p className="text-[11px] text-slate-500 mb-3 leading-relaxed">
                رقم الإيداع مستقل عن رقم الفهرس السنوي. الفهرس الرسمي دفتر متسلسل بلا فراغ (مرسوم 08-244). زر الفهرسة يقترح الرقم التالي ويثبّت النسخة الأصلية.
              </p>

              <div className="space-y-3">
                <div>
                  <label className="block text-[11px] font-medium text-slate-700 mb-1">اختر العقد الحالي أو من الأرشيف:</label>
                  <input
                    type="text"
                    placeholder="ابحث في العقود..."
                    value={searchDocTerm}
                    onChange={(e) => setSearchDocTerm(e.target.value)}
                    className="w-full px-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-[#0F2744]"
                  />
                </div>

                <div className="max-h-48 overflow-y-auto space-y-1.5 border border-slate-100 rounded-lg p-1.5 bg-slate-50/50">
                  {filteredDocs.map((doc) => (
                    <div
                      key={doc.id}
                      onClick={() => onOpenDocument(doc)}
                      className={`flex items-center justify-between p-2 rounded-lg cursor-pointer transition-colors text-xs ${
                        activeDocument?.id === doc.id
                          ? 'bg-blue-50 border border-blue-200 text-blue-900 font-semibold'
                          : 'bg-white border border-slate-200 hover:bg-slate-100 text-slate-800'
                      }`}
                    >
                      <span className="truncate">{doc.title}</span>
                      <span className="text-[10px] text-slate-400 shrink-0">
                        {doc.updatedAt.slice(0, 10)}
                      </span>
                    </div>
                  ))}
                </div>

                {activeDocument && (
                  <div className="bg-emerald-50/70 border border-emerald-200 rounded-lg p-3 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-emerald-900">العقد النشط: {activeDocument.title}</span>
                      <span className="text-[10px] bg-emerald-200/70 text-emerald-900 px-2 py-0.5 rounded-full font-medium">جاهز للفهرسة</span>
                    </div>

                    <div className="grid grid-cols-2 gap-2 pt-1">
                      <div>
                        <label className="block text-[10px] font-medium text-emerald-800 mb-0.5">رقم الإيداع:</label>
                        <input
                          type="text"
                          value={depositNoInput || '35/2026'}
                          onChange={(e) => setDepositNoInput(e.target.value)}
                          className="w-full px-2 py-1 text-xs bg-white border border-emerald-300 rounded font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-medium text-emerald-800 mb-0.5">رقم الفهرس السنوي (مقترح):</label>
                        <input
                          type="text"
                          value={indexNoInput || '142/2026'}
                          onChange={(e) => setIndexNoInput(e.target.value)}
                          className="w-full px-2 py-1 text-xs bg-white border border-emerald-300 rounded font-mono font-bold text-emerald-900"
                        />
                      </div>
                    </div>

                    <button
                      onClick={async () => {
                        await saveDocumentRecord({
                          ...activeDocument,
                          title: `${activeDocument.title} [فهرس: ${indexNoInput || '142/2026'}]`,
                          updatedAt: new Date().toISOString(),
                        });
                        showToast(`تمت فهرسة العقد برقم فهرس (${indexNoInput || '142/2026'}) وإغلاق الأصل للتوثيق الرسمي`);
                      }}
                      className="w-full mt-2 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow-2xs transition-colors"
                    >
                      <FileCheck className="w-4 h-4" />
                      <span>اعتماد وفهرسة العقد رسمياً (ختم الأصل)</span>
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: REGISTRATION BATCHES (مفتشية التسجيل) */}
        {hubTab === 'registration' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between bg-white rounded-xl border border-slate-200 p-3 shadow-2xs">
              <div>
                <h3 className="text-xs font-bold text-slate-900">دفعات الإيداع لدى مفتشية التسجيل والطابع</h3>
                <p className="text-[11px] text-slate-500">متابعة إرسال وفهارس العقود وحالة الوصل والصكوك.</p>
              </div>
              <button
                onClick={() => setIsNewBatchModalOpen(true)}
                className="px-3 py-1.5 bg-[#0F2744] hover:bg-[#1E3A8A] text-white rounded-lg text-xs font-bold flex items-center gap-1 shadow-2xs transition-colors"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>دفعة جديدة</span>
              </button>
            </div>

            <div className="space-y-3">
              {batches.map((batch) => (
                <div key={batch.id} className="bg-white rounded-xl border border-slate-200 p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <div className="flex items-center gap-2">
                      <Calendar className="w-3.5 h-3.5 text-blue-900" />
                      <span className="text-xs font-bold text-slate-900">تاريخ الإرسال: {batch.statusDate}</span>
                    </div>
                    <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold ${
                      batch.status === 'registered' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                    }`}>
                      {batch.status === 'registered' ? 'مسجل' : 'قيد الانتظار'}
                    </span>
                  </div>

                  <div className="grid grid-cols-2 gap-2 text-xs">
                    <div>
                      <span className="text-slate-500 block text-[10px]">فهارس العقود:</span>
                      <span className="font-mono font-bold text-slate-800">{batch.indexNumbers?.join(', ') || '-'}</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">قيمة الرسوم (دج):</span>
                      <span className="font-mono font-bold text-emerald-700">{batch.totalFees.toLocaleString()} دج</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">رقم الوصل:</span>
                      <span className="font-mono text-slate-800">{batch.receiptNumber || 'لم يُحدد بعد'}</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">الصك / وسيلة الدفع:</span>
                      <span className="font-mono text-slate-800">{batch.chequeRef || batch.paymentMethod}</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                    <button
                      onClick={async () => {
                        await deleteRegistrationBatch(batch.id);
                        setBatches(await loadRegistrationBatches());
                        showToast('تم حذف دفعة التسجيل');
                      }}
                      className="p-1 text-red-600 hover:bg-red-50 rounded transition-colors"
                      title="حذف"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              ))}

              {batches.length === 0 && (
                <div className="text-center py-8 bg-white rounded-xl border border-slate-200 text-slate-400 text-xs">
                  لا توجد دفعات إيداع مسجلة بعد.
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 3: LAND CONSERVATION & PUBLICATION (الشهر العقاري) */}
        {hubTab === 'conservation' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between bg-white rounded-xl border border-slate-200 p-3 shadow-2xs">
              <div>
                <h3 className="text-xs font-bold text-slate-900">متابعة الشهر العقاري وجداول الإرسال</h3>
                <p className="text-[11px] text-slate-500">حسب المحافظة العقارية والولاية ورسوم ومراجع الشهر.</p>
              </div>
              <button
                onClick={() => setIsNewTrackModalOpen(true)}
                className="px-3 py-1.5 bg-[#0F2744] hover:bg-[#1E3A8A] text-white rounded-lg text-xs font-bold flex items-center gap-1 shadow-2xs transition-colors"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>إضافة سجل</span>
              </button>
            </div>

            <div className="space-y-3">
              {tracks.map((trk) => (
                <div key={trk.id} className="bg-white rounded-xl border border-slate-200 p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <div className="flex items-center gap-2">
                      <Building2 className="w-3.5 h-3.5 text-blue-900" />
                      <span className="text-xs font-bold text-slate-900">فهرس رقم: {trk.indexNumber}</span>
                    </div>
                    <span className="text-[10px] bg-blue-50 text-blue-900 px-2 py-0.5 rounded font-medium">
                      {trk.conservationOffice} ({trk.wilaya})
                    </span>
                  </div>

                  <div className="grid grid-cols-2 gap-2 text-xs">
                    <div>
                      <span className="text-slate-500 block text-[10px]">رسوم الشهر:</span>
                      <span className="font-mono text-slate-800">{trk.fees.toLocaleString()} دج ({trk.feesPaid ? 'مدفوعة' : 'غير مدفوعة'})</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">تاريخ الإيداع:</span>
                      <span className="font-mono text-slate-800">{trk.depositDate || '-'}</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">المجلد / الرقم:</span>
                      <span className="font-mono font-bold text-slate-800">مجلد {trk.volume || '-'} رقم {trk.number || '-'}</span>
                    </div>
                    <div>
                      <span className="text-slate-500 block text-[10px]">الحالة:</span>
                      <span className={`font-bold ${trk.status === 'published' ? 'text-emerald-700' : 'text-amber-700'}`}>
                        {trk.status === 'published' ? 'تم الشهر' : 'قيد الإيداع'}
                      </span>
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                    <button
                      onClick={async () => {
                        await deletePublicationTrack(trk.id);
                        setTracks(await loadPublicationTracks());
                        showToast('تم حذف السجل');
                      }}
                      className="p-1 text-red-600 hover:bg-red-50 rounded transition-colors"
                      title="حذف"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              ))}

              {tracks.length === 0 && (
                <div className="text-center py-8 bg-white rounded-xl border border-slate-200 text-slate-400 text-xs">
                  لا توجد سجلات شهر عقاري حالياً.
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* NEW REGISTRATION BATCH MODAL */}
      {isNewBatchModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 space-y-3 shadow-xl">
            <h3 className="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">تسجيل دفعة إيداع جديدة لدى المفتشية</h3>
            <form onSubmit={handleSaveBatch} className="space-y-3 text-xs">
              <div>
                <label className="block text-slate-700 mb-1">تاريخ الإرسال:</label>
                <input
                  type="date"
                  value={batchDate}
                  onChange={(e) => setBatchDate(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-200 rounded-lg"
                />
              </div>
              <div>
                <label className="block text-slate-700 mb-1">أرقام الفهارس (مفصولة بفاصلة):</label>
                <input
                  type="text"
                  placeholder="140/2026, 141/2026"
                  value={batchIndexes}
                  onChange={(e) => setBatchIndexes(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                  required
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-700 mb-1">مجموع الرسوم (دج):</label>
                  <input
                    type="number"
                    placeholder="25000"
                    value={batchValue}
                    onChange={(e) => setBatchValue(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-700 mb-1">حالة التسجيل:</label>
                  <select
                    value={batchStatus}
                    onChange={(e) => setBatchStatus(e.target.value as any)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg bg-white"
                  >
                    <option value="pending">قيد الانتظار</option>
                    <option value="registered">مسجل</option>
                  </select>
                </div>
              </div>
              <div>
                <label className="block text-slate-700 mb-1">رقم وصل المفتشية:</label>
                <input
                  type="text"
                  placeholder="وصل رقم ..."
                  value={batchRegRef}
                  onChange={(e) => setBatchRegRef(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                />
              </div>
              <div>
                <label className="block text-slate-700 mb-1">رقم الصك البنكي:</label>
                <input
                  type="text"
                  placeholder="صك رقم ..."
                  value={batchCheckNo}
                  onChange={(e) => setBatchCheckNo(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsNewBatchModalOpen(false)}
                  className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-medium"
                >
                  إلغاء
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-[#0F2744] hover:bg-[#1E3A8A] text-white rounded-lg font-bold"
                >
                  حفظ الدفعة
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* NEW PUBLICATION TRACK MODAL */}
      {isNewTrackModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 space-y-3 shadow-xl">
            <h3 className="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">إضافة سجل شهر عقاري جديد</h3>
            <form onSubmit={handleSaveTrack} className="space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-700 mb-1">رقم الفهرس:</label>
                  <input
                    type="text"
                    placeholder="142/2026"
                    value={trackIndexNo}
                    onChange={(e) => setTrackIndexNo(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono font-bold"
                    required
                  />
                </div>
                <div>
                  <label className="block text-slate-700 mb-1">رسوم الشهر (دج):</label>
                  <input
                    type="number"
                    placeholder="15000"
                    value={trackFees}
                    onChange={(e) => setTrackFees(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-700 mb-1">تاريخ الإيداع:</label>
                  <input
                    type="date"
                    value={trackDepositDate}
                    onChange={(e) => setTrackDepositDate(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-slate-700 mb-1">تاريخ الشهر:</label>
                  <input
                    type="date"
                    value={trackPubDate}
                    onChange={(e) => setTrackPubDate(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg"
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-700 mb-1">المجلد:</label>
                  <input
                    type="text"
                    placeholder="12"
                    value={trackVolume}
                    onChange={(e) => setTrackVolume(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-700 mb-1">الرقم:</label>
                  <input
                    type="text"
                    placeholder="345"
                    value={trackNumber}
                    onChange={(e) => setTrackNumber(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-700 mb-1">الولاية:</label>
                  <input
                    type="text"
                    value={trackWilaya}
                    onChange={(e) => setTrackWilaya(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-slate-700 mb-1">المحافظة العقارية:</label>
                  <input
                    type="text"
                    value={trackOffice}
                    onChange={(e) => setTrackOffice(e.target.value)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-lg"
                  />
                </div>
              </div>
              <label className="flex items-center gap-2 cursor-pointer pt-1">
                <input
                  type="checkbox"
                  checked={trackFeesPaid}
                  onChange={(e) => setTrackFeesPaid(e.target.checked)}
                  className="rounded border-slate-300 text-blue-900"
                />
                <span className="font-medium text-slate-800">رسوم الشهر مدفوعة</span>
              </label>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsNewTrackModalOpen(false)}
                  className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-medium"
                >
                  إلغاء
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-[#0F2744] hover:bg-[#1E3A8A] text-white rounded-lg font-bold"
                >
                  حفظ السجل
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
