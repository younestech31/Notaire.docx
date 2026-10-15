'use client';

import dynamic from 'next/dynamic';

const NotaryEditorApp = dynamic(() => import('@/components/NotaryEditorApp'), {
  ssr: false,
  loading: () => (
    <div className="min-h-screen flex items-center justify-center bg-[#F7F5F0]">
      <div className="flex flex-col items-center gap-3">
        <div className="w-8 h-8 border-4 border-[#0F2744] border-t-transparent rounded-full animate-spin"></div>
        <p className="text-sm font-medium text-[#0F2744]">جاري تحميل محرر الموثق الرقمي...</p>
      </div>
    </div>
  ),
});

export default function HomePage() {
  return <NotaryEditorApp />;
}
