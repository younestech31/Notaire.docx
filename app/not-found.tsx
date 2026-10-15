import Link from 'next/link';

export default function NotFound() {
  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-4 text-center bg-[#F7F5F0]">
      <h2 className="text-2xl font-bold text-[#0F2744] mb-2">الصفحة غير موجودة</h2>
      <p className="text-gray-600 mb-4">الصفحة المطلوبة غير متوفرة في منصة الموثق الرقمي.</p>
      <Link
        href="/"
        className="px-4 py-2 bg-[#1E3A8A] text-white rounded-md hover:bg-[#152a65] transition-colors"
      >
        العودة للرئيسية
      </Link>
    </div>
  );
}
