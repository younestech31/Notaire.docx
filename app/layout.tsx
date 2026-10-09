import type {Metadata} from 'next';
import {Cairo, IBM_Plex_Mono} from 'next/font/google';
import './globals.css';

const cairo = Cairo({
  subsets: ['arabic', 'latin'],
  weight: ['400', '500', '600', '700'],
  variable: '--font-cairo',
  display: 'swap',
});

const ibmPlexMono = IBM_Plex_Mono({
  subsets: ['latin'],
  weight: ['400', '500', '600'],
  variable: '--font-mono-num',
  display: 'swap',
});

export const metadata: Metadata = {
  title: 'الموثق الرقمي — محرر العقود التوثيقية',
  description:
    'منصة مكتبية متخصصة لتحرير العقود التوثيقية باللغة العربية مع استيراد وتصدير ملفات Word (.docx) وإدارة قوالب المكتب وجداول الوصف التقسيمي للعقارات.',
  openGraph: {
    title: 'الموثق الرقمي — محرر العقود التوثيقية',
    description:
      'منصة مكتبية متخصصة لتحرير العقود التوثيقية باللغة العربية مع استيراد وتصدير ملفات Word (.docx) وإدارة قوالب المكتب وجداول الوصف التقسيمي للعقارات.',
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: 'الموثق الرقمي — محرر العقود التوثيقية',
    description:
      'منصة مكتبية متخصصة لتحرير العقود التوثيقية باللغة العربية مع استيراد وتصدير ملفات Word (.docx) وإدارة قوالب المكتب وجداول الوصف التقسيمي للعقارات.',
  },
};

export default function RootLayout({children}: {children: React.ReactNode}) {
  return (
    <html lang="ar" dir="rtl" className={`${cairo.variable} ${ibmPlexMono.variable}`}>
      <body
        suppressHydrationWarning
        className="bg-slate-100 text-slate-900 antialiased selection:bg-blue-900 selection:text-white"
        style={{fontFamily: 'var(--font-cairo), sans-serif'}}
      >
        {children}
      </body>
    </html>
  );
}
