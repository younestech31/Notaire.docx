import type {Metadata, Viewport} from 'next';
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

export const viewport: Viewport = {
  themeColor: '#0F2744',
};

export const metadata: Metadata = {
  title: 'الموثق الرقمي — محرر العقود التوثيقية',
  description:
    'منصة مكتبية متخصصة لتحرير العقود التوثيقية باللغة العربية مع استيراد وتصدير ملفات Word (.docx) وإدارة قوالب المكتب وجداول الوصف التقسيمي للعقارات.',
  manifest: '/site.webmanifest',
  icons: {
    icon: [
      {url: '/favicon.svg', type: 'image/svg+xml'},
      {url: '/favicon.ico', sizes: 'any'},
      {url: '/icons/icon-16.png', sizes: '16x16', type: 'image/png'},
      {url: '/icons/icon-32.png', sizes: '32x32', type: 'image/png'},
      {url: '/icons/icon-48.png', sizes: '48x48', type: 'image/png'},
      {url: '/icons/icon-192.png', sizes: '192x192', type: 'image/png'},
      {url: '/icons/icon-512.png', sizes: '512x512', type: 'image/png'},
    ],
    apple: [{url: '/icons/apple-touch-icon.png', sizes: '180x180', type: 'image/png'}],
  },
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
        className="bg-[#F7F5F0] text-[#0F2744] antialiased selection:bg-[#1E3A8A] selection:text-white"
        style={{fontFamily: 'var(--font-cairo), sans-serif'}}
      >
        {children}
      </body>
    </html>
  );
}
