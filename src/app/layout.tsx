
import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'POD Management Tool | Delhivery · Palam Vihar RPC',
  description: 'Field operation tool for EOD rejection management and remark replacer. Built for efficiency.',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
