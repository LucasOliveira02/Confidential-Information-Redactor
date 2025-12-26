import './globals.css';
import type { Metadata } from 'next';
import Script from 'next/script';
import React from 'react';
import { Polyfill } from './components/polyfill';

export const metadata: Metadata = {
    title: 'Confidential Information Redactor',
};

export default function RootLayout({
    children,
}: {
    children: React.ReactNode;
}) {
    return (
        <html lang="en" suppressHydrationWarning>
            <head>
                <Polyfill />
            </head>
            <body suppressHydrationWarning>
                <Script
                    src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
                    strategy="afterInteractive"
                />
                {children}
            </body>
        </html>
    );
}
