/*
 * This file defines the top‑level layout for the Next.js application.
 * It imports the global stylesheet, sets up basic metadata (title and
 * favicon), and injects the Office.js script required for the Word add‑in.
 *
 * The layout is deliberately minimal – it only renders the HTML skeleton
 * and the `children` prop that contains the rest of the UI (e.g. the
 * `TaskPane`). Keeping the layout simple ensures fast hydration and
 * avoids unnecessary re‑renders.
 */
import './globals.css';
import type { Metadata } from 'next';
import Script from 'next/script';
import React from 'react';
import { Polyfill } from './components/polyfill';

/**
 * Global metadata for the application.
 *
 * - `title`: Shown in the browser tab and used by SEO tools.
 * - `icons`: Provides the favicon for the add‑in.
 */
export const metadata: Metadata = {
    title: 'Confidential Information Redactor',
    icons: {
        icon: '/icon.png',
    },
};

/**
 * RootLayout – wraps the entire application.
 *
 * @param children – The page‑level React nodes rendered inside the layout.
 * @returns The HTML document structure with the Office.js script injected.
 */
export default function RootLayout({
    children,
}: {
    children: React.ReactNode;
}) {
    return (
        <html lang="en" suppressHydrationWarning>
            <head>
                {/* Polyfill component ensures required Office.js APIs are available */}
                <Polyfill />
            </head>
            <body suppressHydrationWarning>
                {/* Load Office.js – required for Word add‑in functionality */}
                <Script
                    src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
                    strategy="afterInteractive"
                />
                {children}
            </body>
        </html>
    );
}
