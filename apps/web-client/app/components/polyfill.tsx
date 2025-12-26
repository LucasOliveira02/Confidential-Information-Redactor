"use client";

import React from 'react';

/**
 * A component that injects critical polyfills strictly before Next.js hydration.
 * patches:
 * 1. console.error: Suppresses known non-critical React warnings.
 * 2. window.history: Prevents "SecurityError" on restricted Office runtime environments.
 */
export function Polyfill() {
    const polyfillScript = `
(function() {
    try {
        if (typeof window === 'undefined') return;
        
        // 1. Console Patching (Suppress known persistent React warnings)
        const originalConsoleError = console.error;
        if (originalConsoleError) {
            console.error = function(...args) {
                 if (args[0] && typeof args[0] === 'string' && 
                    (args[0].includes("Can't perform a React state update") || 
                     args[0].includes("Cannot read properties of null (reading 'bind')"))) {
                    return; 
                }
                originalConsoleError.apply(console, args);
            };
        }

        // 2. History API Patching (Fix for Next.js router in Office Add-ins)
        const noop = function() {};
        
        // Strategy: Force-patch history methods to no-ops to prevent SecurityErrors
        try {
            if (!window.history) {
                window.history = {};
            }
            
            // Try direct assignment first
            window.history.pushState = noop;
            window.history.replaceState = noop;
            
        } catch (e) {
            // If direct assignment fails (read-only), try defining property
            try {
                Object.defineProperty(window.history, 'pushState', { value: noop, writable: true, configurable: true });
                Object.defineProperty(window.history, 'replaceState', { value: noop, writable: true, configurable: true });
            } catch (e2) {
                // If that fails, try patching the prototype
                 try {
                    if (window.History && window.History.prototype) {
                        window.History.prototype.pushState = noop;
                        window.History.prototype.replaceState = noop;
                    }
                } catch (e3) {}
            }
        }

    } catch (e) {
        // Fail silently
    }
})();
`;

    return (
        <script
            dangerouslySetInnerHTML={{
                __html: polyfillScript,
            }}
        />
    );
}
