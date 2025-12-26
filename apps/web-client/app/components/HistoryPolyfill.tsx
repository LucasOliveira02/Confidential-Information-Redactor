"use client";

import React from 'react';

/**
 * A component that injects a critical polyfill for window.history and console.error
 * strictly before Next.js hydration kicks in.
 */
export function HistoryPolyfill() {
    const polyfillScript = `
(function() {
    try {
        if (typeof window === 'undefined') return;
        
        console.log('[Polyfill] Starting initialization...');
        
        // 1. Console Patching
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

        // 2. History API Patching
        function noop() { 
            console.log('[Polyfill] History method called (no-op)');
        }
        
        const mockHistory = {
            pushState: noop,
            replaceState: noop,
            go: noop,
            back: noop,
            forward: noop,
            state: {},
            length: 1
        };

        // Strategy 1: Direct Assignment (if writable)
        try {
            if (!window.history) window.history = mockHistory;
            
            window.history.pushState = noop;
            window.history.replaceState = noop;
            console.log('[Polyfill] Direct assignment attempted');
        } catch (e) {
            console.log('[Polyfill] Direct assignment failed');
        }
            
        // Strategy 2: defineProperty on Instance (forces read-only overwrite)
        try {
            Object.defineProperty(window.history, 'pushState', { value: noop, writable: true, configurable: true });
            Object.defineProperty(window.history, 'replaceState', { value: noop, writable: true, configurable: true });
            console.log('[Polyfill] Instance defineProperty attempted');
        } catch (e) {
             console.log('[Polyfill] Instance defineProperty failed');
        }
        
        // Strategy 3: Prototype Patching
        try {
            if (window.History && window.History.prototype) {
                window.History.prototype.pushState = noop;
                window.History.prototype.replaceState = noop;
                console.log('[Polyfill] Prototype patched');
            }
        } catch (e) {}

        // Strategy 4: Delete and Replace (Nuclear)
        // Check if we succeeded yet
        if (typeof window.history.replaceState !== 'function') {
            console.log('[Polyfill] replaceState still missing, attempting Nuclear replacement');
            try {
                // Try to delete the property first
                try { delete window.history; } catch(e){}
                
                // Then redefine it on window
                Object.defineProperty(window, 'history', {
                    value: mockHistory,
                    writable: true,
                    configurable: true
                });
                console.log('[Polyfill] Window property redefined');
            } catch(e) {
                console.error('[Polyfill] Nuclear option failed', e);
            }
        }
        
        // Final Verification
        if (typeof window.history.replaceState === 'function') {
            console.log('[Polyfill] Success: replaceState is a function');
        } else {
            console.error('[Polyfill] FAILURE: replaceState is NOT a function');
        }

    } catch (e) {
        console.error('[Polyfill] Fatal error', e);
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
