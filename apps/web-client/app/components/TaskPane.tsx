"use client";

import * as React from "react";
import { formatAndRedactDocument } from "../../lib/word-interaction";

/* global Word, Office */

/**
 * Main TaskPane component for the Redactor Add-in.
 * Handles the UI state, user interactions, and communication with the Office JavaScript API.
 * 
 * @returns {JSX.Element} The rendered TaskPane component.
 */
export default function TaskPane() {
  // State for Office API availability and UI feedback
  const [isOfficeInitialized, setIsOfficeInitialized] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);

  // Ref to track if the component is mounted to prevent state updates on unmount
  const isMounted = React.useRef(false);

  React.useEffect(() => {
    isMounted.current = true;

    // Defer Office initialization to ensure mount completion
    // This checks if the host is explicitly Word before enabling functionality
    const timer = setTimeout(() => {
      if (typeof Office !== 'undefined') {
        Office.onReady((info) => {
          if (isMounted.current && info.host === Office.HostType.Word) {
            setIsOfficeInitialized(true);
          }
        });
      }
    }, 0);

    return () => {
      isMounted.current = false;
      clearTimeout(timer);
    };
  }, []);

  /**
   * Wrapper function to initiate the document redaction process.
   * Manages the UI processing state and invokes the Word interaction logic.
   * 
   * @returns {Promise<void>} A promise that resolves when the operation is complete.
   */
  const handleRedactWrapper = async () => {
    console.log(">>> Scan & Redact button clicked");
    setStatus("Starting...");
    setIsProcessing(true);
    try {
      await Word.run(async (context) => {
        await formatAndRedactDocument(context, (msg) => {
          if (isMounted.current) setStatus(msg);
        });
      });
      if (isMounted.current) setStatus("Redaction Complete!");
    } catch (error: unknown) {
      console.error(error);
      const msg = error instanceof Error ? error.message : "Unknown error";
      if (isMounted.current) setStatus(`Error: ${msg}`);
    } finally {
      if (isMounted.current) setIsProcessing(false);
    }
  };

  return (
    <div className="taskpane-container">

      <header className="header">
        <div className="logo-container">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 2L3 7V17L12 22L21 17V7L12 2Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M12 22V12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M21 7L12 12L3 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
          <h1>Redactor</h1>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="content">
        <div className="hero-section">
          <h2>Secure Your Document</h2>
          <p className="description">
            Remove all types of sensitive information from your document, including PII, SPI, PHI, and more !
          </p>
        </div>

        <div className="action-card">
          <div className="status-area">
            {/* Progress Bar Animation */}
            {isProcessing && (
              <div className="progress-bar">
                <div className="progress-fill"></div>
              </div>
            )}
            {status && <p className="status-message">{status}</p>}
          </div>

          <button
            className={`primary-button ${isProcessing ? 'loading' : ''}`}
            onClick={handleRedactWrapper}
            disabled={isProcessing || (!isOfficeInitialized && typeof Office !== 'undefined')}
          >
            <span className="button-text">
              {isProcessing ? "Redacting Content..." : "Scan & Redact"}
            </span>
            {isProcessing && <span className="spinner"></span>}
          </button>
        </div>

      </main>

      <style jsx>{`
        .taskpane-container {
          display: flex;
          flex-direction: column;
          height: 100vh;
          background: var(--bg-color);
        }

        .header {
          padding: 24px;
          border-bottom: 1px solid var(--border-color);
          background: var(--bg-color);
          position: sticky;
          top: 0;
          z-index: 10;
        }

        .logo-container {
          display: flex;
          align-items: center;
          gap: 12px;
          color: var(--primary);
        }

        .logo-container h1 {
          margin: 0;
          font-size: 18px;
          font-weight: 600;
          color: var(--text-main);
          letter-spacing: -0.01em;
        }

        .content {
          padding: 24px;
          display: flex;
          flex-direction: column;
          gap: 32px;
          flex: 1;
          overflow-y: auto;
        }

        .hero-section h2 {
          font-size: 24px;
          font-weight: 700;
          margin-bottom: 8px;
          color: var(--text-main);
          line-height: 1.2;
        }

        .description {
          font-size: 14px;
          color: var(--text-secondary);
          line-height: 1.6;
        }

        .action-card {
          /* Removed white box styling */
          display: flex;
          flex-direction: column;
          gap: 20px;
        }

        .status-area {
          min-height: 48px;
          display: flex;
          flex-direction: column;
          gap: 8px;
        }

        .status-message {
          font-size: 13px;
          color: var(--primary);
          font-weight: 500;
          text-align: center;
        }

        .progress-bar {
          height: 4px;
          background: var(--border-color);
          border-radius: 2px;
          overflow: hidden;
        }

        .progress-fill {
          height: 100%;
          background: var(--primary);
          width: 30%;
          animation: progressMove 2s infinite ease-in-out;
        }

        @keyframes progressMove {
          0% { transform: translateX(-100%); width: 30%; }
          50% { width: 60%; }
          100% { transform: translateX(400%); width: 30%; }
        }

        .primary-button {
          position: relative;
          background: var(--primary);
          color: white;
          border: none;
          padding: 14px 24px;
          font-size: 15px;
          font-weight: 600;
          cursor: pointer;
          border-radius: var(--radius);
          transition: var(--transition);
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 12px;
          width: 100%;
          overflow: hidden;
        }

        .primary-button:hover:not(:disabled) {
          background: var(--primary-hover);
          transform: translateY(-1px);
          box-shadow: 0 4px 12px rgba(0, 120, 212, 0.25);
        }

        .primary-button:active:not(:disabled) {
          transform: translateY(0);
        }

        .primary-button:disabled {
          background: var(--border-color);
          color: var(--text-secondary);
          cursor: not-allowed;
        }

        .spinner {
          width: 18px;
          height: 18px;
          border: 2px solid rgba(255, 255, 255, 0.3);
          border-radius: 50%;
          border-top-color: white;
          animation: spin 0.8s linear infinite;
        }

        @keyframes spin {
          to { transform: rotate(360deg); }
        }


      `}</style>
    </div>
  );
}
