"use client";

import * as React from "react";
import { formatAndRedactDocument } from "../../lib/word-interaction";

/**
 * TaskPane: The main interface for our Office Add-in.
 * 
 * This component manages the lifecycle of the add-in, from initializing
 * the Office environment to handling the redaction process and maintaining
 * a responsive layout that matches our "Operation Blackout" branding.
 */
export default function TaskPane() {
  // --- State & Refs ---
  const [isOfficeInitialized, setIsOfficeInitialized] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);
  const isMounted = React.useRef(false);

  // Measure the branding section to sync the button width dynamically
  const brandRef = React.useRef<HTMLDivElement>(null);
  const [buttonWidth, setButtonWidth] = React.useState<number | undefined>(undefined);

  // --- Initialization ---
  React.useEffect(() => {
    isMounted.current = true;

    // We wait to ensure Office.js is ready
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

  // Dynamic Layout Handling
  React.useEffect(() => {
    if (!brandRef.current) return;

    // Watch for size changes in the header to keep the UI aligned
    const updateWidth = () => {
      if (brandRef.current) {
        setButtonWidth(brandRef.current.offsetWidth);
      }
    };

    updateWidth();
    const observer = new ResizeObserver(updateWidth);
    observer.observe(brandRef.current);

    return () => observer.disconnect();
  }, []);

  // Redaction Logic
  const handleRedact = async () => {
    // Don't double-start or run without Word
    if (isProcessing || (!isOfficeInitialized && typeof Office !== 'undefined')) return;

    setStatus("Scanning for sensitive data...");
    setIsProcessing(true);

    try {
      await Word.run(async (context) => {
        // formatAndRedactDocument handles the heavy lifting & AI communication
        await formatAndRedactDocument(context, (msg) => {
          if (isMounted.current) setStatus(msg);
        });
      });

      if (isMounted.current) setStatus("Document secured!");
    } catch (error: any) {
      console.error("Redaction error:", error);
      if (isMounted.current) {
        setStatus(error.message || "Something went wrong. Please try again.");
      }
    } finally {
      if (isMounted.current) setIsProcessing(false);
    }
  };

  return (
    <div className="taskpane-container">
      {/* Branding Header */}
      <header className="header">
        <div className="brand-lockup" ref={brandRef}>
          <h1 className="brand-title">Operation Blackout</h1>
          <img
            src="/icon.png"
            alt="Logo"
            className="brand-icon"
            width="64"
            height="64"
          />
        </div>
      </header>

      {/* Main Action Area */}
      <main className="content">
        <div className="info-section" style={{ width: buttonWidth ? `${buttonWidth}px` : '100%' }}>
          <h2 className="info-title">Protect Your Data</h2>
          <p className="description">
            Confide your document by hiding its sensitive data and adding a confidential header.
            You can revert the redactions by rejecting changes in Word.
          </p>
        </div>

        <div className="action-container">
          <button
            className={`primary-button ${isProcessing ? 'loading' : ''}`}
            onClick={handleRedact}
            disabled={isProcessing || (!isOfficeInitialized && typeof Office !== 'undefined')}
            style={{ width: buttonWidth ? `${buttonWidth}px` : '100%' }}
          >
            <span className="button-text">
              {isProcessing ? "Anonymizing..." : "Anonymize Document"}
            </span>
            {isProcessing && <span className="spinner"></span>}
          </button>
        </div>

        {/* Footer / Feedback */}
        <div className="feedback-section">
          <p className="feedback-text">
            Your feedback is important to improve this Add-In
            <br />
            <a href="#" className="feedback-link" onClick={(e) => e.preventDefault()}>
              Feedback Survey
            </a>
          </p>
        </div>
      </main>

      <style jsx>{`
        /* Global Layout */
        .taskpane-container {
          display: flex;
          flex-direction: column;
          height: 100vh;
          background: var(--bg-color);
          overflow-x: hidden;
        }

        /* Branding Segment */
        .header {
          padding: 24px;
          background: var(--header-bg);
          display: flex;
          justify-content: center;
        }

        .brand-lockup {
          display: flex;
          align-items: center;
          gap: 12px;
          justify-content: center;
        }

        .brand-icon {
          border-radius: 12px;
          width: clamp(40px, 15vw, 64px);
          height: clamp(40px, 15vw, 64px);
          object-fit: contain;
        }

        .brand-title {
          margin: 0;
          font-size: clamp(20px, 8vw, 32px);
          font-weight: 800;
          color: var(--text-main);
          letter-spacing: -0.02em;
          white-space: nowrap;
        }

        /* Content Area */
        .content {
          padding: 32px 24px;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 28px;
          flex: 1;
        }

        .info-section {
          text-align: center;
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .info-title {
          font-size: clamp(20px, 6vw, 26px);
          font-weight: 700;
          color: white;
          margin: 0;
          line-height: 1.2;
          white-space: nowrap;
        }

        .description {
          font-size: clamp(12px, 4vw, 14px);
          color: white;
          line-height: 1.6;
          margin: 0;
          opacity: 0.9;
        }

        /* Interaction Elements */
        .action-container {
          width: 100%;
          display: flex;
          justify-content: center;
        }

        .primary-button {
          position: relative;
          background: var(--primary);
          color: white;
          border: none;
          padding: 16px 24px;
          font-size: clamp(17px, 5vw, 22px);
          font-weight: 700;
          cursor: pointer;
          border-radius: var(--radius);
          transition: var(--transition);
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 12px;
          width: 100%;
          overflow: hidden;
          white-space: nowrap;
          box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
        }

        .primary-button:hover:not(:disabled) {
          background: var(--primary-hover);
          transform: translateY(-2px);
          box-shadow: 0 6px 15px rgba(0, 120, 212, 0.3);
        }

        .primary-button:active:not(:disabled) {
          transform: translateY(0);
        }

        .primary-button:disabled {
          background: var(--border-color);
          color: var(--text-secondary);
          cursor: not-allowed;
          opacity: 0.7;
        }

        /* Feedback Section */
        .feedback-section {
          margin-top: auto;
          padding: 32px 0 16px;
          text-align: center;
          width: 100%;
        }

        .feedback-text {
          font-size: clamp(11px, 3.5vw, 13px);
          color: var(--text-secondary);
          margin: 0;
          line-height: 1.6;
        }

        .feedback-link {
          color: var(--primary);
          text-decoration: none;
          font-weight: 600;
          transition: var(--transition);
        }

        .feedback-link:hover {
          text-decoration: underline;
          color: var(--primary-hover);
        }

        /* Utilities */
        .spinner {
          width: 20px;
          height: 20px;
          border: 2.5px solid rgba(255, 255, 255, 0.3);
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
