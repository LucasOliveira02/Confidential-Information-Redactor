/*
 * This is the main UI for the Office Add‑in. It orchestrates the
 * initialization of the Office environment, handles user interactions
 * (redact & revert), and displays a live status feed that reflects the
 * progress of the redaction workflow.
 *
 * The component is deliberately split into logical sections with clear
 * comments so future contributors can quickly understand the flow:
 *   1. State & refs – React state for Office readiness, processing flag,
 *      and a mutable ref to track component mount status.
 *   2. Effects – One effect waits for Office.js to become ready; a second
 *      watches the branding header size to keep button widths in sync.
 *   3. Handlers – `handleRedact` starts the redaction process, while
 *      `handleReject` rolls back all changes via the `rejectAllChanges`
 *      helper.
 *   4. Render – The UI consists of a branding header, two action blocks
 *      (protect & revert), a status‑feed, and a feedback footer.
 */
"use client";

import * as React from "react";
import { formatAndRedactDocument, rejectAllChanges } from "../../lib/word-interaction";

export default function TaskPane() {
  // ---------- State & Refs ----------
  const [isOfficeInitialized, setIsOfficeInitialized] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);
  const isMounted = React.useRef(false);

  // Reference to the branding header – used to sync button width.
  const brandRef = React.useRef<HTMLDivElement>(null);
  const [buttonWidth, setButtonWidth] = React.useState<number | undefined>(undefined);

  // ---------- Office Initialization Effect ----------
  // ---------- Office Initialization Effect ----------
  // We use a robust polling mechanism instead of a single check because `Office.js`
  // loads asynchronously and might not be defined immediately when the component mounts.
  React.useEffect(() => {
    isMounted.current = true;
    let attempts = 0;

    const checkOffice = () => {
      // If Office is ready, hook into `onReady`
      if (typeof Office !== "undefined") {
        Office.onReady((info) => {
          if (isMounted.current && info.host === Office.HostType.Word) {
            setIsOfficeInitialized(true);
          }
        });
        return true; // Found it!
      }
      return false; // Not yet...
    };

    // Attempt 1: Immediate check
    if (checkOffice()) return;

    // Attempt 2: Polling every 100ms for up to 5 seconds
    const intervalId = setInterval(() => {
      attempts++;
      if (checkOffice() || attempts > 50) {
        clearInterval(intervalId);
      }
    }, 100);

    return () => {
      isMounted.current = false;
      clearInterval(intervalId);
    };
  }, []);

  // ---------- Dynamic Layout Effect (Button Width) ----------
  React.useEffect(() => {
    if (!brandRef.current) return;
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

  // ---------- Redaction Handler ----------
  const handleRedact = async () => {
    if (isProcessing || (!isOfficeInitialized && typeof Office !== "undefined")) return;
    setStatus("Scanning for sensitive data...");
    setIsProcessing(true);
    try {
      await Word.run(async (context) => {
        await formatAndRedactDocument(context, (msg) => {
          if (isMounted.current) setStatus(msg);
        });
      });
      if (isMounted.current) setStatus("Document secured!");
    } catch (error: any) {
      console.error("Redaction error:", error);
      if (isMounted.current) {
        setStatus(error.message || "Something went wrong, please try again.");
      }
    } finally {
      if (isMounted.current) setIsProcessing(false);
    }
  };

  // ---------- Revert Handler ----------
  const handleReject = async () => {
    if (isProcessing || (!isOfficeInitialized && typeof Office !== "undefined")) return;
    setStatus("Reverting changes...");
    setIsProcessing(true);
    try {
      await Word.run(async (context) => {
        await rejectAllChanges(context);
      });
      if (isMounted.current) setStatus("Changes rejected, document restored.");
    } catch (error: any) {
      console.error("Reject error:", error);
      if (isMounted.current) {
        setStatus(error.message || "Failed to reject changes, please do it manually.");
      }
    } finally {
      if (isMounted.current) setIsProcessing(false);
    }
  };

  // ---------- Render ----------
  return (
    <div className="taskpane-container">
      {/* Branding Header */}
      <header className="header">
        <div className="brand-lockup" ref={brandRef}>
          <h1 className="brand-title">Operation Blackout</h1>
          <img src="/icon.png" alt="Logo" className="brand-icon" width="64" height="64" />
        </div>
      </header>

      {/* Main Action Area */}
      <main className="content">
        {/* Protect Section */}
        <div className="action-block">
          <div className="info-section" style={{ width: buttonWidth ? `${buttonWidth}px` : "100%" }}>
            <h2 className="info-title">Protect Your Data</h2>
            <p className="description">
              Confide your document by hiding its sensitive data and adding a confidential header.
            </p>
          </div>
          <button
            className={`primary-button ${isProcessing ? "loading" : ""}`}
            onClick={handleRedact}
            disabled={isProcessing || (!isOfficeInitialized && typeof Office !== "undefined")}
            style={{ width: buttonWidth ? `${buttonWidth}px` : "100%" }}
          >
            <span className="button-text">
              {isProcessing && !status.includes("Reverting") ? "Anonymizing..." : "Anonymize Document"}
            </span>
            {isProcessing && !status.includes("Reverting") && <span className="spinner"></span>}
          </button>
        </div>

        {/* Status Feed */}
        <div className={`status-feed ${status ? "visible" : ""}`} style={{ width: buttonWidth ? `${buttonWidth}px` : "100%" }}>
          {status && <span className="status-text">{status.toUpperCase()}</span>}
        </div>

        {/* Revert Section */}
        <div className="action-block">
          <div className="info-section" style={{ width: buttonWidth ? `${buttonWidth}px` : "100%" }}>
            <h2 className="info-title">Revert Changes</h2>
            <p className="description">
              You can revert specific redactions by rejecting changes in Word or revert all redactions by clicking the button.
            </p>
          </div>
          <button
            className={`secondary-button ${isProcessing ? "loading" : ""}`}
            onClick={handleReject}
            disabled={isProcessing || (!isOfficeInitialized && typeof Office !== "undefined")}
            style={{ width: buttonWidth ? `${buttonWidth}px` : "100%" }}
          >
            <span className="button-text">
              {isProcessing && status.includes("Reverting") ? "Reverting..." : "Revert All Changes"}
            </span>
            {isProcessing && status.includes("Reverting") && <span className="spinner"></span>}
          </button>
        </div>

        {/* Footer / Feedback */}
        <div className="feedback-section">
          <p className="feedback-text">
            Your feedback is important to improve this Add‑In<br />
            <a href="#" className="feedback-link" onClick={(e) => e.preventDefault()}>
              Please fill out the feedback survey
            </a>
          </p>
        </div>
      </main>

      {/* Scoped Styles – kept inline for simplicity */}
      <style jsx>{`
        .taskpane-container { display: flex; flex-direction: column; height: 100vh; background: var(--bg-color); overflow-x: hidden; }
        .header { padding: 12px 24px; background: var(--header-bg); display: flex; justify-content: center; }
        .brand-lockup { display: flex; align-items: center; gap: 12px; justify-content: center; }
        .brand-icon { border-radius: 12px; width: clamp(40px, 15vw, 64px); height: clamp(40px, 15vw, 64px); object-fit: contain; }
        .brand-title { margin: 0; font-size: clamp(20px, 8vw, 32px); font-weight: 800; color: var(--text-main); letter-spacing: -0.02em; white-space: nowrap; }
        .content { padding: 12px 24px 32px; display: flex; flex-direction: column; align-items: center; gap: 20px; flex: 1; }
        .info-section { text-align: center; display: flex; flex-direction: column; align-items: center; gap: 12px; }
        .info-title { font-size: clamp(20px, 6vw, 26px); font-weight: 700; color: white; margin: 0; line-height: 1.2; white-space: nowrap; }
        .description { font-size: clamp(12px, 4vw, 14px); color: white; line-height: 1.6; margin: 0; opacity: 0.9; }
        .action-block { width: 100%; display: flex; flex-direction: column; align-items: center; gap: 16px; justify-content: center; }
        .primary-button, .secondary-button { position: relative; padding: 16px 24px; font-size: clamp(17px, 5vw, 22px); font-weight: 700; border: none; border-radius: var(--radius); transition: var(--transition); display: flex; align-items: center; justify-content: center; gap: 12px; width: 100%; overflow: hidden; white-space: nowrap; box-shadow: 0 4px 6px rgba(0,0,0,0.2); }
        .primary-button { background: var(--primary); color: white; }
        .primary-button:hover:not(:disabled) { background: var(--primary-hover); transform: translateY(-2px); box-shadow: 0 6px 15px rgba(0,120,212,0.3); }
        .secondary-button { background: #ff6b6b; color: white; }
        .secondary-button:hover:not(:disabled) { background: #ff5252; transform: translateY(-2px); box-shadow: 0 6px 15px rgba(255,107,107,0.3); }
        .primary-button:disabled, .secondary-button:disabled { background: var(--border-color); color: var(--text-secondary); cursor: not-allowed; opacity: 0.7; }
        .primary-button.loading, .secondary-button.loading { opacity: 0.8; cursor: wait; pointer-events: none; animation: pulse 1.5s ease-in-out infinite; }
        @keyframes pulse { 0%,100% { transform: scale(1); } 50% { transform: scale(0.98); opacity: 0.7; } }
        .status-feed { height: 24px; display: flex; align-items: center; justify-content: center; opacity: 0; transition: var(--transition); margin: -8px 0; }
        .status-feed.visible { opacity: 1; }
        .status-text { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; font-size: 10px; font-weight: 600; color: var(--text-secondary); letter-spacing: 0.15em; text-align: center; animation: status-blink 2s ease-in-out infinite; }
        @keyframes status-blink { 0%,100% { opacity: 1; } 50% { opacity: 0.4; } }
        .feedback-section { margin-top: auto; padding: 32px 0 16px; }
        .feedback-text { color: var(--text-secondary); text-align: center; font-size: clamp(11px, 3.5vw, 13px); }
        .feedback-link { color: var(--primary); text-decoration: underline; cursor: pointer; }
      `}</style>
    </div>
  );
}
