/* global Word, Office */

/**
 * Enable Track Changes in the current Word document.
 *
 * Track Changes lets the user review every modification we make –
 * they can accept or reject each redaction later. This is safer than
 * silently deleting text because it leaves a clear audit trail.
 */
export async function enableTrackChanges(context: Word.RequestContext) {
    // Word API 1.5 introduced the `changeTrackingMode` property.
    if (Office.context.requirements.isSetSupported('WordApi', '1.5')) {
        // Load the current tracking mode so we can inspect it.
        context.document.load('changeTrackingMode');
        await context.sync();

        // Turn tracking on only if it isn't already active.
        if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
        }
    } else {
        console.warn("Tracking Changes isn't supported in this Word environment.");
    }
}

/**
 * Insert a bold, orange‑red header that marks the document as confidential.
 *
 * The header is added to the first section only. If the header already
 * exists we skip insertion – this makes the function idempotent.
 */
export async function addConfidentialHeader(context: Word.RequestContext) {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();

    if (sections.items.length > 0) {
        try {
            const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);
            // Use search to check if the header already exists – this avoids
            // loading the entire header text which can fail in some contexts.
            const searchResults = header.search('CONFIDENTIAL DOCUMENT', { matchCase: true });
            searchResults.load('items');
            await context.sync();

            if (searchResults.items.length === 0) {
                const paragraph = header.insertParagraph('CONFIDENTIAL DOCUMENT', 'Start');
                paragraph.font.bold = true;
                paragraph.font.color = '#D83B01'; // orange‑red for high visibility
                paragraph.font.size = 14;
                paragraph.alignment = Word.Alignment.centered;
                await context.sync();
            }
        } catch (error) {
            console.warn("Could not add confidential header:", error);
            // Suppress error so redaction (the main task) can proceed
        }
    }
}

/**
 * Core redaction workflow.
 *
 *  1. Prepare the document (enable tracking + add header).
 *  2. Pull the full text and ask the AI service to identify PII.
 *  3. For each piece of sensitive data, replace it with a redacted token.
 *    While we process each item we update the UI via `onProgress` so the
 *    user sees a live "mission log".
 */
export async function formatAndRedactDocument(
    context: Word.RequestContext,
    onProgress: (msg: string) => void
) {
    // Step 1: Preparation
    onProgress('Setting up document tracking...');
    await enableTrackChanges(context);
    await addConfidentialHeader(context);

    // Step 2: Load document text
    onProgress('Reading document content...');
    const body = context.document.body;
    body.load('text');
    await context.sync();

    const fullText = body.text.trim();
    if (!fullText) {
        onProgress('The document appears to be empty.');
        return;
    }

    // Step 3: Ask AI for PII
    onProgress('AI is identifying sensitive data...');
    try {
        const response = await fetch('/api/redact', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: fullText })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error || 'AI Service Unavailable');
        }

        const { pii = [] } = await response.json();
        if (pii.length === 0) {
            onProgress("No sensitive data found. You're clear!");
            return;
        }

        // Step 4: Redact each item
        let count = 0;
        for (const item of pii) {
            if (!item.trim()) continue;
            count++;
            // Show a concise progress update (ellipsis matches other messages).
            onProgress(`Protecting data (${count}/${pii.length})...`);

            const searchResults = body.search(item, { matchCase: false, ignorePunct: false });
            searchResults.load('items');
            await context.sync();

            for (const found of searchResults.items) {
                // Hyperlink handling
                // Turn tracking off temporarily so the hyperlink removal isn't logged.
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                found.hyperlink = '';
                // Turn tracking back on for the visible redaction.
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

                // Replace the found text with a blacked‑out token.
                const redacted = found.insertText('[REDACTED]', Word.InsertLocation.replace);
                redacted.font.highlightColor = '#000000'; // black background
                redacted.font.color = '#FFFFFF'; // white text
                redacted.font.bold = false;
            }
        }

        await context.sync();
        onProgress(`Mission accomplished! ${pii.length} items secured.`);
    } catch (err: any) {
        console.error('Redaction workflow failed:', err);
        throw err;
    }

    // Final status when everything is done.
    onProgress('Done!');
}

/**
 * Undo all redactions and remove the confidential header.
 *
 * This function is useful if the user decides the automated changes
 * were too aggressive. It simply rejects every tracked change in the
 * body and the header, restoring the original document state.
 */
export async function rejectAllChanges(context: Word.RequestContext) {
    // Word API 1.6 introduced document‑level rejection of tracked changes.
    if (Office.context.requirements.isSetSupported('WordApi', '1.6')) {
        // ---- Clean the body ----
        const bodyChanges = context.document.body.getTrackedChanges();
        bodyChanges.rejectAll();

        // ---- Clean the header ----
        const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
        header.getTrackedChanges().rejectAll();

        await context.sync();
    } else {
        // Fallback for older Word versions – the user must revert manually.
        throw new Error(
            "Your version of Word doesn't support automatic rejection. " +
            "You can still reject changes manually from the 'Review' tab."
        );
    }
}
