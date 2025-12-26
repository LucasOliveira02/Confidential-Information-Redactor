/* global Word, Office */

/**
 * enableTrackChanges: Ensures "Track Changes" is active.
 * 
 * We use Track Changes so the user can review and either accept or 
 * reject the redactions we found. This is safer than just deleting 
 * text permanently without a trail.
 */
export async function enableTrackChanges(context: Word.RequestContext) {
    // We target Requirement Set 1.5 as it contains the changeTrackingMode API
    if (Office.context.requirements.isSetSupported('WordApi', '1.5')) {
        context.document.load("changeTrackingMode");
        await context.sync();

        // If it's already on, we don't need to do anything
        if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
        }
    } else {
        console.warn("Tracking Changes isn't supported in this Word environment.");
    }
}

/**
 * addConfidentialHeader: Marks the document as confidential.
 * 
 * We insert a clear, bold header so there's no ambiguity about the 
 * document's status after processing.
 */
export async function addConfidentialHeader(context: Word.RequestContext) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    if (sections.items.length > 0) {
        const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);
        const range = header.getRange();
        range.load("text");
        await context.sync();

        // Avoid double-marking if the user runs the tool multiple times
        if (!range.text.includes("CONFIDENTIAL DOCUMENT")) {
            const p = header.insertParagraph("CONFIDENTIAL DOCUMENT", "Start");
            p.font.bold = true;
            p.font.color = "#D83B01"; // A respectful, standout orange-red
            p.font.size = 14;
            p.alignment = Word.Alignment.centered;
            await context.sync();
        }
    }
}

/**
 * formatAndRedactDocument: The "brains" of the operation.
 * 
 * 1. Prepares the document (Track Changes + Header).
 * 2. Fetches the text and asks our AI (Gemini) to find sensitive info.
 * 3. Iterates through the document to replace found items with redact cards.
 */
export async function formatAndRedactDocument(context: Word.RequestContext, onProgress: (msg: string) => void) {

    // --- Step 1: Prep ---
    onProgress("Setting up document tracking...");
    await enableTrackChanges(context);
    await addConfidentialHeader(context);

    // --- Step 2: Read ---
    onProgress("Reading document content...");
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const fullText = body.text.trim();
    if (!fullText) {
        onProgress("The document appears to be empty.");
        return;
    }

    // --- Step 3: Analyze (AI) ---
    onProgress("AI is identifying sensitive data...");
    try {
        const response = await fetch('/api/redact', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: fullText })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error || "AI Service Unavailable");
        }

        const { pii = [] } = await response.json();

        if (pii.length === 0) {
            onProgress("No sensitive data found. You're clear!");
            return;
        }

        // --- Step 4: Redact ---
        onProgress(`Protecting your data (${pii.length} unique types found)...`);

        for (const item of pii) {
            if (!item.trim()) continue;

            const searchResults = body.search(item, { matchCase: false, ignorePunct: false });
            searchResults.load("items");
            await context.sync();

            for (const found of searchResults.items) {

                // --- Hyperlink Strategy ---
                // We want redactions to be "plain". If an item is a link, 
                // we strip it silently first (untracked) so the black box
                // isn't clickable.
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                found.hyperlink = "";

                // Switch tracking BACK ON for the visible redaction
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

                // Actual replacement with the "Blackout" look
                const redacted = found.insertText("[REDACTED]", Word.InsertLocation.replace);
                redacted.font.highlightColor = "#000000";
                redacted.font.color = "#FFFFFF";
                redacted.font.bold = false;
            }
        }

        await context.sync();
        onProgress(`Finished! Redacted ${pii.length} sensitive items.`);

    } catch (err: any) {
        console.error("Redaction workflow failed:", err);
        throw err;
    }

    onProgress("Done!");
}
