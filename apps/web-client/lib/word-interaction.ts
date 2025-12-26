/* global Word */

/**
 * Enables "Track Changes" for the document to ensure all redactions are recorded.
 * Uses the Word Requirement Set 1.5 to check for support before enabling.
 * 
 * @param context - The Word RequestContext.
 */
export async function enableTrackChanges(context: Word.RequestContext) {
    // Check if Tracking Changes is supported (Requirement Set 1.5)
    if (Office.context.requirements.isSetSupported('WordApi', '1.5')) {
        context.document.load("changeTrackingMode");
        await context.sync();

        // Only set if not already tracked to avoid unnecessary operations
        if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
        }
    } else {
        console.warn("Tracking Changes is not supported in this version of Word.");
    }
}

/**
 * Inserts a "CONFIDENTIAL DOCUMENT" header if one does not already exist.
 * This ensures the document is clearly marked after processing.
 * 
 * @param context - The Word RequestContext.
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

        // Check if header already exists to avoid duplication
        if (!range.text.includes("CONFIDENTIAL DOCUMENT")) {
            const p = header.insertParagraph("CONFIDENTIAL DOCUMENT", "Start");
            p.font.bold = true;
            p.font.color = "#D83B01"; // Professional dark orange/red
            p.font.size = 14;
            p.alignment = Word.Alignment.centered;
            await context.sync();
        }
    }
}

/**
 * Main function for document redaction.
 * 1. Enables Track Changes and Header.
 * 2. Scans the entire document body text.
 * 3. Sends text to the secure AI endpoint to identify sensitive data.
 * 4. Iteratively redacts found items while managing hyperlink removal.
 * 
 * @param context - The Word RequestContext.
 * @param onProgress - Callback function to update the UI with real-time status.
 */
export async function formatAndRedactDocument(context: Word.RequestContext, onProgress: (msg: string) => void) {
    // 1. Enable Track Changes & Header
    onProgress("Enabling Track Changes & Header...");
    await enableTrackChanges(context);
    await addConfidentialHeader(context);

    // 2. Scan Document Body
    onProgress("Reading document content...");
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const fullText = body.text.trim();
    if (!fullText) {
        onProgress("Document is empty.");
        return;
    }

    onProgress("Identifying sensitive information...");
    try {
        const response = await fetch('/api/redact', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: fullText })
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.error || response.statusText);
        }

        const data = await response.json();
        const piiList: string[] = data.pii || [];

        if (piiList.length > 0) {
            onProgress(`Redacting ${piiList.length} unique items...`);

            // Apply redaction to all occurrences in the document
            for (const pii of piiList) {
                if (!pii.trim()) continue;

                const searchResults = body.search(pii, { matchCase: false, ignorePunct: false });
                searchResults.load("items");
                await context.sync();

                for (let j = 0; j < searchResults.items.length; j++) {
                    const found = searchResults.items[j];

                    /**
                     * We need to remove the clickable link from the text without
                     * creating a confusing "deletion" entry in the history.
                     * Strategy:
                     * 1. Temporarily disable Track Changes.
                     * 2. Remove the hyperlink property (so Word doesn't keep the link
                     * and make the [REDACTED] card clickable).
                     * 3. Re-enable Track Changes.
                     * 4. Perform the visible text replacement (this is tracked).
                     */

                    // 1. Turn OFF Track Changes to remove hyperlink silently
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                    found.hyperlink = ""; // Strips the link property

                    // 2. Turn Track Changes BACK ON for the actual redaction
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

                    // 3. Apply Redaction (Tracked)
                    // Replaces the original text with [REDACTED] within a black box
                    const redacted = found.insertText("[REDACTED]", Word.InsertLocation.replace);
                    redacted.font.highlightColor = "#000000"; // Solid Black background
                    redacted.font.color = "#FFFFFF"; // White text for contrast
                    redacted.font.bold = false;
                }
            }
            await context.sync();
            onProgress(`Successfully redacted ${piiList.length} items!`);
        } else {
            onProgress("No sensitive information found.");
        }

    } catch (e: any) {
        console.error("Error processing document", e);
        throw e;
    }

    onProgress("Done!");
}
