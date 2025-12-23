
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("redact-btn")!.onclick = redact;
    }
});

/* global Word console */

export async function redact() {
    return Word.run(async (context) => {
        // 1. Enable Track Changes
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        // 2. Insert Confidential Header
        const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
        header.clear();
        const headerPara = header.insertParagraph("CONFIDENTIAL DOCUMENT", "Start");
        headerPara.font.bold = true;
        headerPara.font.color = "red";
        headerPara.alignment = Word.Alignment.centered;

        // 3. Perform Redaction
        const sensitivePatterns = [
            // Email: standard pattern
            /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g,
            // Phone: (123) 456-7890, 123-456-7890, etc.
            /(?:\+?1[-. ]?)?\(?([0-9]{3})\)?[-. ]?([0-9]{3})[-. ]?([0-9]{4})/g,
            // Credit Card: 4 groups of 4 digits
            /\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b/g,
            // SSN: xxx-xx-xxxx
            /\b\d{3}-\d{2}-\d{4}\b/g,
            // Dates (DOB context): MM/DD/YYYY
            /\b\d{1,2}\/\d{1,2}\/\d{4}\b/g,
            // Specific IDs from challenge doc
            /\b(?:EMP|MRN|INS)-[A-Z0-9-]+\b/g,
            // "Age" followed by digits (simple context)
            /\bAge\s+\d{2,3}\b/gi
        ];

        const body = context.document.body;
        // Optimization: search in ranges to avoiding splitting runs improperly, 
        // strictly speaking Word.run works best when we search then replace.
        // For simplicity and performance in this demo, we search the whole body.

        // We use split ranges to handle multiple occurrences
        // However, Word API search is safer than getting text and replacing it blindly
        // because getting text -> replacing invalidates the document structure.

        for (const pattern of sensitivePatterns) {
            // Convert Regex to string for Word search (limited regex support in Word search API)
            // Word's 'search' method accepts strings and wildcards, but not full JS Regex.
            // WORKAROUND: We iterate text ranges? 
            // Better: use 'wildcards' feature of Word Search where possible, OR
            // get extracted text, find matches, and then search for those specific matches.

            // Simpler approach for this task:
            // Since Word API search with regex is limited, we'll try to map JS regex to Word Wildcards 
            // OR use a different strategy: Read paragraph, if match, replace ranges.

            // Strategy: Iterate paragraphs to keep batch size reasonable
            const paragraphs = body.paragraphs;
            context.load(paragraphs, 'text');
            await context.sync();

            for (let i = 0; i < paragraphs.items.length; i++) {
                const para = paragraphs.items[i];
                let text = para.text;
                let match;

                // Reset lastIndex for the new string
                // Note: regex source needs 'g' flag which we have in the array
                // The pattern object is reused

                // We need to find matches in the original text
                const matches: { index: number, length: number }[] = [];
                // manual iteration for global regex
                while ((match = pattern.exec(text)) !== null) {
                    matches.push({ index: match.index, length: match[0].length });
                }

                // We process matches in reverse order to not mess up indices when replacing?
                // Actually, we should assume we replace strictly.
                // However, directly calling 'para.insertText' might be tricky if we want to replace *exact* range.
                // The most robust way in Word JS is using `searchResults`.
                // BUT we can't search with complex regex.

                // ALTERNATIVE: Use the specific text we found to search *uniquely* in the paragraph?
                // Risk: if the same text appears twice.

                // BETTER: Iterate all ranges.
                // Given the constraints and typical solutions for this in Word JS:
                // We'll search for the specific text found by regex.

                if (matches.length > 0) {
                    // For each unique match string found, we search it in the paragraph
                    // and replace. This is slightly inefficient but safer for Word API.
                    // To avoid replacing non-PII identical text (rare for emails/phones),
                    // we can try to be more specific.

                    // Actually, let's try a simpler approach for the challenge:
                    // Use Word's wildcards if possible? No, too complex to map.

                    // Let's implement the "Finding content via regex in JS, then search specific string" loop.
                    for (const m of matches) {
                        const matchedText = text.substring(m.index, m.index + m.length);
                        // Escape special chars for standard search
                        const searchResults = para.search(matchedText, { matchCase: false, ignorePunct: false });
                        context.load(searchResults);
                        await context.sync();

                        // This replaces ALL instances of that specific string in the paragraph.
                        // Usually safe for emails/phones. Riskier for "32" (Age).
                        // Mitigation: Check context if possible or accept risk for this demo.
                        for (let j = 0; j < searchResults.items.length; j++) {
                            searchResults.items[j].insertText("[REDACTED]", Word.InsertLocation.replace);
                            searchResults.items[j].font.highlightColor = "black";
                            searchResults.items[j].font.color = "white";
                        }
                    }
                }
            }
        }

        await context.sync();
    });
}
