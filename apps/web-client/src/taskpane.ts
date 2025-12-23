/**
 * This function initializes the Office Add-in for Word. It waits for the
 * Office JS library to initialize, then checks if the host is Word. If so,
 * it sets up the event handlers.
 * 
 * @param info - Information object provided by Office.onReady
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("redact-btn")!.onclick = redact;
    }
});

/**
 * This function implements the redaction process for the Word document.
 * It is triggered when the "Redact Document" button is clicked and executes
 * the operations that redact sensitive information from the document.
 * 
 * Functionality:
 * - Enabling "Track Changes"
 * - Inserting "CONFIDENTIAL DOCUMENT" header
 * - DESCRIBE REDACTING LOGIC HERE
 * - Disabling "Track Changes"
 * 
 * @returns COMPLETE RETURN STATEMENT HERE
 */
export async function redact() {
    return Word.run(async (context) => {
        // TODO: Implement actual redaction logic
        console.log("Redaction started");

        // Temporary feedback
        const body = context.document.body;
        body.insertParagraph("Redaction process initiated...", Word.InsertLocation.end);

        await context.sync();
    });
}
