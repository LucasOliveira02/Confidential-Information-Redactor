
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("redact-btn")!.onclick = redact;
    }
});

/* global Word console */

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
