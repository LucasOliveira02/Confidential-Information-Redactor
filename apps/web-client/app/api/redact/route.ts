/*
 * This endpoint receives a block of text, forwards it to Google's Gemini
 * generative AI model, and returns a JSON array of detected sensitive
 * strings (PII, SPI, PHI, etc.). The response is deliberately minimal –
 * just the array – so the client can iterate over it and perform redaction.
 *
 * Expected request body:
 *   { "text": "...document contents..." }
 *
 * Successful response shape:
 *   { "pii": ["John Doe", "123-45-6789", ...] }
 *
 * Errors are returned with a 500 status and an `error` field describing the
 * problem. In development you will also see the raw AI response under the
 * `raw` key to aid debugging.
 */
import { GoogleGenerativeAI } from "@google/generative-ai";
import { NextResponse } from "next/server";

/**
 * POST handler for the /api/redact endpoint.
 *
 * 1. Validate the request payload and ensure the Gemini API key is set.
 * 2. Initialise the Generative AI client and configure the model.
 * 3. Build a system prompt that asks the model to list every piece of
 *    sensitive information in the supplied text.
 * 4. Parse the model's response, sanitising any markdown wrappers.
 * 5. Return the extracted array, or an error if parsing fails.
 */
export async function POST(req: Request) {
    console.log(">>> API called");
    try {
        const { text } = await req.json();

        // ---------------------------------------------------------------------
        // 1. Environment validation
        // ---------------------------------------------------------------------
        const apiKey = process.env.GEMINI_API_KEY;
        if (!apiKey) {
            console.error("GEMINI_API_KEY is not set");
            return NextResponse.json(
                { error: "GEMINI_API_KEY is not set" },
                { status: 500 }
            );
        }

        // ---------------------------------------------------------------------
        // 2. Initialise Gemini client
        // ---------------------------------------------------------------------
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel(
            {
                model: "gemini-2.0-flash",
                generationConfig: { temperature: 0 },
            },
            { apiVersion: "v1beta" }
        );

        // ---------------------------------------------------------------------
        // 3. AI Prompt
        // ---------------------------------------------------------------------
        const prompt = `Identify all sensitive information in the following text that should be redacted for confidentiality. Include but do not limit to:

1. PII (names, addresses, emails, phone numbers, DOB, gender, race, religion)
2. SPI (biometrics, sexual orientation, criminal records, union membership)
3. PHI (medical records, health insurance, diagnoses, prescriptions)
4. Financial data (credit cards, bank accounts, salaries, tax IDs)
5. Government IDs (SSN, passport, driver’s license, employee IDs)
6. Business secrets (project names, code names, internal URLs/IPs)

Return ONLY a JSON array of the exact strings found. If none are found, return an empty array.

Text to analyse:
"${text}"`;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        let rawText = response.text();

        // ---------------------------------------------------------------------
        // 4. Clean up the response – strip any markdown fences and extract []
        // ---------------------------------------------------------------------
        const startIdx = rawText.indexOf('[');
        const endIdx = rawText.lastIndexOf(']');
        const cleaned =
            startIdx !== -1 && endIdx !== -1
                ? rawText.substring(startIdx, endIdx + 1)
                : rawText.replace(/^```json/i, "").replace(/```$/, "").trim();

        let pii: string[] = [];
        try {
            pii = JSON.parse(cleaned);
            if (!Array.isArray(pii)) pii = [];
        } catch {
            console.error("Failed to parse JSON", rawText);
            return NextResponse.json(
                { error: "Failed to parse AI response. Expected a list.", raw: rawText },
                { status: 500 }
            );
        }

        // ---------------------------------------------------------------------
        // 5. Return the list of detected items
        // ---------------------------------------------------------------------
        return NextResponse.json({ pii });
    } catch (error: unknown) {
        console.error("AI Error:", error);
        const message = error instanceof Error ? error.message : "Unknown error";
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
