import { GoogleGenerativeAI } from "@google/generative-ai";
import { NextResponse } from "next/server";

/**
 * POST Handler for the /api/redact endpoint.
 * Receives document text, sends it to Google's Gemini AI for PII analysis,
 * and returns a list of sensitive strings to be redacted.
 * 
 * @param req - The incoming HTTP request containing the JSON body with `text`.
 * @returns {NextResponse} JSON response containing the array of PII strings or an error message.
 */
export async function POST(req: Request) {
    console.log(">>> API called");
    try {
        const { text } = await req.json();

        // Use env var for key
        const apiKey = process.env.GEMINI_API_KEY;
        if (!apiKey) {
            console.error("GEMINI_API_KEY is not set");
            return NextResponse.json({ error: "GEMINI_API_KEY is not set" }, { status: 500 });
        }

        // Initialize Google Generative AI with the API key
        const genAI = new GoogleGenerativeAI(apiKey);

        // Use the gemini-2.0-flash model (v1beta) for optimal speed and reasoning
        const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" }, { apiVersion: "v1beta" });

        /**
         * AI SYSTEM PROMPT
         * This prompt is carefully engineered to:
         * 1. Define broad categories of sensitive data (PII, SPI, PHI, etc.).
         * 2. Demand a raw JSON array format for programmatic parsing.
         */
        const prompt = `Identify all sensitive information in the following text that should be redacted for confidentiality. This includes but is not limited to:
    
    1. PII (Personally Identifiable Information): Names in every context, physical addresses, email addresses, phone numbers, dates of birth, age, gender, race, origin, religion.
    2. SPI (Sensitive Personal Information): Biometric data, sexual orientation, criminal records, trade union membership.
    3. PHI (Protected Health Information): Medical records, health insurance info, treatments, diagnoses, prescriptions.
    4. Financial Data: Credit card numbers, bank account numbers, salary/compensation, tax IDs, credit scores.
    5. Government ID Numbers: SSNs, passport numbers, driver's licenses, employee IDs, national ID numbers.
    6. Business/Proprietary Data: Internal project names, proprietary code names, trade secrets, internal URLs/IPs.

    Return ONLY a JSON array of the exact strings found. If none, return an empty array.
    Do not include any explanation or markdown formatting, just the raw JSON array.
    
    Text to analyze:
    "${text}"`;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        let textResponse = response.text();

        // Clean up potential markdown formatting (e.g. ```json ... ```)
        textResponse = textResponse.replace(/^```json/i, '').replace(/```$/, '').trim();

        let pii: string[] = [];
        try {
            // Parse the cleaned text response into a JSON array
            pii = JSON.parse(textResponse);
        } catch {
            console.error("Failed to parse JSON", textResponse);
            return NextResponse.json({ error: "Failed to parse AI response", raw: textResponse }, { status: 500 });
        }

        return NextResponse.json({ pii });
    } catch (error: unknown) {
        console.error("AI Error:", error);
        const errorMessage = error instanceof Error ? error.message : "Unknown error";
        return NextResponse.json({ error: errorMessage }, { status: 500 });
    }
}
