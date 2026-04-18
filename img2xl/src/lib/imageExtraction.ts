import OpenAI from "openai";
import { cleanExtractedData } from "./dataCleaning";
import { jsonrepair } from "jsonrepair";

// GitHub token is injected via import.meta.env.VITE_GITHUB_TOKEN
const client = new OpenAI({
  baseURL: "https://models.github.ai/inference",
  apiKey: import.meta.env.VITE_GITHUB_TOKEN || "",
  dangerouslyAllowBrowser: true,
});

const MODEL = "gpt-4.1";

export interface ExtractionOptions {
  detectMultipleTables?: boolean;
  autoCleanData?: boolean;
}

export async function extractTableFromImage(
  base64Data: string,
  mimeType: string,
  options: ExtractionOptions = {}
) {
  let prompt = `Extract all tabular data from this ${mimeType.includes("pdf") ? "document" : "image"}.
Return a JSON object with:
- "headers": Array of strings (column headers).
- "rows": Array of arrays, where each inner array contains the row values in the same order as headers.

Rules:
1. Extract EVERY SINGLE ROW found in the document. Do not summarize. Target 100% data fidelity.
2. If headers are not explicitly labeled, infer logical headers or use "Column 1", "Column 2" etc.
3. Ensure proper JSON formatting. No markdown, no explanations.
4. This is a high-volume extraction task. Process as many rows as possible.`;

  if (options.detectMultipleTables) {
    prompt +=
      "\n5. If multiple distinct tables are present, merge them into one schema or sequential blocks.";
  }

  if (options.autoCleanData) {
    prompt +=
      "\n6. Automatically clean messy OCR data, normalize values, and handle noise.";
  }

  try {
    const response = await client.chat.completions.create({
      model: MODEL,
      max_tokens: 16384,
      temperature: 0.1,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "user",
          content: [
            { type: "text", text: prompt },
            {
              type: "image_url",
              image_url: {
                url: `data:${mimeType};base64,${base64Data}`,
              },
            },
          ],
        },
      ],
    });

    const text = response.choices[0]?.message?.content;
    if (!text) throw new Error("No response from AI model.");

    const jsonStr = text.replace(/```json/g, "").replace(/```/g, "").trim();

    let rawData: any;
    try {
      rawData = JSON.parse(jsonStr);
    } catch (e) {
      console.warn("Malformed JSON detected, attempting repair...", e);
      try {
        const repaired = jsonrepair(jsonStr);
        rawData = JSON.parse(repaired);
      } catch (repairError) {
        throw new Error(
          `Failed to parse AI response even after repair: ${
            repairError instanceof Error
              ? repairError.message
              : String(repairError)
          }`
        );
      }
    }

    let data: any[] = [];

    // Handle efficient format: { headers: [], rows: [[]] }
    if (
      rawData &&
      typeof rawData === "object" &&
      "headers" in rawData &&
      "rows" in rawData
    ) {
      const headers = rawData.headers as string[];
      const rows = rawData.rows as any[][];
      data = rows.map((row) => {
        const obj: any = {};
        headers.forEach((header, idx) => {
          obj[header || `Column ${idx + 1}`] = row[idx];
        });
        return obj;
      });
    }
    // Handle fallback format: [{ ... }]
    else if (Array.isArray(rawData)) {
      data = rawData;
      if (data.length > 0 && Array.isArray(data[0])) {
        data = data.map((row) => {
          const obj: any = {};
          row.forEach((val: any, idx: number) => {
            obj[`Column ${idx + 1}`] = val;
          });
          return obj;
        });
      }
    } else {
      throw new Error(
        "Extracted data is not in a recognized format (array or object with headers/rows)."
      );
    }

    if (options.autoCleanData) {
      data = cleanExtractedData(data);
    }

    return data;
  } catch (error) {
    console.error("GitHub Models Extraction Error:", error);
    throw error;
  }
}