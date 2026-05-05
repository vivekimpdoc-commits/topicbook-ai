import { GoogleGenAI, Type } from "@google/genai";
import { BookTopic } from "../types";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export async function generateTopicsFromText(filename: string, text: string): Promise<BookTopic[]> {
  const prompt = `You are an expert editorial synthesizer and archivist. Analyze the following text extracted from a document named "${filename}".

  Your MISSION is to perform a MAXIMUM depth extraction. Do not skip details. Do not generalize. Create a rich, comprehensive archive that captures 100% of the useful information in a sophisticated "book" format.
  
  LANGUAGE SUPPORT:
  - This document may contain Hindi, English, or a mix of both.
  - Your response (Title, Summary, Content) MUST respect the primary language of the document.
  - If the document is bilingual, use the most natural professional mix to maintain authenticity.
  
  CRITICAL GUIDELINES FOR HIGH-END ARCHIVING:
  - MASTER CHAPTERS: Organize into broad, meaningful "Master Chapters".
  - INTENSITY OF CONTENT: The "content" field for each chapter must be a DEEP synthesis. Aim for maximum verbosity and detail. Use professional, editorial language. We want long-form, high-quality prose that feels like a published book.
  - DO NOT SUMMARIZE ONLY: Elaborate, explain, and synthesize connections between various sections of the source text.
  - PROCESS FLOW: Extract every single methodology, step-by-step process, or workflow into the "processSteps" array. This is crucial for the conceptual flow visualization.
  - SENTIMENT: Precisely classify the tone as one of: 'analytical', 'creative', 'technical', 'narrative'.
  - TAGS: Generate at least 5-8 precise keywords per chapter.

  Format the output strictly as a JSON array.

  Text to process (Synthesize EVERYTHING):
  ${text.substring(0, 80000)}
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-1.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING },
              summary: { type: Type.STRING },
              content: { type: Type.STRING },
              sentiment: { type: Type.STRING, enum: ['analytical', 'creative', 'technical', 'narrative'] },
              wordCount: { type: Type.NUMBER },
              processSteps: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: { type: Type.STRING },
                    description: { type: Type.STRING }
                  },
                  required: ["title", "description"]
                }
              },
              tags: {
                type: Type.ARRAY,
                items: { type: Type.STRING }
              }
            },
            required: ["title", "summary", "content", "tags"]
          }
        }
      }
    });

    const topics = JSON.parse(response.text || "[]");
    return topics.map((t: any, index: number) => ({
      ...t,
      id: `${filename}-${index}`,
      sourceFile: filename
    }));
  } catch (error) {
    console.error("Gemini Error:", error);
    return [];
  }
}
