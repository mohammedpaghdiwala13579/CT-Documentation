import { executeGeminiChat } from "../_lib/gemini.js";

export default async function handler(req: any, res: any) {
  // Enable CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const { messages, model, role, documentContext } = body;

    if (!Array.isArray(messages) || messages.length === 0) {
      return res.status(400).json({ error: "Messages array is required and cannot be empty." });
    }

    const result = await executeGeminiChat({
      messages,
      model,
      role,
      documentContext,
    });

    return res.status(200).json(result);
  } catch (error: any) {
    console.error("Error in Vercel api/gemini/chat:", error);
    return res.status(500).json({
      error: error?.message || "Failed to process chat with Gemini AI.",
    });
  }
}
