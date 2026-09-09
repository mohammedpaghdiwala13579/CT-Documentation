import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

app.use(express.json({ limit: "10mb" }));

// Initialize Google GenAI client
let aiClient: GoogleGenAI | null = null;
function getGenAI(): GoogleGenAI {
  if (!aiClient) {
    const apiKey = process.env.GEMINI_API_KEY || "";
    aiClient = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiClient;
}

// Role-based System Instructions for Maritime Ship Chandler Operations in Bangladesh
const ROLE_SYSTEM_INSTRUCTIONS: Record<string, string> = {
  maritime_analyst: `You are the Senior Maritime Operations & Ship Chandler Specialist for Comilla Traders, a government-licensed ship chandler and marine contractor operating across all major ports of Bangladesh:
- Chittagong (Chattogram) Port (CPA jetties, Karnaphuli river berths, and Outer Anchorage Alpha/Bravo/Charlie, Kutubdia lighterage).
- Mongla Port (Pashur river & Harbaria anchorage).
- Matarbari Deep Sea Port & Payra Port.

You have comprehensive knowledge of the BANGLADESHI MARITIME & LOCAL SUPPLY MARKET:
1. SOURCING HUBS & CURRENT BANGLADESH PRICING:
   - Fresh & Dry Provisions (Khatunganj & Chaktai wholesale commodity market, Chittagong):
     * Fresh Halal Beef: ~750-850 BDT/kg (~$6.20-$7.00 USD/kg)
     * Fresh Mutton: ~1,100-1,250 BDT/kg (~$9.00-$10.20 USD/kg)
     * Fresh Broiler Chicken: ~180-220 BDT/kg (~$1.50-$1.80 USD/kg)
     * Farm Fresh Eggs: ~130-150 BDT/dozen (~$1.10-$1.25 USD/dozen)
     * Premium Basmati Rice: ~140-180 BDT/kg (~$1.15-$1.50 USD/kg); Miniket: ~75-85 BDT/kg
     * Seasonal Vegetables (Potatoes, Onions, Cabbage, Tomatoes): ~45-65 BDT/kg (~$0.38-$0.55 USD/kg)
     * Bottled Mineral Water (1.5L x 12 cases): ~220-250 BDT/case (~$1.85-$2.10 USD/case)
     * Bulk Fresh Drinking Water via Supply Barge to Outer Anchorage: ~$20-$28 USD/metric ton
   - Engine & Deck Marine Stores (Sadarghat, Strand Road & Agrabad, Chittagong):
     * 220m 24mm/28mm 8-Strand Polypropylene Mooring Rope: ~$480-$680 USD/coil
     * Galvanized Wire Ropes (6x36 WS IWRC): ~$3.00-$4.50 USD/meter
     * SOLAS Approved Adult Lifejackets with light & whistle: ~$24-$32 USD/pc
     * SOLAS Immersion Suits (MED certified): ~$98-$130 USD/pc
     * Zinc Anodes for hull & ballast tanks: ~$8.00-$9.50 USD/kg
     * Industrial Cotton Cleaning Rags/Waste: ~65-80 BDT/kg (~$0.55-$0.68 USD/kg)
   - Ship Breaking Yard OEM Machinery & Spares (Bhatiary & Sitakunda, Chittagong):
     * Genuine reconditioned marine valves (JIS / DIN bronze globe, gate, storm valves), pump impellers, purifiers (Alfa Laval/Mitsubishi), diesel engine spares (Daihatsu, Yanmar, MAN B&W) available at 40-70% savings compared to new imports.

2. SPEED & ACCURACY DIRECTIVES:
   - Provide answers FAST, accurately, and straight to the point.
   - Always quote current estimated prices in BOTH Bangladeshi Taka (BDT) and US Dollars (USD) (using current exchange benchmark: 1 USD ≈ 120-123 BDT).
   - Format item lists clearly in markdown tables with: Item Description, IMPA/ISSA Code (if applicable), Standard Unit, and Estimated Unit Price in BDT & USD.
   - Specify whether the pricing is for Port Berth delivery or Outer Anchorage delivery (including supply boat / launch lighterage costs: ~$280-$450 USD per launch trip from Ghat 15).`,

  pricing_negotiator: `You are the Commercial Director & Senior Pricing Strategist for Comilla Traders in Bangladesh.
Your expertise covers:
- Current Bangladeshi ship supply commercial margins (typically 15%-25% on provisions, 20%-35% on technical deck/engine stores, 30%-50% on reconditioned Sitakunda shipyard spares).
- Outer Anchorage lighterage launch boat hire tariffs (~$280-$450 USD/trip from Sadarghat/Ghat 15 to Chittagong Outer Anchorage depending on sea state and waiting time).
- Multi-currency conversions (USD, BDT, EUR, SGD, AED) with proper exchange rate volatility buffers (1 USD ≈ 120-123 BDT).
- Payment terms in shipping: Cash on Delivery (COD), Cash Against Documents (CAD), 30-day DA with foreign owners, and Master's General Receipt.
- National Board of Revenue (NBR) Bangladesh VAT guidelines: supplies to foreign-flagged vessels under bonded customs supply enjoy export zero-rated VAT status.
Give fast, sharp, numbers-driven advice that maximizes profit margin while offering competitive quotes to ship owners and vessel managers.`,

  vessel_auditor: `You are the Quality Control & Maritime Compliance Auditor for Comilla Traders, Chittagong, Bangladesh.
Your duties are:
- Auditing quotations, delivery challans, and commercial invoices against Chittagong Port Authority (CPA) and international shipping standards.
- Verifying complete vessel details (Vessel Name, Call Sign, IMO Number, Port/Berth, Requisition No, and Purchase Order).
- Ensuring quantities, standard marine units (PCS, KGS, COIL, MTR, SET, LTR, DRUM), unit prices, and extended amounts calculate with zero mathematical discrepancies.
- Validating delivery documentation for Master / Chief Engineer / Chief Officer sign-off and ship stamp.
Highlight missing vessel info, unit mismatches, or pricing errors in a clean, high-speed checklist format.`,
};

// Health Check
app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", service: "Comilla Traders Enterprise API" });
});

// Multi-Turn Chat Endpoint
app.post("/api/gemini/chat", async (req, res) => {
  try {
    const { messages, model, role, documentContext } = req.body;

    if (!Array.isArray(messages) || messages.length === 0) {
      return res.status(400).json({ error: "Messages array is required and cannot be empty." });
    }

    // Supported models: gemini-3.8-flash (fast, accurate, latest default), gemini-3.1-flash-lite (ultra fast), gemini-3.1-pro-preview (complex)
    const validModels = ["gemini-3.8-flash", "gemini-3.1-flash-lite", "gemini-3.1-pro-preview", "gemini-3.5-flash"];
    const chosenModel = validModels.includes(model) ? model : "gemini-3.8-flash";

    const baseInstruction = ROLE_SYSTEM_INSTRUCTIONS[role] || ROLE_SYSTEM_INSTRUCTIONS.maritime_analyst;
    let fullSystemInstruction = baseInstruction;

    if (documentContext && typeof documentContext === "object") {
      fullSystemInstruction += `\n\n[CURRENT ACTIVE DOCUMENT CONTEXT IN APP]:
Document Type: ${documentContext.docType || "Quotation"}
Client/Messers: ${documentContext.messers || "Not specified"}
Vessel Name: ${documentContext.vesselName || "Not specified"}
Address / Port: ${documentContext.address || "Chittagong Port, Bangladesh"}
Currency: ${documentContext.currency || "USD"}
Total Items: ${documentContext.itemCount ?? 0}
Total Amount: ${documentContext.currency || "USD"} ${documentContext.totalAmount ?? 0}
Requisition No: ${documentContext.requisitionNo || "None"}
Challan/Invoice No: ${documentContext.docNumber || "None"}
Use this context to provide accurate, context-aware pricing and advice.`;
    }

    // Convert client conversation history to Gemini contents format
    const contents = messages.map((m: { role: string; content: string }) => ({
      role: m.role === "user" ? "user" : "model",
      parts: [{ text: m.content || "" }],
    }));

    const ai = getGenAI();

    // Candidate model priority list for high reliability & zero downtime
    const candidateModels = [chosenModel, "gemini-3.1-flash-lite", "gemini-3.5-flash"].filter(
      (m, idx, arr) => arr.indexOf(m) === idx
    );

    let response: any = null;
    let successfulModel = chosenModel;

    for (const m of candidateModels) {
      try {
        // Try with real-time Google Search grounding for live Bangladeshi market data
        response = await ai.models.generateContent({
          model: m,
          contents,
          config: {
            systemInstruction: fullSystemInstruction,
            temperature: 0.3,
            tools: [{ googleSearch: {} }],
          },
        });
        successfulModel = m;
        break;
      } catch (searchError: any) {
        // Retry without tools if search grounding was the cause
        try {
          response = await ai.models.generateContent({
            model: m,
            contents,
            config: {
              systemInstruction: fullSystemInstruction,
              temperature: 0.3,
            },
          });
          successfulModel = m;
          break;
        } catch (modelError: any) {
          console.warn(`Model ${m} unavailable (${modelError?.message}), attempting next candidate...`);
        }
      }
    }

    if (!response) {
      throw new Error("Gemini AI models are temporarily busy. Please retry in a moment.");
    }

    const replyText = response.text || "";
    return res.json({
      reply: replyText,
      modelUsed: successfulModel,
      roleUsed: role || "maritime_analyst",
    });
  } catch (error: any) {
    console.error("Error in /api/gemini/chat:", error);
    return res.status(500).json({
      error: error?.message || "Failed to process chat with Gemini AI.",
    });
  }
});

// Smart RFQ / Unstructured Text Item Extractor
app.post("/api/gemini/extract-items", async (req, res) => {
  try {
    const { text, model } = req.body;

    if (!text || typeof text !== "string" || text.trim() === "") {
      return res.status(400).json({ error: "Text content is required for item extraction." });
    }

    const chosenModel = model === "gemini-3.1-flash-lite" ? "gemini-3.1-flash-lite" : "gemini-3.8-flash";
    const candidateModels = [chosenModel, "gemini-3.1-flash-lite", "gemini-3.5-flash"].filter(
      (m, idx, arr) => arr.indexOf(m) === idx
    );

    const ai = getGenAI();
    let response: any = null;

    for (const m of candidateModels) {
      try {
        response = await ai.models.generateContent({
          model: m,
          contents: `Extract all maritime supply items, equipment, provisions, and spare parts from this request or text into structured items:
\n\n"""\n${text}\n"""`,
          config: {
            systemInstruction: `You are an expert maritime item parser for ship chandlers operating at Chittagong & Mongla ports in Bangladesh. 
Analyze the input text (which might be an email, RFQ, WhatsApp list, or captain's requisition).
Extract each distinct line item with its description (including sizes, dimensions, or IMPA/ISSA code if present), quantity (as a string or number), standard unit (e.g. PCS, KGS, SET, MTR, LTR, BDL, PKT, BOX, PAIR, ROLL, COIL, DRUM), and price (if mentioned or suggested for Bangladeshi market, otherwise empty string "").
Return an array of JSON objects matching the schema with high speed and zero extra narrative.`,
            responseMimeType: "application/json",
            responseSchema: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  desc: { type: Type.STRING, description: "Detailed item name, spec, or IMPA code" },
                  qty: { type: Type.STRING, description: "Quantity requested" },
                  unit: { type: Type.STRING, description: "Standard maritime unit of measurement" },
                  price: { type: Type.STRING, description: "Unit price if given or suggested, else empty string" },
                },
                required: ["desc", "qty", "unit"],
              },
            },
          },
        });
        if (response) break;
      } catch (err: any) {
        console.warn(`Extraction failed on model ${m} (${err?.message}), trying fallback...`);
      }
    }

    if (!response) {
      throw new Error("Unable to extract items at this moment.");
    }

    let items = [];
    try {
      items = JSON.parse(response.text || "[]");
    } catch {
      items = [];
    }

    return res.json({ items });
  } catch (error: any) {
    console.error("Error in /api/gemini/extract-items:", error);
    return res.status(500).json({
      error: error?.message || "Failed to extract items with Gemini AI.",
    });
  }
});

// Document Auditor Endpoint
app.post("/api/gemini/audit-document", async (req, res) => {
  try {
    const { document, model } = req.body;

    if (!document || !Array.isArray(document.rows)) {
      return res.status(400).json({ error: "Valid document object with rows is required." });
    }

    const chosenModel = model === "gemini-3.1-pro-preview" ? "gemini-3.1-pro-preview" : "gemini-3.5-flash";
    const ai = getGenAI();

    const prompt = `Conduct a rigorous quality, commercial, and maritime compliance audit on this ship chandler document:
Document Type: ${document.docType}
Date: ${document.dateVal || "N/A"}
Customer (Messers): ${document.messers || "N/A"}
Vessel: ${document.vesselName || "N/A"}
Address: ${document.address || "N/A"}
Requisition No: ${document.requisitionNo || "N/A"}
Challan/Invoice No: ${document.challanNo || document.invoiceNo || "N/A"}
VAT %: ${document.vatPercent || 0}%
Transportation/Boat Hire: ${document.transportationFee || 0}
Rows count: ${document.rows.length}
Items sample:
${JSON.stringify(document.rows.slice(0, 30), null, 2)}

Provide:
1. Overall Audit Score (0-100)
2. Critical Issues / Missing Mandatory Shipping Fields (e.g. Missing Vessel name, zero prices, missing units)
3. Pricing & Calculation Verification
4. Recommendations for immediate correction
5. Recommended professional cover remark for the Ship Master or Superintendent.`;

    const response = await ai.models.generateContent({
      model: chosenModel,
      contents: prompt,
      config: {
        systemInstruction: ROLE_SYSTEM_INSTRUCTIONS.vessel_auditor,
        temperature: 0.4,
      },
    });

    return res.json({
      auditResult: response.text || "",
      modelUsed: chosenModel,
    });
  } catch (error: any) {
    console.error("Error in /api/gemini/audit-document:", error);
    return res.status(500).json({
      error: error?.message || "Failed to audit document with Gemini AI.",
    });
  }
});

// Vite Middleware integration for development / Static serving for production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (_req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Comilla Traders server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
