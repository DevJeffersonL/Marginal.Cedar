import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API routes
  app.post("/api/analyze", async (req, res) => {
    try {
      const { trades } = req.body;

      if (!process.env.GEMINI_API_KEY) {
        return res.status(500).json({ error: "Gemini API key is missing on the server." });
      }

      if (!trades || !Array.isArray(trades)) {
        return res.status(400).json({ error: "Invalid trades data provided." });
      }

      if (trades.length === 0) {
        return res.json({ analysis: "No trades recorded yet. Start adding deals to get AI-powered insights." });
      }

      const tradesContext = trades.map(t => {
        const buyPrice = t.buyAmount || 0;
        const sellPrice = t.sellAmount || 0;
        return {
          buyDate: t.buyDate || "N/A",
          buyAmount: buyPrice,
          sellDate: t.sellDate || "N/A",
          sellAmount: sellPrice,
          profit: sellPrice - buyPrice,
          type: t.type || 'pair'
        };
      });

      const prompt = `
        Analyze the following trade data for "Marginal Trade Tracker". 
        Provide a professional, concise "Profit Leakage Analysis".
        Highlight strengths, weaknesses, and a specific actionable tip to improve the user's marginal efficiency.
        Keep the tone technical yet encouraging.
        
        Trades:
        ${JSON.stringify(tradesContext, null, 2)}
      `;

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt
      });
      
      res.json({ analysis: response.text });
    } catch (error) {
      console.error("Gemini Analysis Error:", error);
      res.status(500).json({ error: "Unable to perform AI analysis at this moment." });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on port ${PORT}`);
  });
}

startServer();
