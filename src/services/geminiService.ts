import { GoogleGenAI } from "@google/genai";

const API_KEY = process.env.GEMINI_API_KEY;

export async function analyzeTrades(trades: any[]) {
  if (!API_KEY) {
    throw new Error("Gemini API key is missing. Please set GEMINI_API_KEY in your secrets.");
  }

  if (trades.length === 0) {
    return "No trades recorded yet. Start adding deals to get AI-powered insights.";
  }

  const ai = new GoogleGenAI({ apiKey: API_KEY });

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

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
    });
    return response.text;
  } catch (error) {
    console.error("Gemini Analysis Error:", error);
    return "Unable to perform AI analysis at this moment. Please check your connection or API configuration.";
  }
}
