import { GoogleGenAI } from "@google/genai";
import { Delivery } from "../types";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export async function askAssistant(prompt: string, context: Delivery[]) {
  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: [
      {
        parts: [
          {
            text: `Eres un experto asistente logístico. 
            Tienes acceso a los datos actuales de entregas de la empresa.
            Responde de forma concisa y profesional en español.
            Solo responde sobre logística y los datos proporcionados.
            
            Contexto actual de entregas:
            ${JSON.stringify(context.slice(0, 50))} // Top 50 deliveries
            `
          },
          { text: prompt }
        ]
      }
    ]
  });
  
  return response.text;
}
