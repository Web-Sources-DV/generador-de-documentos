import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';

dotenv.config();

let aiClient: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI {
  if (!aiClient) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error('GEMINI_API_KEY environment variable is not configured.');
    }
    aiClient = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        },
      },
    });
  }
  return aiClient;
}

// Models to try in order of preference if primary experiences high demand (503/429)
const CANDIDATE_MODELS = [
  'gemini-3.7-flash',
  'gemini-flash-latest',
  'gemini-3.1-flash-lite',
  'gemini-3.1-pro-preview',
];

async function generateWithFallbackAndRetry(
  ai: GoogleGenAI,
  cleanBase64: string,
  mimeType: string,
  prompt: string
) {
  let lastError: any = null;

  for (const modelName of CANDIDATE_MODELS) {
    const maxRetries = 2;
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          // Exponential backoff wait
          await new Promise((resolve) => setTimeout(resolve, 800 * Math.pow(2, attempt - 1)));
        }

        const response = await ai.models.generateContent({
          model: modelName,
          contents: {
            parts: [
              {
                inlineData: {
                  mimeType,
                  data: cleanBase64,
                },
              },
              {
                text: prompt,
              },
            ],
          },
          config: {
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.OBJECT,
              properties: {
                firstName: {
                  type: Type.STRING,
                  description: 'Nombres de pila de la persona (ej. CARLOS ANDRÉS)',
                },
                lastName: {
                  type: Type.STRING,
                  description: 'Apellidos completos de la persona (ej. PÉREZ GÓMEZ)',
                },
                fullName: {
                  type: Type.STRING,
                  description: 'Nombre completo (Nombres y Apellidos)',
                },
                passportNumber: {
                  type: Type.STRING,
                  description: 'Número de pasaporte o documento (ej. PA12345678, A1234567)',
                },
                nationality: {
                  type: Type.STRING,
                  description: 'Nacionalidad (ej. ESPAÑOLA, MEXICANA, COLOMBIANA, ESTADOUNIDENSE)',
                },
                documentType: {
                  type: Type.STRING,
                  description: 'Tipo de documento detectado: "Pasaporte", "Cédula", "DNI", "NIE" o "Documento de Identidad"',
                },
                issuingCountry: {
                  type: Type.STRING,
                  description: 'País emisor del pasaporte o documento (ej. ESPAÑA, MÉXICO, COLOMBIA, EE. UU.)',
                },
                birthDate: {
                  type: Type.STRING,
                  description: 'Fecha de nacimiento (ej. 1988-05-14 o 14/05/1988)',
                },
                expiryDate: {
                  type: Type.STRING,
                  description: 'Fecha de vencimiento o caducidad del pasaporte (ej. 2030-10-22)',
                },
                issueDate: {
                  type: Type.STRING,
                  description: 'Fecha de emisión o expedición si figura en el documento',
                },
                sex: {
                  type: Type.STRING,
                  description: 'Sexo o género (M, F, X u Otro)',
                },
                personalNumber: {
                  type: Type.STRING,
                  description: 'Número de identificación nacional, DNI, NIE, CURP, Cédula o RUT si aparece',
                },
                placeOfBirth: {
                  type: Type.STRING,
                  description: 'Lugar o ciudad de nacimiento si figura',
                },
                mrzLine1: {
                  type: Type.STRING,
                  description: 'Primera línea del código MRZ inferior (si está presente)',
                },
                mrzLine2: {
                  type: Type.STRING,
                  description: 'Segunda línea del código MRZ inferior (si está presente)',
                },
                confidenceScore: {
                  type: Type.NUMBER,
                  description: 'Nivel estimado de legibilidad y confianza de 0 a 100',
                },
                notes: {
                  type: Type.STRING,
                  description: 'Observaciones sobre la calidad de la foto o peculiaridades del documento',
                },
              },
              required: ['firstName', 'lastName', 'passportNumber', 'nationality'],
            },
          },
        });

        return {
          response,
          modelUsed: modelName,
        };
      } catch (err: any) {
        lastError = err;
        const errMsg = String(err?.message || '');
        const is503OrRateLimit = errMsg.includes('503') || errMsg.includes('UNAVAILABLE') || errMsg.includes('high demand') || errMsg.includes('429') || errMsg.includes('RESOURCE_EXHAUSTED');

        console.warn(`[AI Extraction] Attempt ${attempt + 1} with model ${modelName} failed:`, errMsg);

        if (!is503OrRateLimit && attempt === 0) {
          // If it's not a temporary overload (e.g. invalid payload), break to try next model or fail
          break;
        }
      }
    }
  }

  throw lastError || new Error('No se pudo procesar el documento con los modelos de IA disponibles.');
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Increase payload limit for base64 passport images
  app.use(express.json({ limit: '30mb' }));
  app.use(express.urlencoded({ extended: true, limit: '30mb' }));

  // API Route: Health check
  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
  });

  // API Route: AI Passport & ID Data Extraction
  app.post('/api/extract-passport', async (req, res) => {
    try {
      const { imageBase64, mimeType = 'image/jpeg' } = req.body;

      if (!imageBase64) {
        return res.status(400).json({
          success: false,
          error: 'No se proporcionó la imagen del documento o pasaporte.',
        });
      }

      // Clean base64 string if it contains data URI prefix
      const cleanBase64 = imageBase64.replace(/^data:image\/[a-zA-Z0-9+]+;base64,/, '');

      const ai = getGeminiClient();

      const prompt = `Analiza detalladamente esta fotografía de un pasaporte o documento de identidad oficial.
Extrae con la máxima precisión todos los campos de información, prestando especial atención a la zona de lectura mecánica (MRZ en la parte inferior) y la zona visual.
Devuelve los datos estrictamente en español en el formato JSON especificado.
Si algún campo no está visible o no existe en el documento, déjalo como cadena vacía ("").
Asegúrate de formatear las fechas preferentemente en formato estándar YYYY-MM-DD o DD/MM/YYYY.`;

      const { response, modelUsed } = await generateWithFallbackAndRetry(
        ai,
        cleanBase64,
        mimeType,
        prompt
      );

      const responseText = response.text || '{}';
      let parsedData: any = {};
      try {
        parsedData = JSON.parse(responseText);
      } catch (parseErr) {
        console.error('Error parsing JSON from Gemini response:', responseText);
        parsedData = {};
      }

      return res.json({
        success: true,
        data: parsedData,
        modelUsed,
        extractedAt: new Date().toISOString(),
      });
    } catch (error: any) {
      console.error('Error during passport extraction:', error);
      const isUnavailable = String(error?.message || '').includes('503') || String(error?.message || '').includes('UNAVAILABLE') || String(error?.message || '').includes('high demand');
      
      return res.status(isUnavailable ? 503 : 500).json({
        success: false,
        isUnavailable,
        error: error.message || 'Error al procesar la imagen con IA',
      });
    }
  });

  // Vite development middleware or static production serving
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`SQP Legal Consulting server running on http://localhost:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
