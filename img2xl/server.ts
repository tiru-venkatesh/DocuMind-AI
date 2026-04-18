import 'dotenv/config';
import express from 'express';
import { createServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';
import OpenAI from "openai";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = process.env.PORT || 3000;

  app.use(express.json({ limit: '50mb' }));

  // API Route for Chatbot using Groq (free)
  app.post('/api/chat', async (req, res) => {
    try {
      const { messages, context } = req.body;
      const token = process.env["GROQ_API_KEY"];
      
      if (!token) {
        return res.status(500).json({ error: 'GROQ_API_KEY is not configured on the server.' });
      }

      const client = new OpenAI({
        baseURL: "https://api.groq.com/openai/v1",
        apiKey: token,
      });

      let contextPrompt = "";
      if (context) {
        try {
          const ctx = JSON.parse(context);
          contextPrompt = `
CURRENT LIVE VIEW CONTEXT:
- Filename: ${ctx.filename}
- Processing Status: ${ctx.status}
- Rows Extracted: ${ctx.rowCount}
- Columns Count: ${ctx.columnCount}
${ctx.error ? `- ERROR DETECTED: ${ctx.error}` : ""}

${ctx.dataSample ? `DATA SAMPLE (First 50 Rows):\n${JSON.stringify(ctx.dataSample, null, 2)}` : "No data has been extracted yet for this file or extraction is in progress."}

Analyze the data sample above if available. If there is an error, explain it to the user. If the file is still processing, tell them to wait a moment.`;
        } catch (e) {
          contextPrompt = `The user is viewing a file. Context preview: ${context}`;
        }
      }

      const systemMessage = {
        role: "system" as const,
        content: `You are Img2XL Bot, a helpful assistant for the Img2XL app. 
Img2XL uses AI to convert images and PDFs into structured Excel data.

Capabilities to explain:
- **Unlimited PDF Size**: We chunk large PDFs (8 pages/chunk) to handle huge files.
- **Batch Mode**: Upload multiple files to see them in the sidebar.
- **Data Cleanup**: We detect tables, fix headers, and normalize numbers automatically.

${contextPrompt}

Always be concise, analytical when data is present, and supportive. Use markdown formatting.`
      };

      // Strip any extra fields (e.g. timestamp) that Groq doesn't accept
      const cleanMessages = messages.map(({ role, content }: { role: string; content: string }) => ({ role, content }));

const response = await client.chat.completions.create({
  model: "llama-3.3-70b-versatile",
  max_tokens: 1024,
  messages: [systemMessage, ...cleanMessages],
});

      res.json({ content: response.choices[0].message.content });
    } catch (error: any) {
      console.error('Chatbot error:', error);
      const message = error?.message || error?.details || JSON.stringify(error) || 'Failed to get response from chatbot';
      res.status(500).json({ error: message });
    }
  });

  // API Route for Excel generation from JSON
  app.post('/api/generate-excel', (req, res) => {
    try {
      const { data, filename } = req.body;
      if (!data || !Array.isArray(data)) {
        return res.status(400).json({ error: 'Invalid data format' });
      }

      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      
      const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

      const cleanFilename = (filename || 'image-data').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${cleanFilename}.xlsx"`);
      res.send(buffer);
    } catch (error) {
      console.error('Excel generation error:', error);
      res.status(500).json({ error: 'Failed to generate Excel file' });
    }
  });

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok' });
  });

  // Vite middleware setup
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createServer({
      server: { 
        middlewareMode: true,
        hmr: process.env.DISABLE_HMR !== 'true'
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
}

startServer();
