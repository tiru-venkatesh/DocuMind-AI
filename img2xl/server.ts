import 'dotenv/config';
import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';
import OpenAI from 'openai';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

// ── Auto-repair truncated / malformed JSON ───────────────────────────────────
function repairAndParseJSON(raw: string): any {
  // Strip markdown fences
  let s = raw.replace(/```json\s*/gi, '').replace(/```\s*/gi, '').trim();

  // 1. Try direct parse first
  try { return JSON.parse(s); } catch {}

  // 2. Try extracting { ... } block
  const objStart = s.indexOf('{');
  const objEnd   = s.lastIndexOf('}');
  if (objStart !== -1 && objEnd !== -1 && objEnd > objStart) {
    try { return JSON.parse(s.slice(objStart, objEnd + 1)); } catch {}
  }

  // 3. JSON is truncated — auto-close open brackets/braces
  const base = objStart !== -1 ? s.slice(objStart) : s;
  try {
    const fixed = closeTruncatedJSON(base);
    const result = JSON.parse(fixed);
    console.log('⚠️  Used truncation repair');
    return result;
  } catch {}

  throw new Error('Cannot parse AI JSON. First 300 chars: ' + s.slice(0, 300));
}

function closeTruncatedJSON(s: string): string {
  let result = s.trim();

  // Remove trailing comma + any incomplete last item
  result = result.replace(/,\s*$/, '');
  // Remove incomplete last array item (starts with [ or " but not closed)
  result = result.replace(/,\s*\[[^\]]*$/, '');
  result = result.replace(/,\s*"[^"]*$/, '');

  // Count unclosed brackets and braces
  let braces = 0, brackets = 0;
  let inStr = false, escape = false;
  for (const ch of result) {
    if (escape)          { escape = false; continue; }
    if (ch === '\\')     { escape = true;  continue; }
    if (ch === '"')      { inStr = !inStr; continue; }
    if (inStr)           continue;
    if (ch === '{')      braces++;
    else if (ch === '}') braces--;
    else if (ch === '[') brackets++;
    else if (ch === ']') brackets--;
  }

  // Close in reverse order
  for (let i = 0; i < brackets; i++) result += ']';
  for (let i = 0; i < braces;   i++) result += '}';

  return result;
}

async function startServer() {
  const app  = express();
  const PORT = 3000;
  app.use(express.json({ limit: '50mb' }));

  // ── /api/extract ─────────────────────────────────────────────────────────
  app.post('/api/extract', async (req, res) => {
    try {
      const { base64Image, mimeType } = req.body;
      if (!base64Image || !mimeType)
        return res.status(400).json({ error: 'base64Image and mimeType required.' });

      const token = process.env.GROQ_API_KEY;
      if (!token) return res.status(500).json({ error: 'GROQ_API_KEY not set.' });

      const client = new OpenAI({ baseURL: 'https://api.groq.com/openai/v1', apiKey: token });

      const response = await client.chat.completions.create({
        model:       'meta-llama/llama-4-scout-17b-16e-instruct',
        max_tokens:  8192,
        temperature: 0.1,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'text',
              text: `Extract ALL tabular data from this image.
Return ONLY this JSON format, nothing else, no markdown:
{"headers":["Col1","Col2"],"rows":[["val1","val2"],["val3","val4"]]}
- Include every single row without skipping
- Use "Column 1", "Column 2" if no headers visible
- Ensure the JSON is complete and properly closed`,
            },
            {
              type: 'image_url',
              image_url: { url: `data:${mimeType};base64,${base64Image}` },
            },
          ],
        }],
      });

      const rawText    = response.choices[0]?.message?.content || '';
      const finishReason = response.choices[0]?.finish_reason;
      console.log(`AI finish_reason: ${finishReason} | response length: ${rawText.length}`);
      console.log('AI raw (first 500):', rawText.slice(0, 500));

      let parsed: any;
      try { parsed = repairAndParseJSON(rawText); }
      catch (e: any) {
        return res.status(422).json({ error: `Could not parse AI response: ${e.message}` });
      }

      let data: any[] = [];
      if (parsed?.headers && parsed?.rows) {
        const headers: string[] = parsed.headers;
        data = (parsed.rows as any[][]).map(row => {
          const obj: Record<string, any> = {};
          headers.forEach((h, i) => { obj[h || `Column ${i+1}`] = row[i] ?? ''; });
          return obj;
        });
      } else if (Array.isArray(parsed)) {
        if (parsed.length > 0 && Array.isArray(parsed[0])) {
          data = parsed.map((row: any[]) => {
            const obj: Record<string, any> = {};
            row.forEach((v, i) => { obj[`Column ${i+1}`] = v; });
            return obj;
          });
        } else {
          data = parsed;
        }
      } else {
        return res.status(422).json({ error: 'Unexpected AI response format.' });
      }

      if (!data.length) {
        return res.status(422).json({ error: 'No table data found in this image.' });
      }

      console.log(`✅ Extracted ${data.length} rows, ${Object.keys(data[0]).length} columns`);
      res.json({ success: true, data });
    } catch (err: any) {
      console.error('Extract error:', err);
      res.status(500).json({ error: err?.message || 'Extraction failed.' });
    }
  });

  // ── /api/chat ─────────────────────────────────────────────────────────────
  app.post('/api/chat', async (req, res) => {
    try {
      const { messages } = req.body;
      const token = process.env.GROQ_API_KEY;
      if (!token) return res.status(500).json({ error: 'GROQ_API_KEY not set.' });

      const client = new OpenAI({ baseURL: 'https://api.groq.com/openai/v1', apiKey: token });

      const response = await client.chat.completions.create({
        model:      'llama-3.3-70b-versatile',
        max_tokens: 1024,
        messages: [
          {
            role: 'system',
            content: `You are Img2XL Bot. Img2XL converts images and PDFs with tables into Excel files using AI.
Supported formats: PNG, JPG, JPEG, WEBP. Users can rename headers, merge columns, download .xlsx.
Be concise, friendly, and helpful.`,
          },
          ...(messages || []).map(({ role, content }: any) => ({
            role: role === 'model' ? 'assistant' : role,
            content,
          })),
        ],
      });

      res.json({ content: response.choices[0].message.content });
    } catch (err: any) {
      console.error('Chat error:', err);
      res.status(500).json({ error: err?.message || 'Chat failed.' });
    }
  });

  // ── /api/generate-excel ───────────────────────────────────────────────────
  app.post('/api/generate-excel', (req, res) => {
    try {
      const { data, filename } = req.body;
      if (!data || !Array.isArray(data)) return res.status(400).json({ error: 'Invalid data.' });
      const ws   = XLSX.utils.json_to_sheet(data);
      const wb   = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      const buf  = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      const name = (filename || 'export').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${name}.xlsx"`);
      res.send(buf);
    } catch { res.status(500).json({ error: 'Excel generation failed.' }); }
  });

  app.get('/api/health', (_req, res) => res.json({ status: 'ok' }));

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true, hmr: process.env.DISABLE_HMR !== 'true' },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const dist = path.join(process.cwd(), 'dist');
    app.use(express.static(dist));
    app.get('*', (_req, res) => res.sendFile(path.join(dist, 'index.html')));
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n✅  Img2XL → http://localhost:${PORT}`);
    console.log(`   GROQ: ${process.env.GROQ_API_KEY ? '✓ loaded' : '✗ MISSING'}\n`);
  });
}

startServer();
