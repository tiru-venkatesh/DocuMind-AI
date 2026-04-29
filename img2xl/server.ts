import 'dotenv/config';
import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';
import OpenAI from 'openai';
import mammoth from 'mammoth';
import { createRequire } from 'module';

const require    = createRequire(import.meta.url);
const pdfParse   = require('pdf-parse');
const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const mkClient = () => new OpenAI({
  baseURL: 'https://api.groq.com/openai/v1',
  apiKey: process.env.GROQ_API_KEY || '',
});

// ── chunk large text into safe sizes ─────────────────────────────────────────
function chunkText(text: string, maxChars = 18000): string[] {
  const chunks: string[] = [];
  for (let i = 0; i < text.length; i += maxChars) {
    chunks.push(text.slice(i, i + maxChars));
  }
  return chunks;
}

// ── extract table rows from a text chunk ─────────────────────────────────────
async function chunkToTable(chunk: string, ai: OpenAI): Promise<any[]> {
  try {
    const r = await ai.chat.completions.create({
      model: 'llama-3.3-70b-versatile',
      max_tokens: 2048,
      temperature: 0.1,
      messages: [{
        role: 'user',
        content: 'Extract ALL tabular/structured data from this text.\nReturn ONLY a JSON array of objects. No markdown. No explanation. Return [] if no table.\nExample: [{"Name":"Alice","Score":"95"}]\n\nTEXT:\n' + chunk,
      }],
    });
    const raw = (r.choices[0]?.message?.content || '[]')
      .replace(/```json\s*/gi, '').replace(/```\s*/gi, '').trim();
    try {
      const d = JSON.parse(raw);
      return Array.isArray(d) ? d : [];
    } catch { return []; }
  } catch { return []; }
}

// ── OCR one image buffer via Groq vision ─────────────────────────────────────
async function ocrImageBuffer(imgBase64: string, mimeType: string, label: string, ai: OpenAI): Promise<string> {
  try {
    const r = await ai.chat.completions.create({
      model: 'meta-llama/llama-4-scout-17b-16e-instruct',
      max_tokens: 2048,
      temperature: 0.1,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'text',
            text: label + ' — You are an OCR engine. Extract ALL visible text exactly as it appears. Output raw text only, no summaries.',
          },
          {
            type: 'image_url',
            image_url: { url: 'data:' + mimeType + ';base64,' + imgBase64 },
          },
        ],
      }],
    });
    return r.choices[0]?.message?.content?.trim() || '';
  } catch (e: any) {
    console.warn('OCR failed for ' + label + ':', e.message);
    return '';
  }
}

// ── render PDF pages to images and OCR each ──────────────────────────────────
async function ocrScannedPdf(buf: Buffer, ai: OpenAI): Promise<{ rawText: string; tableData: any[]; pages: number }> {
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.js');
  const canvasMod = await import('canvas');
  const createCanvas = canvasMod.createCanvas;

  const pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  const totalPages = pdfDoc.numPages;
  const maxPages = Math.min(totalPages, 50);

  console.log('Scanned PDF — OCR processing ' + maxPages + ' of ' + totalPages + ' pages...');

  const pageTexts: string[] = [];
  const allRows: any[] = [];

  for (let pageNum = 1; pageNum <= maxPages; pageNum++) {
    try {
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: 1.8 });
      const canvas = createCanvas(viewport.width, viewport.height);
      const ctx = canvas.getContext('2d') as any;

      await page.render({ canvasContext: ctx, viewport }).promise;

      const imgBase64 = canvas.toBuffer('image/png').toString('base64');
      const pageText = await ocrImageBuffer(imgBase64, 'image/png', 'Page ' + pageNum + ' of ' + totalPages, ai);

      if (pageText) {
        pageTexts.push('--- Page ' + pageNum + ' ---\n' + pageText);
        console.log('  Page ' + pageNum + ': ' + pageText.length + ' chars');
        const rows = await chunkToTable(pageText, ai);
        allRows.push(...rows);
      }

      if (pageNum < maxPages) {
        await new Promise(r => setTimeout(r, 150));
      }
    } catch (e: any) {
      console.warn('Page ' + pageNum + ' failed:', e.message);
    }
  }

  let rawText = pageTexts.join('\n\n');
  if (totalPages > maxPages) {
    rawText += '\n\n[Note: OCR limited to first ' + maxPages + ' of ' + totalPages + ' pages]';
  }

  return { rawText, tableData: allRows, pages: totalPages };
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '200mb' }));

  // ── POST /api/extract-full ─────────────────────────────────────────────────
  app.post('/api/extract-full', async (req, res) => {
    try {
      const { base64, mimeType, fileName = '' } = req.body;
      if (!base64) return res.status(400).json({ error: 'base64 required.' });

      const buf  = Buffer.from(base64, 'base64');
      const name = fileName.toLowerCase();
      const ai   = mkClient();

      let rawText  = '';
      let pages    = 0;
      let tableData: any[] = [];
      let fileType = 'unknown';

      // ── IMAGE ──────────────────────────────────────────────────────────────
      if (mimeType.startsWith('image/')) {
        fileType = 'image';
        pages    = 1;

        // OCR — extract all text
        rawText = await ocrImageBuffer(base64, mimeType, 'Image', ai);
        console.log('Image OCR: ' + rawText.length + ' chars');

        // Extract table from OCR text
        if (rawText.trim()) {
          tableData = await chunkToTable(rawText, ai);
        }

        // Fallback: direct vision table extraction
        if (tableData.length === 0) {
          try {
            const r = await ai.chat.completions.create({
              model: 'meta-llama/llama-4-scout-17b-16e-instruct',
              max_tokens: 4096,
              temperature: 0.1,
              messages: [{
                role: 'user',
                content: [
                  {
                    type: 'text',
                    text: 'Extract ALL tabular data from this image.\nReturn ONLY a JSON array of objects. No markdown. Return [] if no table.\nExample: [{"Item":"Coffee","Price":"$2.50"}]',
                  },
                  {
                    type: 'image_url',
                    image_url: { url: 'data:' + mimeType + ';base64,' + base64 },
                  },
                ],
              }],
            });
            const raw2 = (r.choices[0]?.message?.content || '[]')
              .replace(/```json\s*/gi, '').replace(/```\s*/gi, '').trim();
            try {
              const d = JSON.parse(raw2);
              if (Array.isArray(d) && d.length > 0) tableData = d;
            } catch {}
          } catch {}
        }
      }

      // ── PDF ────────────────────────────────────────────────────────────────
      else if (mimeType.includes('pdf') || name.endsWith('.pdf')) {
        fileType = 'pdf';

        // Try text extraction first
        const pdf = await pdfParse(buf);
        pages   = pdf.numpages || 1;
        rawText = (pdf.text || '').trim();
        console.log('PDF: ' + pages + ' pages, ' + rawText.length + ' chars of text');

        if (!rawText || rawText.length < 50) {
          // Scanned PDF — OCR each page via vision
          const result = await ocrScannedPdf(buf, ai);
          rawText   = result.rawText;
          tableData = result.tableData;
          pages     = result.pages;
        } else {
          // Text PDF — chunk and extract table
          const chunks  = chunkText(rawText, 18000);
          console.log(chunks.length + ' chunks to process...');
          const results = await Promise.all(chunks.map(c => chunkToTable(c, ai)));
          tableData = results.flat();
        }
      }

      // ── DOCX ───────────────────────────────────────────────────────────────
      else if (name.endsWith('.docx') || mimeType.includes('wordprocessingml')) {
        fileType = 'docx';
        const doc = await mammoth.extractRawText({ buffer: buf });
        rawText   = doc.value || '';
        pages     = Math.ceil(rawText.length / 3000);
        const results = await Promise.all(chunkText(rawText, 18000).map(c => chunkToTable(c, ai)));
        tableData = results.flat();
      }

      // ── EXCEL ──────────────────────────────────────────────────────────────
      else if (name.endsWith('.xlsx') || name.endsWith('.xls') || mimeType.includes('spreadsheetml') || mimeType.includes('ms-excel')) {
        fileType = 'excel';
        const wb = XLSX.read(buf, { type: 'buffer' });
        let allText = '';
        for (const sh of wb.SheetNames) {
          const ws = wb.Sheets[sh];
          tableData.push(...XLSX.utils.sheet_to_json(ws, { defval: '' }));
          allText += '\n=== Sheet: ' + sh + ' ===\n' + XLSX.utils.sheet_to_csv(ws);
        }
        rawText = allText;
        pages   = wb.SheetNames.length;
      }

      // ── CSV ────────────────────────────────────────────────────────────────
      else if (name.endsWith('.csv') || mimeType === 'text/csv') {
        fileType = 'csv';
        const wb  = XLSX.read(buf.toString('utf-8'), { type: 'string' });
        const ws  = wb.Sheets[wb.SheetNames[0]];
        tableData = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rawText   = XLSX.utils.sheet_to_csv(ws);
        pages     = 1;
      }

      // ── TSV ────────────────────────────────────────────────────────────────
      else if (name.endsWith('.tsv') || mimeType === 'text/tab-separated-values') {
        fileType = 'tsv';
        const wb  = XLSX.read(buf.toString('utf-8'), { type: 'string', FS: '\t' });
        const ws  = wb.Sheets[wb.SheetNames[0]];
        tableData = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rawText   = buf.toString('utf-8');
        pages     = 1;
      }

      // ── TXT ────────────────────────────────────────────────────────────────
      else if (name.endsWith('.txt') || mimeType === 'text/plain') {
        fileType = 'txt';
        rawText  = buf.toString('utf-8');
        pages    = Math.ceil(rawText.length / 3000);
        const results = await Promise.all(chunkText(rawText, 18000).map(c => chunkToTable(c, ai)));
        tableData = results.flat();
      }

      // ── JSON ───────────────────────────────────────────────────────────────
      else if (name.endsWith('.json') || mimeType === 'application/json') {
        fileType = 'json';
        const parsed = JSON.parse(buf.toString('utf-8'));
        tableData = Array.isArray(parsed) ? parsed : [parsed];
        rawText   = JSON.stringify(parsed, null, 2);
        pages     = 1;
      }

      // ── PPTX ───────────────────────────────────────────────────────────────
      else if (name.endsWith('.pptx') || mimeType.includes('presentationml')) {
        fileType = 'pptx';
        const AdmZip = (await import('adm-zip')).default;
        const zip    = new AdmZip(buf);
        const slides = zip.getEntries().filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName));
        pages = slides.length;
        let txt = '';
        for (const s of slides) {
          const xml = s.getData().toString('utf-8');
          txt += (xml.match(/<a:t>([^<]*)<\/a:t>/g) || [])
            .map(t => t.replace(/<[^>]+>/g, '')).join(' ') + '\n';
        }
        rawText = txt;
        const results = await Promise.all(chunkText(rawText, 18000).map(c => chunkToTable(c, ai)));
        tableData = results.flat();
      }

      else {
        return res.status(415).json({ error: 'Unsupported file type: ' + mimeType });
      }

      console.log('✅ ' + fileType.toUpperCase() + ' | ' + pages + 'p | ' + rawText.length + 'ch | ' + tableData.length + ' rows');

      res.json({
        success:   true,
        fileType,
        pages,
        rawText,
        tableData,
        charCount: rawText.length,
        rowCount:  tableData.length,
      });

    } catch (err: any) {
      console.error('Extract error:', err.message);
      res.status(500).json({ error: err.message || 'Extraction failed.' });
    }
  });

  // ── POST /api/chat-doc ─────────────────────────────────────────────────────
  app.post('/api/chat-doc', async (req, res) => {
    try {
      const { messages, context = '' } = req.body;
      const ai  = mkClient();
      const ctx = context.slice(0, 8000);

      const r = await ai.chat.completions.create({
        model: 'llama-3.3-70b-versatile',
        max_tokens: 1024,
        messages: [
          {
            role: 'system',
            content: 'You are a smart document assistant. The user uploaded a document. Here is its full content:\n\n---\n' + ctx + '\n---\n\nAnswer questions accurately and helpfully.',
          },
          ...(messages || []).map(({ role, content }: any) => ({
            role: role === 'model' ? 'assistant' : role,
            content,
          })),
        ],
      });

      res.json({ content: r.choices[0].message.content });
    } catch (err: any) {
      console.error('Chat error:', err.message);
      res.status(500).json({ error: err.message });
    }
  });

  // ── POST /api/download ─────────────────────────────────────────────────────
  app.post('/api/download', (req, res) => {
    try {
      const { data = [], format = 'xlsx', filename = 'export', rawText = '' } = req.body;

      if (format === 'xlsx') {
        const ws  = XLSX.utils.json_to_sheet(data);
        const wb  = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Data');
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '.xlsx"');
        return res.send(buf);
      }
      if (format === 'csv') {
        const ws  = XLSX.utils.json_to_sheet(data);
        const csv = XLSX.utils.sheet_to_csv(ws);
        res.setHeader('Content-Type', 'text/csv');
        res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '.csv"');
        return res.send(csv);
      }
      if (format === 'json') {
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '.json"');
        return res.send(JSON.stringify(data, null, 2));
      }
      if (format === 'txt') {
        res.setHeader('Content-Type', 'text/plain');
        res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '.txt"');
        return res.send(rawText || JSON.stringify(data, null, 2));
      }

      res.status(400).json({ error: 'Unknown format: ' + format });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get('/api/health', (_req, res) => res.json({ status: 'ok', version: '5.0' }));

  // ── Vite dev / static prod ─────────────────────────────────────────────────
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

  const PORT = 3000;
  app.listen(PORT, '0.0.0.0', () => {
    console.log('\n✅  Img2XL v5.0 → http://localhost:' + PORT);
    console.log('   GROQ: ' + (process.env.GROQ_API_KEY ? '✓ loaded' : '✗ MISSING'));
    console.log('   Inputs: Images PDF(OCR) DOCX XLSX CSV TSV TXT JSON PPTX');
    console.log('   Outputs: XLSX CSV JSON TXT\n');
  });
}

startServer();