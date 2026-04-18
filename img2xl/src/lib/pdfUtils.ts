import { PDFDocument, rgb, degrees, StandardFonts } from 'pdf-lib';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

export interface WatermarkOptions {
  text: string;
  opacity: number; // 0 to 1
  size: number;
  rotation: number; // degrees
  color: { r: number, g: number, b: number };
}

/**
 * Generates a PDF from table data using jsPDF and applies a watermark using pdf-lib
 */
export async function generateWatermarkedPdf(
  data: any[], 
  filename: string, 
  watermark: WatermarkOptions
): Promise<Uint8Array> {
  // 1. Generate base PDF with table using jsPDF
  const doc = new jsPDF();
  const headers = Object.keys(data[0] || {});
  const rows = data.map(item => Object.values(item));

  autoTable(doc, {
    head: [headers],
    body: rows,
    theme: 'grid',
    styles: { fontSize: 8 },
    headStyles: { fillColor: [59, 130, 246] } // Accent color #3B82F6
  });

  const pdfArrayBuffer = doc.output('arraybuffer');

  // 2. Load into pdf-lib to apply watermark features
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const pages = pdfDoc.getPages();
  const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  for (const page of pages) {
    const { width, height } = page.getSize();
    
    // Draw watermark text in the center
    page.drawText(watermark.text, {
      x: width / 2 - (watermark.text.length * watermark.size / 4), // Rough centering
      y: height / 2,
      size: watermark.size,
      font: font,
      color: rgb(watermark.color.r / 255, watermark.color.g / 255, watermark.color.b / 255),
      opacity: watermark.opacity,
      rotate: degrees(watermark.rotation),
    });
  }

  return await pdfDoc.save();
}

/**
 * Efficiently converts Uint8Array to Base64 string without stack overflow
 */
function bufferToBase64(buffer: Uint8Array): string {
  const CHUNK_SIZE = 0x8000;
  let index = 0;
  const length = buffer.length;
  let result = '';
  while (index < length) {
    const slice = buffer.subarray(index, Math.min(index + CHUNK_SIZE, length));
    result += String.fromCharCode.apply(null, slice as unknown as number[]);
    index += CHUNK_SIZE;
  }
  return btoa(result);
}

/**
 * Splits a large PDF into smaller base64 chunks for processing
 * This allows "unlimited" size support by bypassing the 50MB inline limit
 * Sequential processing ensures the AI handles each part within its context window
 */
export async function splitPdfIntoChunks(
  file: File, 
  pagesPerChunk: number = 5,
  onProgress?: (chunk: number, total: number) => void
): Promise<{ data: string, mimeType: string }[]> {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(arrayBuffer);
  const totalPages = pdfDoc.getPageCount();
  const chunks: { data: string, mimeType: string }[] = [];
  const totalChunks = Math.ceil(totalPages / pagesPerChunk);

  for (let i = 0; i < totalPages; i += pagesPerChunk) {
    const chunkIndex = Math.floor(i / pagesPerChunk);
    if (onProgress) onProgress(chunkIndex + 1, totalChunks);

    const newDoc = await PDFDocument.create();
    const end = Math.min(i + pagesPerChunk, totalPages);
    
    // Copy pages
    const pageIndices = Array.from({ length: end - i }, (_, k) => i + k);
    const copiedPages = await newDoc.copyPages(pdfDoc, pageIndices);
    copiedPages.forEach((page) => newDoc.addPage(page));

    const pdfBytes = await newDoc.save();
    const base64 = bufferToBase64(pdfBytes);
    
    chunks.push({
      data: base64,
      mimeType: 'application/pdf'
    });
  }

  return chunks;
}
