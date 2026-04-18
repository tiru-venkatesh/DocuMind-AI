/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { 
  Upload, FileSpreadsheet, Trash2, Loader2, Download, AlertCircle, 
  Image as ImageIcon, CheckCircle2, Table as TableIcon, Sparkles,
  RefreshCw, Files, ArrowRight, ChevronRight, X, FileText, Package,
  ShieldCheck, Zap, Globe, Database, Workflow, ClipboardList, Layers,
  Cpu, FileDown, Settings2, Plus, ChevronDown, BarChart3, Activity,
  Eye, Lock, Columns as ColumnsIcon, Rows as RowsIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import JSZip from 'jszip';
import ReactMarkdown from 'react-markdown';
import { extractTableFromImage, ExtractionOptions } from './lib/imageExtraction';
import { optimizeImage } from './lib/optimization';
import { splitPdfIntoChunks, generateWatermarkedPdf, WatermarkOptions } from './lib/pdfUtils';
import { cn } from './lib/utils';

interface FileItem {
  id: string;
  file: File;
  preview: string | null;
  status: 'pending' | 'processing' | 'completed' | 'error';
  progress: number;
  statusMessage: string;
  extractedData: any[] | null;
  error?: string;
}

export default function App() {
  const [items, setItems] = useState<FileItem[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [isProcessingAll, setIsProcessingAll] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [downloadDropdownOpen, setDownloadDropdownOpen] = useState(false);
  const [globalError, setGlobalError] = useState<string | null>(null);
  const [options, setOptions] = useState<ExtractionOptions>({ detectMultipleTables: false, autoCleanData: true });
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 50;
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [chatMessages, setChatMessages] = useState<{ role: 'user' | 'assistant', content: string, timestamp: string }[]>(() => {
    const saved = localStorage.getItem('img2xl_chat_history');
    if (!saved) return [];
    try { return JSON.parse(saved).map((m: any) => ({ ...m, timestamp: m.timestamp || new Date().toISOString() })); }
    catch { return []; }
  });
  const [chatInput, setChatInput] = useState('');
  const [isChatLoading, setIsChatLoading] = useState(false);
  const [isPdfModalOpen, setIsPdfModalOpen] = useState(false);
  const [watermark] = useState<WatermarkOptions>({ text: 'CONFIDENTIAL', opacity: 0.3, size: 45, rotation: 45, color: { r: 150, g: 150, b: 150 } });

  const fileInputRef = useRef<HTMLInputElement>(null);
  const chatScrollRef = useRef<HTMLDivElement>(null);
  const activeItem = items.find(item => item.id === activeId);

  const getFilteredData = (data: any[]) => {
    if (!data) return [];
    return data.filter(row => Object.entries(filters).every(([key, value]) => {
      if (!value) return true;
      return String(row[key] || '').toLowerCase().includes(String(value).toLowerCase());
    }));
  };

  const handleFilterChange = (column: string, value: string) => {
    setFilters(prev => ({ ...prev, [column]: value }));
    setCurrentPage(1);
  };

  const handleSendMessage = async () => {
    if (!chatInput.trim() || isChatLoading) return;
    const userMsg = { role: 'user' as const, content: chatInput, timestamp: new Date().toISOString() };
    setChatMessages(prev => [...prev, userMsg]);
    setChatInput('');
    setIsChatLoading(true);
    try {
      const liveViewContext = activeItem ? {
        filename: activeItem.file.name, status: activeItem.status, error: activeItem.error,
        rowCount: activeItem.extractedData?.length || 0,
        columnCount: activeItem.extractedData?.[0] ? Object.keys(activeItem.extractedData[0]).length : 0,
        dataSample: activeItem.extractedData ? activeItem.extractedData.slice(0, 50) : null
      } : null;
      const response = await fetch('/api/chat', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: chatMessages.map(({ role, content }) => ({ role, content })), context: liveViewContext ? JSON.stringify(liveViewContext) : undefined })
      });
      if (!response.ok) throw new Error('Chat failed');
      const data = await response.json();
      setChatMessages(prev => [...prev, { role: 'assistant', content: data.content, timestamp: new Date().toISOString() }]);
    } catch (err: any) {
      setChatMessages(prev => [...prev, { role: 'assistant', content: "Error: " + err.message, timestamp: new Date().toISOString() }]);
    } finally { setIsChatLoading(false); }
  };

  useEffect(() => { if (chatScrollRef.current) chatScrollRef.current.scrollTop = chatScrollRef.current.scrollHeight; }, [chatMessages, isChatLoading]);
  useEffect(() => { localStorage.setItem('img2xl_chat_history', JSON.stringify(chatMessages)); }, [chatMessages]);
  useEffect(() => {
    const handler = (e: MouseEvent) => { if (downloadDropdownOpen && !(e.target as Element).closest('.dl-dropdown')) setDownloadDropdownOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [downloadDropdownOpen]);

  const clearChatHistory = () => { setChatMessages([]); localStorage.removeItem('img2xl_chat_history'); setShowClearConfirm(false); };

  const exportChatHistory = () => {
    if (!chatMessages.length) return;
    const blob = new Blob([`IMG2XL CHAT EXPORT\n${'─'.repeat(50)}\n\n` + chatMessages.map(m => `[${new Date(m.timestamp).toLocaleString()}] ${m.role.toUpperCase()}\n${m.content}\n\n`).join('─'.repeat(30) + '\n\n')], { type: 'text/plain' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `img2xl_chat_${Date.now()}.txt` });
    a.click();
  };

  const processFiles = (selectedFiles: FileList | File[]) => {
    const newItems: FileItem[] = [];
    Array.from(selectedFiles).forEach(file => {
      const isImage = file.type.startsWith('image/');
      const isPDF = file.type === 'application/pdf';
      if (!isImage && !isPDF) { setGlobalError(`"${file.name}" is not supported. Use PNG, JPEG, or PDF.`); return; }
      if (file.size > 20 * 1024 * 1024) { setGlobalError(`"${file.name}" exceeds 20MB limit.`); return; }
      const id = Math.random().toString(36).substr(2, 9);
      const item: FileItem = { id, file, preview: null, status: 'pending', progress: 0, statusMessage: 'Ready', extractedData: null };
      if (isImage) { const r = new FileReader(); r.onload = e => setItems(prev => prev.map(p => p.id === id ? { ...p, preview: e.target?.result as string } : p)); r.readAsDataURL(file); }
      else item.preview = 'PDF_PLACEHOLDER';
      newItems.push(item);
    });
    if (newItems.length) { setItems(prev => [...prev, ...newItems]); if (!activeId) setActiveId(newItems[0].id); setGlobalError(null); }
  };

  const removeItem = (id: string) => setItems(prev => { const f = prev.filter(i => i.id !== id); if (activeId === id) setActiveId(f.length ? f[f.length - 1].id : null); return f; });

  const updateCell = (rowIndex: number, column: string, value: string) => {
    if (!activeId) return;
    setItems(prev => prev.map(item => item.id === activeId && item.extractedData ? { ...item, extractedData: item.extractedData.map((r, i) => i === rowIndex ? { ...r, [column]: value } : r) } : item));
  };

  const updateHeader = (oldH: string, newH: string) => {
    if (!activeId || !newH || oldH === newH) return;
    setItems(prev => prev.map(item => item.id === activeId && item.extractedData ? {
      ...item, extractedData: item.extractedData.map(row => Object.fromEntries(Object.entries(row).map(([k, v]) => [k === oldH ? newH : k, v])))
    } : item));
  };

  const addRow = (index: number) => {
    if (!activeId) return;
    setItems(prev => prev.map(item => {
      if (item.id === activeId && item.extractedData) {
        const headers = Object.keys(item.extractedData[0] || {});
        const newRow = headers.reduce((a, h) => ({ ...a, [h]: '' }), {});
        const d = [...item.extractedData]; d.splice(index + 1, 0, newRow);
        return { ...item, extractedData: d };
      }
      return item;
    }));
  };

  const deleteRow = (index: number) => {
    if (!activeId) return;
    setItems(prev => prev.map(item => item.id === activeId && item.extractedData ? { ...item, extractedData: item.extractedData.filter((_, i) => i !== index) } : item));
  };

  const extractSingle = async (id: string) => {
    const item = items.find(i => i.id === id);
    if (!item || item.status === 'completed' || item.status === 'processing') return;
    const upd = (s: Partial<FileItem>) => setItems(prev => prev.map(i => i.id === id ? { ...i, ...s } : i));
    upd({ status: 'processing', progress: 10, statusMessage: 'Preprocessing…', error: undefined });
    try {
      let all: any[] = [];
      if (item.file.type === 'application/pdf') {
        upd({ progress: 15, statusMessage: 'Reading PDF…' });
        const chunks = await splitPdfIntoChunks(item.file, 8, (c, t) => upd({ progress: 15 + Math.floor((c / t) * 5), statusMessage: `Splitting… (${c}/${t})` }));
        for (let i = 0; i < chunks.length; i++) {
          upd({ progress: 25 + Math.floor((i / chunks.length) * 65), statusMessage: `Analysing part ${i + 1}/${chunks.length}…` });
          all = [...all, ...await extractTableFromImage(chunks[i].data, chunks[i].mimeType, options)];
        }
      } else {
        upd({ progress: 15, statusMessage: 'Optimising image…' });
        const base64 = await optimizeImage(item.file);
        if (base64.length > 50 * 1024 * 1024) throw new Error('File exceeds 50 MB limit after optimisation.');
        upd({ progress: 40, statusMessage: 'AI Analysis…' });
        all = await extractTableFromImage(base64, item.file.type, options);
      }
      upd({ status: 'completed', progress: 100, statusMessage: 'Complete', extractedData: all });
    } catch (err: any) {
      let msg = err.message || 'Extraction failed';
      if (msg.includes('Safety') || msg.includes('blocked')) msg = 'Content blocked by safety filters.';
      else if (msg.includes('429') || msg.includes('quota')) msg = 'Rate limit exceeded. Please wait and retry.';
      upd({ status: 'error', progress: 0, statusMessage: 'Failed', error: msg });
    }
  };

  const handleExtractAll = async () => {
    const pending = items.filter(i => i.status === 'pending' || i.status === 'error');
    if (!pending.length) return;
    setIsProcessingAll(true);
    for (const item of pending) await extractSingle(item.id);
    setIsProcessingAll(false);
  };

  const downloadExcel = async (item: FileItem) => {
    if (!item.extractedData) return;
    try {
      const r = await fetch('/api/generate-excel', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ data: item.extractedData, filename: item.file.name.split('.')[0] }) });
      if (!r.ok) throw new Error();
      const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(await r.blob()), download: `${item.file.name.split('.')[0]}.xlsx` });
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
    } catch { setGlobalError(`Failed to download "${item.file.name}"`); }
  };

  const downloadCSV = (item: FileItem) => {
    if (!item.extractedData?.length) return;
    const headers = Object.keys(item.extractedData[0]);
    const csv = [headers.join(','), ...item.extractedData.map(row => headers.map(h => { const c = String(row[h] || '').replace(/"/g, '""'); return c.includes(',') || c.includes('"') || c.includes('\n') ? `"${c}"` : c; }).join(','))].join('\n');
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([csv], { type: 'text/csv' })), download: `${item.file.name.split('.')[0]}.csv` });
    a.click();
  };

  const downloadJSON = (item: FileItem) => {
    if (!item.extractedData) return;
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([JSON.stringify(item.extractedData, null, 2)], { type: 'application/json' })), download: `${item.file.name.split('.')[0]}.json` });
    a.click();
  };

  const downloadAllAsZip = async () => {
    const done = items.filter(i => i.status === 'completed' && i.extractedData);
    if (!done.length) return;
    const zip = new JSZip();
    for (const item of done) {
      try {
        const r = await fetch('/api/generate-excel', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ data: item.extractedData, filename: item.file.name.split('.')[0] }) });
        if (r.ok) zip.file(`${item.file.name.split('.')[0]}.xlsx`, await r.blob());
      } catch {}
    }
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(await zip.generateAsync({ type: 'blob' })), download: `img2xl_${Date.now()}.zip` });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  };

  const totalRows = items.reduce((a, i) => a + (i.extractedData?.length || 0), 0);
  const completedCount = items.filter(i => i.status === 'completed').length;

  return (
    <div className="min-h-screen flex flex-col font-sans overflow-hidden" style={{ background: '#080810', fontFamily: "'DM Sans', 'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@300;400;500&family=Playfair+Display:ital,wght@0,700;1,700&display=swap');
        :root {
          --gold: #C9A84C;
          --gold-dim: #8B6914;
          --cream: #F5EDD6;
          --ink: #0A0A12;
          --glass: rgba(255,255,255,0.03);
          --border: rgba(255,255,255,0.06);
          --accent: #C9A84C;
        }
        .scrollbar-hide::-webkit-scrollbar { display: none; }
        .scrollbar-hide { -ms-overflow-style: none; scrollbar-width: none; }
        .thin-scroll::-webkit-scrollbar { width: 2px; }
        .thin-scroll::-webkit-scrollbar-track { background: transparent; }
        .thin-scroll::-webkit-scrollbar-thumb { background: var(--gold-dim); border-radius: 2px; }
        .serif { font-family: 'Playfair Display', Georgia, serif; }
        .mono { font-family: 'DM Mono', 'Courier New', monospace; }
        .panel { background: var(--glass); border: 1px solid var(--border); backdrop-filter: blur(20px); border-radius: 16px; }
        .gold-glow { box-shadow: 0 0 30px rgba(201,168,76,0.15); }
        .upload-zone { border: 1px dashed rgba(201,168,76,0.2); border-radius: 12px; transition: all 0.2s; }
        .upload-zone:hover { border-color: var(--gold); background: rgba(201,168,76,0.04); }
        .file-row { transition: all 0.15s; border: 1px solid transparent; border-radius: 12px; }
        .file-row:hover { background: rgba(255,255,255,0.03); }
        .file-row.active { background: rgba(201,168,76,0.07); border-color: rgba(201,168,76,0.25); }
        .btn-gold { background: var(--gold); color: #000; font-weight: 700; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; border: none; transition: all 0.15s; }
        .btn-gold:hover { background: #DFC06A; transform: translateY(-1px); }
        .btn-gold:active { transform: translateY(0); }
        .btn-ghost { background: transparent; border: 1px solid var(--border); color: rgba(255,255,255,0.5); font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; font-weight: 600; transition: all 0.15s; }
        .btn-ghost:hover { border-color: var(--gold); color: var(--gold); }
        .tag { font-size: 9px; letter-spacing: 0.15em; text-transform: uppercase; font-weight: 700; padding: 3px 8px; border-radius: 4px; }
        .tag-green { background: rgba(52,211,153,0.1); color: #34D399; border: 1px solid rgba(52,211,153,0.2); }
        .tag-gold { background: rgba(201,168,76,0.1); color: var(--gold); border: 1px solid rgba(201,168,76,0.2); }
        .tag-red { background: rgba(239,68,68,0.1); color: #F87171; border: 1px solid rgba(239,68,68,0.2); }
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th { background: #0D0D18; border-bottom: 1px solid var(--border); padding: 12px 16px; font-size: 9px; letter-spacing: 0.15em; text-transform: uppercase; color: rgba(255,255,255,0.35); font-weight: 600; text-align: left; position: sticky; top: 0; z-index: 10; }
        .data-table td { border-bottom: 1px solid rgba(255,255,255,0.03); padding: 10px 16px; font-size: 11px; color: rgba(255,255,255,0.65); vertical-align: top; }
        .data-table tr:hover td { background: rgba(201,168,76,0.03); }
        .chat-user { background: rgba(201,168,76,0.1); border: 1px solid rgba(201,168,76,0.15); border-radius: 12px 12px 2px 12px; padding: 12px 16px; }
        .chat-ai { background: rgba(255,255,255,0.03); border: 1px solid var(--border); border-radius: 12px 12px 12px 2px; padding: 12px 16px; }
        .divider { width: 1px; height: 16px; background: var(--border); }
        .stat-card { background: rgba(255,255,255,0.02); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }
        .noise-bg { position: relative; }
        .noise-bg::before { content: ''; position: absolute; inset: 0; background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.03'/%3E%3C/svg%3E"); opacity: 0.4; pointer-events: none; border-radius: inherit; }
      `}</style>

      {/* Top Nav */}
      <nav style={{ borderBottom: '1px solid var(--border)', background: 'rgba(8,8,16,0.8)', backdropFilter: 'blur(24px)' }} className="px-8 py-5 flex justify-between items-center shrink-0 sticky top-0 z-50">
        <div className="flex items-center gap-10">
          <div className="flex items-center gap-3">
            <div style={{ background: 'var(--gold)', width: 32, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: 8 }}>
              <TableIcon size={18} color="#000" />
            </div>
            <div>
              <div style={{ fontFamily: 'DM Sans', fontWeight: 700, fontSize: 16, letterSpacing: '-0.02em', color: '#fff' }}>
                IMG<span style={{ color: 'var(--gold)' }}>2XL</span>
              </div>
              <div style={{ fontSize: 8, letterSpacing: '0.2em', color: 'var(--gold-dim)', fontFamily: 'DM Mono', marginTop: -2 }}>PRO · v4.2</div>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-6" style={{ fontSize: 11, letterSpacing: '0.08em', color: 'rgba(255,255,255,0.3)', fontWeight: 500 }}>
            {['Documentation', 'API', 'Pricing'].map(l => <a key={l} href="#" style={{ transition: 'color 0.15s' }} onMouseEnter={e => (e.target as HTMLElement).style.color = 'var(--gold)'} onMouseLeave={e => (e.target as HTMLElement).style.color = 'rgba(255,255,255,0.3)'}>{l}</a>)}
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-4" style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border)', borderRadius: 8, padding: '8px 14px' }}>
            <div className="flex items-center gap-2">
              <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#34D399', boxShadow: '0 0 8px #34D399' }} />
              <span style={{ fontSize: 9, letterSpacing: '0.12em', color: 'rgba(255,255,255,0.4)', fontFamily: 'DM Mono' }}>SYSTEM NOMINAL</span>
            </div>
            <div style={{ width: 1, height: 12, background: 'var(--border)' }} />
            <div className="flex items-center gap-1.5">
              <Zap size={11} color="var(--gold)" />
              <span style={{ fontSize: 9, letterSpacing: '0.12em', color: 'var(--gold)', fontFamily: 'DM Mono' }}>28ms</span>
            </div>
          </div>
          <button className="btn-gold px-6 py-2 rounded-lg">Upgrade</button>
        </div>
      </nav>

      <main className="flex-1 grid gap-5 px-5 pb-5 pt-5 overflow-hidden" style={{ gridTemplateColumns: '320px 1fr 340px', height: 'calc(100vh - 73px)' }}>
        
        {/* LEFT PANEL */}
        <aside className="flex flex-col gap-4 overflow-hidden">
          {/* Headline */}
          <div className="shrink-0 px-1 pt-2">
            <h1 style={{ fontFamily: 'Playfair Display', fontSize: 42, fontWeight: 700, lineHeight: 0.92, letterSpacing: '-0.03em', color: '#fff' }}>
              Visual<br /><em style={{ color: 'var(--gold)' }}>data,</em><br />structured.
            </h1>
            <p style={{ marginTop: 12, fontSize: 12, color: 'rgba(255,255,255,0.3)', lineHeight: 1.6, fontWeight: 400 }}>
              Enterprise extraction from images and PDFs. Instant Excel output.
            </p>
          </div>

          {/* File Queue */}
          <div className="panel flex-1 flex flex-col overflow-hidden noise-bg" style={{ padding: '20px' }}>
            <div className="flex items-center justify-between mb-4 shrink-0">
              <div className="flex items-center gap-2">
                <Files size={13} color="var(--gold)" />
                <span style={{ fontSize: 10, letterSpacing: '0.12em', color: 'rgba(255,255,255,0.4)', fontWeight: 600, textTransform: 'uppercase' }}>Queue</span>
                <span style={{ fontSize: 10, background: 'rgba(201,168,76,0.1)', color: 'var(--gold)', padding: '1px 7px', borderRadius: 4, fontFamily: 'DM Mono', border: '1px solid rgba(201,168,76,0.2)' }}>{items.length}</span>
              </div>
              {items.length > 0 && <button onClick={() => setItems([])} style={{ fontSize: 9, color: 'rgba(239,68,68,0.5)', letterSpacing: '0.1em', textTransform: 'uppercase', fontWeight: 700, cursor: 'pointer', background: 'none', border: 'none', transition: 'color 0.15s' }} onMouseEnter={e => (e.target as HTMLElement).style.color = '#F87171'} onMouseLeave={e => (e.target as HTMLElement).style.color = 'rgba(239,68,68,0.5)'}>Clear all</button>}
            </div>

            <div className="flex-1 overflow-y-auto thin-scroll space-y-2">
              <AnimatePresence initial={false}>
                {items.map(item => (
                  <motion.div key={item.id} initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }}
                    onClick={() => setActiveId(item.id)}
                    className={cn('file-row p-3 flex items-center gap-3 cursor-pointer group', activeId === item.id && 'active')}
                  >
                    <div style={{ width: 36, height: 36, background: 'rgba(255,255,255,0.04)', border: '1px solid var(--border)', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                      {item.file.type === 'application/pdf' ? <FileText size={16} color="var(--gold)" /> : <ImageIcon size={16} color="var(--gold)" />}
                    </div>
                    <div className="flex-1 min-w-0">
                      <p style={{ fontSize: 11, fontWeight: 600, color: activeId === item.id ? '#fff' : 'rgba(255,255,255,0.55)', letterSpacing: '-0.01em', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.file.name}</p>
                      <div className="mt-1">
                        {item.status === 'processing' && (
                          <div>
                            <div className="flex items-center gap-1.5 mb-1">
                              <Loader2 size={9} color="var(--gold)" className="animate-spin" />
                              <span style={{ fontSize: 9, color: 'var(--gold)', letterSpacing: '0.1em', textTransform: 'uppercase', fontWeight: 700 }}>{item.statusMessage}</span>
                            </div>
                            <div style={{ height: 2, background: 'rgba(255,255,255,0.06)', borderRadius: 1, overflow: 'hidden' }}>
                              <motion.div style={{ height: '100%', background: 'var(--gold)' }} animate={{ width: `${item.progress}%` }} transition={{ duration: 0.3 }} />
                            </div>
                          </div>
                        )}
                        {item.status === 'completed' && <span className="tag tag-green">Done</span>}
                        {item.status === 'error' && <div className="flex items-center gap-2"><span className="tag tag-red">Error</span><button onClick={e => { e.stopPropagation(); extractSingle(item.id); }}><RefreshCw size={10} color="var(--gold)" /></button></div>}
                        {item.status === 'pending' && <span style={{ fontSize: 9, color: 'rgba(255,255,255,0.2)', letterSpacing: '0.08em', textTransform: 'uppercase' }}>Pending</span>}
                      </div>
                    </div>
                    <button onClick={e => { e.stopPropagation(); removeItem(item.id); }} style={{ opacity: 0, transition: 'opacity 0.15s' }} className="group-hover:opacity-100" onMouseEnter={e => (e.currentTarget as HTMLElement).style.opacity = '1'} onMouseLeave={e => (e.currentTarget as HTMLElement).style.opacity = '0'}>
                      <X size={13} color="rgba(239,68,68,0.6)" />
                    </button>
                  </motion.div>
                ))}
              </AnimatePresence>

              <div className="upload-zone p-8 text-center cursor-pointer mt-2" onClick={() => fileInputRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); processFiles(e.dataTransfer.files); }}>
                <input type="file" ref={fileInputRef} onChange={e => e.target.files && processFiles(e.target.files)} className="hidden" accept="image/*,application/pdf" multiple />
                <Upload size={20} color="var(--gold)" style={{ margin: '0 auto 10px', opacity: 0.6 }} />
                <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.3)', letterSpacing: '0.08em', textTransform: 'uppercase', fontWeight: 600 }}>Drop files or click</p>
                <p style={{ fontSize: 9, color: 'rgba(255,255,255,0.15)', marginTop: 4 }}>PNG, JPEG, PDF · max 20 MB</p>
              </div>
            </div>

            {items.length > 0 && (
              <div style={{ borderTop: '1px solid var(--border)', marginTop: 16, paddingTop: 16 }} className="shrink-0 space-y-3">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="checkbox" checked={options.detectMultipleTables} onChange={e => setOptions(p => ({ ...p, detectMultipleTables: e.target.checked }))} style={{ accentColor: 'var(--gold)', width: 13, height: 13 }} />
                  <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.4)', letterSpacing: '0.08em', textTransform: 'uppercase', fontWeight: 600 }}>Detect multiple tables</span>
                </label>
                <button onClick={handleExtractAll} disabled={isProcessingAll || items.every(i => i.status === 'completed')} className="btn-gold w-full py-3 rounded-xl" style={{ opacity: (isProcessingAll || items.every(i => i.status === 'completed')) ? 0.4 : 1, cursor: (isProcessingAll || items.every(i => i.status === 'completed')) ? 'not-allowed' : 'pointer' }}>
                  {isProcessingAll ? `Processing… (${completedCount}/${items.length})` : 'Extract All Files'}
                </button>
              </div>
            )}
          </div>

          {/* Stats */}
          <div className="grid grid-cols-2 gap-3 shrink-0">
            <div className="stat-card">
              <div style={{ fontSize: 8, letterSpacing: '0.15em', color: 'rgba(255,255,255,0.25)', textTransform: 'uppercase', fontWeight: 700, marginBottom: 6 }}>Total Rows</div>
              <div style={{ fontSize: 28, fontWeight: 700, color: '#fff', lineHeight: 1, fontFamily: 'DM Mono' }}>{totalRows.toLocaleString()}</div>
            </div>
            <div className="stat-card">
              <div style={{ fontSize: 8, letterSpacing: '0.15em', color: 'rgba(255,255,255,0.25)', textTransform: 'uppercase', fontWeight: 700, marginBottom: 6 }}>Files Done</div>
              <div style={{ fontSize: 28, fontWeight: 700, color: 'var(--gold)', lineHeight: 1, fontFamily: 'DM Mono' }}>{completedCount}/{items.length || '—'}</div>
            </div>
          </div>

          {items.some(i => i.status === 'completed') && (
            <button onClick={downloadAllAsZip} className="btn-ghost w-full py-3 rounded-xl flex items-center justify-center gap-2 shrink-0">
              <Package size={13} /> Download All as ZIP
            </button>
          )}

          {globalError && (
            <div style={{ background: 'rgba(239,68,68,0.08)', border: '1px solid rgba(239,68,68,0.2)', borderRadius: 10, padding: '10px 14px' }} className="flex items-start gap-2 shrink-0">
              <AlertCircle size={13} color="#F87171" style={{ marginTop: 1, flexShrink: 0 }} />
              <p style={{ fontSize: 10, color: '#F87171', lineHeight: 1.5 }}>{globalError}</p>
            </div>
          )}
        </aside>

        {/* MIDDLE PANEL */}
        <section className="flex flex-col overflow-hidden h-full">
          <AnimatePresence mode="wait">
            {!activeItem ? (
              <motion.div key="empty" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
                className="panel flex-1 overflow-y-auto thin-scroll noise-bg"
              >
                {/* Hero */}
                <div style={{ padding: '60px 64px', position: 'relative', overflow: 'hidden' }}>
                  <div style={{ position: 'absolute', top: -60, right: -60, opacity: 0.04 }}>
                    <TableIcon size={400} color="var(--gold)" />
                  </div>
                  <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.1 }}>
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: 'rgba(201,168,76,0.08)', border: '1px solid rgba(201,168,76,0.2)', borderRadius: 20, padding: '6px 14px', marginBottom: 28 }}>
                      <Sparkles size={12} color="var(--gold)" />
                      <span style={{ fontSize: 10, color: 'var(--gold)', letterSpacing: '0.15em', textTransform: 'uppercase', fontWeight: 700 }}>AI-Powered · Stateless · Instant</span>
                    </div>
                    <h1 style={{ fontFamily: 'Playfair Display', fontSize: 72, fontWeight: 700, lineHeight: 0.88, letterSpacing: '-0.03em', color: '#fff', marginBottom: 28 }}>
                      From<br /><em style={{ color: 'var(--gold)' }}>image</em><br />to data.
                    </h1>
                    <p style={{ fontSize: 15, color: 'rgba(255,255,255,0.4)', lineHeight: 1.7, maxWidth: 480, marginBottom: 40, fontWeight: 400 }}>
                      Img2XL uses computer vision and language models to extract every cell, header, and row from your images and PDFs — and exports them directly to Excel.
                    </p>
                    <div className="flex items-center gap-4 flex-wrap">
                      <button onClick={() => fileInputRef.current?.click()} className="btn-gold px-10 py-4 rounded-xl" style={{ fontSize: 12 }}>Upload a file</button>
                      <div className="flex items-center gap-3" style={{ fontSize: 11, color: 'rgba(255,255,255,0.25)' }}>
                        <ShieldCheck size={13} color="rgba(255,255,255,0.2)" /> Stateless · No data stored
                      </div>
                    </div>
                  </motion.div>
                </div>

                {/* Feature Grid */}
                <div style={{ padding: '0 64px 64px', display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 16 }}>
                  {[
                    { icon: Zap, title: 'Fast Extraction', desc: 'Complex tables extracted in seconds with our parallelised pipeline.' },
                    { icon: Globe, title: '40+ Languages', desc: 'Full support for RTL scripts, Asian typography, and mixed-language docs.' },
                    { icon: Database, title: 'Smart Schema', desc: 'Headers inferred, data types normalised, noise removed automatically.' },
                    { icon: Lock, title: 'Zero Storage', desc: 'Images processed in memory and discarded. Nothing stored on our servers.' },
                    { icon: Layers, title: 'Multi-Format', desc: 'PNG, JPEG, PDF — all processed through the same precision pipeline.' },
                    { icon: Activity, title: 'Batch Mode', desc: 'Upload entire document sets and process them in one queued operation.' },
                  ].map((f, i) => (
                    <motion.div key={i} whileHover={{ y: -4 }} style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid var(--border)', borderRadius: 14, padding: '24px', transition: 'border-color 0.2s' }}
                      onMouseEnter={e => (e.currentTarget as HTMLElement).style.borderColor = 'rgba(201,168,76,0.3)'}
                      onMouseLeave={e => (e.currentTarget as HTMLElement).style.borderColor = 'var(--border)'}
                    >
                      <div style={{ width: 40, height: 40, background: 'rgba(201,168,76,0.08)', border: '1px solid rgba(201,168,76,0.15)', borderRadius: 10, display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 16 }}>
                        <f.icon size={18} color="var(--gold)" />
                      </div>
                      <h3 style={{ fontSize: 13, fontWeight: 700, color: '#fff', marginBottom: 8, letterSpacing: '-0.01em' }}>{f.title}</h3>
                      <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.3)', lineHeight: 1.6 }}>{f.desc}</p>
                    </motion.div>
                  ))}
                </div>
              </motion.div>

            ) : activeItem.status === 'processing' ? (
              <motion.div key="loading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
                className="panel flex-1 flex flex-col items-center justify-center noise-bg"
              >
                <div style={{ position: 'relative', marginBottom: 32 }}>
                  <div style={{ width: 80, height: 80, border: '2px solid rgba(255,255,255,0.06)', borderTopColor: 'var(--gold)', borderRadius: '50%' }} className="animate-spin" />
                  <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                    <Sparkles size={24} color="var(--gold)" />
                  </div>
                </div>
                <h3 style={{ fontFamily: 'Playfair Display', fontSize: 24, color: '#fff', marginBottom: 8 }}>{activeItem.statusMessage}</h3>
                <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.3)' }}>{activeItem.file.name}</p>
              </motion.div>

            ) : activeItem.status === 'error' ? (
              <motion.div key="error" initial={{ opacity: 0 }} animate={{ opacity: 1 }}
                className="panel flex-1 flex flex-col items-center justify-center noise-bg"
              >
                <div style={{ background: 'rgba(239,68,68,0.08)', border: '1px solid rgba(239,68,68,0.2)', borderRadius: 20, padding: 32, marginBottom: 24 }}>
                  <AlertCircle size={40} color="#F87171" />
                </div>
                <h3 style={{ fontFamily: 'Playfair Display', fontSize: 24, color: '#fff', marginBottom: 8 }}>Extraction Failed</h3>
                <p style={{ fontSize: 12, color: '#F87171', marginBottom: 24, maxWidth: 340, textAlign: 'center', lineHeight: 1.6 }}>{activeItem.error}</p>
                <button onClick={() => extractSingle(activeItem.id)} className="btn-gold px-8 py-3 rounded-xl">Retry</button>
              </motion.div>

            ) : activeItem.extractedData && (
              <motion.div key="results" initial={{ opacity: 0 }} animate={{ opacity: 1 }}
                className="panel flex-1 flex flex-col overflow-hidden noise-bg"
              >
                {/* Results Header */}
                <div style={{ padding: '16px 24px', borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 }}>
                  <div className="flex items-center gap-3">
                    <span className="tag tag-green">Extracted</span>
                    <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.3)', fontFamily: 'DM Mono', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 200 }}>{activeItem.file.name}</span>
                    <div className="divider" />
                    <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.25)', fontFamily: 'DM Mono' }}>{getFilteredData(activeItem.extractedData).length} / {activeItem.extractedData.length} rows · {Object.keys(activeItem.extractedData[0] || {}).length} cols</span>
                  </div>
                  <div className="flex items-center gap-2">
                    {Object.values(filters).some(f => f) && <button onClick={() => setFilters({})} className="btn-ghost px-3 py-1.5 rounded-lg flex items-center gap-1"><Trash2 size={10} /> Clear filters</button>}
                    <button onClick={() => setIsEditMode(!isEditMode)} style={{ fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase', fontWeight: 700, padding: '6px 14px', borderRadius: 8, border: '1px solid', transition: 'all 0.15s', cursor: 'pointer', background: isEditMode ? 'var(--gold)' : 'transparent', color: isEditMode ? '#000' : 'rgba(255,255,255,0.4)', borderColor: isEditMode ? 'var(--gold)' : 'var(--border)' }}>
                      <span className="flex items-center gap-1.5"><Settings2 size={11} />{isEditMode ? 'Editing' : 'Edit'}</span>
                    </button>
                  </div>
                </div>

                {/* Table */}
                <div className="flex-1 overflow-auto thin-scroll" style={{ background: '#050508', margin: '12px', borderRadius: 12, border: '1px solid var(--border)' }}>
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th style={{ width: 44, textAlign: 'center', borderRight: '1px solid var(--border)' }}>#</th>
                        {Object.keys(activeItem.extractedData[0] || {}).map(header => (
                          <th key={header} style={{ minWidth: 160 }}>
                            {isEditMode ? (
                              <input type="text" value={header} onChange={e => updateHeader(header, e.target.value)}
                                style={{ background: 'transparent', border: 'none', color: 'rgba(255,255,255,0.6)', fontSize: 9, letterSpacing: '0.12em', textTransform: 'uppercase', fontWeight: 600, outline: 'none', width: '100%', fontFamily: 'DM Mono' }} />
                            ) : (
                              <div>
                                <div style={{ marginBottom: 6 }}>{header}</div>
                                <input type="text" placeholder={`Filter…`} value={filters[header] || ''} onChange={e => handleFilterChange(header, e.target.value)}
                                  style={{ width: '100%', background: 'rgba(255,255,255,0.04)', border: '1px solid var(--border)', borderRadius: 4, padding: '3px 8px', fontSize: 9, color: '#fff', outline: 'none', fontFamily: 'DM Mono' }} />
                              </div>
                            )}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const filtered = getFilteredData(activeItem.extractedData);
                        const start = (currentPage - 1) * rowsPerPage;
                        return filtered.slice(start, start + rowsPerPage).map((row, i) => (
                          <tr key={start + i}>
                            <td style={{ textAlign: 'center', borderRight: '1px solid var(--border)', color: 'rgba(255,255,255,0.2)', fontFamily: 'DM Mono', fontSize: 10 }}>
                              {isEditMode ? (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                  <button onClick={() => deleteRow(start + i)} style={{ color: 'rgba(239,68,68,0.4)', background: 'none', border: 'none', cursor: 'pointer', padding: 2 }}><X size={10} /></button>
                                  <button onClick={() => addRow(start + i)} style={{ color: 'rgba(201,168,76,0.4)', background: 'none', border: 'none', cursor: 'pointer', padding: 2 }}><Plus size={10} /></button>
                                </div>
                              ) : start + i + 1}
                            </td>
                            {Object.entries(row).map(([key, val]: [string, any], j) => (
                              <td key={j} style={{ borderRight: '1px solid rgba(255,255,255,0.02)' }}>
                                {isEditMode ? (
                                  <textarea value={val !== null && val !== undefined ? String(val) : ''} onChange={e => updateCell(start + i, key, e.target.value)}
                                    style={{ background: 'transparent', border: 'none', color: 'rgba(255,255,255,0.65)', fontSize: 11, outline: 'none', width: '100%', resize: 'none', minHeight: 36, fontFamily: 'DM Sans', padding: 0 }} rows={1} />
                                ) : val !== null && val !== undefined ? String(val) : <span style={{ color: 'rgba(255,255,255,0.1)' }}>—</span>}
                              </td>
                            ))}
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>

                {/* Pagination */}
                {getFilteredData(activeItem.extractedData).length > rowsPerPage && (
                  <div style={{ borderTop: '1px solid var(--border)', padding: '10px 24px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 12, flexShrink: 0 }}>
                    <button onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1} style={{ opacity: currentPage === 1 ? 0.2 : 1, background: 'none', border: 'none', cursor: 'pointer', color: '#fff' }}><ChevronRight size={16} style={{ transform: 'rotate(180deg)' }} /></button>
                    <span style={{ fontSize: 10, fontFamily: 'DM Mono', color: 'rgba(255,255,255,0.35)', letterSpacing: '0.08em' }}>Page {currentPage} of {Math.ceil(getFilteredData(activeItem.extractedData).length / rowsPerPage)}</span>
                    <button onClick={() => setCurrentPage(p => Math.min(Math.ceil(getFilteredData(activeItem.extractedData).length / rowsPerPage), p + 1))} disabled={currentPage === Math.ceil(getFilteredData(activeItem.extractedData).length / rowsPerPage)} style={{ opacity: currentPage === Math.ceil(getFilteredData(activeItem.extractedData).length / rowsPerPage) ? 0.2 : 1, background: 'none', border: 'none', cursor: 'pointer', color: '#fff' }}><ChevronRight size={16} /></button>
                  </div>
                )}

                {/* Action Bar */}
                <div style={{ padding: '16px 24px', borderTop: '1px solid var(--border)', display: 'flex', gap: 10, flexShrink: 0, background: 'rgba(0,0,0,0.2)' }}>
                  <div className="dl-dropdown" style={{ position: 'relative', flex: 1 }}>
                    <div style={{ display: 'flex', borderRadius: 10, overflow: 'hidden' }}>
                      <button onClick={() => downloadExcel(activeItem)} className="btn-gold flex-1 py-3 flex items-center justify-center gap-2" style={{ borderRadius: 0 }}>
                        <FileSpreadsheet size={13} /> Download .XLSX
                      </button>
                      <button onClick={() => setDownloadDropdownOpen(!downloadDropdownOpen)} className="btn-gold py-3 px-3" style={{ borderRadius: 0, borderLeft: '1px solid rgba(0,0,0,0.2)' }}>
                        <ChevronDown size={13} style={{ transform: downloadDropdownOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s' }} />
                      </button>
                    </div>
                    <AnimatePresence>
                      {downloadDropdownOpen && (
                        <motion.div initial={{ opacity: 0, y: 6 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 6 }}
                          style={{ position: 'absolute', bottom: '100%', left: 0, marginBottom: 6, width: '100%', background: '#111118', border: '1px solid var(--border)', borderRadius: 12, padding: 6, zIndex: 100 }}
                        >
                          {[{ label: 'Excel (.XLSX)', icon: FileSpreadsheet, fn: () => { downloadExcel(activeItem); setDownloadDropdownOpen(false); } },
                            { label: 'CSV (.CSV)', icon: FileText, fn: () => { downloadCSV(activeItem); setDownloadDropdownOpen(false); } },
                            { label: 'JSON (.JSON)', icon: Database, fn: () => { downloadJSON(activeItem); setDownloadDropdownOpen(false); } }
                          ].map(opt => (
                            <button key={opt.label} onClick={opt.fn} style={{ width: '100%', textAlign: 'left', padding: '10px 14px', borderRadius: 8, fontSize: 11, color: 'rgba(255,255,255,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: 'none', border: 'none', cursor: 'pointer', transition: 'all 0.1s' }}
                              onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = 'rgba(201,168,76,0.08)'; (e.currentTarget as HTMLElement).style.color = 'var(--gold)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = 'none'; (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.55)'; }}
                            >
                              <span style={{ fontWeight: 600, letterSpacing: '0.06em' }}>{opt.label}</span>
                              <opt.icon size={12} />
                            </button>
                          ))}
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </div>
                  <button onClick={() => { setItems(prev => prev.map(i => i.id === activeId ? { ...i, status: 'pending', extractedData: null } : i)); extractSingle(activeItem.id); }}
                    className="btn-ghost px-4 py-3 rounded-xl flex items-center gap-2">
                    <RefreshCw size={13} />
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </section>

        {/* RIGHT PANEL — Chat */}
        <aside className="hidden xl:flex flex-col overflow-hidden h-full">
          <div className="panel flex-1 flex flex-col overflow-hidden noise-bg">
            {/* Chat Header */}
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0, background: 'rgba(255,255,255,0.01)' }}>
              <div className="flex items-center gap-3">
                <div style={{ width: 8, height: 8, borderRadius: '50%', background: 'var(--gold)', boxShadow: '0 0 12px var(--gold)' }} />
                <span style={{ fontSize: 11, fontWeight: 700, color: '#fff', letterSpacing: '0.08em', textTransform: 'uppercase' }}>AI Assistant</span>
                {activeItem && (
                  <div style={{ background: 'rgba(201,168,76,0.1)', border: '1px solid rgba(201,168,76,0.2)', borderRadius: 6, padding: '2px 8px', display: 'flex', alignItems: 'center', gap: 4 }}>
                    <div style={{ width: 5, height: 5, borderRadius: '50%', background: 'var(--gold)' }} className="animate-pulse" />
                    <span style={{ fontSize: 8, color: 'var(--gold)', letterSpacing: '0.12em', textTransform: 'uppercase', fontWeight: 700 }}>Live</span>
                  </div>
                )}
              </div>
              {chatMessages.length > 0 && !showClearConfirm && (
                <div className="flex items-center gap-1">
                  <button onClick={exportChatHistory} style={{ padding: 6, borderRadius: 6, background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.3)', transition: 'color 0.15s' }} onMouseEnter={e => (e.currentTarget as HTMLElement).style.color = 'var(--gold)'} onMouseLeave={e => (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.3)'}><FileDown size={13} /></button>
                  <button onClick={() => setShowClearConfirm(true)} style={{ padding: 6, borderRadius: 6, background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.3)', transition: 'color 0.15s' }} onMouseEnter={e => (e.currentTarget as HTMLElement).style.color = '#F87171'} onMouseLeave={e => (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.3)'}><Trash2 size={13} /></button>
                </div>
              )}
              {showClearConfirm && (
                <div className="flex items-center gap-2" style={{ background: 'rgba(239,68,68,0.08)', border: '1px solid rgba(239,68,68,0.15)', borderRadius: 8, padding: '4px 8px' }}>
                  <span style={{ fontSize: 9, color: '#F87171', letterSpacing: '0.1em', textTransform: 'uppercase' }}>Clear?</span>
                  <button onClick={clearChatHistory} style={{ fontSize: 9, fontWeight: 700, color: '#fff', background: '#EF4444', border: 'none', borderRadius: 4, padding: '2px 8px', cursor: 'pointer' }}>Yes</button>
                  <button onClick={() => setShowClearConfirm(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.4)' }}><X size={11} /></button>
                </div>
              )}
            </div>

            {/* Messages */}
            <div className="flex-1 overflow-y-auto thin-scroll" style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: 12 }}>
              {chatMessages.length === 0 && (
                <div style={{ textAlign: 'center', marginTop: 48 }}>
                  <Sparkles size={28} color="var(--gold)" style={{ margin: '0 auto 12px', opacity: 0.4 }} />
                  <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.25)', letterSpacing: '0.06em', lineHeight: 1.6, maxWidth: 200, margin: '0 auto' }}>
                    Ask me to analyse your extracted data
                  </p>
                  <div style={{ marginTop: 20, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    {['Summarise the data', 'Find the highest value', 'Count unique entries'].map(s => (
                      <button key={s} onClick={() => { setChatInput(s); }} style={{ fontSize: 10, color: 'var(--gold)', background: 'rgba(201,168,76,0.06)', border: '1px solid rgba(201,168,76,0.15)', borderRadius: 8, padding: '7px 12px', cursor: 'pointer', textAlign: 'left', transition: 'all 0.15s' }}
                        onMouseEnter={e => (e.currentTarget as HTMLElement).style.background = 'rgba(201,168,76,0.1)'}
                        onMouseLeave={e => (e.currentTarget as HTMLElement).style.background = 'rgba(201,168,76,0.06)'}
                      >{s}</button>
                    ))}
                  </div>
                </div>
              )}
              {chatMessages.map((msg, i) => (
                <div key={i} style={{ display: 'flex', flexDirection: 'column', alignItems: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '88%', alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start' }}>
                  <div className={msg.role === 'user' ? 'chat-user' : 'chat-ai'}>
                    {msg.role === 'assistant' && <div style={{ fontSize: 8, color: 'var(--gold)', letterSpacing: '0.15em', textTransform: 'uppercase', fontWeight: 700, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 4 }}><Cpu size={9} /> Img2XL AI</div>}
                    <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.75)', lineHeight: 1.65, fontWeight: 400 }}>
                      <ReactMarkdown>{msg.content}</ReactMarkdown>
                    </div>
                  </div>
                  <span style={{ fontSize: 8, color: 'rgba(255,255,255,0.15)', marginTop: 4, fontFamily: 'DM Mono' }}>{new Date(msg.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                </div>
              ))}
              {isChatLoading && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '12px 16px', background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border)', borderRadius: '12px 12px 12px 2px', alignSelf: 'flex-start' }}>
                  {[0, 1, 2].map(i => <div key={i} style={{ width: 5, height: 5, borderRadius: '50%', background: 'var(--gold)', opacity: 0.7 }} className="animate-bounce" />)}
                </div>
              )}
              <div ref={chatScrollRef} />
            </div>

            {/* Input */}
            <div style={{ padding: '14px 16px', borderTop: '1px solid var(--border)', flexShrink: 0, background: 'rgba(0,0,0,0.2)' }}>
              <form onSubmit={e => { e.preventDefault(); handleSendMessage(); }} style={{ display: 'flex', gap: 8 }}>
                <input type="text" value={chatInput} onChange={e => setChatInput(e.target.value)} placeholder={activeItem?.extractedData ? `Ask about ${activeItem.file.name}…` : 'Upload a file to get started…'}
                  style={{ flex: 1, background: 'rgba(255,255,255,0.04)', border: '1px solid var(--border)', borderRadius: 12, padding: '10px 16px', fontSize: 12, color: '#fff', outline: 'none', fontFamily: 'DM Sans', transition: 'border-color 0.15s' }}
                  onFocus={e => (e.target as HTMLElement).style.borderColor = 'rgba(201,168,76,0.4)'}
                  onBlur={e => (e.target as HTMLElement).style.borderColor = 'var(--border)'}
                />
                <button type="submit" disabled={isChatLoading} className="btn-gold px-4 py-2 rounded-xl" style={{ opacity: isChatLoading ? 0.5 : 1, cursor: isChatLoading ? 'not-allowed' : 'pointer' }}>
                  <ArrowRight size={16} />
                </button>
              </form>
            </div>
          </div>
        </aside>
      </main>

      {/* Footer */}
      <footer style={{ borderTop: '1px solid var(--border)', padding: '14px 32px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'rgba(0,0,0,0.4)', backdropFilter: 'blur(20px)', flexShrink: 0 }}>
        <div className="flex items-center gap-6" style={{ fontSize: 10, color: 'rgba(255,255,255,0.2)', fontFamily: 'DM Mono', letterSpacing: '0.08em' }}>
          <span>© {new Date().getFullYear()} IMG2XL</span>
          <span style={{ color: 'var(--gold-dim)', fontWeight: 700 }}>HIMESH & TIRU</span>
        </div>
        <div className="flex items-center gap-4" style={{ fontSize: 9, color: 'rgba(255,255,255,0.15)', letterSpacing: '0.1em', textTransform: 'uppercase', fontFamily: 'DM Mono' }}>
          <div className="flex items-center gap-1.5"><Package size={11} /> Batch enabled</div>
          <div className="flex items-center gap-1.5"><ShieldCheck size={11} /> Stateless mode</div>
        </div>
      </footer>
    </div>
  );
}