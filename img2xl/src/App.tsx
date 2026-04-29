/**
 * Img2XL v5.0 — Universal Document → Excel/Data Extractor
 * Built by Himesh & Tiru
 * Auth: 2 free uses (name required), 3rd+ needs login
 */
import logo from './assets/logo.png';
import { Table } from 'lucide-react';
import React, {
  useState, useRef, useEffect, useContext, createContext,
  useMemo, useCallback,
} from 'react';
import {
  Upload, FileSpreadsheet, Trash2, Loader2, Download, AlertCircle,
  CheckCircle2, Table as TableIcon, Sparkles, RefreshCw, X,
  LogOut, History, Mail, Lock, Eye, EyeOff, ChevronDown,
  ShieldCheck, MessageSquare, Send, UserCircle2, Copy,
  FileText, File, ChevronRight, User, Edit2,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword,
  createUserWithEmailAndPassword, signOut, GoogleAuthProvider,
  signInWithPopup, updateProfile, User as FirebaseUser,
} from 'firebase/auth';
import {
  getFirestore, collection, addDoc, query, where, getDocs,
  deleteDoc, doc, orderBy, serverTimestamp, setDoc, getDoc, Timestamp,
} from 'firebase/firestore';

// ─── Firebase ─────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyCDKY9B4j4HMt39LhwgujNZTM0NNX-l-ts",
  authDomain: "my-project-1436-1754940582084.firebaseapp.com",
  projectId: "my-project-1436-1754940582084",
  storageBucket: "my-project-1436-1754940582084.firebasestorage.app",
  messagingSenderId: "917682303442",
  appId: "1:917682303442:web:9dd8a887a1f203458e180b",
};
const fbApp = initializeApp(firebaseConfig);
const auth  = getAuth(fbApp);
const db    = getFirestore(fbApp);

// ─── Design tokens ────────────────────────────────────────────────────────────
const G = {
  bg:'#0a0a0a', bgCard:'#111111', bgCard2:'#161616', border:'#222222',
  gold:'#C9A84C', goldL:'#E2C97E', goldD:'#8B6914',
  white:'#FFFFFF', gray:'#888888', dim:'#444444',
  ok:'#4ADE80', err:'#F87171', info:'#60A5FA',
};

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,700;0,900;1,700&family=DM+Sans:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{background:${G.bg};color:${G.white};font-family:'DM Sans',sans-serif;overflow-x:hidden}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:${G.bg}}
::-webkit-scrollbar-thumb{background:${G.goldD};border-radius:3px}
::selection{background:${G.gold}33;color:${G.goldL}}
.fd{font-family:'Playfair Display',serif}
.fm{font-family:'JetBrains Mono',monospace}
.gt{background:linear-gradient(135deg,${G.gold},${G.goldL},${G.gold});-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.bg{background:linear-gradient(135deg,${G.gold},${G.goldL});color:#000;font-weight:700;font-family:'DM Sans',sans-serif;letter-spacing:.08em;text-transform:uppercase;font-size:11px;border:none;cursor:pointer;transition:all .2s;border-radius:8px;display:inline-flex;align-items:center;gap:8px}
.bg:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 8px 24px ${G.gold}33}
.bg:active{transform:translateY(0)}
.bg:disabled{opacity:.5;pointer-events:none}
.gh{background:transparent;color:${G.gray};border:1px solid ${G.border};font-weight:600;font-family:'DM Sans',sans-serif;letter-spacing:.06em;text-transform:uppercase;font-size:11px;cursor:pointer;transition:all .2s;border-radius:8px;display:inline-flex;align-items:center;gap:8px}
.gh:hover{border-color:${G.gold}66;color:${G.goldL};background:${G.gold}0A}
.gh:disabled{opacity:.4;pointer-events:none}
.nb{font-size:13px;font-weight:500;color:${G.gray};cursor:pointer;background:none;border:none;transition:color .2s;padding:0}
.nb:hover,.nb.on{color:${G.white}}
.if{width:100%;background:${G.bgCard2};border:1px solid ${G.border};border-radius:8px;padding:12px 16px;color:${G.white};font-size:14px;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .2s}
.if:focus{border-color:${G.gold}66}
.if::placeholder{color:${G.dim}}
.dz{border:1.5px dashed ${G.border};border-radius:12px;cursor:pointer;transition:all .25s;background:${G.bgCard2};display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:48px 24px;text-align:center}
.dz:hover,.dz.ov{border-color:${G.gold}88;background:${G.gold}07}
.tg{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:4px;font-size:9px;font-family:'JetBrains Mono',monospace;letter-spacing:.08em;text-transform:uppercase;font-weight:500;background:${G.gold}15;color:${G.gold};border:1px solid ${G.gold}33}
.sl{font-size:10px;font-family:'JetBrains Mono',monospace;letter-spacing:.15em;text-transform:uppercase;color:${G.gold};font-weight:500}
.mb{position:fixed;inset:0;background:rgba(0,0,0,.88);backdrop-filter:blur(12px);z-index:400;display:flex;align-items:center;justify-content:center;padding:24px}
.dm{position:absolute;right:0;top:calc(100%+8px);width:210px;background:${G.bgCard};border:1px solid ${G.border};border-radius:10px;overflow:hidden;box-shadow:0 16px 48px rgba(0,0,0,.6);z-index:200}
.di{width:100%;background:none;border:none;padding:12px 16px;color:${G.gray};font-size:12px;font-family:'DM Sans',sans-serif;font-weight:500;cursor:pointer;display:flex;align-items:center;gap:10px;transition:all .15s;text-align:left}
.di:hover{background:${G.gold}0D;color:${G.goldL}}
.di.dg:hover{background:${G.err}15;color:${G.err}}
.ub{background:${G.ok}22;border:1px solid ${G.ok}33;border-radius:12px 12px 4px 12px}
.ab{background:${G.bgCard2};border:1px solid ${G.border};border-radius:12px 12px 12px 4px}
.tr:hover{background:${G.gold}08}
@keyframes pd{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(.85)}}
@keyframes sh{0%{transform:translateX(-100%)}100%{transform:translateX(100%)}}
@keyframes sp{to{transform:rotate(360deg)}}
.spin{animation:sp 1s linear infinite}
.pb{height:3px;background:${G.border};border-radius:2px;overflow:hidden;width:100%}
.pf{height:100%;background:linear-gradient(90deg,${G.goldD},${G.gold},${G.goldL});border-radius:2px;transition:width .5s ease;position:relative;overflow:hidden}
.pf::after{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,.3),transparent);animation:sh 1.5s infinite}
@media(max-width:768px){.hm{display:none!important}}
`;

// ─── Types ────────────────────────────────────────────────────────────────────
interface ExtractionResult {
  fileType: string; pages: number; rawText: string;
  tableData: any[]; charCount: number; rowCount: number; fileName: string;
}
interface ChatMsg { role: 'user' | 'model'; content: string; }
interface Toast { id: string; type: 'success' | 'error' | 'info'; msg: string; }
interface Profile { displayName: string; totalConversions: number; plan: string; createdAt: string; }

// ─── Auth ─────────────────────────────────────────────────────────────────────
interface AuthCtx {
  user: FirebaseUser | null; profile: Profile | null;
  loading: boolean; usageCount: number; guestName: string;
  setGuestName: (n: string) => void; incUsage: () => void;
}
const Ctx = createContext<AuthCtx>({
  user: null, profile: null, loading: true, usageCount: 0,
  guestName: '', setGuestName: () => {}, incUsage: () => {},
});
const useAuth = () => useContext(Ctx);

const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user,       setUser]       = useState<FirebaseUser | null>(null);
  const [profile,    setProfile]    = useState<Profile | null>(null);
  const [loading,    setLoading]    = useState(true);
  const [usageCount, setUsageCount] = useState<number>(() => parseInt(localStorage.getItem('img2xl_usage') || '0'));
  const [guestName,  setGuestNameS] = useState<string>(() => localStorage.getItem('img2xl_name') || '');

  const setGuestName = (n: string) => { setGuestNameS(n); localStorage.setItem('img2xl_name', n); };
  const incUsage     = () => { const n = usageCount + 1; setUsageCount(n); localStorage.setItem('img2xl_usage', String(n)); };

  const loadProfile = async (uid: string) => {
    try {
      const snap = await getDoc(doc(db, 'users', uid));
      if (snap.exists()) { setProfile(snap.data() as Profile); return; }
      const p: Profile = { displayName: auth.currentUser?.displayName || 'User', totalConversions: 0, plan: 'free', createdAt: new Date().toISOString() };
      await setDoc(doc(db, 'users', uid), p);
      setProfile(p);
    } catch (e) { console.error(e); }
  };

  useEffect(() => onAuthStateChanged(auth, async u => {
    setUser(u);
    if (u) await loadProfile(u.uid); else setProfile(null);
    setLoading(false);
  }), []);

  const val = useMemo(() => ({ user, profile, loading, usageCount, guestName, setGuestName, incUsage }), [user, profile, loading, usageCount, guestName]);
  return <Ctx.Provider value={val}>{children}</Ctx.Provider>;
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const Logo = ({ sm = false }: { sm?: boolean }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
    
    <img
      src={logo}
      alt="Img2XL Logo"
      style={{
        width: sm ? 28 : 34,
        height: sm ? 28 : 34,
        objectFit: "contain"
      }}
    />

    <div>
      <div style={{
        fontWeight: 700,
        fontSize: sm ? 14 : 17,
        color: '#fff'
      }}>
        IMG<span style={{ color: '#C9A84C' }}>2XL</span>
      </div>
    </div>
  </div>
);

const Toasts = ({ items, dismiss }: { items: Toast[]; dismiss: (id: string) => void }) => (
  <div style={{ position: 'fixed', top: 70, right: 20, zIndex: 500, display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 320 }}>
    <AnimatePresence>
      {items.map(t => (
        <motion.div key={t.id} initial={{ opacity: 0, x: 40 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 40 }}
          style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '11px 15px', borderRadius: 8, background: t.type === 'success' ? `${G.ok}15` : t.type === 'error' ? `${G.err}15` : `${G.gold}15`, border: `1px solid ${t.type === 'success' ? G.ok : t.type === 'error' ? G.err : G.gold}33` }}>
          {t.type === 'success' ? <CheckCircle2 size={13} color={G.ok} /> : t.type === 'error' ? <AlertCircle size={13} color={G.err} /> : <Sparkles size={13} color={G.gold} />}
          <span style={{ fontSize: 12, color: G.white, flex: 1 }}>{t.msg}</span>
          <button onClick={() => dismiss(t.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.gray }}><X size={11} /></button>
        </motion.div>
      ))}
    </AnimatePresence>
  </div>
);

// ─── Navbar ───────────────────────────────────────────────────────────────────
const Navbar = ({ tab, setTab, onLogin, onLogout }: { tab: string; setTab: (t: string) => void; onLogin: () => void; onLogout: () => void }) => {
  const { user, profile, guestName } = useAuth();
  const [sc, setSc] = useState(false);
  const [dd, setDd] = useState(false);

  useEffect(() => { const h = () => setSc(window.scrollY > 10); window.addEventListener('scroll', h); return () => window.removeEventListener('scroll', h); }, []);

  const displayName = user ? (profile?.displayName || user.email?.split('@')[0] || 'User') : guestName;

  return (
    <nav style={{ position: 'fixed', top: 0, left: 0, right: 0, zIndex: 100, background: sc ? `${G.bg}EE` : 'transparent', backdropFilter: sc ? 'blur(20px)' : 'none', borderBottom: sc ? `1px solid ${G.border}` : '1px solid transparent', transition: 'all .3s' }}>
      <div style={{ maxWidth: 1300, margin: '0 auto', padding: '0 24px', height: 58, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <button onClick={() => setTab('home')} style={{ background: 'none', border: 'none', cursor: 'pointer' }}><Logo /></button>

        <div className="hm" style={{ display: 'flex', alignItems: 'center', gap: 28 }}>
          {[{ id: 'home', l: 'Home' }, { id: 'convert', l: 'Convert' }, ...(user ? [{ id: 'history', l: 'History' }] : []), { id: 'about', l: 'About' }].map(l => (
            <button key={l.id} className={`nb ${tab === l.id ? 'on' : ''}`} onClick={() => setTab(l.id)} style={{ color: tab === l.id ? G.white : G.gray }}>{l.l}</button>
          ))}
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div className="hm" style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 9px', borderRadius: 4, background: `${G.ok}15`, border: `1px solid ${G.ok}33` }}>
            <div style={{ width: 5, height: 5, borderRadius: '50%', background: G.ok }} />
            <span className="fm" style={{ fontSize: 8, color: G.ok, letterSpacing: '0.08em' }}>ONLINE</span>
          </div>

          {(user || guestName) ? (
            <div style={{ position: 'relative' }}>
              <button onClick={() => setDd(d => !d)} style={{ display: 'flex', alignItems: 'center', gap: 8, background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 8, padding: '5px 11px 5px 5px', cursor: 'pointer' }}>
                <div style={{ width: 26, height: 26, borderRadius: 6, background: `linear-gradient(135deg,${G.goldD},${G.gold})`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, color: '#000' }}>
                  {displayName[0]?.toUpperCase() || 'U'}
                </div>
                <span className="hm" style={{ fontSize: 12, color: G.white, fontWeight: 500 }}>{displayName}</span>
                <ChevronDown size={11} color={G.gray} />
              </button>
              <AnimatePresence>
                {dd && (
                  <motion.div className="dm" initial={{ opacity: 0, y: -8, scale: .96 }} animate={{ opacity: 1, y: 0, scale: 1 }} exit={{ opacity: 0, y: -8, scale: .96 }} transition={{ duration: .15 }}>
                    <div style={{ padding: '11px 15px', borderBottom: `1px solid ${G.border}` }}>
                      <div style={{ fontSize: 12, color: G.white, fontWeight: 600 }}>{displayName}</div>
                      {user && <div style={{ fontSize: 11, color: G.gray, marginTop: 2 }}>{user.email}</div>}
                      {!user && <div style={{ fontSize: 10, color: G.gold, marginTop: 3 }}>Guest · Sign in for full access</div>}
                    </div>
                    {user && <>
                      <button className="di" onClick={() => { setTab('account'); setDd(false); }}><UserCircle2 size={13} /> Account</button>
                      <button className="di" onClick={() => { setTab('history'); setDd(false); }}><History size={13} /> History</button>
                      <div style={{ borderTop: `1px solid ${G.border}`, margin: '4px 0' }} />
                    </>}
                    <button className="di dg" onClick={() => { onLogout(); setDd(false); }}><LogOut size={13} /> {user ? 'Logout' : 'Switch Account'}</button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          ) : (
            <button className="bg" style={{ padding: '7px 18px' }} onClick={onLogin}>Sign In</button>
          )}
        </div>
      </div>
    </nav>
  );
};

// ─── Name Gate Modal (for guests) ─────────────────────────────────────────────
const NameGate = ({ onDone, requireLogin }: { onDone: (name: string) => void; requireLogin: boolean }) => {
  const [name, setName] = useState('');
  const [err,  setErr]  = useState('');

  const submit = () => {
    if (!name.trim()) { setErr('Please enter your name.'); return; }
    onDone(name.trim());
  };

  return (
    <div className="mb">
      <motion.div initial={{ opacity: 0, scale: .95, y: 20 }} animate={{ opacity: 1, scale: 1, y: 0 }}
        style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 16, padding: 36, width: '100%', maxWidth: 380 }}>
        <Logo />
        <div style={{ marginTop: 24, marginBottom: 24 }}>
          <h2 style={{ fontSize: 20, fontWeight: 700, color: G.white, marginBottom: 8 }}>
            {requireLogin ? 'Login Required' : 'What\'s your name?'}
          </h2>
          <p style={{ fontSize: 13, color: G.gray, lineHeight: 1.6 }}>
            {requireLogin
              ? 'You\'ve used 2 free conversions. Create an account to continue.'
              : 'Enter your name to get started. No account needed for first 2 uses.'}
          </p>
        </div>

        {!requireLogin && (
          <>
            <input className="if" value={name} onChange={e => setName(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && submit()}
              placeholder="Your name" autoFocus />
            {err && <p style={{ fontSize: 12, color: G.err, marginTop: 6 }}>{err}</p>}
            <button className="bg" style={{ width: '100%', justifyContent: 'center', padding: '12px 0', marginTop: 16 }} onClick={submit}>
              Continue
            </button>
            <div style={{ textAlign: 'center', marginTop: 14, fontSize: 12, color: G.dim }}>
              Already have an account? <span style={{ color: G.gold, cursor: 'pointer' }} onClick={() => onDone('__LOGIN__')}>Sign in</span>
            </div>
          </>
        )}

        {requireLogin && (
          <button className="bg" style={{ width: '100%', justifyContent: 'center', padding: '12px 0' }} onClick={() => onDone('__LOGIN__')}>
            Sign In / Create Account
          </button>
        )}
      </motion.div>
    </div>
  );
};

// ─── Auth Modal ───────────────────────────────────────────────────────────────
const AuthModal = ({ open, onClose, onSuccess }: { open: boolean; onClose: () => void; onSuccess: () => void }) => {
  const [mode,   setMode]   = useState<'login' | 'signup'>('login');
  const [email,  setEmail]  = useState('');
  const [pw,     setPw]     = useState('');
  const [name,   setName]   = useState('');
  const [busy,   setBusy]   = useState(false);
  const [err,    setErr]    = useState<string | null>(null);
  const [showPw, setShowPw] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault(); setBusy(true); setErr(null);
    try {
      if (mode === 'signup') {
        const c = await createUserWithEmailAndPassword(auth, email, pw);
        await updateProfile(c.user, { displayName: name });
      } else {
        await signInWithEmailAndPassword(auth, email, pw);
      }
      onClose(); onSuccess();
    } catch (e: any) { setErr(e.message); }
    finally { setBusy(false); }
  };

  const google = async () => {
    try { await signInWithPopup(auth, new GoogleAuthProvider()); onClose(); onSuccess(); }
    catch (e: any) { setErr(e.message); }
  };

  if (!open) return null;
  return (
    <div className="mb" onClick={onClose}>
      <motion.div initial={{ opacity: 0, scale: .95, y: 20 }} animate={{ opacity: 1, scale: 1, y: 0 }}
        onClick={e => e.stopPropagation()}
        style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 16, padding: 36, width: '100%', maxWidth: 400, position: 'relative' }}>
        <button onClick={onClose} style={{ position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', cursor: 'pointer', color: G.gray }}><X size={16} /></button>
        <div style={{ marginBottom: 28 }}>
          <Logo />
          <h2 style={{ fontSize: 18, fontWeight: 700, color: G.white, marginTop: 18 }}>{mode === 'login' ? 'Welcome back' : 'Create account'}</h2>
          <p style={{ fontSize: 12, color: G.gray, marginTop: 5 }}>{mode === 'login' ? 'Sign in to continue' : 'Free forever — start converting'}</p>
        </div>

        {err && <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', background: `${G.err}12`, border: `1px solid ${G.err}33`, borderRadius: 8, marginBottom: 16 }}><AlertCircle size={13} color={G.err} /><span style={{ fontSize: 12, color: G.err }}>{err}</span></div>}

        <form onSubmit={submit} style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {mode === 'signup' && <input className="if" required value={name} onChange={e => setName(e.target.value)} placeholder="Full name" />}
          <input className="if" type="email" required value={email} onChange={e => setEmail(e.target.value)} placeholder="Email" />
          <div style={{ position: 'relative' }}>
            <input className="if" type={showPw ? 'text' : 'password'} required value={pw} onChange={e => setPw(e.target.value)} placeholder="Password" style={{ paddingRight: 40 }} />
            <button type="button" onClick={() => setShowPw(p => !p)} style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: G.gray }}>
              {showPw ? <EyeOff size={14} /> : <Eye size={14} />}
            </button>
          </div>
          <button className="bg" style={{ padding: '12px 0', justifyContent: 'center', marginTop: 4 }} disabled={busy} type="submit">
            {busy ? <Loader2 size={14} className="spin" /> : (mode === 'login' ? 'Sign In' : 'Create Account')}
          </button>
        </form>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10, margin: '18px 0' }}>
          <div style={{ flex: 1, height: 1, background: G.border }} /><span style={{ fontSize: 11, color: G.dim }}>or</span><div style={{ flex: 1, height: 1, background: G.border }} />
        </div>

        <button className="gh" style={{ width: '100%', justifyContent: 'center', padding: '11px 0' }} onClick={google}>
          <svg width="13" height="13" viewBox="0 0 24 24"><path fill={G.gray} d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" /><path fill={G.gray} d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" /><path fill={G.gray} d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z" /><path fill={G.gray} d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" /></svg>
          Continue with Google
        </button>

        <button onClick={() => setMode(m => m === 'login' ? 'signup' : 'login')}
          style={{ display: 'block', width: '100%', textAlign: 'center', marginTop: 16, background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: G.gold }}>
          {mode === 'login' ? "Don't have an account? Sign up" : 'Already have an account? Sign in'}
        </button>
      </motion.div>
    </div>
  );
};

// ─── Section 1: Extracted Text ────────────────────────────────────────────────
// ── Auto Summary Component ────────────────────────────────────────────────────
const AutoSummary = ({ result }: { result: ExtractionResult }) => {
  const [summary, setSummary] = useState('');
  const [busy, setBusy]       = useState(false);
  const [done, setDone]       = useState(false);

  const generate = async () => {
    setBusy(true);
    try {
      const r = await fetch('/api/chat-doc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: [{ role: 'user', content: 'Give me a clear, structured summary of this document in bullet points. Include: main topics, key numbers/data, important names, and any conclusions.' }],
          context: result.rawText,
        }),
      });
      const j = await r.json();
      setSummary(j.content || 'No summary generated.');
      setDone(true);
    } catch (e: any) { setSummary('Error: ' + e.message); }
    finally { setBusy(false); }
  };

  return (
    <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, overflow: 'hidden' }}>
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: G.bgCard2 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 18 }}>✨</span>
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: G.white }}>AI Auto Summary</div>
            <div style={{ fontSize: 10, color: G.gray }}>One-click document overview</div>
          </div>
        </div>
        {!done && (
          <button className="bg" style={{ padding: '7px 16px', fontSize: 10 }} onClick={generate} disabled={busy}>
            {busy ? <Loader2 size={12} className="spin" /> : <Sparkles size={12} />}
            {busy ? 'Generating…' : 'Generate Summary'}
          </button>
        )}
        {done && <button className="gh" style={{ padding: '6px 12px', fontSize: 10 }} onClick={() => { setDone(false); setSummary(''); }}>Regenerate</button>}
      </div>
      {!done && !busy && (
        <div style={{ padding: '24px 20px', textAlign: 'center', color: G.dim }}>
          <p style={{ fontSize: 13 }}>Click "Generate Summary" to get an AI overview of your document.</p>
        </div>
      )}
      {busy && (
        <div style={{ padding: '24px 20px', textAlign: 'center' }}>
          <Loader2 size={24} color={G.gold} className="spin" style={{ margin: '0 auto 12px' }} />
          <p style={{ fontSize: 13, color: G.gray }}>Reading your document…</p>
        </div>
      )}
      {done && summary && (
        <div style={{ padding: 20, maxHeight: 300, overflowY: 'auto' }}>
          <pre style={{ fontSize: 13, color: G.gray, lineHeight: 1.8, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontFamily: "'DM Sans',sans-serif" }}>
            {summary}
          </pre>
        </div>
      )}
    </div>
  );
};

// ── Stats Card ────────────────────────────────────────────────────────────────
const StatsCard = ({ result }: { result: ExtractionResult }) => {
  const words = result.rawText.trim() ? result.rawText.trim().split(/\s+/).length : 0;
  const readMins = Math.ceil(words / 200);
  const stats = [
    { label: 'Pages',    value: result.pages,                   icon: '📄' },
    { label: 'Words',    value: words.toLocaleString(),         icon: '🔤' },
    { label: 'Rows',     value: result.rowCount.toLocaleString(), icon: '📊' },
    { label: 'Read Time',value: readMins + ' min',              icon: '⏱️' },
  ];
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 12 }}>
      {stats.map(s => (
        <div key={s.label} style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 10, padding: '14px 16px', textAlign: 'center' }}>
          <div style={{ fontSize: 20, marginBottom: 6 }}>{s.icon}</div>
          <div style={{ fontSize: 18, fontWeight: 700, color: G.gold }}>{s.value}</div>
          <div style={{ fontSize: 10, color: G.gray, marginTop: 3, fontFamily: "'JetBrains Mono',monospace", textTransform: 'uppercase', letterSpacing: '0.08em' }}>{s.label}</div>
        </div>
      ))}
    </div>
  );
};

// ── Batch Multi-file Processor ────────────────────────────────────────────────
const BatchStatus = ({ files, onClear }: { files: Array<{ name: string; status: string; rows: number }>; onClear: () => void }) => (
  <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, overflow: 'hidden' }}>
    <div style={{ padding: '12px 18px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: G.bgCard2 }}>
      <span style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Batch Processing — {files.length} files</span>
      <button className="gh" style={{ padding: '5px 10px', fontSize: 10 }} onClick={onClear}>Clear</button>
    </div>
    <div style={{ maxHeight: 200, overflowY: 'auto' }}>
      {files.map((f, i) => (
        <div key={i} style={{ padding: '10px 18px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
          <span style={{ fontSize: 12, color: G.white, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</span>
          <span className="tg" style={{ background: f.status === 'done' ? `${G.ok}15` : f.status === 'error' ? `${G.err}15` : `${G.gold}15`, color: f.status === 'done' ? G.ok : f.status === 'error' ? G.err : G.gold, borderColor: f.status === 'done' ? G.ok : f.status === 'error' ? G.err : G.gold }}>{f.status}</span>
          {f.rows > 0 && <span style={{ fontSize: 10, color: G.gray, fontFamily: "'JetBrains Mono',monospace" }}>{f.rows} rows</span>}
        </div>
      ))}
    </div>
  </div>
);

const ExtractedTextSection = ({ result }: { result: ExtractionResult }) => {
  const [copied, setCopied] = useState(false);
  const [search, setSearch] = useState('');

  const copy = () => {
    navigator.clipboard.writeText(result.rawText);
    setCopied(true); setTimeout(() => setCopied(false), 2000);
  };

  const FILE_ICONS: Record<string, string> = {
    pdf: '📄', excel: '📊', image: '🖼️', docx: '📝', csv: '📋',
    txt: '🔤', json: '🗂️', pptx: '📑', tsv: '📋',
  };

  return (
    <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: G.bgCard2 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <span style={{ fontSize: 20 }}>{FILE_ICONS[result.fileType] || '📄'}</span>
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: G.white }}>{result.fileName}</div>
            <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
              <span className="tg">{result.fileType}</span>
              <span className="tg">{result.pages} page{result.pages !== 1 ? 's' : ''}</span>
              <span className="tg">{result.charCount.toLocaleString()} chars</span>
            </div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <input className="if" value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Search in text…" style={{ padding: '6px 12px', fontSize: 11, width: 160 }} />
          <button className="gh" style={{ padding: '6px 12px', fontSize: 10 }} onClick={copy}>
            {copied ? <CheckCircle2 size={12} color={G.ok} /> : <Copy size={12} />}
            {copied ? 'Copied' : 'Copy'}
          </button>
        </div>
      </div>

      {/* Text content */}
      <div style={{ padding: 20, maxHeight: 400, overflowY: 'auto' }}>
        {result.rawText.trim() ? (
          <pre style={{ fontSize: 12, color: G.gray, lineHeight: 1.7, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontFamily: "'JetBrains Mono',monospace" }}>
            {search.trim()
              ? result.rawText.split(new RegExp('(' + search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi')).map((part, i) =>
                  part.toLowerCase() === search.toLowerCase()
                    ? <mark key={i} style={{ background: G.gold + '44', color: G.white, borderRadius: 2 }}>{part}</mark>
                    : part
                )
              : result.rawText}
          </pre>
        ) : (
          <div style={{ textAlign: 'center', padding: 40, color: G.dim }}>
            <FileText size={32} style={{ margin: '0 auto 12px' }} />
            <p style={{ fontSize: 13 }}>No plain text extracted from this file.</p>
          </div>
        )}
      </div>
    </div>
  );
};

// ─── Section 2: Document Chatbot ──────────────────────────────────────────────
const DocChatSection = ({ result }: { result: ExtractionResult }) => {
  const [msgs,  setMsgs]  = useState<ChatMsg[]>([{
    role: 'model',
    content: `I've read your document "${result.fileName}" (${result.pages} page${result.pages !== 1 ? 's' : ''}, ${result.charCount.toLocaleString()} characters). Ask me anything about it!`,
  }]);
  const [input, setInput] = useState('');
  const [busy,  setBusy]  = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => { if (ref.current) ref.current.scrollTop = ref.current.scrollHeight; }, [msgs, busy]);

  const quick = [
    'Summarize this document',
    'What are the key data points?',
    'List all tables found',
    'What is the total/sum?',
  ];

  const send = async (text: string) => {
    if (!text.trim() || busy) return;
    const um: ChatMsg = { role: 'user', content: text };
    setMsgs(p => [...p, um]); setInput(''); setBusy(true);
    try {
      const r = await fetch('/api/chat-doc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: [...msgs, um].map(m => ({ role: m.role === 'model' ? 'assistant' : m.role, content: m.content })),
          context: result.rawText,
        }),
      });
      const j = await r.json();
      if (!r.ok) throw new Error(j.error);
      setMsgs(p => [...p, { role: 'model', content: j.content }]);
    } catch (e: any) {
      setMsgs(p => [...p, { role: 'model', content: `Error: ${e.message}` }]);
    }
    finally { setBusy(false); }
  };

  return (
    <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: G.bgCard2 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 30, height: 30, borderRadius: 8, background: `linear-gradient(135deg,${G.goldD},${G.gold})`, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <MessageSquare size={14} color="#000" />
          </div>
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Ask AI About This Document</div>
            <div style={{ fontSize: 10, color: G.gold, fontFamily: "'JetBrains Mono',monospace" }}>Full document context loaded</div>
          </div>
        </div>
        <button onClick={() => setMsgs([{ role: 'model', content: 'Chat cleared. What would you like to know?' }])}
          style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.gray }}><RefreshCw size={13} /></button>
      </div>

      {/* Messages */}
      <div ref={ref} style={{ height: 320, overflowY: 'auto', padding: '16px', display: 'flex', flexDirection: 'column', gap: 10 }}>
        {msgs.map((m, i) => (
          <div key={i} style={{ display: 'flex', justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
            <div className={m.role === 'user' ? 'ub' : 'ab'} style={{ maxWidth: '85%', padding: '10px 14px', fontSize: 13, lineHeight: 1.6, color: G.white }}>
              {m.content}
            </div>
          </div>
        ))}
        {busy && (
          <div style={{ display: 'flex', gap: 4, padding: '10px 14px' }}>
            {[0, 1, 2].map(i => <div key={i} style={{ width: 6, height: 6, borderRadius: '50%', background: G.gold, animation: `pd 1s ${i * .2}s infinite` }} />)}
          </div>
        )}
      </div>

      {/* Quick replies */}
      <div style={{ padding: '0 14px 10px', display: 'flex', flexWrap: 'wrap', gap: 6 }}>
        {quick.map(q => (
          <button key={q} onClick={() => send(q)} style={{ background: `${G.gold}10`, border: `1px solid ${G.gold}33`, borderRadius: 6, padding: '5px 10px', fontSize: 10, color: G.gold, cursor: 'pointer', fontFamily: "'DM Sans',sans-serif" }}>
            {q}
          </button>
        ))}
      </div>

      {/* Input */}
      <div style={{ padding: '10px 14px', borderTop: `1px solid ${G.border}`, display: 'flex', gap: 8 }}>
        <input className="if" value={input} onChange={e => setInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && send(input)}
          placeholder="Ask anything about your document…"
          style={{ flex: 1, padding: '9px 13px', fontSize: 13 }} />
        <button className="bg" style={{ padding: '9px 13px', flexShrink: 0 }} onClick={() => send(input)} disabled={busy}>
          <Send size={13} />
        </button>
      </div>
    </div>
  );
};

// ─── Section 3: Table Preview + Download ──────────────────────────────────────
const TableDownloadSection = ({ result }: { result: ExtractionResult }) => {
  const [format,    setFormat]    = useState<'xlsx' | 'csv' | 'json' | 'txt'>('xlsx');
  const [editH,     setEditH]     = useState<string | null>(null);
  const [draft,     setDraft]     = useState('');
  const [tableData, setTableData] = useState<any[]>(result.tableData);
  const [selCols,   setSelCols]   = useState<string[]>([]);
  const [busy,      setBusy]      = useState(false);

  const headers = tableData.length > 0 ? Object.keys(tableData[0]) : [];

  const toggleCol = (c: string) => setSelCols(p => p.includes(c) ? p.filter(x => x !== c) : [...p, c]);

  const merge = () => {
    if (selCols.length < 2) return;
    const key = selCols.join(' + ');
    const data = tableData.map(row => {
      const m = selCols.map(c => row[c]).filter(Boolean).join(' ');
      const r = { ...row }; selCols.forEach(c => delete r[c]);
      return { ...r, [key]: m };
    });
    setTableData(data); setSelCols([]);
  };

  const rename = (old: string) => {
    if (!draft.trim()) { setEditH(null); return; }
    const data = tableData.map(row => { const r = { ...row }; r[draft.trim()] = r[old]; delete r[old]; return r; });
    setTableData(data); setEditH(null);
  };

  const download = async () => {
    setBusy(true);
    try {
      const r = await fetch('/api/download', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: tableData, format, filename: result.fileName.replace(/\.[^.]+$/, ''), rawText: result.rawText }),
      });
      if (!r.ok) throw new Error('Download failed');
      const blob = await r.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = `${result.fileName.replace(/\.[^.]+$/, '')}_export.${format}`;
      a.click(); URL.revokeObjectURL(url);
    } catch (e: any) { alert(e.message); }
    finally { setBusy(false); }
  };

  const formats: { id: 'xlsx' | 'csv' | 'json' | 'txt'; label: string; icon: string }[] = [
    { id: 'xlsx', label: 'Excel .xlsx', icon: '📊' },
    { id: 'csv',  label: 'CSV',         icon: '📋' },
    { id: 'json', label: 'JSON',        icon: '🗂️' },
    { id: 'txt',  label: 'Plain Text',  icon: '🔤' },
  ];

  return (
    <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, overflow: 'hidden' }}>
      {/* Header */}
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${G.border}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10, background: G.bgCard2 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <FileSpreadsheet size={18} color={G.gold} />
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Extracted Table Data</div>
            <div className="fm" style={{ fontSize: 10, color: G.gray, marginTop: 2 }}>{tableData.length} rows · {headers.length} columns</div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          {selCols.length >= 2 && (
            <button className="gh" style={{ padding: '6px 10px', fontSize: 10 }} onClick={merge}>Merge Cols</button>
          )}
        </div>
      </div>

      {tableData.length === 0 ? (
        <div style={{ padding: 40, textAlign: 'center', color: G.dim }}>
          <Table size={32} style={{ margin: '0 auto 12px' }} />
          <p style={{ fontSize: 13 }}>No structured table data found. Use the chat to ask about the content.</p>
        </div>
      ) : (
        <div style={{ overflowX: 'auto', maxHeight: 320 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>
                {headers.map((h, i) => (
                  <th key={i} style={{ padding: '9px 12px', textAlign: 'left', background: G.bgCard2, borderBottom: `1px solid ${G.border}`, position: 'sticky', top: 0 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                      <input type="checkbox" checked={selCols.includes(h)} onChange={() => toggleCol(h)} style={{ width: 11, height: 11, accentColor: G.gold, cursor: 'pointer' }} />
                      {editH === h ? (
                        <div style={{ display: 'flex', gap: 4 }}>
                          <input autoFocus value={draft} onChange={e => setDraft(e.target.value)}
                            onKeyDown={e => { if (e.key === 'Enter') rename(h); if (e.key === 'Escape') setEditH(null); }}
                            style={{ background: G.bgCard, border: `1px solid ${G.gold}66`, borderRadius: 4, padding: '2px 6px', color: G.white, fontSize: 11, outline: 'none', width: 90 }} />
                          <button onClick={() => rename(h)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.ok }}><CheckCircle2 size={11} /></button>
                          <button onClick={() => setEditH(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.err }}><X size={11} /></button>
                        </div>
                      ) : (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                          <span className="fm" style={{ fontSize: 9, color: G.gold, letterSpacing: '0.08em', textTransform: 'uppercase' }}>{h}</span>
                          <button onClick={() => { setEditH(h); setDraft(h); }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.dim, padding: 1 }}><Edit2 size={9} /></button>
                        </div>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tableData.slice(0, 100).map((row, i) => (
                <tr key={i} className="tr" style={{ borderBottom: `1px solid ${G.border}` }}>
                  {headers.map((h, j) => (
                    <td key={j} style={{ padding: '8px 12px', color: G.gray, maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{String(row[h] ?? '')}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          {tableData.length > 100 && (
            <div className="fm" style={{ padding: '10px 16px', fontSize: 10, color: G.dim, borderTop: `1px solid ${G.border}`, background: G.bgCard2 }}>
              Showing 100 of {tableData.length} rows — all rows included in download
            </div>
          )}
        </div>
      )}

      {/* Download bar */}
      <div style={{ padding: '16px 20px', borderTop: `1px solid ${G.border}`, background: G.bgCard2 }}>
        <div style={{ fontSize: 11, color: G.gray, marginBottom: 12, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.08em' }}>Export Format</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 14 }}>
          {formats.map(f => (
            <button key={f.id} onClick={() => setFormat(f.id)} style={{
              padding: '8px 14px', borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: 'pointer', transition: 'all .2s',
              background: format === f.id ? `${G.gold}20` : G.bgCard,
              border: format === f.id ? `1px solid ${G.gold}66` : `1px solid ${G.border}`,
              color: format === f.id ? G.gold : G.gray,
            }}>
              {f.icon} {f.label}
            </button>
          ))}
        </div>
        <button className="bg" style={{ padding: '11px 28px', fontSize: 12 }} onClick={download} disabled={busy}>
          {busy ? <Loader2 size={14} className="spin" /> : <Download size={14} />}
          {busy ? 'Preparing…' : `Download as ${format.toUpperCase()}`}
        </button>
      </div>
    </div>
  );
};

// ─── Main Convert Page ────────────────────────────────────────────────────────
const ConvertPage = ({ addToast, onNeedAuth }: { addToast: (t: 'success' | 'error' | 'info', m: string) => void; onNeedAuth: () => void }) => {
  const { user, guestName, usageCount, incUsage } = useAuth();
  const [drag,     setDrag]     = useState(false);
  const [result,   setResult]   = useState<ExtractionResult | null>(null);
  const [busy,     setBusy]     = useState(false);
  const [progress, setProgress] = useState(0);
  const [step,     setStep]     = useState('');

  const canUse = user || (guestName && usageCount < 2);

  const process = async (file: File) => {
    if (!canUse) { onNeedAuth(); return; }
    setBusy(true); setResult(null); setProgress(10);
    setStep('Reading file…');

    try {
      const b64 = await new Promise<string>(res => {
        const r = new FileReader();
        r.onload = e => res((e.target?.result as string).split(',')[1]);
        r.readAsDataURL(file);
      });

      setProgress(30); setStep('Sending to AI…');

      const resp = await fetch('/api/extract-full', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ base64: b64, mimeType: file.type, fileName: file.name }),
      });

      setProgress(80); setStep('Processing…');
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error);

      setProgress(100); setStep('Done!');
      setResult({ ...data, fileName: file.name });

      if (!user) incUsage();

      // Save to Firebase if logged in
      if (user) {
        await addDoc(collection(db, 'history'), {
          fileName: file.name, fileType: data.fileType, pages: data.pages,
          rowCount: data.rowCount, charCount: data.charCount,
          timestamp: serverTimestamp(), userId: user.uid,
        });
      }

      addToast('success', `${file.name} extracted — ${data.rowCount} rows!`);
    } catch (e: any) {
      addToast('error', e.message);
      setStep('');
    } finally {
      setBusy(false);
      setTimeout(() => setProgress(0), 1000);
    }
  };

  const onDrop = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    process(files[0]);
  };

  const ACCEPT = "image/*,.pdf,.docx,.xlsx,.xls,.csv,.tsv,.txt,.json,.pptx";

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto', padding: '80px 24px 60px' }}>
      <div style={{ marginBottom: 32 }}>
        <p className="sl" style={{ marginBottom: 8 }}>Convert</p>
        <h2 className="fd" style={{ fontSize: 34, fontWeight: 700, color: G.white, marginBottom: 8 }}>Upload Any Document</h2>
        <p style={{ fontSize: 14, color: G.gray }}>Images, PDF (any size), Word, Excel, CSV, PowerPoint, TXT, JSON → extracted data</p>
      </div>

      {/* Usage indicator for guests */}
      {!user && guestName && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 16px', background: usageCount >= 2 ? `${G.err}12` : `${G.gold}12`, border: `1px solid ${usageCount >= 2 ? G.err : G.gold}33`, borderRadius: 8, marginBottom: 20 }}>
          <User size={13} color={usageCount >= 2 ? G.err : G.gold} />
          <span style={{ fontSize: 12, color: usageCount >= 2 ? G.err : G.gold }}>
            {usageCount >= 2 ? 'Free uses exhausted — please sign in to continue.' : `${guestName} · ${2 - usageCount} free use${2 - usageCount !== 1 ? 's' : ''} remaining`}
          </span>
          {usageCount >= 2 && <button className="bg" style={{ padding: '4px 12px', marginLeft: 'auto' }} onClick={onNeedAuth}>Sign In</button>}
        </div>
      )}

      {/* Drop zone */}
      {!busy && !result && (
        <div className={`dz${drag ? ' ov' : ''}`} style={{ minHeight: 260 }}
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); onDrop(e.dataTransfer.files); }}
          onClick={() => document.getElementById('fi-main')?.click()}>
          <div style={{ width: 52, height: 52, borderRadius: 12, background: `${G.gold}15`, border: `1px solid ${G.gold}33`, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Upload size={22} color={G.gold} />
          </div>
          <div>
            <p style={{ fontSize: 15, fontWeight: 600, color: G.white, marginBottom: 6 }}>Drop any file or click to upload</p>
            <p style={{ fontSize: 12, color: G.gray }}>Images · PDF (1000+ pages) · Word · Excel · CSV · PowerPoint · TXT · JSON</p>
          </div>
          <input id="fi-main" type="file" accept={ACCEPT} style={{ display: 'none' }} onChange={e => onDrop(e.target.files)} />
        </div>
      )}

      {/* Progress */}
      {busy && (
        <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, padding: 40, textAlign: 'center' }}>
          <Loader2 size={32} color={G.gold} className="spin" style={{ margin: '0 auto 20px' }} />
          <p style={{ fontSize: 15, fontWeight: 600, color: G.white, marginBottom: 6 }}>{step}</p>
          <p style={{ fontSize: 12, color: G.gray, marginBottom: 20 }}>Large files may take a moment — chunking and processing…</p>
          <div className="pb" style={{ maxWidth: 400, margin: '0 auto' }}>
            <div className="pf" style={{ width: `${progress}%` }} />
          </div>
          <div className="fm" style={{ fontSize: 10, color: G.dim, marginTop: 8 }}>{progress}%</div>
        </motion.div>
      )}

      {/* Result — 3 sections */}
      {result && !busy && (
        <motion.div initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }} style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
          {/* Upload another */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
            <div className="tg"><CheckCircle2 size={10} /> Extraction complete</div>
            <label className="gh" style={{ padding: '8px 16px', cursor: 'pointer' }}>
              <Upload size={12} /> Upload Another
              <input type="file" accept={ACCEPT} style={{ display: 'none' }} onChange={e => { setResult(null); onDrop(e.target.files); }} />
            </label>
          </div>

          {/* Section 1 — Raw text */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
              <div style={{ width: 24, height: 24, borderRadius: 6, background: G.gold, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, color: '#000' }}>1</div>
              <span style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Extracted Text</span>
            </div>
            <ExtractedTextSection result={result} />
          </div>

          {/* Section 2 — Chat */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
              <div style={{ width: 24, height: 24, borderRadius: 6, background: G.gold, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, color: '#000' }}>2</div>
              <span style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Chat With Your Document</span>
            </div>
            <DocChatSection result={result} />
          </div>

          {/* Section 3 — Download */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
              <div style={{ width: 24, height: 24, borderRadius: 6, background: G.gold, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, color: '#000' }}>3</div>
              <span style={{ fontSize: 13, fontWeight: 600, color: G.white }}>Download as Excel / CSV / JSON / TXT</span>
            </div>
            <TableDownloadSection result={result} />
          </div>
        </motion.div>
      )}
    </div>
  );
};

// ─── History Page ─────────────────────────────────────────────────────────────
const HistoryPage = ({ goConvert }: { goConvert: () => void }) => {
  const { user } = useAuth();
  const [hist,    setHist]    = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  const load = async () => {
    if (!user) return; setLoading(true);
    try {
      const q    = query(collection(db, 'history'), where('userId', '==', user.uid), orderBy('timestamp', 'desc'));
      const snap = await getDocs(q);
      setHist(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    } catch (e) { console.error(e); }
    finally { setLoading(false); }
  };

  useEffect(() => { load(); }, [user]);

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto', padding: '80px 24px 60px' }}>
      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', marginBottom: 32, flexWrap: 'wrap', gap: 16 }}>
        <div><p className="sl" style={{ marginBottom: 8 }}>History</p><h2 className="fd" style={{ fontSize: 34, fontWeight: 700, color: G.white }}>Past Conversions</h2></div>
        <button className="gh" style={{ padding: '10px 14px' }} onClick={load}><RefreshCw size={13} /> Refresh</button>
      </div>

      {loading ? (
        <div style={{ display: 'flex', justifyContent: 'center', padding: 80 }}><Loader2 size={24} color={G.gold} className="spin" /></div>
      ) : hist.length === 0 ? (
        <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 12, padding: 56, textAlign: 'center' }}>
          <History size={32} color={G.dim} style={{ margin: '0 auto 16px' }} />
          <h3 style={{ fontSize: 16, fontWeight: 600, color: G.white, marginBottom: 8 }}>No conversions yet</h3>
          <button className="bg" style={{ padding: '10px 24px', marginTop: 16 }} onClick={goConvert}><Upload size={13} /> Start Converting</button>
        </div>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(250px,1fr))', gap: 14 }}>
          {hist.map(h => (
            <div key={h.id} style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 10, padding: 20, transition: 'border-color .2s' }}
              onMouseEnter={e => (e.currentTarget.style.borderColor = `${G.gold}44`)}
              onMouseLeave={e => (e.currentTarget.style.borderColor = G.border)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 12 }}>
                <FileSpreadsheet size={18} color={G.gold} />
                <button onClick={async () => { await deleteDoc(doc(db, 'history', h.id)); setHist(p => p.filter(x => x.id !== h.id)); }}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: G.dim }}><Trash2 size={13} /></button>
              </div>
              <p style={{ fontSize: 12, fontWeight: 600, color: G.white, marginBottom: 4, wordBreak: 'break-all' }}>{h.fileName}</p>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 8 }}>
                <span className="tg">{h.fileType}</span>
                <span className="tg">{h.rowCount || 0} rows</span>
              </div>
              <div className="fm" style={{ fontSize: 9, color: G.dim }}>{h.timestamp?.toDate?.().toLocaleString()}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

// ─── Landing Page ─────────────────────────────────────────────────────────────
const Landing = ({ onStart }: { onStart: () => void }) => (
  <div style={{ paddingTop: 60 }}>
    <section style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', padding: '0 24px', maxWidth: 1280, margin: '0 auto', gap: 80, flexWrap: 'wrap' }}>
      <motion.div initial={{ opacity: 0, x: -30 }} animate={{ opacity: 1, x: 0 }} transition={{ duration: .7 }} style={{ flex: 1, minWidth: 280 }}>
        <div className="tg" style={{ marginBottom: 28 }}><Sparkles size={10} /> AI · ANY FILE · INSTANT</div>
        <h1 className="fd" style={{ fontSize: 'clamp(44px,6vw,80px)', lineHeight: .92, fontWeight: 900, color: G.white, marginBottom: 24 }}>
          ANY FILE.<br /><em className="gt" style={{ fontStyle: 'italic' }}>EXTRACT.</em><br />EXCEL READY.
        </h1>
        <p style={{ fontSize: 15, color: G.gray, lineHeight: 1.7, maxWidth: 420, marginBottom: 36 }}>
          Upload a 1000-page PDF, a scanned receipt, an Excel sheet, a PowerPoint deck — and get back clean structured data plus an AI that knows your document.
        </p>

        {/* Supported types grid */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 8, maxWidth: 380, marginBottom: 36 }}>
          {[['📄', 'PDF'], ['🖼️', 'Images'], ['📝', 'Word'], ['📊', 'Excel'], ['📋', 'CSV'], ['📑', 'PPT'], ['🔤', 'TXT'], ['🗂️', 'JSON']].map(([icon, label]) => (
            <div key={label} style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 8, padding: '8px 6px', textAlign: 'center' }}>
              <div style={{ fontSize: 18, marginBottom: 3 }}>{icon}</div>
              <div style={{ fontSize: 9, color: G.gray, fontFamily: "'JetBrains Mono',monospace", textTransform: 'uppercase' }}>{label}</div>
            </div>
          ))}
        </div>

        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
          <button className="bg" style={{ padding: '14px 28px', fontSize: 12 }} onClick={onStart}><Upload size={14} /> Start Converting</button>
        </div>
        <p style={{ marginTop: 16, fontSize: 12, color: G.dim }}><ShieldCheck size={11} color={G.dim} style={{ display: 'inline', marginRight: 4 }} />First 2 conversions free · No credit card needed</p>
      </motion.div>

      {/* Feature cards */}
      <motion.div className="hm" initial={{ opacity: 0, x: 30 }} animate={{ opacity: 1, x: 0 }} transition={{ duration: .7, delay: .15 }} style={{ flex: 1, maxWidth: 420, display: 'flex', flexDirection: 'column', gap: 14 }}>
        {[
          { icon: '📄', title: 'PDF — Any Size', desc: 'Even 1000-page PDFs — we chunk and process every page automatically.' },
          { icon: '💬', title: 'Chat With Your Doc', desc: 'Ask the AI anything about the document. Summarize, find data, analyze.' },
          { icon: '📥', title: 'Your Format', desc: 'Download as Excel, CSV, JSON, or plain text — your choice.' },
        ].map((c, i) => (
          <motion.div key={i} initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: .3 + i * .1 }}
            style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 10, padding: '18px 20px', display: 'flex', gap: 14, alignItems: 'flex-start' }}>
            <span style={{ fontSize: 24, marginTop: 2 }}>{c.icon}</span>
            <div>
              <div style={{ fontSize: 14, fontWeight: 600, color: G.white, marginBottom: 4 }}>{c.title}</div>
              <div style={{ fontSize: 12, color: G.gray, lineHeight: 1.6 }}>{c.desc}</div>
            </div>
          </motion.div>
        ))}
      </motion.div>
    </section>
  </div>
);

// ─── Root App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [tab,       setTab]       = useState('home');
  const [showAuth,  setShowAuth]  = useState(false);
  const [showGate,  setShowGate]  = useState(false);
  const [toasts,    setToasts]    = useState<Toast[]>([]);

  const addToast = (type: 'success' | 'error' | 'info', msg: string) => {
    const id = Math.random().toString(36).substr(2, 9);
    setToasts(p => [...p, { id, type, msg }]);
    setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), 4500);
  };

  const logout = async () => {
    await signOut(auth); setTab('home');
    localStorage.removeItem('img2xl_usage');
    localStorage.removeItem('img2xl_name');
    addToast('info', 'Signed out.');
  };

  const handleNeedAuth = () => {
    setShowGate(true);
  };

  const handleGateDone = (name: string) => {
    if (name === '__LOGIN__') { setShowGate(false); setShowAuth(true); }
    else { /* AuthProvider setGuestName will be called via context */ setShowGate(false); }
  };

  const handleConvertClick = () => {
    setTab('convert');
  };

  return (
    <AuthProvider>
      <style>{CSS}</style>
      <div style={{ minHeight: '100vh', background: G.bg, color: G.white, overflowX: 'hidden' }}>
        <Navbar tab={tab} setTab={setTab} onLogin={() => setShowAuth(true)} onLogout={logout} />

        <AnimatePresence mode="wait">
          {tab === 'home' && (
            <motion.div key="home" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={{ duration: .25 }}>
              <Landing onStart={handleConvertClick} />
            </motion.div>
          )}
          {tab === 'convert' && (
            <motion.div key="conv" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }} transition={{ duration: .25 }}>
              <InnerConvert addToast={addToast} onNeedAuth={handleNeedAuth} />
            </motion.div>
          )}
          {tab === 'history' && (
            <motion.div key="hist" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }} transition={{ duration: .25 }}>
              <HistoryPage goConvert={() => setTab('convert')} />
            </motion.div>
          )}
          {tab === 'about' && (
            <motion.div key="about" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={{ duration: .25 }}>
              <div style={{ maxWidth: 800, margin: '0 auto', padding: '80px 24px' }}>
                <p className="sl" style={{ marginBottom: 12 }}>About</p>
                <h2 className="fd" style={{ fontSize: 38, fontWeight: 700, color: G.white, marginBottom: 20 }}>Built by Himesh & Tiru.</h2>
                <p style={{ fontSize: 15, color: G.gray, lineHeight: 1.8, marginBottom: 36 }}>
                  Img2XL eliminates the pain of manually copying data from documents into spreadsheets. Upload any file — PDF, image, Word, Excel, CSV, PowerPoint — and get structured data instantly.
                </p>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 36 }}>
                  {['React', 'TypeScript', 'Firebase', 'Groq AI', 'SheetJS', 'pdf-parse', 'mammoth'].map(t => <span key={t} className="tg">{t}</span>)}
                </div>
                <a href="mailto:himeshandtiru@gmail.com" className="gh" style={{ padding: '10px 20px', textDecoration: 'none' }}>✉ himeshandtiru@gmail.com</a>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Name gate */}
        <AnimatePresence>
          {showGate && (
            <InnerGate onDone={handleGateDone} onLoginRequest={() => { setShowGate(false); setShowAuth(true); }} />
          )}
        </AnimatePresence>

        <AuthModal open={showAuth} onClose={() => setShowAuth(false)} onSuccess={() => { setShowAuth(false); setTab('convert'); addToast('success', 'Signed in!'); }} />
        <Toasts items={toasts} dismiss={id => setToasts(p => p.filter(t => t.id !== id))} />
      </div>
    </AuthProvider>
  );
}

// ─── Inner wrappers that can access context ───────────────────────────────────
function InnerConvert({ addToast, onNeedAuth }: { addToast: any; onNeedAuth: () => void }) {
  const { user, guestName, usageCount } = useAuth();

  // If no identity yet — show gate
  const [gated, setGated] = useState(!user && !guestName);

  if (gated) {
    return (
      <div style={{ paddingTop: 80 }}>
        <GateInPage onDone={() => setGated(false)} onLogin={onNeedAuth} />
      </div>
    );
  }

  return <ConvertPage addToast={addToast} onNeedAuth={onNeedAuth} />;
}

function GateInPage({ onDone, onLogin }: { onDone: () => void; onLogin: () => void }) {
  const { setGuestName } = useAuth();
  const [name, setName] = useState('');
  const [err,  setErr]  = useState('');

  const go = () => {
    if (!name.trim()) { setErr('Enter your name to continue.'); return; }
    setGuestName(name.trim());
    onDone();
  };

  return (
    <div style={{ maxWidth: 420, margin: '60px auto', padding: '0 24px' }}>
      <div style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 16, padding: 40 }}>
        <Logo />
        <h2 style={{ fontSize: 20, fontWeight: 700, color: G.white, marginTop: 24, marginBottom: 8 }}>What's your name?</h2>
        <p style={{ fontSize: 13, color: G.gray, marginBottom: 24, lineHeight: 1.6 }}>
          Your first 2 conversions are free — no account needed. From your 3rd conversion, a free account is required.
        </p>
        <input className="if" value={name} onChange={e => setName(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && go()} placeholder="Your name" autoFocus />
        {err && <p style={{ fontSize: 12, color: G.err, marginTop: 6 }}>{err}</p>}
        <button className="bg" style={{ width: '100%', justifyContent: 'center', padding: '12px 0', marginTop: 16 }} onClick={go}>
          Continue Free
        </button>
        <div style={{ textAlign: 'center', marginTop: 14 }}>
          <button onClick={onLogin} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: G.gold }}>
            Already have an account? Sign in
          </button>
        </div>
      </div>
    </div>
  );
}

function InnerGate({ onDone, onLoginRequest }: { onDone: (n: string) => void; onLoginRequest: () => void }) {
  const { usageCount, setGuestName } = useAuth();
  const requireLogin = usageCount >= 2;
  const [name, setName] = useState('');
  const [err,  setErr]  = useState('');

  const go = () => {
    if (requireLogin) { onLoginRequest(); return; }
    if (!name.trim()) { setErr('Enter your name.'); return; }
    setGuestName(name.trim());
    onDone(name.trim());
  };

  return (
    <div className="mb">
      <motion.div initial={{ opacity: 0, scale: .95 }} animate={{ opacity: 1, scale: 1 }}
        style={{ background: G.bgCard, border: `1px solid ${G.border}`, borderRadius: 16, padding: 36, width: '100%', maxWidth: 380 }}>
        <Logo />
        <h2 style={{ fontSize: 19, fontWeight: 700, color: G.white, marginTop: 20, marginBottom: 10 }}>
          {requireLogin ? 'Login Required' : "What's your name?"}
        </h2>
        <p style={{ fontSize: 13, color: G.gray, marginBottom: 20, lineHeight: 1.6 }}>
          {requireLogin ? 'You\'ve used your 2 free conversions. Create a free account to continue.' : 'Enter your name to get your 2 free conversions.'}
        </p>
        {!requireLogin && (
          <>
            <input className="if" value={name} onChange={e => setName(e.target.value)} onKeyDown={e => e.key === 'Enter' && go()} placeholder="Your name" autoFocus />
            {err && <p style={{ fontSize: 12, color: G.err, marginTop: 6 }}>{err}</p>}
          </>
        )}
        <button className="bg" style={{ width: '100%', justifyContent: 'center', padding: '12px 0', marginTop: 16 }} onClick={go}>
          {requireLogin ? 'Sign In / Sign Up' : 'Continue'}
        </button>
      </motion.div>
    </div>
  );
}