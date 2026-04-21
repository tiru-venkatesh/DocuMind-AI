/**
 * Img2XL — Convert Images to Excel Instantly
 * Built by Himesh & Tiru | Powered by Gemini AI + Firebase
 * Theme: Premium Black & Gold
 * @license Apache-2.0
 */

import React, {
  useState, useRef, useEffect, useContext, createContext, useMemo,
} from 'react';
import {
  Upload, FileSpreadsheet, Trash2, Loader2, Download, AlertCircle,
  CheckCircle2, Table as TableIcon, Sparkles, RefreshCw, X,
  Plus, Edit2, Columns as ColumnsIcon, LogOut, History,
  User, Github, Mail, Lock, Eye, EyeOff, ChevronDown,
  ShieldCheck, Zap, MessageSquare, Info, CreditCard,
  Menu, Send, UserCircle2, FileDown, Copy,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
// AI calls go through /api/extract and /api/chat on the server
import { initializeApp } from 'firebase/app';
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword,
  createUserWithEmailAndPassword, signOut, GoogleAuthProvider,
  signInWithPopup, updateProfile, updatePassword, deleteUser,
  User as FirebaseUser,
} from 'firebase/auth';
import {
  getFirestore, collection, addDoc, query, where, getDocs,
  deleteDoc, doc, orderBy, serverTimestamp, setDoc, getDoc,
  Timestamp,
} from 'firebase/firestore';
const firebaseConfig = {
  apiKey: "AIzaSyCDKY9B4j4HMt39LhwgujNZTM0NNX-l-ts",
  authDomain: "my-project-1436-1754940582084.firebaseapp.com",
  projectId: "my-project-1436-1754940582084",
  storageBucket: "my-project-1436-1754940582084.firebasestorage.app",
  messagingSenderId: "917682303442",
  appId: "1:917682303442:web:9dd8a887a1f203458e180b"
};
// ─── Firebase & Gemini ────────────────────────────────────────────────────────
const firebaseApp = initializeApp(firebaseConfig);
const auth        = getAuth(firebaseApp);
const db          = getFirestore(firebaseApp);


// ─── Design Tokens ───────────────────────────────────────────────────────────
const G = {
  bg:        '#0a0a0a',
  bgCard:    '#111111',
  bgCard2:   '#161616',
  border:    '#222222',
  gold:      '#C9A84C',
  goldLight: '#E2C97E',
  goldDark:  '#8B6914',
  white:     '#FFFFFF',
  gray:      '#888888',
  grayDim:   '#444444',
  success:   '#4ADE80',
  error:     '#F87171',
  info:      '#60A5FA',
};

// ─── Global CSS injected once ────────────────────────────────────────────────
const GLOBAL_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,700;0,900;1,700&family=DM+Sans:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{background:${G.bg};color:${G.white};font-family:'DM Sans',sans-serif;overflow-x:hidden}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:${G.bg}}
::-webkit-scrollbar-thumb{background:${G.goldDark};border-radius:2px}
::selection{background:${G.gold}33;color:${G.goldLight}}

.fd{font-family:'Playfair Display',serif}
.fm{font-family:'JetBrains Mono',monospace}

.gold-text{background:linear-gradient(135deg,${G.gold},${G.goldLight},${G.gold});-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}

.btn-gold{background:linear-gradient(135deg,${G.gold},${G.goldLight});color:#000;font-weight:700;font-family:'DM Sans',sans-serif;letter-spacing:.08em;text-transform:uppercase;font-size:11px;border:none;cursor:pointer;transition:all .2s;border-radius:8px;display:inline-flex;align-items:center;gap:8px}
.btn-gold:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 8px 24px ${G.gold}33}
.btn-gold:active{transform:translateY(0)}
.btn-gold:disabled{opacity:.5;pointer-events:none}

.btn-ghost{background:transparent;color:${G.gray};border:1px solid ${G.border};font-weight:600;font-family:'DM Sans',sans-serif;letter-spacing:.06em;text-transform:uppercase;font-size:11px;cursor:pointer;transition:all .2s;border-radius:8px;display:inline-flex;align-items:center;gap:8px}
.btn-ghost:hover{border-color:${G.gold}66;color:${G.goldLight};background:${G.gold}0A}

.nav-btn{font-size:13px;font-weight:500;color:${G.gray};cursor:pointer;background:none;border:none;transition:color .2s;padding:0}
.nav-btn:hover,.nav-btn.active{color:${G.white}}

.glass{background:${G.bgCard};border:1px solid ${G.border};border-radius:12px}

.input-field{width:100%;background:${G.bgCard2};border:1px solid ${G.border};border-radius:8px;padding:12px 16px;color:${G.white};font-size:14px;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .2s}
.input-field:focus{border-color:${G.gold}66}
.input-field::placeholder{color:${G.grayDim}}

.drop-zone{border:1.5px dashed ${G.border};border-radius:12px;cursor:pointer;transition:all .25s;background:${G.bgCard2};display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:48px 24px;text-align:center}
.drop-zone:hover,.drop-zone.over{border-color:${G.gold}88;background:${G.gold}07}

.tag{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:4px;font-size:10px;font-family:'JetBrains Mono',monospace;letter-spacing:.08em;text-transform:uppercase;font-weight:500;background:${G.gold}15;color:${G.gold};border:1px solid ${G.gold}33}

.dot{width:7px;height:7px;border-radius:50%;display:inline-block}
.dot-ok{background:${G.success};box-shadow:0 0 6px ${G.success}}
.dot-proc{background:${G.gold};animation:pulse-dot 1.2s infinite}
.dot-err{background:${G.error}}
.dot-idle{background:${G.grayDim}}

.progress-bar{height:2px;background:${G.border};border-radius:1px;overflow:hidden;width:100%}
.progress-fill{height:100%;background:linear-gradient(90deg,${G.goldDark},${G.gold},${G.goldLight});border-radius:1px;transition:width .4s ease;position:relative;overflow:hidden}
.progress-fill::after{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,.3),transparent);animation:shimmer 1.5s infinite}

.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.85);backdrop-filter:blur(10px);z-index:400;display:flex;align-items:center;justify-content:center;padding:24px}

.dd-menu{position:absolute;right:0;top:calc(100% + 8px);width:200px;background:${G.bgCard};border:1px solid ${G.border};border-radius:10px;overflow:hidden;box-shadow:0 16px 48px rgba(0,0,0,.6);z-index:200}
.dd-item{width:100%;background:none;border:none;padding:12px 16px;color:${G.gray};font-size:12px;font-family:'DM Sans',sans-serif;font-weight:500;cursor:pointer;display:flex;align-items:center;gap:10px;transition:all .15s;text-align:left}
.dd-item:hover{background:${G.gold}0D;color:${G.goldLight}}
.dd-item.danger:hover{background:${G.error}15;color:${G.error}}

.tr-row:hover{background:${G.gold}08}
.chat-user{background:${G.gold}22;border:1px solid ${G.gold}33;border-radius:12px 12px 4px 12px}
.chat-bot{background:${G.bgCard2};border:1px solid ${G.border};border-radius:12px 12px 12px 4px}
.sl{font-size:10px;font-family:'JetBrains Mono',monospace;letter-spacing:.15em;text-transform:uppercase;color:${G.gold};font-weight:500}

@keyframes pulse-dot{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(.85)}}
@keyframes shimmer{0%{transform:translateX(-100%)}100%{transform:translateX(100%)}}
@keyframes spin{to{transform:rotate(360deg)}}
.spin{animation:spin 1s linear infinite}

@media(max-width:768px){.hide-m{display:none!important}}
`;

// ─── Types ────────────────────────────────────────────────────────────────────
interface ConversionHistory { id:string; fileName:string; timestamp:Timestamp; extractedData:any[]; rowCount:number; userId:string; }
interface FileItem { id:string; file:File; preview:string|null; status:'pending'|'processing'|'completed'|'error'; progress:number; statusMessage:string; extractedData:any[]|null; error?:string; }
interface ChatMsg { role:'user'|'model'; content:string; ts:number; }
interface Toast { id:string; type:'success'|'error'|'info'; msg:string; }
interface Profile { displayName:string; photoURL:string; totalConversions:number; plan:'free'|'pro'; createdAt:string; }
interface AuthCtx { user:FirebaseUser|null; profile:Profile|null; loading:boolean; refresh:()=>Promise<void>; }

// ─── Auth Context ─────────────────────────────────────────────────────────────
const Ctx = createContext<AuthCtx>({ user:null, profile:null, loading:true, refresh:async()=>{} });
const useAuth = () => useContext(Ctx);

const AuthProvider:React.FC<{children:React.ReactNode}> = ({children}) => {
  const [user, setUser]       = useState<FirebaseUser|null>(null);
  const [profile, setProfile] = useState<Profile|null>(null);
  const [loading, setLoading] = useState(true);

  const loadProfile = async (uid:string) => {
    try {
      const snap = await getDoc(doc(db,'users',uid));
      if (snap.exists()) { setProfile(snap.data() as Profile); return; }
      const p:Profile = { displayName:auth.currentUser?.displayName||'User', photoURL:'', totalConversions:0, plan:'free', createdAt:new Date().toISOString() };
      await setDoc(doc(db,'users',uid), p);
      setProfile(p);
    } catch(e) { console.error(e); }
  };

  useEffect(() => onAuthStateChanged(auth, async u => {
    setUser(u);
    if (u) await loadProfile(u.uid); else setProfile(null);
    setLoading(false);
  }),[]);

  const val = useMemo(()=>({ user, profile, loading, refresh: async()=>{ if(user) await loadProfile(user.uid); } }),[user,profile,loading]);
  return <Ctx.Provider value={val}>{children}</Ctx.Provider>;
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const Logo = ({sm=false}:{sm?:boolean}) => (
  <div style={{display:'flex',alignItems:'center',gap:10}}>
    <div style={{width:sm?30:36,height:sm?30:36,background:`linear-gradient(135deg,${G.goldDark},${G.gold})`,borderRadius:8,display:'flex',alignItems:'center',justifyContent:'center'}}>
      <TableIcon size={sm?15:18} color="#000" strokeWidth={2.5}/>
    </div>
    <div>
      <div style={{fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:sm?15:18,color:G.white,letterSpacing:'-0.02em',lineHeight:1}}>
        IMG<span style={{color:G.gold}}>2XL</span>
      </div>
      {!sm && <div style={{fontFamily:"'JetBrains Mono',monospace",fontSize:8,color:G.gold,letterSpacing:'0.15em',textTransform:'uppercase',marginTop:2}}>PRO · v4.2</div>}
    </div>
  </div>
);

// ─── Toasts ───────────────────────────────────────────────────────────────────
const Toasts = ({items,dismiss}:{items:Toast[];dismiss:(id:string)=>void}) => (
  <div style={{position:'fixed',top:72,right:20,zIndex:500,display:'flex',flexDirection:'column',gap:8,maxWidth:320}}>
    <AnimatePresence>
      {items.map(t => (
        <motion.div key={t.id} initial={{opacity:0,x:40}} animate={{opacity:1,x:0}} exit={{opacity:0,x:40}}
          style={{display:'flex',alignItems:'center',gap:10,padding:'12px 16px',borderRadius:8,
            background:t.type==='success'?`${G.success}15`:t.type==='error'?`${G.error}15`:`${G.gold}15`,
            border:`1px solid ${t.type==='success'?G.success:t.type==='error'?G.error:G.gold}33`}}>
          {t.type==='success'?<CheckCircle2 size={13} color={G.success}/>:t.type==='error'?<AlertCircle size={13} color={G.error}/>:<Info size={13} color={G.gold}/>}
          <span style={{fontSize:12,color:G.white,flex:1}}>{t.msg}</span>
          <button onClick={()=>dismiss(t.id)} style={{background:'none',border:'none',cursor:'pointer',color:G.gray}}><X size={11}/></button>
        </motion.div>
      ))}
    </AnimatePresence>
  </div>
);

// ─── Navbar ───────────────────────────────────────────────────────────────────
const Navbar = ({tab,setTab,onLogin,onLogout}:{tab:string;setTab:(t:string)=>void;onLogin:()=>void;onLogout:()=>void}) => {
  const {user,profile} = useAuth();
  const [sc,setSc] = useState(false);
  const [dd,setDd] = useState(false);
  const [mb,setMb] = useState(false);

  useEffect(()=>{const h=()=>setSc(window.scrollY>20); window.addEventListener('scroll',h); return()=>window.removeEventListener('scroll',h);},[]);

  const pubLinks = [{id:'landing',l:'Home'},{id:'about',l:'About'},{id:'pricing',l:'Pricing'}];
  const priLinks = [{id:'convert',l:'Convert'},{id:'history',l:'History'}];
  const links    = user ? [pubLinks[0],...priLinks,...pubLinks.slice(1)] : pubLinks;

  return (
    <nav style={{position:'fixed',top:0,left:0,right:0,zIndex:100,background:sc?`${G.bg}EE`:'transparent',backdropFilter:sc?'blur(20px)':'none',borderBottom:sc?`1px solid ${G.border}`:'1px solid transparent',transition:'all .3s'}}>
      <div style={{maxWidth:1280,margin:'0 auto',padding:'0 24px',height:60,display:'flex',alignItems:'center',justifyContent:'space-between'}}>
        <button onClick={()=>setTab('landing')} style={{background:'none',border:'none',cursor:'pointer'}}><Logo/></button>

        {/* Desktop */}
        <div className="hide-m" style={{display:'flex',alignItems:'center',gap:32}}>
          {links.map(l=>(
            <button key={l.id} className={`nav-btn ${tab===l.id?'active':''}`} onClick={()=>setTab(l.id)} style={{color:tab===l.id?G.white:G.gray}}>{l.l}</button>
          ))}
        </div>

        <div style={{display:'flex',alignItems:'center',gap:12}}>
          {/* Status badge */}
          <div className="hide-m" style={{display:'flex',alignItems:'center',gap:6,padding:'5px 10px',borderRadius:4,background:`${G.success}15`,border:`1px solid ${G.success}33`}}>
            <span className="dot dot-ok" style={{width:6,height:6}}/>
            <span className="fm" style={{fontSize:9,color:G.success,letterSpacing:'0.08em'}}>SYSTEM NOMINAL</span>
          </div>

          {user ? (
            <div style={{position:'relative'}}>
              <button onClick={()=>setDd(d=>!d)} style={{display:'flex',alignItems:'center',gap:8,background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:8,padding:'6px 12px 6px 6px',cursor:'pointer'}}>
                <div style={{width:28,height:28,borderRadius:6,background:`linear-gradient(135deg,${G.goldDark},${G.gold})`,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,fontWeight:700,color:'#000'}}>
                  {(user.email||'U')[0].toUpperCase()}
                </div>
                <span className="hide-m" style={{fontSize:12,color:G.white,fontWeight:500}}>{profile?.displayName||user.email?.split('@')[0]}</span>
                <ChevronDown size={12} color={G.gray}/>
              </button>
              <AnimatePresence>
                {dd && (
                  <motion.div className="dd-menu" initial={{opacity:0,y:-8,scale:.96}} animate={{opacity:1,y:0,scale:1}} exit={{opacity:0,y:-8,scale:.96}} transition={{duration:.15}}>
                    <div style={{padding:'12px 16px',borderBottom:`1px solid ${G.border}`}}>
                      <div style={{fontSize:12,color:G.white,fontWeight:600}}>{profile?.displayName||'User'}</div>
                      <div style={{fontSize:11,color:G.gray,marginTop:2}}>{user.email}</div>
                      <div className="tag" style={{marginTop:6}}>{profile?.plan==='pro'?'⭐ Pro':'Free'}</div>
                    </div>
                    <button className="dd-item" onClick={()=>{setTab('account');setDd(false)}}><UserCircle2 size={14}/> Account</button>
                    <button className="dd-item" onClick={()=>{setTab('history');setDd(false)}}><History size={14}/> History</button>
                    <div style={{borderTop:`1px solid ${G.border}`,margin:'4px 0'}}/>
                    <button className="dd-item danger" onClick={()=>{onLogout();setDd(false)}}><LogOut size={14}/> Logout</button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          ) : (
            <button className="btn-gold" style={{padding:'8px 20px'}} onClick={onLogin}>Sign In</button>
          )}

          <button onClick={()=>setMb(m=>!m)} style={{background:'none',border:'none',cursor:'pointer',color:G.gray,padding:4,display:'none'}} className="show-m">
            {mb?<X size={20}/>:<Menu size={20}/>}
          </button>
        </div>
      </div>

      {/* Mobile menu */}
      <AnimatePresence>
        {mb && (
          <motion.div initial={{height:0}} animate={{height:'auto'}} exit={{height:0}} style={{overflow:'hidden',background:G.bgCard,borderTop:`1px solid ${G.border}`}}>
            <div style={{padding:16,display:'flex',flexDirection:'column',gap:4}}>
              {links.map(l=>(
                <button key={l.id} onClick={()=>{setTab(l.id);setMb(false);}} style={{background:tab===l.id?`${G.gold}15`:'none',border:'none',borderRadius:8,padding:'12px 16px',color:tab===l.id?G.gold:G.gray,fontFamily:"'DM Sans',sans-serif",fontSize:14,fontWeight:500,textAlign:'left',cursor:'pointer'}}>{l.l}</button>
              ))}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </nav>
  );
};

// ─── Landing ──────────────────────────────────────────────────────────────────
const Landing = ({onStart}:{onStart:()=>void}) => {
  const feats = [
    {icon:<Zap size={20} color={G.gold}/>, t:'Instant Extraction', d:'Gemini AI reads every cell, header, and row from your image in seconds.'},
    {icon:<FileSpreadsheet size={20} color={G.gold}/>, t:'Excel-Ready Output', d:'Download a clean, structured .xlsx — open straight in Excel or Sheets.'},
    {icon:<ShieldCheck size={20} color={G.gold}/>, t:'Secure & Private', d:'Stateless processing. No image data stored beyond your own history.'},
  ];
  const steps = [
    {n:'01',t:'Upload',d:'Drop any receipt, invoice, or table image.'},
    {n:'02',t:'Extract',d:'AI scans and structures all tabular data instantly.'},
    {n:'03',t:'Download',d:'Get your Excel file, rename headers if needed.'},
  ];

  return (
    <div style={{paddingTop:60}}>
      {/* Hero */}
      <section style={{minHeight:'100vh',display:'flex',alignItems:'center',padding:'0 24px',maxWidth:1280,margin:'0 auto',gap:80,flexWrap:'wrap'}}>
        <motion.div initial={{opacity:0,x:-30}} animate={{opacity:1,x:0}} transition={{duration:.7}} style={{flex:1,minWidth:280}}>
          <div className="tag" style={{marginBottom:28}}><Sparkles size={10}/> AI-POWERED · STATELESS · INSTANT</div>
          <h1 className="fd" style={{fontSize:'clamp(48px,7vw,88px)',lineHeight:.92,fontWeight:900,color:G.white,marginBottom:28}}>
            VISUAL<br/><em className="gold-text" style={{fontStyle:'italic'}}>DATA,</em><br/>STRUCTURED.
          </h1>
          <p style={{fontSize:15,color:G.gray,lineHeight:1.7,maxWidth:400,marginBottom:40}}>
            Enterprise extraction from images and PDFs.<br/>Instant Excel output.
          </p>
          <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
            <button className="btn-gold" style={{padding:'14px 28px',fontSize:12}} onClick={onStart}><Upload size={14}/> Upload a File</button>
            <button className="btn-ghost" style={{padding:'14px 24px'}} onClick={()=>document.getElementById('hiw')?.scrollIntoView({behavior:'smooth'})}>How It Works</button>
          </div>
          <div style={{marginTop:32,display:'flex',alignItems:'center',gap:6,color:G.grayDim,fontSize:12}}>
            <ShieldCheck size={12} color={G.grayDim}/> Stateless · No data stored
          </div>
        </motion.div>

        {/* Mock preview */}
        <motion.div className="hide-m" initial={{opacity:0,x:30}} animate={{opacity:1,x:0}} transition={{duration:.7,delay:.15}} style={{flex:1,maxWidth:460}}>
          <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:16,overflow:'hidden',boxShadow:`0 40px 80px rgba(0,0,0,.5)`}}>
            <div style={{padding:'12px 18px',borderBottom:`1px solid ${G.border}`,display:'flex',alignItems:'center',gap:12,background:G.bgCard2}}>
              <div style={{display:'flex',gap:5}}>
                {['#F87171','#FBBF24','#4ADE80'].map(c=><div key={c} style={{width:9,height:9,borderRadius:'50%',background:c}}/>)}
              </div>
              <div style={{flex:1,height:7,background:G.border,borderRadius:4}}/>
            </div>
            <div style={{padding:20}}>
              <div className="tag" style={{marginBottom:14}}>Extracted Data · 18 rows · 3 columns</div>
              <div style={{borderRadius:8,overflow:'hidden',border:`1px solid ${G.border}`}}>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',padding:'9px 14px',background:G.bgCard2,gap:16}}>
                  {['Category','Subcategory','Amount'].map(h=>(
                    <span key={h} className="fm" style={{fontSize:9,color:G.gold,letterSpacing:'0.1em',textTransform:'uppercase'}}>{h}</span>
                  ))}
                </div>
                {[['Food','Restaurant','$42.50'],['Transport','Uber','$18.00'],['Utilities','Electric','$120.00'],['Shopping','Groceries','$85.75']].map(([a,b,c],i)=>(
                  <div key={i} className="tr-row" style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',padding:'9px 14px',gap:16,borderTop:`1px solid ${G.border}`,fontSize:12,color:G.gray}}>
                    <span style={{color:G.white}}>{a}</span><span>{b}</span>
                    <span className="fm" style={{color:G.gold}}>{c}</span>
                  </div>
                ))}
              </div>
              <button className="btn-gold" style={{marginTop:14,padding:'10px 0',width:'100%',justifyContent:'center',fontSize:11}}><Download size={12}/> Download Excel</button>
            </div>
          </div>
        </motion.div>
      </section>

      {/* Features */}
      <section style={{padding:'100px 24px',maxWidth:1280,margin:'0 auto'}}>
        <p className="sl" style={{textAlign:'center',marginBottom:16}}>Core Features</p>
        <h2 className="fd" style={{fontSize:38,fontWeight:700,textAlign:'center',color:G.white,marginBottom:56}}>Built for precision.</h2>
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))',gap:20}}>
          {feats.map((f,i)=>(
            <motion.div key={i} initial={{opacity:0,y:20}} whileInView={{opacity:1,y:0}} viewport={{once:true}} transition={{delay:i*.1}}
              style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:32,transition:'border-color .2s'}}
              whileHover={{borderColor:`${G.gold}44`} as any}>
              <div style={{width:42,height:42,borderRadius:10,background:`${G.gold}15`,border:`1px solid ${G.gold}33`,display:'flex',alignItems:'center',justifyContent:'center',marginBottom:20}}>{f.icon}</div>
              <h3 style={{fontSize:16,fontWeight:600,color:G.white,marginBottom:10}}>{f.t}</h3>
              <p style={{fontSize:13,color:G.gray,lineHeight:1.7}}>{f.d}</p>
            </motion.div>
          ))}
        </div>
      </section>

      {/* How it works */}
      <section id="hiw" style={{padding:'100px 24px',borderTop:`1px solid ${G.border}`,borderBottom:`1px solid ${G.border}`,background:G.bgCard2}}>
        <div style={{maxWidth:800,margin:'0 auto',textAlign:'center'}}>
          <p className="sl" style={{marginBottom:16}}>How It Works</p>
          <h2 className="fd" style={{fontSize:38,fontWeight:700,color:G.white,marginBottom:56}}>Three steps to structured data.</h2>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:32}}>
            {steps.map((s,i)=>(
              <motion.div key={i} initial={{opacity:0,y:20}} whileInView={{opacity:1,y:0}} viewport={{once:true}} transition={{delay:i*.15}} style={{textAlign:'center'}}>
                <div className="fm" style={{fontSize:32,color:`${G.gold}44`,marginBottom:14}}>{s.n}</div>
                <h3 style={{fontSize:16,fontWeight:600,color:G.white,marginBottom:8}}>{s.t}</h3>
                <p style={{fontSize:13,color:G.gray,lineHeight:1.7}}>{s.d}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* About */}
      <section id="about" style={{padding:'100px 24px',maxWidth:900,margin:'0 auto'}}>
        <p className="sl" style={{marginBottom:16}}>About</p>
        <h2 className="fd" style={{fontSize:38,fontWeight:700,color:G.white,marginBottom:24}}>Built by Himesh & Tiru.</h2>
        <p style={{fontSize:15,color:G.gray,lineHeight:1.8,maxWidth:620,marginBottom:40}}>
          Img2XL was created to eliminate the time wasted manually re-typing data from images into spreadsheets.
          We combined Google Gemini's vision AI with a clean, fast interface — so you go from image to Excel in seconds.
        </p>
        <div style={{display:'flex',flexWrap:'wrap',gap:10,marginBottom:40}}>
          {['React','TypeScript','Firebase','Gemini AI','SheetJS','Tailwind'].map(t=><div key={t} className="tag">{t}</div>)}
        </div>
        <div style={{display:'flex',gap:14,flexWrap:'wrap'}}>
          <a href="mailto:himeshandtiru@gmail.com" className="btn-ghost" style={{padding:'10px 20px',textDecoration:'none'}}><Mail size={13}/> himeshandtiru@gmail.com</a>
          <a href="https://github.com" target="_blank" rel="noreferrer" className="btn-ghost" style={{padding:'10px 20px',textDecoration:'none'}}><Github size={13}/> GitHub</a>
        </div>
      </section>

      {/* Pricing */}
      <section id="pricing" style={{padding:'100px 24px',borderTop:`1px solid ${G.border}`,background:G.bgCard2}}>
        <div style={{maxWidth:900,margin:'0 auto'}}>
          <p className="sl" style={{marginBottom:16,textAlign:'center'}}>Pricing</p>
          <h2 className="fd" style={{fontSize:38,fontWeight:700,color:G.white,marginBottom:56,textAlign:'center'}}>Start free. Scale up.</h2>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))',gap:20}}>
            {/* Free */}
            <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:36}}>
              <h3 style={{fontSize:20,fontWeight:600,color:G.white,marginBottom:6}}>Free</h3>
              <div style={{fontSize:40,fontWeight:700,color:G.white,marginBottom:24}}>$0<span style={{fontSize:13,color:G.gray,fontWeight:400}}>/mo</span></div>
              {['5 conversions/day','Excel download','Session history','Editable headers'].map(f=>(
                <div key={f} style={{display:'flex',alignItems:'center',gap:10,marginBottom:12}}><CheckCircle2 size={13} color={G.gold}/><span style={{fontSize:13,color:G.gray}}>{f}</span></div>
              ))}
              <button className="btn-ghost" style={{width:'100%',justifyContent:'center',padding:'12px 0',marginTop:20}} onClick={onStart}>Get Started Free</button>
            </div>
            {/* Pro */}
            <div style={{background:`linear-gradient(135deg,${G.gold}12,${G.bgCard})`,border:`1px solid ${G.gold}44`,borderRadius:12,padding:36,position:'relative'}}>
              <div style={{position:'absolute',top:16,right:16,background:G.gold,color:'#000',fontSize:9,fontWeight:700,letterSpacing:'0.1em',padding:'3px 8px',borderRadius:4,textTransform:'uppercase'}}>Recommended</div>
              <h3 style={{fontSize:20,fontWeight:600,color:G.white,marginBottom:6}}>Pro</h3>
              <div style={{fontSize:40,fontWeight:700,color:G.white,marginBottom:24}}>$12<span style={{fontSize:13,color:G.gray,fontWeight:400}}>/mo</span></div>
              {['Unlimited conversions','Cloud history vault','Batch processing','Priority AI access','API access'].map(f=>(
                <div key={f} style={{display:'flex',alignItems:'center',gap:10,marginBottom:12}}><CheckCircle2 size={13} color={G.gold}/><span style={{fontSize:13,color:G.gray}}>{f}</span></div>
              ))}
              <button className="btn-gold" style={{width:'100%',justifyContent:'center',padding:'12px 0',marginTop:20}}>Upgrade to Pro</button>
            </div>
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer style={{borderTop:`1px solid ${G.border}`,padding:'40px 24px',display:'flex',flexWrap:'wrap',alignItems:'center',justifyContent:'space-between',gap:16,maxWidth:1280,margin:'0 auto'}}>
        <Logo sm/>
        <div style={{display:'flex',gap:24}}>
          {['About','Pricing','Privacy'].map(l=><span key={l} style={{fontSize:12,color:G.grayDim,cursor:'pointer'}}>{l}</span>)}
        </div>
        <div className="fm" style={{fontSize:11,color:G.grayDim}}>© 2026 IMG2XL · Built by Himesh & Tiru</div>
      </footer>
    </div>
  );
};

// ─── File Card ────────────────────────────────────────────────────────────────
const FileCard = ({file,onRemove,onDownload,onUpdate}:{file:FileItem;onRemove:()=>void;onDownload:()=>void;onUpdate:(f:FileItem)=>void}) => {
  const [editH, setEditH] = useState<string|null>(null);
  const [draft, setDraft] = useState('');
  const [selCols, setSelCols] = useState<string[]>([]);
  const [copied, setCopied] = useState(false);

  const toggleCol = (c:string) => setSelCols(p=>p.includes(c)?p.filter(x=>x!==c):[...p,c]);

  const merge = () => {
    if (selCols.length<2||!file.extractedData) return;
    const key = selCols.join(' + ');
    const data = file.extractedData.map(row=>{ const m=selCols.map(c=>row[c]).filter(Boolean).join(' '); const r={...row}; selCols.forEach(c=>delete r[c]); return {...r,[key]:m}; });
    onUpdate({...file,extractedData:data}); setSelCols([]);
  };

  const rename = (old:string) => {
    if (!draft.trim()||!file.extractedData) { setEditH(null); return; }
    const data = file.extractedData.map(row=>{ const r={...row}; r[draft.trim()]=r[old]; delete r[old]; return r; });
    onUpdate({...file,extractedData:data}); setEditH(null);
  };

  const copy = () => {
    if (!file.extractedData) return;
    const hs = Object.keys(file.extractedData[0]);
    const rows = file.extractedData.map(r=>hs.map(h=>r[h]).join('\t'));
    navigator.clipboard.writeText([hs.join('\t'),...rows].join('\n'));
    setCopied(true); setTimeout(()=>setCopied(false),2000);
  };

  const sc = file.status==='completed'?G.success:file.status==='error'?G.error:G.gold;
  const sl = file.status==='completed'?'Complete':file.status==='processing'?'Processing…':file.status==='error'?'Failed':'Queued';

  return (
    <motion.div layout initial={{opacity:0,y:16}} animate={{opacity:1,y:0}} style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,overflow:'hidden'}}>
      {/* Bar */}
      <div style={{padding:'13px 18px',borderBottom:`1px solid ${G.border}`,display:'flex',alignItems:'center',justifyContent:'space-between',gap:12,flexWrap:'wrap'}}>
        <div style={{display:'flex',alignItems:'center',gap:10,minWidth:0}}>
          <span className={`dot dot-${file.status==='completed'?'ok':file.status==='processing'?'proc':file.status==='error'?'err':'idle'}`}/>
          <span style={{fontSize:13,fontWeight:600,color:G.white,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:240}}>{file.file.name}</span>
          <span className="fm" style={{fontSize:10,color:sc}}>{sl}</span>
        </div>
        <div style={{display:'flex',gap:8,alignItems:'center',flexShrink:0}}>
          {selCols.length>=2 && <button className="btn-ghost" style={{padding:'6px 10px',fontSize:10}} onClick={merge}><ColumnsIcon size={11}/> Merge</button>}
          {file.status==='completed' && <>
            <button className="btn-ghost" style={{padding:'6px 9px'}} onClick={copy} title="Copy table">{copied?<CheckCircle2 size={13} color={G.success}/>:<Copy size={13}/>}</button>
            <button className="btn-gold" style={{padding:'7px 14px',fontSize:10}} onClick={onDownload}><Download size={12}/> Download Excel</button>
          </>}
          <button onClick={onRemove} style={{background:'none',border:`1px solid ${G.border}`,borderRadius:6,padding:'6px 8px',cursor:'pointer',color:G.gray,display:'flex'}}><Trash2 size={13}/></button>
        </div>
      </div>

      {file.status==='processing' && <div style={{padding:'0 18px'}}><div className="progress-bar" style={{margin:'10px 0'}}><div className="progress-fill" style={{width:`${file.progress}%`}}/></div></div>}

      {/* Content */}
      <div style={{display:'grid',gridTemplateColumns:file.extractedData?'180px 1fr':'1fr'}}>
        {file.preview && (
          <div style={{background:G.bgCard2,borderRight:`1px solid ${G.border}`,display:'flex',alignItems:'center',justifyContent:'center',minHeight:120,padding:10}}>
            <img src={file.preview} alt="preview" style={{maxWidth:'100%',maxHeight:160,objectFit:'contain',borderRadius:6}}/>
          </div>
        )}
        {file.status==='completed'&&file.extractedData && (
          <div style={{overflowX:'auto',maxHeight:280}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
              <thead>
                <tr>
                  {Object.keys(file.extractedData[0]).map((h,i)=>(
                    <th key={i} style={{padding:'9px 12px',textAlign:'left',background:G.bgCard2,borderBottom:`1px solid ${G.border}`,position:'sticky',top:0}}>
                      <div style={{display:'flex',alignItems:'center',gap:6}}>
                        <input type="checkbox" checked={selCols.includes(h)} onChange={()=>toggleCol(h)} style={{width:11,height:11,accentColor:G.gold,cursor:'pointer'}}/>
                        {editH===h ? (
                          <div style={{display:'flex',alignItems:'center',gap:4}}>
                            <input autoFocus value={draft} onChange={e=>setDraft(e.target.value)}
                              onKeyDown={e=>{if(e.key==='Enter')rename(h);if(e.key==='Escape')setEditH(null);}}
                              style={{background:G.bgCard,border:`1px solid ${G.gold}66`,borderRadius:4,padding:'2px 6px',color:G.white,fontSize:11,outline:'none',width:90}}/>
                            <button onClick={()=>rename(h)} style={{background:'none',border:'none',cursor:'pointer',color:G.success}}><CheckCircle2 size={11}/></button>
                            <button onClick={()=>setEditH(null)} style={{background:'none',border:'none',cursor:'pointer',color:G.error}}><X size={11}/></button>
                          </div>
                        ) : (
                          <div style={{display:'flex',alignItems:'center',gap:4}}>
                            <span className="fm" style={{fontSize:9,color:G.gold,letterSpacing:'0.08em',textTransform:'uppercase'}}>{h}</span>
                            <button onClick={()=>{setEditH(h);setDraft(h);}} style={{background:'none',border:'none',cursor:'pointer',color:G.grayDim,padding:1}}><Edit2 size={9}/></button>
                          </div>
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {file.extractedData.map((row,i)=>(
                  <tr key={i} className="tr-row" style={{borderBottom:`1px solid ${G.border}`}}>
                    {Object.keys(row).map((h,j)=>(
                      <td key={j} style={{padding:'8px 12px',color:G.gray,maxWidth:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{row[h]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {file.status==='error' && (
          <div style={{padding:20,display:'flex',alignItems:'center',gap:12,background:`${G.error}08`}}>
            <AlertCircle size={16} color={G.error}/>
            <div><div style={{fontSize:12,color:G.error,fontWeight:600}}>Extraction failed</div><div style={{fontSize:11,color:G.gray,marginTop:3}}>{file.error}</div></div>
          </div>
        )}
      </div>

      {file.status==='completed'&&file.extractedData && (
        <div style={{padding:'8px 18px',borderTop:`1px solid ${G.border}`,display:'flex',gap:20,background:G.bgCard2}}>
          <span className="fm" style={{fontSize:10,color:G.gray}}>{file.extractedData.length} rows</span>
          <span className="fm" style={{fontSize:10,color:G.gray}}>{Object.keys(file.extractedData[0]).length} columns</span>
          <span className="fm" style={{fontSize:10,color:G.gold,marginLeft:'auto'}}>✓ Ready</span>
        </div>
      )}
    </motion.div>
  );
};

// ─── Convert Page ─────────────────────────────────────────────────────────────
const Convert = ({files,setFiles,addToast}:{files:FileItem[];setFiles:React.Dispatch<React.SetStateAction<FileItem[]>>;addToast:(t:'success'|'error'|'info',m:string)=>void}) => {
  const {user} = useAuth();
  const [drag, setDrag] = useState(false);

  const process = async (item:FileItem) => {
    setFiles(p=>p.map(f=>f.id===item.id?{...f,status:'processing',progress:25,statusMessage:'Sending to Gemini…'}:f));
    try {
      const b64 = await new Promise<string>(res=>{ const r=new FileReader(); r.onload=e=>res((e.target?.result as string).split(',')[1]); r.readAsDataURL(item.file); });
      setFiles(p=>p.map(f=>f.id===item.id?{...f,progress:60,statusMessage:'Extracting data…'}:f));
      const resp = await fetch('/api/extract', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ base64Image: b64, mimeType: item.file.type }),
      });
      const result = await resp.json();
      if (!resp.ok) throw new Error(result.error || 'Extraction failed');
      const data = result.data;
      setFiles(p=>p.map(f=>f.id===item.id?{...f,status:'completed',progress:100,statusMessage:'Done',extractedData:data}:f));
      if (user) {
        await addDoc(collection(db,'history'),{fileName:item.file.name,timestamp:serverTimestamp(),extractedData:data,rowCount:data.length,userId:user.uid});
        const ref=doc(db,'users',user.uid); const snap=await getDoc(ref);
        if(snap.exists()) await setDoc(ref,{totalConversions:(snap.data().totalConversions||0)+1},{merge:true});
      }
      addToast('success',`${item.file.name} extracted!`);
    } catch(e:any) {
      setFiles(p=>p.map(f=>f.id===item.id?{...f,status:'error',error:e.message,statusMessage:'Failed'}:f));
      addToast('error',`Failed: ${e.message}`);
    }
  };

  const handleFiles = (fs:File[]) => {
    const items:FileItem[] = fs.map(f=>({id:Math.random().toString(36).substr(2,9),file:f,preview:URL.createObjectURL(f),status:'pending',progress:0,statusMessage:'Queued',extractedData:null}));
    setFiles(p=>[...items,...p]);
    items.forEach(process);
  };

  const dl = (f:FileItem) => {
    if(!f.extractedData) return;
    const ws=XLSX.utils.json_to_sheet(f.extractedData);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Data');
    XLSX.writeFile(wb,`${f.file.name.replace(/\.[^.]+$/,'')}_export.xlsx`);
  };

  return (
    <div style={{maxWidth:1000,margin:'0 auto',padding:'80px 24px 60px'}}>
      <div style={{display:'flex',alignItems:'flex-end',justifyContent:'space-between',marginBottom:32,flexWrap:'wrap',gap:16}}>
        <div><p className="sl" style={{marginBottom:8}}>Convert</p><h2 className="fd" style={{fontSize:34,fontWeight:700,color:G.white}}>Image to Excel</h2></div>
        <div style={{display:'flex',gap:10}}>
          <label className="btn-ghost" style={{padding:'10px 16px',cursor:'pointer'}}><Plus size={13}/> Add Files<input type="file" multiple accept="image/*" style={{display:'none'}} onChange={e=>e.target.files&&handleFiles(Array.from(e.target.files))}/></label>
          {files.length>0&&<button className="btn-ghost" style={{padding:'10px 16px',color:G.error,borderColor:`${G.error}44`}} onClick={()=>setFiles([])}><Trash2 size={13}/> Clear All</button>}
        </div>
      </div>

      {files.length===0 ? (
        <div className={`drop-zone${drag?' over':''}`} style={{minHeight:240}}
          onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
          onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(Array.from(e.dataTransfer.files).filter(f=>f.type.startsWith('image/')));}}
          onClick={()=>document.getElementById('fi-main')?.click()}>
          <div style={{width:48,height:48,borderRadius:10,background:`${G.gold}15`,border:`1px solid ${G.gold}33`,display:'flex',alignItems:'center',justifyContent:'center'}}><Upload size={20} color={G.gold}/></div>
          <div><p style={{fontSize:14,fontWeight:600,color:G.white,marginBottom:6}}>Drop files or click to upload</p><p style={{fontSize:12,color:G.gray}}>PNG, JPEG, PDF · max 20 MB</p></div>
          <input id="fi-main" type="file" multiple accept="image/*" style={{display:'none'}} onChange={e=>e.target.files&&handleFiles(Array.from(e.target.files))}/>
        </div>
      ) : (
        <div style={{display:'flex',flexDirection:'column',gap:14}}>
          {files.map(f=><FileCard key={f.id} file={f} onRemove={()=>setFiles(p=>p.filter(x=>x.id!==f.id))} onDownload={()=>dl(f)} onUpdate={u=>setFiles(p=>p.map(x=>x.id===u.id?u:x))}/>)}
          <div className={`drop-zone${drag?' over':''}`} style={{minHeight:70,padding:20,flexDirection:'row',gap:10}}
            onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
            onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(Array.from(e.dataTransfer.files).filter(f=>f.type.startsWith('image/')));}}
            onClick={()=>document.getElementById('fi-more')?.click()}>
            <Plus size={14} color={G.grayDim}/><span style={{fontSize:12,color:G.grayDim}}>Add more files</span>
            <input id="fi-more" type="file" multiple accept="image/*" style={{display:'none'}} onChange={e=>e.target.files&&handleFiles(Array.from(e.target.files))}/>
          </div>
        </div>
      )}
    </div>
  );
};

// ─── History Page ─────────────────────────────────────────────────────────────
const HistoryPage = ({goConvert}:{goConvert:()=>void}) => {
  const {user}   = useAuth();
  const [hist,   setHist]   = useState<ConversionHistory[]>([]);
  const [loading,setLoading]= useState(true);

  const load = async () => {
    if(!user) return; setLoading(true);
    try {
      const q=query(collection(db,'history'),where('userId','==',user.uid),orderBy('timestamp','desc'));
      const snap=await getDocs(q);
      setHist(snap.docs.map(d=>({id:d.id,...d.data()} as ConversionHistory)));
    } catch(e){console.error(e);}
    finally{setLoading(false);}
  };

  useEffect(()=>{load();},[user]);

  const del = async (id:string) => { await deleteDoc(doc(db,'history',id)); setHist(p=>p.filter(h=>h.id!==id)); };

  const dl = (h:ConversionHistory) => {
    const ws=XLSX.utils.json_to_sheet(h.extractedData);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Data');
    XLSX.writeFile(wb,`${h.fileName.replace(/\.[^.]+$/,'')}_recovered.xlsx`);
  };

  return (
    <div style={{maxWidth:1000,margin:'0 auto',padding:'80px 24px 60px'}}>
      <div style={{display:'flex',alignItems:'flex-end',justifyContent:'space-between',marginBottom:36,flexWrap:'wrap',gap:16}}>
        <div><p className="sl" style={{marginBottom:8}}>History</p><h2 className="fd" style={{fontSize:34,fontWeight:700,color:G.white}}>Past Conversions</h2></div>
        <button className="btn-ghost" style={{padding:'10px 14px'}} onClick={load}><RefreshCw size={13}/> Refresh</button>
      </div>

      {loading ? (
        <div style={{display:'flex',justifyContent:'center',padding:80}}><Loader2 size={26} color={G.gold} className="spin"/></div>
      ) : hist.length===0 ? (
        <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:56,textAlign:'center'}}>
          <History size={34} color={G.grayDim} style={{margin:'0 auto 18px'}}/>
          <h3 style={{fontSize:17,fontWeight:600,color:G.white,marginBottom:8}}>No conversions yet</h3>
          <p style={{fontSize:13,color:G.gray,marginBottom:24}}>Upload your first image to get started.</p>
          <button className="btn-gold" style={{padding:'10px 24px'}} onClick={goConvert}><Upload size={13}/> Go to Convert</button>
        </div>
      ) : (
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(270px,1fr))',gap:16}}>
          {hist.map(h=>(
            <motion.div key={h.id} initial={{opacity:0,scale:.97}} animate={{opacity:1,scale:1}}
              style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:22,transition:'border-color .2s'}}
              whileHover={{borderColor:`${G.gold}44`} as any}>
              <div style={{display:'flex',justifyContent:'space-between',marginBottom:14}}>
                <div style={{width:38,height:38,borderRadius:8,background:`${G.gold}15`,border:`1px solid ${G.gold}33`,display:'flex',alignItems:'center',justifyContent:'center'}}><FileSpreadsheet size={17} color={G.gold}/></div>
                <div style={{display:'flex',gap:6}}>
                  <button onClick={()=>dl(h)} style={{background:'none',border:`1px solid ${G.border}`,borderRadius:6,padding:'5px 7px',cursor:'pointer',color:G.gray,display:'flex'}}><Download size={12}/></button>
                  <button onClick={()=>del(h.id)} style={{background:'none',border:`1px solid ${G.border}`,borderRadius:6,padding:'5px 7px',cursor:'pointer',color:G.gray,display:'flex'}}><Trash2 size={12}/></button>
                </div>
              </div>
              <p style={{fontSize:12,fontWeight:600,color:G.white,marginBottom:4,wordBreak:'break-all'}}>{h.fileName}</p>
              <p className="fm" style={{fontSize:10,color:G.gray,marginBottom:10}}>{h.rowCount} rows extracted</p>
              <div className="fm" style={{fontSize:9,color:G.grayDim}}>{h.timestamp?.toDate().toLocaleString()}</div>
            </motion.div>
          ))}
        </div>
      )}
    </div>
  );
};

// ─── Account Page ─────────────────────────────────────────────────────────────
const AccountPage = ({addToast}:{addToast:(t:'success'|'error'|'info',m:string)=>void}) => {
  const {user,profile} = useAuth();
  const [editN,   setEditN]   = useState(false);
  const [newName, setNewName] = useState(profile?.displayName||'');
  const [pwForm,  setPwForm]  = useState(false);
  const [pw1,     setPw1]     = useState('');
  const [pw2,     setPw2]     = useState('');
  const [confirm, setConfirm] = useState(false);
  const [busy,    setBusy]    = useState(false);

  const saveName = async () => {
    if(!user||!newName.trim()) return;
    await updateProfile(user,{displayName:newName});
    await setDoc(doc(db,'users',user.uid),{displayName:newName},{merge:true});
    setEditN(false); addToast('success','Name updated!');
  };

  const changePw = async () => {
    if(!user||pw1!==pw2){addToast('error','Passwords do not match.');return;}
    setBusy(true);
    try{await updatePassword(user,pw1);addToast('success','Password updated!');setPwForm(false);setPw1('');setPw2('');}
    catch(e:any){addToast('error',e.message);}
    finally{setBusy(false);}
  };

  const delAccount = async () => {
    if(!user) return; setBusy(true);
    try{await deleteDoc(doc(db,'users',user.uid));await deleteUser(user);}
    catch(e:any){addToast('error',e.message);}
    finally{setBusy(false);setConfirm(false);}
  };

  return (
    <div style={{maxWidth:800,margin:'0 auto',padding:'80px 24px 60px'}}>
      <p className="sl" style={{marginBottom:8}}>Account</p>
      <h2 className="fd" style={{fontSize:34,fontWeight:700,color:G.white,marginBottom:36}}>Your Profile</h2>
      <div style={{display:'flex',flexDirection:'column',gap:18}}>

        {/* Profile card */}
        <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:30}}>
          <div style={{display:'flex',alignItems:'center',gap:18,marginBottom:28}}>
            <div style={{width:60,height:60,borderRadius:12,background:`linear-gradient(135deg,${G.goldDark},${G.gold})`,display:'flex',alignItems:'center',justifyContent:'center',fontSize:26,fontWeight:700,color:'#000'}}>{(user?.email||'U')[0].toUpperCase()}</div>
            <div>
              <div style={{fontSize:17,fontWeight:600,color:G.white,marginBottom:4}}>{profile?.displayName||'User'}</div>
              <div style={{fontSize:13,color:G.gray}}>{user?.email}</div>
              <div className="tag" style={{marginTop:8}}>{profile?.plan==='pro'?'⭐ Pro':'Free Plan'}</div>
            </div>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))',gap:14}}>
            {[{l:'Conversions',v:profile?.totalConversions||0},{l:'Member Since',v:profile?.createdAt?new Date(profile.createdAt).toLocaleDateString():'—'},{l:'Plan',v:profile?.plan==='pro'?'Pro':'Free'}].map(s=>(
              <div key={s.l} style={{background:G.bgCard2,border:`1px solid ${G.border}`,borderRadius:8,padding:'14px 18px'}}>
                <div className="fm" style={{fontSize:9,color:G.grayDim,letterSpacing:'0.1em',textTransform:'uppercase',marginBottom:8}}>{s.l}</div>
                <div style={{fontSize:20,fontWeight:700,color:G.gold}}>{s.v}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Name */}
        <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:24}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:14}}>
            <h3 style={{fontSize:14,fontWeight:600,color:G.white}}>Display Name</h3>
            {!editN&&<button className="btn-ghost" style={{padding:'5px 12px',fontSize:10}} onClick={()=>{setEditN(true);setNewName(profile?.displayName||'');}}><Edit2 size={11}/> Edit</button>}
          </div>
          {editN ? (
            <div style={{display:'flex',gap:10}}>
              <input className="input-field" value={newName} onChange={e=>setNewName(e.target.value)} style={{flex:1}}/>
              <button className="btn-gold" style={{padding:'10px 16px'}} onClick={saveName}><CheckCircle2 size={13}/></button>
              <button className="btn-ghost" style={{padding:'10px 12px'}} onClick={()=>setEditN(false)}><X size={13}/></button>
            </div>
          ) : <p style={{fontSize:13,color:G.gray}}>{profile?.displayName||'—'}</p>}
        </div>

        {/* Password */}
        <div style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:12,padding:24}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:14}}>
            <h3 style={{fontSize:14,fontWeight:600,color:G.white}}>Password</h3>
            <button className="btn-ghost" style={{padding:'5px 12px',fontSize:10}} onClick={()=>setPwForm(p=>!p)}><Lock size={11}/> {pwForm?'Cancel':'Change'}</button>
          </div>
          <AnimatePresence>
            {pwForm && (
              <motion.div initial={{height:0,opacity:0}} animate={{height:'auto',opacity:1}} exit={{height:0,opacity:0}} style={{overflow:'hidden'}}>
                <div style={{display:'flex',flexDirection:'column',gap:10}}>
                  <input className="input-field" type="password" placeholder="New password" value={pw1} onChange={e=>setPw1(e.target.value)}/>
                  <input className="input-field" type="password" placeholder="Confirm password" value={pw2} onChange={e=>setPw2(e.target.value)}/>
                  <button className="btn-gold" style={{padding:'10px 20px',alignSelf:'flex-start'}} onClick={changePw} disabled={busy}>{busy?<Loader2 size={13} className="spin"/>:'Update Password'}</button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Danger */}
        <div style={{background:`${G.error}08`,border:`1px solid ${G.error}22`,borderRadius:12,padding:24}}>
          <h3 style={{fontSize:13,fontWeight:600,color:G.error,marginBottom:6}}>Danger Zone</h3>
          <p style={{fontSize:12,color:G.gray,marginBottom:14}}>Permanently delete your account and all data. Cannot be undone.</p>
          <button className="btn-ghost" style={{padding:'8px 16px',color:G.error,borderColor:`${G.error}44`}} onClick={()=>setConfirm(true)}><Trash2 size={13}/> Delete Account</button>
        </div>
      </div>

      {/* Confirm modal */}
      <AnimatePresence>
        {confirm && (
          <motion.div className="modal-bg" initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} onClick={()=>setConfirm(false)}>
            <motion.div initial={{scale:.9}} animate={{scale:1}} exit={{scale:.9}} onClick={e=>e.stopPropagation()}
              style={{background:G.bgCard,border:`1px solid ${G.error}44`,borderRadius:16,padding:36,maxWidth:380,width:'100%'}}>
              <AlertCircle size={28} color={G.error} style={{marginBottom:14}}/>
              <h3 style={{fontSize:17,fontWeight:700,color:G.white,marginBottom:10}}>Delete account?</h3>
              <p style={{fontSize:13,color:G.gray,marginBottom:24,lineHeight:1.6}}>This permanently deletes your account and all history. This cannot be undone.</p>
              <div style={{display:'flex',gap:10}}>
                <button className="btn-ghost" style={{flex:1,justifyContent:'center',padding:'11px 0'}} onClick={()=>setConfirm(false)}>Cancel</button>
                <button className="btn-gold" style={{flex:1,justifyContent:'center',padding:'11px 0',background:G.error,boxShadow:'none'}} onClick={delAccount} disabled={busy}>{busy?<Loader2 size={13} className="spin"/>:'Yes, Delete'}</button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
};

// ─── Auth Modal ───────────────────────────────────────────────────────────────
const AuthModal = ({open,onClose,onSuccess}:{open:boolean;onClose:()=>void;onSuccess:()=>void}) => {
  const [mode,  setMode]  = useState<'login'|'signup'>('login');
  const [email, setEmail] = useState('');
  const [pw,    setPw]    = useState('');
  const [name,  setName]  = useState('');
  const [busy,  setBusy]  = useState(false);
  const [err,   setErr]   = useState<string|null>(null);
  const [showPw,setShowPw]= useState(false);

  const submit = async (e:React.FormEvent) => {
    e.preventDefault(); setBusy(true); setErr(null);
    try {
      if(mode==='signup'){const c=await createUserWithEmailAndPassword(auth,email,pw);await updateProfile(c.user,{displayName:name});}
      else await signInWithEmailAndPassword(auth,email,pw);
      onClose(); onSuccess();
    } catch(e:any){setErr(e.message);}
    finally{setBusy(false);}
  };

  const google = async () => {
    try{await signInWithPopup(auth,new GoogleAuthProvider());onClose();onSuccess();}
    catch(e:any){setErr(e.message);}
  };

  if(!open) return null;

  return (
    <div className="modal-bg" onClick={onClose}>
      <motion.div initial={{opacity:0,scale:.95,y:20}} animate={{opacity:1,scale:1,y:0}} onClick={e=>e.stopPropagation()}
        style={{background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:16,padding:36,width:'100%',maxWidth:400,position:'relative'}}>
        <button onClick={onClose} style={{position:'absolute',top:14,right:14,background:'none',border:'none',cursor:'pointer',color:G.gray}}><X size={16}/></button>

        <div style={{textAlign:'center',marginBottom:28}}>
          <Logo/><br/>
          <h2 style={{fontSize:18,fontWeight:700,color:G.white,marginTop:16}}>{mode==='login'?'Welcome back':'Create account'}</h2>
          <p style={{fontSize:12,color:G.gray,marginTop:5}}>{mode==='login'?'Sign in to access your history':'Start converting images to Excel'}</p>
        </div>

        {err && <div style={{display:'flex',alignItems:'center',gap:8,padding:'10px 14px',background:`${G.error}12`,border:`1px solid ${G.error}33`,borderRadius:8,marginBottom:16}}><AlertCircle size={13} color={G.error}/><span style={{fontSize:12,color:G.error}}>{err}</span></div>}

        <form onSubmit={submit} style={{display:'flex',flexDirection:'column',gap:12}}>
          {mode==='signup' && <input className="input-field" required value={name} onChange={e=>setName(e.target.value)} placeholder="Full name"/>}
          <input className="input-field" type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="Email address"/>
          <div style={{position:'relative'}}>
            <input className="input-field" type={showPw?'text':'password'} required value={pw} onChange={e=>setPw(e.target.value)} placeholder="Password" style={{paddingRight:40}}/>
            <button type="button" onClick={()=>setShowPw(p=>!p)} style={{position:'absolute',right:12,top:'50%',transform:'translateY(-50%)',background:'none',border:'none',cursor:'pointer',color:G.gray}}>
              {showPw?<EyeOff size={14}/>:<Eye size={14}/>}
            </button>
          </div>
          <button className="btn-gold" style={{padding:'12px 0',justifyContent:'center',marginTop:4}} disabled={busy} type="submit">
            {busy?<Loader2 size={14} className="spin"/>:(mode==='login'?'Sign In':'Create Account')}
          </button>
        </form>

        <div style={{display:'flex',alignItems:'center',gap:10,margin:'20px 0'}}>
          <div style={{flex:1,height:1,background:G.border}}/><span style={{fontSize:11,color:G.grayDim}}>or</span><div style={{flex:1,height:1,background:G.border}}/>
        </div>

        <button className="btn-ghost" style={{width:'100%',justifyContent:'center',padding:'11px 0'}} onClick={google}>
          <svg width="14" height="14" viewBox="0 0 24 24"><path fill={G.gray} d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill={G.gray} d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill={G.gray} d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z"/><path fill={G.gray} d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg>
          Continue with Google
        </button>

        <button onClick={()=>setMode(m=>m==='login'?'signup':'login')} style={{display:'block',width:'100%',textAlign:'center',marginTop:18,background:'none',border:'none',cursor:'pointer',fontSize:12,color:G.gold}}>
          {mode==='login'?"Don't have an account? Sign up":'Already have an account? Sign in'}
        </button>
      </motion.div>
    </div>
  );
};

// ─── Chatbot ──────────────────────────────────────────────────────────────────
const Chatbot = () => {
  const [open,  setOpen]  = useState(false);
  const [msgs,  setMsgs]  = useState<ChatMsg[]>([{role:'model',content:"Hi! I'm the Img2XL assistant. How can I help?",ts:Date.now()}]);
  const [input, setInput] = useState('');
  const [busy,  setBusy]  = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(()=>{ if(ref.current) ref.current.scrollTop=ref.current.scrollHeight; },[msgs,busy]);

  const quick = ['How do I upload an image?','What formats are supported?','How to download Excel?'];

  const send = async (text:string) => {
    if(!text.trim()||busy) return;
    const um:ChatMsg={role:'user',content:text,ts:Date.now()};
    setMsgs(p=>[...p,um]); setInput(''); setBusy(true);
    try {
      const res = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: [...msgs, um].map(m => ({ role: m.role === 'model' ? 'assistant' : m.role, content: m.content })),
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Chat failed');
      setMsgs(p=>[...p,{role:'model',content:json.content||'Sorry, no response.',ts:Date.now()}]);
    } catch(e:any) { setMsgs(p=>[...p,{role:'model',content:`Error: ${e.message}`,ts:Date.now()}]); }
    finally { setBusy(false); }
  };

  return (
    <div style={{position:'fixed',bottom:24,right:24,zIndex:150,display:'flex',flexDirection:'column',alignItems:'flex-end',gap:12}}>
      <AnimatePresence>
        {open && (
          <motion.div initial={{opacity:0,scale:.9,y:16}} animate={{opacity:1,scale:1,y:0}} exit={{opacity:0,scale:.9,y:16}}
            style={{width:340,background:G.bgCard,border:`1px solid ${G.border}`,borderRadius:16,overflow:'hidden',boxShadow:`0 24px 64px rgba(0,0,0,.6)`}}>
            {/* Header */}
            <div style={{padding:'12px 16px',borderBottom:`1px solid ${G.border}`,display:'flex',alignItems:'center',justifyContent:'space-between',background:G.bgCard2}}>
              <div style={{display:'flex',alignItems:'center',gap:10}}>
                <div style={{width:26,height:26,borderRadius:7,background:`linear-gradient(135deg,${G.goldDark},${G.gold})`,display:'flex',alignItems:'center',justifyContent:'center'}}><Sparkles size={12} color="#000"/></div>
                <div><div style={{fontSize:12,fontWeight:700,color:G.white}}>AI Assistant</div><div className="fm" style={{fontSize:9,color:G.gold}}>IMG2XL</div></div>
              </div>
              <div style={{display:'flex',gap:6}}>
                <button onClick={()=>setMsgs([{role:'model',content:'Chat cleared. How can I help?',ts:Date.now()}])} style={{background:'none',border:'none',cursor:'pointer',color:G.gray,padding:3}}><RefreshCw size={12}/></button>
                <button onClick={()=>setOpen(false)} style={{background:'none',border:'none',cursor:'pointer',color:G.gray,padding:3}}><X size={13}/></button>
              </div>
            </div>
            {/* Messages */}
            <div ref={ref} style={{height:300,overflowY:'auto',padding:14,display:'flex',flexDirection:'column',gap:10}}>
              {msgs.map((m,i)=>(
                <div key={i} style={{display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start'}}>
                  <div className={m.role==='user'?'chat-user':'chat-bot'} style={{maxWidth:'85%',padding:'9px 12px',fontSize:12,lineHeight:1.6,color:G.white}}>{m.content}</div>
                </div>
              ))}
              {busy && <div style={{display:'flex',gap:4,padding:'9px 12px'}}>{[0,1,2].map(i=><div key={i} style={{width:5,height:5,borderRadius:'50%',background:G.gold,animation:`pulse-dot 1s ${i*.2}s infinite`}}/>)}</div>}
            </div>
            {/* Quick replies */}
            <div style={{padding:'0 10px 8px',display:'flex',flexWrap:'wrap',gap:5}}>
              {quick.map(q=><button key={q} onClick={()=>send(q)} style={{background:`${G.gold}10`,border:`1px solid ${G.gold}33`,borderRadius:5,padding:'4px 8px',fontSize:9,color:G.gold,cursor:'pointer',fontFamily:"'DM Sans',sans-serif"}}>{q}</button>)}
            </div>
            {/* Input */}
            <div style={{padding:'8px 10px',borderTop:`1px solid ${G.border}`,display:'flex',gap:7}}>
              <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&send(input)} className="input-field" placeholder="Ask anything…" style={{flex:1,padding:'8px 12px',fontSize:12}}/>
              <button className="btn-gold" style={{padding:'8px 12px',flexShrink:0}} onClick={()=>send(input)}><Send size={12}/></button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <button className="btn-gold" onClick={()=>setOpen(p=>!p)}
        style={{width:50,height:50,borderRadius:'50%',padding:0,justifyContent:'center',fontSize:0}}>
        <AnimatePresence mode="wait">
          {open
            ? <motion.div key="x"  initial={{scale:0,rotate:-90}} animate={{scale:1,rotate:0}} exit={{scale:0}}><X size={19} color="#000"/></motion.div>
            : <motion.div key="ch" initial={{scale:0,rotate:90}}  animate={{scale:1,rotate:0}} exit={{scale:0}}><MessageSquare size={19} color="#000"/></motion.div>
          }
        </AnimatePresence>
      </button>
    </div>
  );
};

// ─── Root ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab,     setTab]     = useState('landing');
  const [showAuth,setShowAuth]= useState(false);
  const [files,   setFiles]   = useState<FileItem[]>([]);
  const [toasts,  setToasts]  = useState<Toast[]>([]);

  const addToast = (type:'success'|'error'|'info', msg:string) => {
    const id = Math.random().toString(36).substr(2,9);
    setToasts(p=>[...p,{id,type,msg}]);
    setTimeout(()=>setToasts(p=>p.filter(t=>t.id!==id)),4000);
  };

  const logout = async () => {
    await signOut(auth); setTab('landing'); setFiles([]);
    addToast('info','Signed out successfully.');
  };

  const start = () => { if(auth.currentUser) setTab('convert'); else setShowAuth(true); };

  return (
    <AuthProvider>
      <style>{GLOBAL_CSS}</style>
      <div style={{minHeight:'100vh',background:G.bg,color:G.white,overflowX:'hidden'}}>
        <Navbar tab={tab} setTab={setTab} onLogin={()=>setShowAuth(true)} onLogout={logout}/>

        <AnimatePresence mode="wait">
          {tab==='landing' && (
            <motion.div key="land" initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} transition={{duration:.25}}>
              <Landing onStart={start}/>
            </motion.div>
          )}
          {tab==='convert' && (
            <motion.div key="conv" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} exit={{opacity:0}} transition={{duration:.25}}>
              <Convert files={files} setFiles={setFiles} addToast={addToast}/>
            </motion.div>
          )}
          {tab==='history' && (
            <motion.div key="hist" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} exit={{opacity:0}} transition={{duration:.25}}>
              <HistoryPage goConvert={()=>setTab('convert')}/>
            </motion.div>
          )}
          {tab==='account' && (
            <motion.div key="acc" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} exit={{opacity:0}} transition={{duration:.25}}>
              <AccountPage addToast={addToast}/>
            </motion.div>
          )}
          {(tab==='about'||tab==='pricing') && (
            <motion.div key="ap" initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} transition={{duration:.25}}>
              <Landing onStart={start}/>
            </motion.div>
          )}
        </AnimatePresence>

        <Chatbot/>
        <AuthModal open={showAuth} onClose={()=>setShowAuth(false)} onSuccess={()=>setTab('convert')}/>
        <Toasts items={toasts} dismiss={id=>setToasts(p=>p.filter(t=>t.id!==id))}/>
      </div>
    </AuthProvider>
  );
}
