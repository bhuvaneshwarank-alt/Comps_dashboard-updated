import React,{useCallback,useEffect,useMemo,useRef,useState} from 'react';
import {createRoot} from 'react-dom/client';
import {api,setCsrf} from './api.js';
import GoogleSignIn from './components/GoogleSignIn.jsx';
import './style.css';

const Icon=({name})=>{
 const paths={
  home:<><rect x="3" y="3" width="8" height="8" rx="2"/><rect x="13" y="3" width="8" height="5" rx="2"/><rect x="13" y="10" width="8" height="11" rx="2"/><rect x="3" y="13" width="8" height="8" rx="2"/></>,
  records:<><path d="M8 6h13M8 12h13M8 18h13"/><path d="M3 6h.01M3 12h.01M3 18h.01"/></>,
  search:<><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></>,
  shield:<><path d="M12 22s8-4 8-11V5l-8-3-8 3v6c0 7 8 11 8 11Z"/><path d="m9 12 2 2 4-4"/></>,
  chevron:<path d="m9 18 6-6-6-6"/>,close:<><path d="m18 6-12 12M6 6l12 12"/></>,
  sun:<><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M4.93 4.93l1.41 1.41m11.32 11.32 1.41 1.41M2 12h2m16 0h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></>,
  logout:<><path d="M10 17l5-5-5-5"/><path d="M15 12H3"/><path d="M12 3h6a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-6"/></>
 };
 return <svg className="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]||paths.records}</svg>;
};
const fmtDate=value=>{if(!value)return '—';const d=new Date(value);return Number.isNaN(d.getTime())?String(value):new Intl.DateTimeFormat('en-IN',{day:'2-digit',month:'short',year:'numeric'}).format(d);};
const labelRisk=value=>String(value||'Not rated');
const labelStatus=value=>String(value||'Unknown');
const safeFields=record=>Object.entries(record).filter(([key,value])=>key&&value!==''&&value!==null&&value!==undefined&&!/^password|otp|token|secret/i.test(key));

function SignIn({config,onLogin}){
 const [identity,setIdentity]=useState(null),[password,setPassword]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const [showPassword,setShowPassword]=useState(false),[persona,setPersona]=useState('admin');
 const googleCredential=useCallback(value=>{setIdentity(value);setError('');},[]);
 async function submit(event){event.preventDefault();setBusy(true);setError('');try{const body=config.mode==='demo'?{persona}:{idToken:identity,password};const result=await api('/auth/login',{method:'POST',body});setCsrf(result.csrf);onLogin(result.access,{demo:result.demo});}catch(e){setError(e.message);setPassword('');}finally{setBusy(false);}}
 return <main className="login-page"><div className="login-orb orb-a"/><div className="login-orb orb-b"/><section className="login-card">
   <div className="brand-mark"><Icon name="shield"/><span>DIGICLARITY</span></div>
   <p className="eyebrow">COMPLIANCE MANAGEMENT PORTAL</p><h1>Welcome back.</h1>
   <p className="login-copy">Your compliance work, organized in one secure place.</p>
   {config.mode==='demo'?<form onSubmit={submit} className="signin-form"><div className="demo-flag"><span className="demo-dot"/> Preview with synthetic data</div><label htmlFor="persona">Choose a demo role</label><select id="persona" value={persona} onChange={e=>setPersona(e.target.value)}><option value="admin">Administrator · full sample register</option><option value="hr">Department User · HR records only</option></select><button className="primary-button wide" disabled={busy}>{busy?'Opening…':'Open demo dashboard'}<Icon name="chevron"/></button></form>:
     <><div className="company-login"><GoogleSignIn clientId={config.googleClientId} onCredential={googleCredential}/>{identity&&<span className="verified"><i/>Company identity verified. Enter your dashboard password below.</span>}</div>{identity&&<form onSubmit={submit} className="signin-form"><label htmlFor="password">Dashboard password</label><div className="password-wrap"><input id="password" autoComplete="current-password" type={showPassword?'text':'password'} maxLength={128} value={password} onChange={e=>setPassword(e.target.value)} required/><button className="show-password" type="button" aria-label={showPassword?'Hide password':'Show password'} onClick={()=>setShowPassword(x=>!x)}>{showPassword?'Hide':'Show'}</button></div><button className="primary-button wide" disabled={busy||password.length===0}>{busy?'Verifying…':'Sign in'}<Icon name="chevron"/></button></form>}</>}
   {error&&<p className="form-error" role="alert">{error}</p>}
   <div className="login-foot"><span>Protected workspace</span><span>·</span><span>Asia / Kolkata</span></div>
 </section><footer className="login-legal">DigiClarity internal compliance workspace</footer></main>;
}

function App(){
 const [config,setConfig]=useState(null),[access,setAccess]=useState(null),[demo,setDemo]=useState(false),[initializing,setInitializing]=useState(true),[page,setPage]=useState('overview');
 const [overview,setOverview]=useState(null),[filters,setFilters]=useState({framework:'',department:'',status:'',risk:'',search:''}),[result,setResult]=useState({items:[],total:0,page:1,pageSize:25});
 const [selected,setSelected]=useState(null),[history,setHistory]=useState([]),[detailTab,setDetailTab]=useState('overview'),[loading,setLoading]=useState(false),[toast,setToast]=useState('');
 const [theme,setTheme]=useState(()=>localStorage.getItem('dc-cloud-theme')||'light');
 const [error,setError]=useState(''),[pageNo,setPageNo]=useState(1);
 const [analytics,setAnalytics]=useState(null),[analyticsError,setAnalyticsError]=useState('');
 const timer=useRef(null),lastActivity=useRef(0),abort=useRef(null);
 const notify=useCallback(value=>{setToast(value);clearTimeout(timer.current);timer.current=setTimeout(()=>setToast(''),3000);},[]);
 const expire=useCallback(message=>{setAccess(null);setOverview(null);setSelected(null);setHistory([]);setCsrf('');setError(message||'Your session ended. Sign in again.');},[]);
 const loadOverview=useCallback(async()=>{const data=await api('/overview');setOverview(data);},[]);
 const loadRecords=useCallback(async(nextPage=1)=>{
   abort.current?.abort();const controller=new AbortController();abort.current=controller;setLoading(true);
   try{const q=new URLSearchParams({page:String(nextPage),pageSize:'25',...Object.fromEntries(Object.entries(filters).filter(([,v])=>v))});const response=await fetch(`/api/records?${q}`,{credentials:'same-origin',signal:controller.signal,cache:'no-store'});const data=await response.json();if(!response.ok)throw Object.assign(new Error(data.message||'Records unavailable.'),{code:data.error});setResult(data);setPageNo(data.page);setError('');}
   catch(e){if(e.name!=='AbortError'){if(e.code?.startsWith('SESSION_'))expire(e.message);else setError(e.message);}}
   finally{if(!controller.signal.aborted)setLoading(false);}
 },[filters,expire]);
 useEffect(()=>{api('/config').then(setConfig).catch(()=>setConfig({mode:'demo',googleClientId:''}));api('/session').then(async data=>{setCsrf(data.csrf);setAccess(data.access);await Promise.all([api('/overview').then(setOverview),api('/records?page=1&pageSize=25').then(setResult)]);}).catch(e=>{if(e.code!=='SESSION_REQUIRED'&&e.status!==401)setError(e.message);}).finally(()=>setInitializing(false));},[]);
 useEffect(()=>{if(!access)return;loadOverview().catch(e=>{if(e.status===401)expire(e.message);});loadRecords(1);},[access,loadOverview,loadRecords,expire]);
 useEffect(()=>{if(!access||!selected)return;let active=true;Promise.all([api(`/records/${encodeURIComponent(selected['Record ID'])}`),api(`/records/${encodeURIComponent(selected['Record ID'])}/history`)]).then(([record,events])=>{if(active){setSelected(record.record);setHistory(events.items);}}).catch(e=>{if(active){if(e.status===401)expire(e.message);else if(e.status!==404)setError(e.message);}});return()=>{active=false;};},[access,selected?.['Record ID'],expire]);
 useEffect(()=>{if(!access)return;const onActivity=()=>{const now=Date.now();if(now-lastActivity.current<30000)return;lastActivity.current=now;api('/session/activity',{method:'POST',body:{at:new Date().toISOString()}}).catch(e=>{if(e.status===401)expire(e.message);});};const events=['pointerdown','keydown','touchstart'];events.forEach(e=>window.addEventListener(e,onActivity,{passive:true}));return()=>events.forEach(e=>window.removeEventListener(e,onActivity));},[access,expire]);
 useEffect(()=>{if(!access)return;const id=setInterval(()=>{api('/session').then(()=>{}).catch(e=>{if(e.status===401)expire(e.message);});},45000);return()=>clearInterval(id);},[access,expire]);
 useEffect(()=>{const onKey=e=>{if(e.key==='Escape')setSelected(null);};window.addEventListener('keydown',onKey);document.documentElement.dataset.theme=theme;localStorage.setItem('dc-cloud-theme',theme);return()=>window.removeEventListener('keydown',onKey);},[theme]);
 const overviewSummary=overview?.summary;
 useEffect(()=>{if(!access||page!=='analytics')return;let active=true;setAnalyticsError('');api('/analytics').then(data=>{if(active)setAnalytics(data);}).catch(e=>{if(active)setAnalyticsError(e.message);});return()=>{active=false;};},[access,page]);
 const showPage=useCallback(target=>{setSelected(null);setPage(target);},[]);
 async function signout(){try{await api('/auth/logout',{method:'POST',body:{}});}catch{}setAccess(null);setCsrf('');setOverview(null);setSelected(null);setHistory([]);setPage('overview');}
 const frameworks=overview?.filters.frameworks||[],departments=overview?.filters.departments||[];
 if(initializing)return <div className="app-loading"><span className="spinner"/><span>Connecting to your workspace…</span></div>;
 if(!access)return <SignIn config={config||{mode:'demo',googleClientId:''}} onLogin={(user,meta)=>{setAccess(user);setDemo(meta.demo);setError('');setPage('overview');}}/>;
 const pageTitle=page==='overview'?'Executive overview':page==='analytics'?'Compliance analytics':'Compliance register';
 const rowsForPage=result.items||[],maxPage=Math.max(1,Math.ceil((result.total||0)/25));
 return <div className="app-shell">
  <aside className="sidebar"><a className="brand-lockup" href="#home" onClick={e=>{e.preventDefault();showPage('overview');}}><span className="brand-glyph"><Icon name="shield"/></span><span><b>DIGICLARITY</b><small>COMPLIANCE PORTAL</small></span></a>
   <div className="nav-section">WORKSPACE</div><nav className="nav-list"><button className={`nav-button ${page==='overview'?'active':''}`} onClick={()=>showPage('overview')}><Icon name="home"/>Executive overview</button><button className={`nav-button ${page==='records'?'active':''}`} onClick={()=>showPage('records')}><Icon name="records"/>Compliance register<span className="nav-count">{overviewSummary?.total??'—'}</span></button><button className={`nav-button ${page==='analytics'?'active':''}`} onClick={()=>showPage('analytics')}><Icon name="shield"/>Compliance analytics</button></nav>
   <div className="sidebar-spacer"/><div className="sidebar-scope"><span className="scope-dot"/><div><b>{access.role}</b><small>{access.departmentScoped?access.departments.join(', '):'All departments'}</small></div></div><div className="sidebar-footer">{demo&&<span className="demo-footer-tag">DEMO DATA</span>}<span>Milestone 1 · Read only</span></div>
  </aside>
  <main className="main-area"><header className="topbar"><div className="breadcrumb"><span>Workspace</span><b>/</b><strong>{pageTitle}</strong></div><div className="top-actions">{demo&&<span className="demo-chip"><i/>Synthetic data</span>}<button className="icon-button" title="Toggle appearance" aria-label="Toggle appearance" onClick={()=>setTheme(t=>t==='light'?'dark':'light')}><Icon name="sun"/></button><div className="account-chip"><span className="avatar">{(access.name||access.email||'U').trim()[0].toUpperCase()}</span><span className="account-copy"><b>{access.name}</b><small>{access.role}</small></span></div><button className="icon-button" onClick={signout} title="Sign out" aria-label="Sign out"><Icon name="logout"/></button></div></header>
   <section className="page-content"><div className="page-heading"><div><p className="eyebrow">COMPLIANCE OPERATIONS</p><h1>{pageTitle}</h1><p className="subheading">A clear view of your assigned compliance work.</p></div><span className="updated-label"><span className="live-dot"/>Connected to {demo?'sample workspace':'Google Sheets'}</span></div>
    {error&&<div className="notice-error" role="alert"><Icon name="shield"/><span>{error}</span><button className="text-button" onClick={()=>{setError('');loadRecords(pageNo);}}>Retry</button></div>}
    {page==='overview'?<>
      <div className="hero-card"><div className="hero-copy"><span className="hero-kicker"><Icon name="shield"/> YOUR COMPLIANCE SNAPSHOT</span><h2>Good work starts<br/>with a clear view.</h2><p>Review open findings, check risk exposure, and keep your evidence moving.</p><button className="hero-button" onClick={()=>showPage('records')}>View register <Icon name="chevron"/></button></div><div className="hero-art" aria-hidden="true"><div className="hero-ring ring-1"/><div className="hero-ring ring-2"/><div className="hero-shield"><Icon name="shield"/></div><span className="art-node node-a"/><span className="art-node node-b"/><span className="art-node node-c"/></div></div>
      <div className="section-intro"><div><h2>At a glance</h2><p>Current records within your access scope</p></div><span className="scope-pill">{access.departmentScoped?access.departments.join(' · '):'All departments'}</span></div>
      <div className="kpi-grid"><Kpi label="Total records" value={overviewSummary?.total} note="in your access scope" tone="blue"/><Kpi label="Open items" value={overviewSummary?.open} note="require follow-up" tone="violet"/><Kpi label="High risk" value={overviewSummary?.highRisk} note="open critical or high" tone="rose"/><Kpi label="Overdue actions" value={overviewSummary?.overdueActions} note="past target completion" tone="rose"/><Kpi label="Pending evidence" value={overviewSummary?.pendingEvidence} note="open items need evidence" tone="violet"/><Kpi label="Closed" value={overviewSummary?.closed} note="completed records" tone="green"/></div>
      <div className="lower-grid"><section className="panel framework-panel"><div className="panel-heading"><div><h2>Framework coverage</h2><p>Records grouped by primary framework</p></div><span className="panel-symbol"><Icon name="shield"/></span></div>{overviewSummary?.byFramework?.length?<div className="framework-list">{overviewSummary.byFramework.map((item,i)=><div className="framework-row" key={item.name}><div className="framework-name"><span className={`framework-marker marker-${i%4}`}/><b>{item.name}</b><span>{item.total} records</span></div><div className="bar-track"><span className={`bar-fill fill-${i%4}`} style={{width:`${Math.max(4,item.total/Math.max(overviewSummary.total,1)*100)}%`}}/></div></div>)}</div>:<Empty text="No framework-linked records are in scope."/>}</section>
      <section className="panel status-panel"><div className="panel-heading"><div><h2>Record status</h2><p>Current workflow distribution</p></div><Icon name="records"/></div>{overviewSummary?.byStatus?.length?<div className="status-list">{overviewSummary.byStatus.slice(0,6).map((item,i)=><div className="status-row" key={item.name}><span className={`status-bullet status-${i%5}`}/><span>{item.name}</span><b>{item.total}</b></div>)}</div>:<Empty text="No status information is available."/>}</section></div>
      <section className="panel recent-panel"><div className="panel-heading"><div><h2>Recent records</h2><p>Latest items visible to your role</p></div><button className="text-button" onClick={()=>showPage('records')}>View all <Icon name="chevron"/></button></div><RecordTable rows={rowsForPage.slice(0,5)} loading={loading} onSelect={setSelected} compact/></section>
    </>:page==='analytics'?<>
      <div className="section-intro"><div><h2>Compliance analytics</h2><p>Read-only distributions from records visible to your role and departments</p></div><span className="read-only-pill">Read only</span></div>
      {analyticsError?<p className="form-error" role="alert">{analyticsError}</p>:!analytics?<div className="table-loading"><span className="spinner"/>Loading analytics…</div>:<>
        <div className="kpi-grid"><Kpi label="Records analyzed" value={analytics.total} note="Within your access scope" tone="blue"/></div>
        <div className="analytics-grid">{[['Department','byDepartment'],['Primary framework','byFramework'],['Record type','byRecordType'],['Record status','byStatus'],['Residual risk','byRisk']].map(([label,key])=><section className="panel analytics-panel" key={key}><div className="panel-heading"><h2>{label}</h2></div>{analytics[key].length?analytics[key].map(({name,total})=><div className="analytics-row" key={name}><div className="analytics-row-label"><span>{name}</span><b>{total}</b></div><div className="bar-track"><span className="bar-fill fill-0" style={{width:`${total/Math.max(analytics.total,1)*100}%`}}/></div></div>):<Empty text="No records in this category."/>}</section>)}</div>
        <p className="analytics-notice">{analytics.notice}</p>
      </>}
    </>:<>
      <div className="register-toolbar"><div className="search-box"><Icon name="search"/><input aria-label="Search records" placeholder="Search ID, title or department" value={filters.search} onChange={e=>{setFilters(f=>({...f,search:e.target.value}));setPageNo(1);}}/></div><button className="filter-button" onClick={()=>setFilters({framework:'',department:'',status:'',risk:'',search:''})}>Clear filters</button><a className="filter-button" href="/api/records-export.csv" download="digiclarity-compliance-records.csv">Export scoped CSV</a></div>
      <div className="filter-grid"><Select label="Framework" value={filters.framework} items={frameworks} onChange={v=>setFilters(f=>({...f,framework:v}))}/><Select label="Department" value={filters.department} items={departments} onChange={v=>setFilters(f=>({...f,department:v}))}/><Select label="Status" value={filters.status} items={overview?.filters.statuses||[]} onChange={v=>setFilters(f=>({...f,status:v}))}/><Select label="Risk" value={filters.risk} items={overview?.filters.risks||[]} onChange={v=>setFilters(f=>({...f,risk:v}))}/></div>
      <section className="panel table-panel"><div className="panel-heading"><div><h2>Compliance records</h2><p>{result.total??0} record{result.total===1?'':'s'} match your view</p></div><span className="read-only-pill"><Icon name="shield"/> Read only</span></div><RecordTable rows={rowsForPage} loading={loading} onSelect={setSelected}/><div className="pagination"><span>Page {pageNo} of {maxPage}</span><div><button disabled={pageNo<=1||loading} onClick={()=>{loadRecords(pageNo-1);window.scrollTo({top:0,behavior:'smooth'});}}>Previous</button><button disabled={pageNo>=maxPage||loading} onClick={()=>{loadRecords(pageNo+1);window.scrollTo({top:0,behavior:'smooth'});}}>Next</button></div></div></section>
    </>}
    <footer className="page-footer"><span>Internal compliance workspace</span><span>Current prototype · Read only</span></footer>
   </section>
  </main>
  {selected&&<RecordDialog record={selected} events={history} tab={detailTab} setTab={setDetailTab} onClose={()=>setSelected(null)}/>}
  <div className={`toast ${toast?'visible':''}`} role="status">{toast}</div>
 </div>;
}
function Kpi({label,value,note,tone}){return <article className="kpi-card"><span className={`kpi-icon tone-${tone}`}><Icon name={tone==='green'?'shield':'records'}/></span><span className="kpi-label">{label}</span><strong>{value??'—'}</strong><small>{note}</small></article>;}
function Empty({text}){return <p className="empty-state">{text}</p>;}
function Select({label,value,items,onChange}){return <label className="filter-select"><span>{label}</span><select value={value} onChange={e=>onChange(e.target.value)}><option value="">All {label.toLowerCase()}s</option>{items.map(item=><option key={item} value={item}>{item}</option>)}</select></label>;}
function RecordTable({rows,loading,onSelect,compact=false}){
 if(loading)return <div className="table-loading"><span className="spinner"/>Loading records…</div>;
 if(!rows.length)return <Empty text="No records match these filters."/>;
 return <div className="table-scroll"><table><thead><tr><th>Record</th><th>Department</th><th>Framework</th><th>Status</th><th>Risk</th>{!compact&&<th>Due date</th>}</tr></thead><tbody>{rows.map(r=><tr key={r['Record ID']} onClick={()=>onSelect(r)} tabIndex="0" onKeyDown={e=>e.key==='Enter'&&onSelect(r)}><td><div className="record-title"><b>{r['Record Title']||'Untitled record'}</b><small>{r['Record ID']} · {r['Record Type']||'Record'}</small></div></td><td>{r.Department||'—'}</td><td>{r['Primary Compliance Framework']||'—'}</td><td><span className={`status-tag ${statusTone(r['Record Status'])}`}>{labelStatus(r['Record Status'])}</span></td><td><span className={`risk-tag ${riskTone(r['Residual Risk Rating']||r['Inherent Risk Rating'])}`}>{labelRisk(r['Residual Risk Rating']||r['Inherent Risk Rating'])}</span></td>{!compact&&<td>{fmtDate(r['Due Date'])}</td>}</tr>)}</tbody></table></div>;
}
function statusTone(v){const s=String(v||'').toLowerCase();if(s.includes('closed')||s.includes('complete'))return 'success';if(s.includes('progress'))return 'progress';if(s.includes('overdue')||s.includes('open')||s.includes('pending'))return 'warning';return 'neutral';}
function riskTone(v){const s=String(v||'').toLowerCase();if(s==='critical')return 'critical';if(s==='high')return 'high';if(s==='medium')return 'medium';return 'low';}
function RecordDialog({record,events,tab,setTab,onClose}){
 const items=safeFields(record).filter(([k])=>!['Record ID','Record Title'].includes(k));
 return <div className="dialog-backdrop" onMouseDown={e=>e.target===e.currentTarget&&onClose()}><section className="detail-dialog" role="dialog" aria-modal="true" aria-labelledby="record-heading"><header className="dialog-header"><div><p className="eyebrow">{record['Record Type']||'COMPLIANCE RECORD'} · {record['Record ID']}</p><h2 id="record-heading">{record['Record Title']||'Untitled record'}</h2></div><button className="icon-button" aria-label="Close record details" onClick={onClose}><Icon name="close"/></button></header><div className="dialog-pills"><span className={`status-tag ${statusTone(record['Record Status'])}`}>{labelStatus(record['Record Status'])}</span><span className={`risk-tag ${riskTone(record['Residual Risk Rating']||record['Inherent Risk Rating'])}`}>{labelRisk(record['Residual Risk Rating']||record['Inherent Risk Rating'])} risk</span><span className="scope-pill">{record.Department||'Unassigned'}</span></div><div className="dialog-tabs" role="tablist"><button role="tab" aria-selected={tab==='overview'} className={tab==='overview'?'active':''} onClick={()=>setTab('overview')}>Record details</button><button role="tab" aria-selected={tab==='history'} className={tab==='history'?'active':''} onClick={()=>setTab('history')}>History <span>{events.length}</span></button></div>{tab==='overview'?<div className="detail-grid">{items.map(([key,value])=><div className="detail-field" key={key}><span>{key}</span><strong>{/^.*date$/i.test(key)?fmtDate(value):String(value)}</strong></div>)}</div>:<div className="history-list">{events.length?events.map((event,i)=><article className="history-event" key={event['Event ID']||`${event.Timestamp}-${i}`}><span className="history-dot"/><div><b>{event.Action||event.Event||'Record updated'}</b><p>{event['Change Reason']||event['Reviewer Comment']||event['New Value']||'—'}</p><small>{fmtDate(event.Timestamp)}{event['Changed By']?` · ${event['Changed By']}`:''}</small></div></article>):<Empty text="No record history is available."/>}</div>}<footer className="dialog-footer"><span><Icon name="shield"/> Read only · edits are disabled in this milestone</span><button className="secondary-button" onClick={onClose}>Close</button></footer></section></div>;
}

createRoot(document.getElementById('root')).render(<React.StrictMode><App/></React.StrictMode>);
