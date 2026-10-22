import 'dotenv/config';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import { SheetsRepository,MemoryRepository } from './repository.js';
import { demoTables } from './demo-data.js';
import { makeAuth } from './auth.js';
import { AppError,check } from './errors.js';
import { filterRecords,objects,recordById,summary,unique,visibleRecords,RECORD_TABLES } from './records.js';
import { AUTH_TABLES } from './access.js';
import { analyticsForVisibleRecords } from './analytics.js';

const mode=process.env.APP_MODE||'demo';
check(['demo','google'].includes(mode),500,'INVALID_CONFIG','APP_MODE must be demo or google.');
if(mode==='demo'&&(process.env.K_SERVICE||process.env.NODE_ENV==='production'))throw Error('Synthetic demo mode is disabled in Cloud Run/production. Set APP_MODE=google and configure Google identity and Sheets.');
if(mode==='google'&&!process.env.GOOGLE_CLIENT_ID)throw Error('GOOGLE_CLIENT_ID is required in google mode.');
const origin=process.env.APP_ORIGIN||'http://localhost:5173';
const domains=(process.env.ALLOWED_DOMAINS||'digiclarity.in,clarityrcm.com').split(',').map(s=>s.trim().toLowerCase()).filter(Boolean);
check(domains.length>0,500,'INVALID_CONFIG','At least one approved domain is required.');
const repo=mode==='demo'?new MemoryRepository(demoTables()):new SheetsRepository(process.env.GOOGLE_SHEETS_ID);
const config={mode,origin,domains,googleClientId:process.env.GOOGLE_CLIENT_ID||'',passwordPepper:process.env.DGS_DASHBOARD_PASSWORD_PEPPER_V1||'',iamStorageOwner:process.env.IAM_STORAGE_OWNER||'',serviceAccount:process.env.CLOUD_RUN_SERVICE_ACCOUNT_EMAIL||''};
if(mode==='google'&&!config.passwordPepper)throw Error('Copy DGS_DASHBOARD_PASSWORD_PEPPER_V1 securely from existing Script Properties. Never create a replacement value.');
const auth=makeAuth({repo,config});
const app=express(), port=Number(process.env.PORT)||8080, cookieSecure=mode==='google'&&!origin.includes('localhost');
app.disable('x-powered-by');
app.use(helmet({contentSecurityPolicy:{directives:{defaultSrc:["'self'"],scriptSrc:["'self'","https://accounts.google.com/gsi/client","'unsafe-inline'"],styleSrc:["'self'","'unsafe-inline'","https://fonts.googleapis.com"],fontSrc:["'self'","data:","https://fonts.gstatic.com"],imgSrc:["'self'","data:","https://lh3.googleusercontent.com"],frameSrc:["https://accounts.google.com"],connectSrc:["'self'","https://accounts.google.com/gsi/"]}}}));
app.use((req,res,next)=>{if(req.path.startsWith('/api/'))res.set('Cache-Control','no-store');next();});
app.use(express.json({limit:'16kb',strict:true}));
app.use('/api/auth/login',rateLimit({windowMs:15*60*1000,limit:30,standardHeaders:true,legacyHeaders:false}));
app.use((req,res,next)=>{if(req.path.startsWith('/api/')&&['POST','PUT','PATCH','DELETE'].includes(req.method)){if(req.headers.origin!==origin)return res.status(403).json({error:'ORIGIN_REJECTED',message:'Refresh the dashboard and retry.'});}next();});
const cookie={httpOnly:true,sameSite:'strict',secure:cookieSecure,path:'/'};
function sendError(res,error){const known=error instanceof AppError;return res.status(known?error.status:500).json({error:known?error.code:'INTERNAL_ERROR',message:known?error.message:'The request could not be completed.'});}
function handle(fn){return async(req,res)=>{try{await fn(req,res);}catch(error){if(!(error instanceof AppError))console.error('Request failed',{route:req.route?.path,code:error?.code||'INTERNAL_ERROR'});sendError(res,error);}};}
async function requestTables(names){return repo.tables(names);}
function assertActivityOrigin(req){check(req.headers.origin===origin,403,'ORIGIN_REJECTED','Refresh the dashboard and retry.');}

app.get('/api/config',(req,res)=>res.json({mode,googleClientId:config.googleClientId}));
app.post('/api/auth/login',handle(async(req,res)=>{
  assertActivityOrigin(req);
  const session=await auth.login(req.body||{});
  res.cookie('dc_session',session.id,{...cookie,maxAge:24*60*60*1000});
  res.json({access:session.access,csrf:session.csrf,demo:mode==='demo'});
}));
app.post('/api/auth/logout',handle(async(req,res)=>{
  assertActivityOrigin(req);
  const {id,session}=await auth.authenticate(req);auth.requireCsrf(req,session);auth.revoke(id);res.clearCookie('dc_session',cookie);res.status(204).end();
}));
app.get('/api/session',handle(async(req,res)=>{
  const {access,session}=await auth.authenticate(req);
  res.json({access,csrf:session.csrf,mode});
}));
app.post('/api/session/activity',handle(async(req,res)=>{
  assertActivityOrigin(req);
  const {access,session}=await auth.authenticate(req);auth.requireCsrf(req,session);auth.touch(session);res.json({ok:true,access});
}));
const signedIn=handle(async(req,res,callback)=>{const ctx=await auth.authenticate(req);await callback(ctx);});
const protectedGet=callback=>handle(async(req,res)=>{const {access}=await auth.authenticate(req);await callback(req,res,access);});
app.get('/api/overview',protectedGet(async(req,res,access)=>{
  check(access.permissions.viewDashboard,403,'ACCESS_DENIED','Your role cannot open the dashboard overview.');
  const tables=await requestTables(RECORD_TABLES),records=visibleRecords(tables,access);
  res.json({summary:summary(records),filters:{frameworks:unique(records,'Primary Compliance Framework'),departments:unique(records,'Department'),statuses:unique(records,'Record Status'),risks:unique(records,'Residual Risk Rating')}});
}));
app.get('/api/analytics',protectedGet(async(req,res,access)=>{
  check(access.permissions.viewDashboard&&access.permissions.viewRecords,403,'ACCESS_DENIED','Analytics are not available to this role.');
  const tables=await requestTables(RECORD_TABLES);
  res.json(analyticsForVisibleRecords(visibleRecords(tables,access)));
}));
app.get('/api/records',protectedGet(async(req,res,access)=>{
  const tables=await requestTables(RECORD_TABLES),all=visibleRecords(tables,access),filtered=filterRecords(all,req.query);
  const page=Math.max(1,Math.min(100000,Number.parseInt(req.query.page||'1',10)||1)),pageSize=Math.max(1,Math.min(100,Number.parseInt(req.query.pageSize||'25',10)||25));
  res.json({items:filtered.slice((page-1)*pageSize,page*pageSize),page,pageSize,total:filtered.length,summary:summary(all)});
}));
// Export is derived only from the server-authorized record set, never from client-provided rows.
// Prefix dangerous spreadsheet formula characters to avoid CSV formula injection.
function csvCell(value){
  const raw=String(value??'');
  const safe=/^[\s]*[=+@-]/.test(raw)?`'${raw}`:raw;
  return `\"${safe.replaceAll('\"','\"\"')}\"`;
}
app.get('/api/records-export.csv',protectedGet(async(req,res,access)=>{
  check(access.permissions.viewRecords===true,403,'ACCESS_DENIED','You cannot export records.');
  // Exports do not include sensitive or unexpected columns added to the source sheet.
  const columns=['Record ID','Record Title','Record Type','Department','Record Status','Primary Compliance Framework','Requirement / Control ID','Residual Risk Rating','Owner','Target Completion Date','Action Status'];
  const tables=await requestTables(RECORD_TABLES);
  const scoped=visibleRecords(tables,access);
  check(scoped.length<=25000,413,'EXPORT_TOO_LARGE','Export exceeds the safe row limit.');
  const rows=[columns,...scoped.map(record=>columns.map(name=>record[name]??''))];
  res.set('Content-Type','text/csv; charset=utf-8');
  res.set('Content-Disposition','attachment; filename=digiclarity-compliance-records.csv');
  res.send('\uFEFF'+rows.map(row=>row.map(csvCell).join(',')).join('\r\n')+'\r\n');
}));
app.get('/api/records/:id',protectedGet(async(req,res,access)=>{
  const tables=await requestTables(RECORD_TABLES),record=recordById(visibleRecords(tables,access),req.params.id);
  res.json({record});
}));
app.get('/api/records/:id/history',protectedGet(async(req,res,access)=>{
  check(access.permissions.viewRecords||access.permissions.viewApprovalHistory,403,'ACCESS_DENIED','Your role cannot view record history.');
  const [base,historyTables]=await Promise.all([requestTables(RECORD_TABLES),requestTables(['Record_History'])]);
  recordById(visibleRecords(base,access),req.params.id);
  if(historyTables.Record_History===null)return res.json({items:[]});
  const items=objects(historyTables.Record_History,'Record_History',['Record ID','Timestamp']).filter(row=>row['Record ID']===req.params.id).sort((a,b)=>String(b.Timestamp).localeCompare(String(a.Timestamp)));
  res.json({items});
}));

if(process.env.NODE_ENV==='production'){
  const dist=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../../dist');
  app.use(express.static(dist,{index:false,immutable:true,maxAge:'1h'}));
  app.get('*splat',(req,res)=>res.sendFile(path.join(dist,'index.html')));
}
app.use((req,res)=>res.status(404).json({error:'NOT_FOUND',message:'Route not found.'}));
app.use((error,req,res,next)=>sendError(res,error));
app.listen(port,'0.0.0.0',()=>console.log(`DigiClarity ${mode} API listening on ${port}`));
