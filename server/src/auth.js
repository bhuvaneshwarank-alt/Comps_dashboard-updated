import { createHmac,randomBytes,timingSafeEqual } from 'node:crypto';
import { OAuth2Client } from 'google-auth-library';
import { check,fail } from './errors.js';
import { AUTH_TABLES,getAccess,requireSchema,text } from './access.js';
import { dashboardPasswordPolicy_ } from './source-rules.js';
export const CREDENTIAL_HEADERS=['Email','Salt','Password Hash','Updated By','Updated At','Password History JSON','Password Expires At','Failed Attempts','Locked Until','Last Login At'];
export const IDLE_MS=5*60*1000;
export const ABSOLUTE_MS=24*60*60*1000;
export function legacyPasswordHash(password,salt,pepper){
  check(Boolean(pepper),503,'PASSWORD_CONFIG_REQUIRED','The existing server password pepper must be configured securely.');
  let value=String(salt)+'\n'+String(password);
  for(let i=0;i<180;i++)value=createHmac('sha256',pepper).update(value,'utf8').digest('base64').replaceAll('+','-').replaceAll('/','_');
  return value;
}
export function safeEqual(a,b){const x=Buffer.from(String(a)),y=Buffer.from(String(b));return x.length===y.length&&timingSafeEqual(x,y);}
function credential(tables,email){
  requireSchema(tables.Dashboard_Credentials,CREDENTIAL_HEADERS,'Dashboard_Credentials');
  const hits=tables.Dashboard_Credentials.slice(1).map((row,i)=>({row,index:i+2})).filter(x=>text(x.row[0]).toLowerCase()===email);
  check(hits.length===1,403,'PASSWORD_SETUP_REQUIRED','Create or repair this account password in the existing dashboard. Password setup is not migrated yet.');
  return hits[0];
}
function expiry(row){
  if(row[6])return Date.parse(row[6]);
  const updated=Date.parse(row[4]);return updated+90*86400000;
}
export async function verifyPassword(repo,tables,access,password,pepper,now=Date.now()){
  const {row,index}=credential(tables,access.email);
  const locked=Date.parse(row[8]||'');
  check(!Number.isFinite(locked)||locked<=now,423,'ACCOUNT_LOCKED','Too many incorrect attempts. Try again after the 15-minute lock.');
  const candidate=legacyPasswordHash(password,row[1],pepper);
  if(!safeEqual(candidate,row[2])){
    const failures=Number(row[7]||0)+1,until=failures>=5?new Date(now+15*60000).toISOString():'';
    await repo.credentialResult(index,failures,until,row[9]||'');
    fail(failures>=5?423:401,failures>=5?'ACCOUNT_LOCKED':'INVALID_PASSWORD',failures>=5?'Five incorrect attempts. This account is locked for 15 minutes.':'Incorrect dashboard password.');
  }
  const expires=expiry(row);
  check(Number.isFinite(expires)&&expires>now,403,'PASSWORD_EXPIRED','Change your expired password in the existing dashboard before continuing.');
  try{dashboardPasswordPolicy_(password,access);}catch{fail(403,'PASSWORD_POLICY_UPGRADE_REQUIRED','Upgrade your password in the existing dashboard to meet the current policy.');}
  await repo.credentialResult(index,0,'',new Date(now).toISOString());
  return String(row[4]);
}
export function cookieValue(req){
  const raw=(req.headers.cookie||'').split(';').map(x=>x.trim()).find(x=>x.startsWith('dc_session='));
  return raw?raw.slice('dc_session='.length):'';
}
export function validateGooglePayload(payload,domains){
  const email=text(payload?.email).toLowerCase(),domain=email.split('@')[1];
  check(payload?.email_verified===true&&Boolean(payload?.sub)&&domains.includes(payload?.hd)&&domain===payload.hd,403,'IDENTITY_DENIED','Select an approved company Google Workspace account.');
  return email;
}
export function makeAuth({repo,config,now=Date.now}){
  const sessions=new Map(),google=new OAuth2Client(config.googleClientId);
  // Serialized credential writes within ONE process. This is not a distributed
  // lock; production scale-out remains a documented migration task.
  let tail=Promise.resolve();
  const serialized=fn=>{const task=tail.then(fn);tail=task.catch(()=>{});return task;};
  async function login(input){
    if(config.mode==='demo'){
      check(['admin','hr'].includes(input.persona),400,'INVALID_PERSONA','Choose a demo account.');
      const email=input.persona==='hr'?'hr@demo.invalid':'admin@demo.invalid';
      const tables=await repo.tables(AUTH_TABLES),access=getAccess(tables,email,config.domains,now(),config);
      return issue(access,'demo');
    }
    check(typeof input.idToken==='string'&&input.idToken.length<10000,400,'IDENTITY_REQUIRED','Sign in with your company Google account first.');
    check(typeof input.password==='string'&&input.password.length<=128,400,'PASSWORD_REQUIRED','Enter your dashboard password.');
    let payload;
    try{payload=(await google.verifyIdToken({idToken:input.idToken,audience:config.googleClientId})).getPayload();}
    catch{fail(401,'IDENTITY_INVALID','Google sign-in expired or could not be verified. Sign in again.');}
    const email=validateGooglePayload(payload,config.domains);
    return serialized(async()=>{
      const tables=await repo.tables(AUTH_TABLES),access=getAccess(tables,email,config.domains,now());
      const credentialVersion=await verifyPassword(repo,tables,access,input.password,config.passwordPepper,now());
      return issue(access,credentialVersion);
    });
  }
  function issue(access,credentialVersion){
    for(const [key,s] of sessions)if(now()-s.lastSeen>=IDLE_MS||now()-s.created>=ABSOLUTE_MS)sessions.delete(key);
    const id=randomBytes(32).toString('hex'),csrf=randomBytes(32).toString('hex');
    sessions.set(id,{email:access.email,created:now(),lastSeen:now(),csrf,credentialVersion});
    return {id,csrf,access};
  }
  async function authenticate(req){
    const id=cookieValue(req),session=sessions.get(id);
    check(Boolean(session),401,'SESSION_REQUIRED','Sign in to open the dashboard.');
    if(now()-session.lastSeen>=IDLE_MS||now()-session.created>=ABSOLUTE_MS||session.lastSeen>now()+60000){sessions.delete(id);fail(401,'SESSION_EXPIRED','Dashboard locked after five minutes of inactivity. Sign in again.');}
    const tables=await repo.tables(AUTH_TABLES),access=getAccess(tables,session.email,config.domains,now(),config);
    if(config.mode!=='demo'){
      const {row}=credential(tables,session.email);
      check(String(row[4])===session.credentialVersion&&expiry(row)>now(),401,'SESSION_REVOKED','Credentials changed or expired. Sign in again.');
    }
    return {access,session,id};
  }
  function requireCsrf(req,session){check(safeEqual(req.headers['x-csrf-token']||'',session.csrf),403,'CSRF_REJECTED','Refresh the dashboard and retry.');}
  return {login,authenticate,requireCsrf,revoke:id=>sessions.delete(id),touch:s=>s.lastSeen=now(),sessionCount:()=>sessions.size};
}
