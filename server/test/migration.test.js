import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {parse} from 'acorn';
import vm from 'node:vm';import net from 'node:net';import {spawn} from 'node:child_process';import {createHmac} from 'node:crypto';
import {getAccess,USER_HEADERS,GRANT_HEADERS,AUTH_TABLES} from '../src/access.js';
import {makeAuth,legacyPasswordHash,validateGooglePayload,verifyPassword,IDLE_MS} from '../src/auth.js';
import {filterRecords,visibleRecords,summary,recordById,objects} from '../src/records.js';
import {rbacBuiltInRoleSeeds_,rbacNormalizeRole_,dashboardPasswordPolicy_,phase3FrameworkKey_} from '../src/source-rules.js';
import {MemoryRepository} from '../src/repository.js';
import {demoTables} from '../src/demo-data.js';
import {AppError} from '../src/errors.js';

test('ported pure rules are byte-for-byte the implementations in the uploaded Code.gs',()=>{
 const raw=fs.readFileSync(new URL('../../reference/apps-script/Code.gs',import.meta.url),'utf8');
 const source=parse(raw,{ecmaVersion:'latest'}),ported=fs.readFileSync(new URL('../src/source-rules.js',import.meta.url),'utf8'),dst=parse(ported,{ecmaVersion:'latest',sourceType:'module'});
 const names=['rbacEmptyPermissions_','rbacNormalizePermissions_','rbacBuiltInRoleSeeds_','rbacSafeJsonArray_','rbacNormalizeRole_','iamParseDepartments_','iamDepartmentKey_','iamCanReadRecord_','dashboardPasswordNormalizeForBlocklist_','dashboardPasswordIdentityTokens_','dashboardPasswordHasObviousSequence_','dashboardPasswordHasDateLikePattern_','dashboardPasswordPolicy_','phase3Value_','phase3Risk_','phase3Status_','phase3IsTerminal_','phase3IsOpen_','phase3Date_','phase3DayKey_','phase3Overdue_','phase3PendingEvidence_','phase3FrameworkKey_','phase3FrameworkLabel_'];
 const find=(ast,name)=>ast.body.find(n=>n.type==='FunctionDeclaration'&&n.id.name===name);
 for(const name of names)assert.equal(ported.slice(find(dst,name).start,find(dst,name).end),raw.slice(find(source,name).start,find(source,name).end),name);
});

test('port retains canonical department and framework normalization rules',()=>{
 assert.equal(phase3FrameworkKey_('SOC 2 Type II'),'soc 2');
 assert.equal(phase3FrameworkKey_('ISO/IEC 27001:2022'),'iso 27001:2022');
 assert.equal(phase3FrameworkKey_('HIPAA Security Rule'),'hipaa');
 const roles=rbacBuiltInRoleSeeds_().map(rbacNormalizeRole_);assert.deepEqual(roles.map(r=>r.name),['Admin','Compliance Manager','Auditor','Department User']);
});

test('register view enforces assigned department scope and excludes deleted records',()=>{
 const tables=demoTables(),hr=getAccess(tables,'hr@demo.invalid',['demo.invalid']),rows=visibleRecords(tables,hr);
 assert.equal(rows.length,2);assert.ok(rows.every(r=>r.Department==='HR'));
 assert.throws(()=>recordById(rows,'AUD-01-2026'),e=>e.status===404);
 tables.Record_Deletion_Register.push(['INC-01-2026','Deleted']);
 assert.equal(visibleRecords(tables,hr).length,1);
});

test('filtering and dashboard totals use the same normalized risk, overdue and evidence rules',()=>{
 const rows=visibleRecords(demoTables(),getAccess(demoTables(),'admin@demo.invalid',['demo.invalid']));
 assert.equal(summary(rows,new Date('2026-10-22T12:00:00.000Z')).total,6);assert.equal(summary(rows,new Date('2026-10-22T12:00:00.000Z')).highRisk,2);assert.equal(summary(rows,new Date('2026-10-22T12:00:00.000Z')).closed,1);
 assert.equal(summary(rows,new Date('2026-10-22T12:00:00.000Z')).overdueActions,1);assert.equal(summary(rows,new Date('2026-10-22T12:00:00.000Z')).pendingEvidence,1);
 assert.equal(filterRecords(rows,{framework:'SOC2'}).length,2);
 assert.equal(filterRecords(rows,{risk:'high'}).length,2);
 assert.equal(filterRecords(rows,{search:'continuity'}).length,1);
});

test('only exact Google Workspace hosted-domain claims for an allowed domain pass',()=>{
 const payload={email:'member@digiclarity.in',email_verified:true,hd:'digiclarity.in',sub:'123'};
 assert.equal(validateGooglePayload(payload,['digiclarity.in','clarityrcm.com']),payload.email);
 for(const bad of [{...payload,hd:undefined},{...payload,email_verified:false},{...payload,hd:'evil.test'}, {...payload,email:'member@gmail.com'}])assert.throws(()=>validateGooglePayload(bad,['digiclarity.in','clarityrcm.com']),e=>e.status===403);
});

test('legacy password derivation is stable, peppered and rejects an absent pepper',()=>{
 const pepper='private test pepper',password='VioletCanoe!83xQ',salt='c2FsdA==';
 const hash=legacyPasswordHash(password,salt,pepper);
 assert.equal(hash,legacyPasswordHash(password,salt,pepper));
 assert.notEqual(hash,legacyPasswordHash(password,salt,'different pepper'));
 assert.throws(()=>legacyPasswordHash('anything','salt',''),e=>e.code==='PASSWORD_CONFIG_REQUIRED');
 const raw=fs.readFileSync(new URL('../../reference/apps-script/Code.gs',import.meta.url),'utf8');const ast=parse(raw,{ecmaVersion:'latest'});
 const helper=ast.body.find(n=>n.type==='FunctionDeclaration'&&n.id.name==='dashboardPasswordHash_');
 const context={Utilities:{computeHmacSha256Signature:(v,k)=>createHmac('sha256',k).update(v,'utf8').digest(),base64EncodeWebSafe:b=>b.toString('base64').replaceAll('+','-').replaceAll('/','_')},dashboardPasswordPepper_:()=>pepper};
 vm.createContext(context);vm.runInContext(raw.slice(helper.start,helper.end)+'\nthis.ported=dashboardPasswordHash_;',context);
 assert.equal(hash,context.ported(password,salt),'The Node hash must verify the original GAS HMAC encoding.');
});

test('role headers and master schema fail closed on an absent or duplicate required column',()=>{
 assert.deepEqual(USER_HEADERS.length,7);assert.equal(GRANT_HEADERS.length,15);
 assert.throws(()=>objects(null,'Compliance_Master_Register',['Record ID']),e=>e.status===503);
 assert.throws(()=>objects([['Record ID','Record ID']], 'Compliance_Master_Register',['Record ID']),e=>e.status===503);
});

test('local demo authentication is explicitly synthetic, department-scoped and idle-bound',async()=>{
 const repo=new MemoryRepository(demoTables()),config={mode:'demo',domains:['demo.invalid'],googleClientId:'',passwordPepper:''};let now=1000000;const auth=makeAuth({repo,config,now:()=>now});
 const login=await auth.login({persona:'hr'});assert.equal(login.access.role,'Department User');
 const req={headers:{cookie:`dc_session=${login.id}`}},found=await auth.authenticate(req);assert.equal(found.access.departments[0],'HR');
 auth.touch(found.session);now+=IDLE_MS+1;
 await assert.rejects(()=>auth.authenticate(req),e=>e.code==='SESSION_EXPIRED');
});

test('protected department role cannot use global write permissions accidentally',()=>{
 const tables=demoTables();tables.Role_Definitions=[['Role Name','Description','Built In','Active','Scope Type','Allowed Departments','Allowed Users','Permissions JSON','Updated By','Updated At'],
 ['HR Viewer','','false','true','DEPARTMENTS','["HR"]','[]',JSON.stringify({viewRecords:true,viewDashboard:true,createRecords:true}), '', '']];
 tables.User_Access.push(['outsider@demo.invalid','Outsider','HR Viewer',true,'','','["IT"]']);
 assert.throws(()=>getAccess(tables,'outsider@demo.invalid',['demo.invalid']),e=>e.status===403);
});

test('Google-backed department access fails closed when Sheets storage ownership is not verified',()=>{
 const tables=demoTables();
 assert.throws(()=>getAccess(tables,'hr@demo.invalid',['demo.invalid'],Date.now(),{mode:'google',iamStorageOwner:'old-owner@company.invalid',serviceAccount:'cloud-run@project.iam.gserviceaccount.com'}),e=>e.code==='STORAGE_OWNER_MISMATCH');
 const allowed=getAccess(tables,'hr@demo.invalid',['demo.invalid'],Date.now(),{mode:'google',iamStorageOwner:'cloud-run@project.iam.gserviceaccount.com',serviceAccount:'cloud-run@project.iam.gserviceaccount.com'});
 assert.equal(allowed.departments[0],'HR');
});

test('forged or cross-domain identity data is never inferred from a typed email',()=>{
 assert.throws(()=>validateGooglePayload({email:'admin@digiclarity.in',email_verified:true,sub:'x'},['digiclarity.in']),e=>e.code==='IDENTITY_DENIED');
});

test('five password failures lock for 15 minutes and a correct existing password resets the counter',async()=>{
 const password='VioletCanoe!83xQ',pepper='test server pepper';const tables=demoTables();
 tables.Dashboard_Credentials.push(['admin@demo.invalid','test-salt',legacyPasswordHash(password,'test-salt',pepper),'Admin','2026-10-22T00:00:00.000Z','[]','2027-01-20T00:00:00.000Z',0,'','']);
 const repo=new MemoryRepository(tables),access=getAccess(tables,'admin@demo.invalid',['demo.invalid']);let now=Date.parse('2026-10-22T10:00:00.000Z');
 for(let i=1;i<=5;i++){const snapshot=await repo.tables(AUTH_TABLES);await assert.rejects(()=>verifyPassword(repo,snapshot,access,'wrong password',pepper,now),e=>e.code===(i===5?'ACCOUNT_LOCKED':'INVALID_PASSWORD'));}
 const locked=await repo.tables(AUTH_TABLES);assert.equal(locked.Dashboard_Credentials[1][7],5);assert.equal(Date.parse(locked.Dashboard_Credentials[1][8])-now,15*60*1000);
 await assert.rejects(()=>verifyPassword(repo,locked,access,password,pepper,now),e=>e.code==='ACCOUNT_LOCKED');
 now+=15*60*1000+1;const unlocked=await repo.tables(AUTH_TABLES);assert.equal(await verifyPassword(repo,unlocked,access,password,pepper,now),'2026-10-22T00:00:00.000Z');
 const reset=await repo.tables(AUTH_TABLES);assert.equal(reset.Dashboard_Credentials[1][7],0);assert.equal(reset.Dashboard_Credentials[1][8],'');
});

test('HTTP preview runs the real React API flow, scoped records, record history and CSRF logout',async t=>{
 const probe=net.createServer();await new Promise(resolve=>probe.listen(0,'127.0.0.1',resolve));const port=probe.address().port;await new Promise(resolve=>probe.close(resolve));
 const child=spawn(process.execPath,['server/src/index.js'],{cwd:new URL('../..',import.meta.url),env:{...process.env,APP_MODE:'demo',NODE_ENV:'test',PORT:String(port),APP_ORIGIN:`http://127.0.0.1:${port}`,ALLOWED_DOMAINS:'demo.invalid'},stdio:['ignore','pipe','pipe']});
 let log='';child.stdout.on('data',d=>log+=d.toString());child.stderr.on('data',d=>log+=d.toString());
 const base=`http://127.0.0.1:${port}`;let ready=false;
 try{for(let i=0;i<70;i++){if(child.exitCode!==null)throw Error(log);try{if((await fetch(base+'/api/config')).ok){ready=true;break;}}catch{}await new Promise(r=>setTimeout(r,50));}assert.ok(ready,'Node API starts');
  const wrongOrigin=await fetch(base+'/api/auth/login',{method:'POST',headers:{origin:'https://attacker.invalid','content-type':'application/json'},body:JSON.stringify({persona:'admin'})});assert.equal(wrongOrigin.status,403);
  async function login(persona){const response=await fetch(base+'/api/auth/login',{method:'POST',headers:{origin:base,'content-type':'application/json'},body:JSON.stringify({persona})});assert.equal(response.status,200);return {data:await response.json(),cookie:response.headers.get('set-cookie').split(';')[0]};}
  const admin=await login('admin'),hr=await login('hr');
  const get=(cookie,path)=>fetch(base+path,{headers:{cookie}});
  const summaryResponse=await get(admin.cookie,'/api/overview'),overview=await summaryResponse.json();assert.equal(overview.summary.total,6);
  const hrRows=await (await get(hr.cookie,'/api/records?page=1&pageSize=25')).json();assert.equal(hrRows.total,2);assert.ok(hrRows.items.every(r=>r.Department==='HR'));
  assert.equal((await get(hr.cookie,'/api/records/AUD-01-2026')).status,404);
  const event=(await get(admin.cookie,'/api/records/AUD-01-2026/history')).status;assert.equal(event,200);
  const badLogout=await fetch(base+'/api/auth/logout',{method:'POST',headers:{origin:base,cookie:hr.cookie,'content-type':'application/json','x-csrf-token':'wrong'},body:'{}'});assert.notEqual(badLogout.status,204);
  const logout=await fetch(base+'/api/auth/logout',{method:'POST',headers:{origin:base,cookie:admin.cookie,'content-type':'application/json','x-csrf-token':admin.data.csrf},body:'{}'});assert.equal(logout.status,204);
  assert.equal((await get(admin.cookie,'/api/overview')).status,401);
 }finally{child.kill('SIGTERM');await new Promise(resolve=>child.once('exit',resolve));}
});

