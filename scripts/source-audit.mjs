import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { parse } from 'acorn';
const root=new URL('../',import.meta.url).pathname;
const dir=path.join(root,'reference/apps-script');
const code=fs.readFileSync(path.join(dir,'Code.gs'),'utf8');
const ast=parse(code,{ecmaVersion:'latest'});
const nodes=new Map(ast.body.filter(n=>n.type==='FunctionDeclaration').map(n=>[n.id.name,n]));
const pure=[
 'rbacEmptyPermissions_','rbacNormalizePermissions_','rbacBuiltInRoleSeeds_',
 'rbacSafeJsonArray_','rbacNormalizeRole_','iamParseDepartments_',
 'iamDepartmentKey_','iamCanReadRecord_','dashboardPasswordNormalizeForBlocklist_',
 'dashboardPasswordIdentityTokens_','dashboardPasswordHasObviousSequence_',
 'dashboardPasswordHasDateLikePattern_','dashboardPasswordPolicy_',
 'phase3Value_','phase3Risk_','phase3Status_','phase3IsTerminal_',
 'phase3IsOpen_','phase3Date_','phase3DayKey_','phase3Overdue_',
 'phase3PendingEvidence_','phase3FrameworkKey_','phase3FrameworkLabel_'
];
const constants=['RBAC_PERMISSION_KEYS','DASHBOARD_PASSWORD_MIN_LENGTH','DASHBOARD_PASSWORD_MAX_LENGTH'];
const constantNodes=ast.body.filter(n=>n.type==='VariableDeclaration'&&n.declarations.some(d=>constants.includes(d.id.name)));
const output='// Pure rules ported verbatim from the uploaded Code.gs. Regenerate: npm run audit:source\n'+
 constantNodes.map(n=>code.slice(n.start,n.end)).join('\n')+'\n'+pure.map(name=>{
 const n=nodes.get(name);if(!n)throw Error('Missing source rule '+name);return code.slice(n.start,n.end);
 }).join('\n\n')+'\nexport { '+[...constants,...pure].join(', ')+' };\n';
fs.writeFileSync(path.join(root,'server/src/source-rules.js'),output);
const implemented=new Set(['getRecords','getRecordById','getRecordHistory','getDashboardData']);
const methods=[...code.slice(code.indexOf('function dashboardSecureInvoke(')).matchAll(/case\s+["']([^"']+)["']\s*:/g)].map(m=>m[1]);
const services=['SpreadsheetApp','DriveApp','Drive','PropertiesService','LockService','Session','Utilities','MailApp','ScriptApp','UrlFetchApp','HtmlService'];
const inventory=fs.readdirSync(dir).sort().map(name=>{
 const bytes=fs.readFileSync(path.join(dir,name));const text=bytes.toString('utf8');
 let functions=[];if(name.endsWith('.gs'))functions=parse(text,{ecmaVersion:'latest'}).body.filter(n=>n.type==='FunctionDeclaration').map(n=>n.id.name);
 return {name,bytes:bytes.length,sha256:crypto.createHash('sha256').update(bytes).digest('hex'),lines:text.split('\n').length,functions,services:services.filter(s=>new RegExp('\\b'+s+'\\.').test(text))};
});
const html=fs.readFileSync(path.join(dir,'Index.html'),'utf8');
const authMethods=['completeDashboardPasswordReset','createDashboardPassword','getDashboardLoginIdentity','getLoginBrandLogo','getLoginGatewayContext','heartbeatDashboardSession','loginDashboardWithCredentials','requestDashboardPasswordResetOtp','resumeDashboardSession','revokeDashboardSession','verifyDashboardLoginEmail','verifyDashboardPassword','verifyDashboardPasswordResetOtp'];
const rpcMethods=[...new Set([...methods,...authMethods])].filter(name=>new RegExp('\\.'+name+'\\s*\\(').test(html));
const legacyInvokable=rpcMethods.sort();
fs.writeFileSync(path.join(root,'docs/source-inventory.json'),JSON.stringify({
 baseline:'User-uploaded Code.gs V65 and Index.html with V112 markers; source hashes are authoritative.',
 stage:'Milestone 1: native React register, detail, summary and history; remaining workflows not migrated.',
 files:inventory,protectedOperations:methods.map(name=>({name,status:implemented.has(name)?(name==='getDashboardData'?'ported_read_summary':'ported_read_operation'):'not_migrated'})),
 frontendRpcMethods:legacyInvokable,
 pureRulesPorted:pure,manifest:JSON.parse(fs.readFileSync(path.join(dir,'appsscript.json'),'utf8'))
},null,2));
console.log(`Audited ${inventory.length} files, ${methods.length} protected operations, ${pure.length} pure rules.`);
