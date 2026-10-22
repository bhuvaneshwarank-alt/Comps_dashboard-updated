import { check } from './errors.js';
import { requirePermission,text } from './access.js';
import { iamCanReadRecord_,phase3FrameworkKey_,phase3FrameworkLabel_,phase3Risk_,phase3Value_,phase3IsOpen_,phase3Overdue_,phase3PendingEvidence_ } from './source-rules.js';
export const RECORD_TABLES=['Compliance_Master_Register','Record_Deletion_Register'];
export function objects(table,name,required){
  check(Array.isArray(table)&&table.length>0,503,'SCHEMA_MISMATCH',`${name} is missing or empty.`);
  const headers=table[0].map(text),nonblank=headers.filter(Boolean);
  check(new Set(nonblank).size===nonblank.length&&required.every(h=>headers.includes(h)),503,'SCHEMA_MISMATCH',`${name} has missing or duplicate column names.`);
  return table.slice(1).filter(r=>r.some(c=>text(c)!=='')).map(row=>Object.fromEntries(headers.flatMap((h,i)=>h?[[h,row[i]??'']]:[])));
}
export function visibleRecords(tables,access){
  requirePermission(access,'viewRecords');
  const records=objects(tables.Compliance_Master_Register,'Compliance_Master_Register',['Record ID','Record Title','Record Type','Department','Record Status','Primary Compliance Framework']);
  const deleted=new Set();
  if(tables.Record_Deletion_Register!==null){
    const events=objects(tables.Record_Deletion_Register,'Record_Deletion_Register',['Record ID','Deletion Status']);
    events.forEach(r=>{const state=text(r['Deletion Status']).toLowerCase();if(state==='deleted')deleted.add(text(r['Record ID']));else if(state==='restored')deleted.delete(text(r['Record ID']));});
  }
  return records.filter(r=>phase3Value_(r['Record ID'])&&!deleted.has(text(r['Record ID']))&&iamCanReadRecord_(access,r));
}
export function filterRecords(records,query={}){
  const search=text(query.search).toLowerCase();
  check(search.length<=200,400,'INVALID_FILTER','Search must be 200 characters or fewer.');
  return records.filter(r=>(!query.framework||phase3FrameworkKey_(r['Primary Compliance Framework'])===phase3FrameworkKey_(query.framework))&&
    (!query.department||text(r.Department).toLowerCase()===text(query.department).toLowerCase())&&
    (!query.status||text(r['Record Status']).toLowerCase()===text(query.status).toLowerCase())&&
    (!query.risk||phase3Risk_(r).toLowerCase()===text(query.risk).toLowerCase())&&
    (!search||['Record ID','Record Title','Department','Record Type'].some(k=>text(r[k]).toLowerCase().includes(search))));
}
export function unique(records,key){return [...new Set(records.map(r=>text(r[key])).filter(Boolean))].sort();}
export function summary(records,now=new Date()){
  const open=records.filter(phase3IsOpen_);
  const frameworks=[...new Set(records.map(r=>phase3FrameworkLabel_(r['Primary Compliance Framework'])).filter(Boolean))];
  const parts=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Kolkata',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(now);
  const todayKey=`${parts.find(p=>p.type==='year').value}-${parts.find(p=>p.type==='month').value}-${parts.find(p=>p.type==='day').value}`;
  return {total:records.length,open:open.length,highRisk:open.filter(r=>['Critical','High'].includes(phase3Risk_(r))).length,
    overdueActions:records.filter(r=>phase3Overdue_(r,todayKey)).length,pendingEvidence:records.filter(phase3PendingEvidence_).length,
    closed:records.filter(r=>text(r['Record Status']).toLowerCase()==='closed').length,
    byFramework:frameworks.map(name=>({name,total:records.filter(r=>phase3FrameworkLabel_(r['Primary Compliance Framework'])===name).length})),
    byStatus:unique(records,'Record Status').map(name=>({name,total:records.filter(r=>r['Record Status']===name).length}))};
}
export function recordById(records,id){const matches=records.filter(r=>text(r['Record ID'])===id);check(matches.length===1,404,'RECORD_UNAVAILABLE','Record unavailable or outside your department scope.');return matches[0];}
