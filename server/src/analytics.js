import {text} from './access.js';
import {phase3Risk_,phase3FrameworkLabel_} from './source-rules.js';

// Accept only records already authorized and department-filtered by visibleRecords().
// Counts describe records, not certified control coverage or audit readiness.
export function analyticsForVisibleRecords(records){
  const group=(getter)=>[...records.reduce((map,record)=>{
    const name=text(getter(record))||'Unspecified';
    map.set(name,(map.get(name)||0)+1);return map;
  },new Map())].map(([name,total])=>({name,total})).sort((a,b)=>b.total-a.total||a.name.localeCompare(b.name));
  return {
    total:records.length,
    byDepartment:group(r=>r.Department),
    byFramework:group(r=>phase3FrameworkLabel_(r['Primary Compliance Framework'])),
    byRecordType:group(r=>r['Record Type']),
    byStatus:group(r=>r['Record Status']),
    byRisk:group(r=>phase3Risk_(r)),
    notice:'Record distributions only. Not an audit opinion, certification score, or control effectiveness assessment.'
  };
}
