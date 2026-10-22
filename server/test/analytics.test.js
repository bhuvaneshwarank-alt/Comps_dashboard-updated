import test from 'node:test';
import assert from 'node:assert/strict';
import {analyticsForVisibleRecords} from '../src/analytics.js';
import {demoTables} from '../src/demo-data.js';
import {objects} from '../src/records.js';

test('analytics yields stable grouped totals over supplied records',()=>{
  const records=objects(demoTables().Compliance_Master_Register,'Compliance_Master_Register',['Record ID']);
  const report=analyticsForVisibleRecords(records);
  assert.equal(report.total,6);
  assert.deepEqual(report.byDepartment.find(x=>x.name==='HR'),{name:'HR',total:2});
  assert.equal(report.byRecordType.reduce((n,x)=>n+x.total,0),6);
  assert.equal(report.byFramework.reduce((n,x)=>n+x.total,0),6);
});
test('analytics never adds records outside the caller-supplied scope',()=>{
  const records=objects(demoTables().Compliance_Master_Register,'Compliance_Master_Register',['Record ID']);
  const hr=analyticsForVisibleRecords(records.filter(r=>r.Department==='HR'));
  assert.equal(hr.total,2);
  assert.deepEqual(hr.byDepartment,[{name:'HR',total:2}]);
  assert.ok(hr.byRecordType.every(x=>x.total<=2));
  assert.equal(analyticsForVisibleRecords([]).total,0);
});
