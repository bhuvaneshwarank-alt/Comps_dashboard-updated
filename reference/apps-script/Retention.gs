/* DigiClarity TC-24 retention governance, prototype V64.
 * Seven years is a DRAFT default, not approved organizational policy.
 * No permanent deletion, evidence deletion, scheduled purge or email is performed.
 * Public access is via dashboardSecureInvoke only; implementation functions end in _.
 */
const TC24_POLICY_SHEET_ = 'TC24_Retention_Policy';
const TC24_HOLD_SHEET_ = 'TC24_Legal_Holds';
const TC24_DISPOSAL_SHEET_ = 'TC24_Disposal_Reviews';
const TC24_LOG_SHEET_ = 'TC24_Retention_Audit';
const TC24_SCHEMAS_ = {
  TC24_Retention_Policy: ['Rule Key','Record Type','Retention Years','Policy State','Updated At','Updated By','Change Reason'],
  TC24_Legal_Holds: ['Hold ID','Record ID','Status','Reason','Created By','Created At','Released By','Released At','Release Reason'],
  TC24_Disposal_Reviews: ['Request ID','Record ID','Requested By','Requested At','Reason','Status','Reviewed By','Reviewed At','Review Note'],
  TC24_Retention_Audit: ['Event ID','Timestamp','Actor','Action','Record ID','Reference ID','Outcome','Details']
};

function tc24Access_(write) {
  const access = rbacRequire_('read');
  const permitted = !access.departmentScoped && (access.isSystemAdmin ||
    (access.role === 'Compliance Manager' && !!access.permissions.reviewChanges));
  if (!permitted) throw new Error('TC24_ACCESS_DENIED: Retention Management requires Admin or Compliance Manager.');
  if (write && !access.isSystemAdmin && access.role !== 'Compliance Manager') {
    throw new Error('TC24_ACCESS_DENIED: Cannot change retention governance.');
  }
  return access;
}
function tc24Sheet_(name, create) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet && create) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1,1,1,TC24_SCHEMAS_[name].length).setValues([TC24_SCHEMAS_[name]]);
    sheet.setFrozenRows(1);
  }
  if (sheet) {
    const required = TC24_SCHEMAS_[name];
    const actual = sheet.getRange(1,1,1,required.length).getDisplayValues()[0];
    if (actual.some(function(v,i){return v !== required[i];})) {
      throw new Error('TC24_SCHEMA_MISMATCH: Unexpected column layout in '+ name +'.');
    }
  }
  return sheet;
}
function tc24Rows_(name) {
  const sheet = tc24Sheet_(name, false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const columns=TC24_SCHEMAS_[name];
  const values=sheet.getRange(2,1,sheet.getLastRow()-1,columns.length).getDisplayValues();
  return values.map(function(row,i){
    const entry={rowNumber:i+2};
    columns.forEach(function(key,j){entry[key]=String(row[j]||'').trim();});
    return entry;
  }).filter(function(row){return !!Object.keys(row).some(function(k){return k!=='rowNumber' && !!row[k];});});
}
function tc24SafeCellText_(input) {
  const value=String(input==null?'':input).trim();
  return /^[=+\-@\t\r]/.test(value) ? '\''+value : value;
}
function tc24WriteLog_(action, recordId, reference, outcome, details) {
  const sheet=tc24Sheet_(TC24_LOG_SHEET_, false);
  if (!sheet) throw new Error('TC24_SETUP_REQUIRED: Retention audit log is missing.');
  sheet.appendRow([Utilities.getUuid(),new Date().toISOString(),rbacIdentity_(),action,
    recordId||'',reference||'',outcome||'SUCCESS',tc24SafeCellText_(String(details||'').slice(0,200))]);
}
function tc24WithLock_(fn) {
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {return fn();} finally {lock.releaseLock();}
}
function tc24Initialize_() {
  const access=tc24Access_(true);
  if (!access.isSystemAdmin) throw new Error('TC24_ACCESS_DENIED: Only Admin may initialize retention sheets.');
  return tc24WithLock_(function(){
    Object.keys(TC24_SCHEMAS_).forEach(function(name){tc24Sheet_(name,true);});
    const sheet=tc24Sheet_(TC24_POLICY_SHEET_,false);
    const defaults=tc24Rows_(TC24_POLICY_SHEET_);
    if (!defaults.some(function(row){return row['Rule Key']==='DEFAULT';})) {
      sheet.appendRow(['DEFAULT','*',7,'DRAFT - PROTOTYPE',new Date().toISOString(),rbacIdentity_(),
        'Provisional default; management policy approval required before production']);
      tc24WriteLog_('RETENTION_SETUP','', 'DEFAULT','SUCCESS','Draft seven-year default set; automatic deletion disabled');
    }
    SpreadsheetApp.flush();
    return {success:true,created:true,policyState:'DRAFT - PROTOTYPE',
      sheets:Object.keys(TC24_SCHEMAS_), deletionEnabled:false};
  });
}
function tc24Policy_() {
  const rows=tc24Rows_(TC24_POLICY_SHEET_);
  const main=rows.find(function(r){return r['Rule Key']==='DEFAULT';});
  if (!main) return null;
  const years=Number(main['Retention Years']);
  if (!Number.isInteger(years)||years<1||years>30) throw new Error('TC24_POLICY_INVALID: Retention years must be 1–30.');
  return {years:years,state:main['Policy State'],updatedAt:main['Updated At'],updatedBy:main['Updated By']};
}
function tc24SavePolicy_(years, reason) {
  tc24Access_(true);
  years=Number(years); reason=String(reason||'').trim();
  if (!Number.isInteger(years)||years<1||years>30) throw new Error('Retention years must be 1–30.');
  if (reason.length<8||reason.length>500) throw new Error('Provide an 8–500 character reason.');
  return tc24WithLock_(function(){
    const sheet=tc24Sheet_(TC24_POLICY_SHEET_,false);
    if (!sheet) throw new Error('TC24_SETUP_REQUIRED: Initialize TC24 before updating policy.');
    const row=tc24Rows_(TC24_POLICY_SHEET_).find(function(r){return r['Rule Key']==='DEFAULT';});
    if (!row) throw new Error('TC24_SETUP_REQUIRED: Default policy missing.');
    const old=Number(row['Retention Years']);
    sheet.getRange(row.rowNumber,1,1,7).setValues([['DEFAULT','*',years,'DRAFT - PROTOTYPE',
      new Date().toISOString(),rbacIdentity_(),tc24SafeCellText_(reason)]]);
    tc24WriteLog_('POLICY_DRAFT_CHANGED','', 'DEFAULT','SUCCESS','Years '+old+' to '+years+'; '+reason);
    SpreadsheetApp.flush();
    return {success:true,years:years,state:'DRAFT - PROTOTYPE',deletionEnabled:false};
  });
}
function tc24Date_(value) {
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? new Date(Date.UTC(value.getUTCFullYear(),value.getUTCMonth(),value.getUTCDate())) : null;
  const raw=String(value||'').trim();
  if (!raw || raw==='N/A') return null;
  let y,m,d,match=raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);
  if (match) {y=+match[1];m=+match[2];d=+match[3];}
  else {
    match=raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (!match) return null;
    // Google Sheets data may be rendered in MM/DD/YYYY or DD/MM/YYYY.
    // Reject ambiguous dates rather than calculating an incorrect legal deadline.
    if (+match[1]<=12 && +match[2]<=12) return null;
    if (+match[1]>12) {d=+match[1];m=+match[2];}
    else {m=+match[1];d=+match[2];}
    y=+match[3];
  }
  const t=new Date(Date.UTC(y,m-1,d));
  return t.getUTCFullYear()===y && t.getUTCMonth()===m-1 && t.getUTCDate()===d ? t : null;
}
function tc24IsoDate_(dt){return dt?dt.toISOString().slice(0,10):'';}
function tc24AddYears_(date,years) {
  const y=date.getUTCFullYear()+years, m=date.getUTCMonth(), day=date.getUTCDate();
  const endOfMonth=new Date(Date.UTC(y,m+1,0)).getUTCDate();
  return new Date(Date.UTC(y,m,Math.min(day,endOfMonth)));
}
function tc24HoldMap_() {
  const map={};
  tc24Rows_(TC24_HOLD_SHEET_).forEach(function(row){
    if (row.Status==='ACTIVE') map[row['Record ID']]={id:row['Hold ID'],reason:row.Reason,
      since:row['Created At'],createdBy:row['Created By']};
  });
  return map;
}
function tc24RecordPolicy_(record,policy,hold,today) {
  const id=String(record['Record ID']||'').trim();
  const type=String(record['Record Type']||'').trim();
  const status=String(record['Record Status']||'').trim();
  const closed=/^(closed|resolved|completed)$/i.test(status);
  const closedDate=tc24Date_(record['Closure Date']) || tc24Date_(record['Incident Closure Date']);
  const explicit=tc24Date_(record['Record Retention Expiry Date']);
  const start=closed ? closedDate : null;
  const expiry=start ? tc24AddYears_(start,policy.years) : null;
  // Fail conservative: master dates may extend retention, but may not
  // silently shorten this seven-year DRAFT minimum.
  const useMasterExpiry=!!(expiry && explicit && explicit.getTime()>expiry.getTime());
  const effective=closed ? (useMasterExpiry?explicit:expiry) : null;
  const ms=24*60*60*1000;
  const days=effective ? Math.ceil((effective.getTime()-today.getTime())/ms) : null;
  let state=closed ? (!start ? 'Missing closure date' : (days<0?'Expired':days<=90?'Expiring soon':'Retained')) : 'Open / not started';
  if (hold) state='Legal hold';
  return {recordId:id,recordType:type,title:String(record['Record Title']||'').slice(0,120),
    recordStatus:status,department:String(record['Department']||''),state:state,
    retentionYears:policy.years,retentionStart:tc24IsoDate_(start),expiryDate:tc24IsoDate_(effective),
    expirySource:useMasterExpiry?'Later master expiry (review before production)':explicit&&start?'Earlier master expiry ignored; draft minimum applies':start?'Calculated (draft rule)':'Not calculated',
    daysRemaining:days,legalHold:!!hold,holdId:hold?hold.id:'',
    disposalEligible:closed&&!!start&&!!effective&&days<0&&!hold};
}
function tc24GetCenter_() {
  tc24Access_(false);
  const policy=tc24Policy_();
  if (!policy) return {needsSetup:true,deletionEnabled:false,canInitialize:rbacRequire_('read').isSystemAdmin};
  const holds=tc24HoldMap_();
  const now=new Date();
  const today=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()));
  const records=readMasterRecords_().records.map(function(r){
    return tc24RecordPolicy_(r,policy,holds[String(r['Record ID']||'').trim()],today);
  });
  const counts={tracked:records.length,open:0,retained:0,expiringSoon:0,expired:0,legalHolds:0,missingClosure:0};
  records.forEach(function(r){
    if(r.state==='Open / not started')counts.open++;
    if(r.state==='Retained')counts.retained++;
    if(r.state==='Expiring soon')counts.expiringSoon++;
    if(r.state==='Expired')counts.expired++;
    if(r.state==='Legal hold')counts.legalHolds++;
    if(r.state==='Missing closure date')counts.missingClosure++;
  });
  const requests=tc24Rows_(TC24_DISPOSAL_SHEET_).slice(-100).reverse().map(function(r){
    return {id:r['Request ID'],recordId:r['Record ID'],requestedBy:r['Requested By'],
      requestedAt:r['Requested At'],status:r.Status,reviewedBy:r['Reviewed By'],
      reviewedAt:r['Reviewed At'],canReview:r.Status==='PENDING_REVIEW' &&
      r['Requested By'].toLowerCase()!==rbacIdentity_()};
  });
  return {needsSetup:false,policy:policy,counts:counts,records:records,
    requests:requests,legalHolds:Object.keys(holds).map(function(key){return Object.assign({recordId:key},holds[key]);}),
    deletionEnabled:false,autoRemindersEnabled:false,retentionPolicyApproved:false,
    canInitialize:rbacRequire_('read').isSystemAdmin};
}
function tc24RecordById_(id){
  const record=readMasterRecords_().records.find(function(r){return String(r['Record ID']||'').trim()===id;});
  if(!record)throw new Error('TC24_RECORD_NOT_FOUND: Record not found or not accessible.');
  return record;
}
function tc24SetLegalHold_(recordId,decision,reason) {
  tc24Access_(true);
  recordId=String(recordId||'').trim();decision=String(decision||'').toUpperCase();reason=String(reason||'').trim();
  if(!recordId||!['PLACE','RELEASE'].includes(decision))throw new Error('Invalid record or legal hold action.');
  if(reason.length<8||reason.length>500) throw new Error('Legal hold reason must contain 8–500 characters.');
  return tc24WithLock_(function(){
    tc24RecordById_(recordId);
    const current=tc24Rows_(TC24_HOLD_SHEET_).find(function(r){return r['Record ID']===recordId&&r.Status==='ACTIVE';});
    const sheet=tc24Sheet_(TC24_HOLD_SHEET_,false), now=new Date().toISOString();
    if(!sheet)throw new Error('TC24_SETUP_REQUIRED');
    if(decision==='PLACE'){
      if(current)throw new Error('TC24_HOLD_EXISTS: Record already under legal hold.');
      const id='LH-'+Utilities.getUuid();
      sheet.appendRow([id,recordId,'ACTIVE',tc24SafeCellText_(reason),rbacIdentity_(),now,'','','']);
      tc24WriteLog_('LEGAL_HOLD_PLACED',recordId,id,'SUCCESS','Legal hold placed; reason retained in protected hold register');
      SpreadsheetApp.flush();return {success:true,id:id,status:'ACTIVE'};
    }
    if(!current)throw new Error('TC24_NO_ACTIVE_HOLD: No active legal hold exists.');
    sheet.getRange(current.rowNumber,3).setValue('RELEASED');
    sheet.getRange(current.rowNumber,7,1,3).setValues([[rbacIdentity_(),now,tc24SafeCellText_(reason)]]);
    tc24WriteLog_('LEGAL_HOLD_RELEASED',recordId,current['Hold ID'],'SUCCESS','Legal hold released');
    SpreadsheetApp.flush();return {success:true,id:current['Hold ID'],status:'RELEASED'};
  });
}
function tc24AssertDeletionAllowed_(recordId){
  // Applies to the existing soft-delete path. Existing retention rules remain
  // draft; all legal holds nevertheless fail closed if the TC24 module is set up.
  const holdSheet=getSpreadsheet_().getSheetByName(TC24_HOLD_SHEET_);
  if(!holdSheet)return;
  if(tc24HoldMap_()[String(recordId||'').trim()])
    throw new Error('TC24_LEGAL_HOLD: This record is subject to a legal hold and cannot be deleted.');
}
function tc24RequestDisposalReview_(recordId,reason) {
  tc24Access_(true);
  recordId=String(recordId||'').trim(); reason=String(reason||'').trim();
  if(reason.length<8||reason.length>500)throw new Error('Explain disposal readiness in 8–500 characters.');
  return tc24WithLock_(function(){
    const policy=tc24Policy_();if(!policy)throw new Error('TC24_SETUP_REQUIRED');
    const record=tc24RecordById_(recordId);
    const today=new Date(),utc=new Date(Date.UTC(today.getUTCFullYear(),today.getUTCMonth(),today.getUTCDate()));
    const item=tc24RecordPolicy_(record,policy,tc24HoldMap_()[recordId],utc);
    if(!item.disposalEligible)throw new Error('TC24_NOT_ELIGIBLE: Record is not expired, lacks a valid closure date, or has a legal hold.');
    const sheet=tc24Sheet_(TC24_DISPOSAL_SHEET_,false);
    const pending=tc24Rows_(TC24_DISPOSAL_SHEET_).some(function(r){return r['Record ID']===recordId&&r.Status==='PENDING_REVIEW';});
    if(pending)throw new Error('TC24_PENDING: A disposal review is already pending for this record.');
    const id='DR-'+Utilities.getUuid();
    sheet.appendRow([id,recordId,rbacIdentity_(),new Date().toISOString(),tc24SafeCellText_(reason),'PENDING_REVIEW','','','']);
    tc24WriteLog_('DISPOSAL_REVIEW_REQUESTED',recordId,id,'SUCCESS','Review only; no deletion is performed');
    SpreadsheetApp.flush();return {success:true,id:id,status:'PENDING_REVIEW',deletionPerformed:false};
  });
}
function tc24ReviewDisposalRequest_(requestId,decision,note) {
  tc24Access_(true);
  requestId=String(requestId||'').trim();decision=String(decision||'').toUpperCase();
  note=String(note||'').trim();
  if(!['APPROVE','REJECT'].includes(decision))throw new Error('Select APPROVE or REJECT.');
  if(note.length<8||note.length>500)throw new Error('Reviewer note must have 8–500 characters.');
  return tc24WithLock_(function(){
    const row=tc24Rows_(TC24_DISPOSAL_SHEET_).find(function(r){return r['Request ID']===requestId;});
    if(!row||row.Status!=='PENDING_REVIEW')throw new Error('TC24_REQUEST_NOT_PENDING');
    if(row['Requested By'].toLowerCase()===rbacIdentity_())throw new Error('TC24_SELF_APPROVAL_BLOCKED: Independent reviewer required.');
    const policy=tc24Policy_(),record=tc24RecordById_(row['Record ID']);
    const now=new Date(),utc=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()));
    const current=tc24RecordPolicy_(record,policy,tc24HoldMap_()[row['Record ID']],utc);
    if(decision==='APPROVE'&&!current.disposalEligible)throw new Error('TC24_RECHECK_FAILED: Retention eligibility changed or legal hold exists.');
    const state=decision==='APPROVE'?'APPROVED_FOR_MANUAL_REVIEW':'REJECTED';
    tc24Sheet_(TC24_DISPOSAL_SHEET_,false).getRange(row.rowNumber,6,1,4)
      .setValues([[state,rbacIdentity_(),now.toISOString(),tc24SafeCellText_(note)]]);
    tc24WriteLog_('DISPOSAL_REVIEW_'+decision,row['Record ID'],requestId,'SUCCESS','Decision recorded; deletion disabled');
    SpreadsheetApp.flush();return {success:true,status:state,deletionPerformed:false};
  });
}
