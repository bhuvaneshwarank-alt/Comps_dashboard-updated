/**
 * DigiClarity Compliance Dashboard — ONE-TIME PRE-PRODUCTION CLEANUP
 * Add as a separate Apps Script file (not to Index.html).
 * Source alignment: V77 HTML / V55 Code.gs.
 *
 * DRY RUN  : goLiveCleanupPreview()
 * BACKUP   : goLiveCleanupPrepare()
 * EXECUTE  : goLiveCleanupExecute('TOKEN', 'ERASE PREPRODUCTION TEST DATA')
 *
 * Does not modify User_Access, Role_Definitions, configuration, or historic
 * security/access audit logs. Clears only explicitly allowlisted data sheets.
 * Does not delete attachments in Google Drive or disable scheduled triggers.
 * NEVER invoke goLiveCleanupExecute from doGet/doPost or the dashboard UI.
 */

const GLC26 = Object.freeze({
  phrase: 'ERASE PREPRODUCTION TEST DATA',
  validityMinutes: 20,
  propKey: 'DGS_GO_LIVE_CLEANUP_PREPARED_V1',
  auditSheet: 'GoLive_Cleanup_Audit',
  // Explicit allowlist: any unknown/new sheet is PRESERVED by default.
  targets: [
    'Compliance_Master_Register', 'Compliance_Draft_Register',
    'Record_History', 'Record_Deletion_Register', 'Change_Requests',
    'Evidence_Register', 'Evidence_Metadata_Changes',
    'Evidence_Requests', 'Evidence_Request_Events', 'Evidence_Request_Notifications',
    'Control_Tests', 'Control_Test_Reviews',
    'Compliance_Reminder_Log', 'Compliance_Reminder_Runs', 'Compliance_Reminder_Archives',
    'Compliance_Monitor_Events', 'Compliance_Monitor_Runs',
    'Audit_Readiness_Exports',
    'IAM_Responses', 'IAM_Response_Reviews',
    'Access_Requests', 'Access_Review_Cycles', 'Access_Review_Items',
    'Temporary_Access_Grants', 'Privileged_Approvals', 'Emergency_Access_Log',
    'Dashboard_Sessions', 'Dashboard_Password_Reset'
  ],
  protected: [
    'Framework_Master', 'Framework_Control_Master', 'Lookup_Master',
    'Cross_Framework_Matrix', 'Approval_Matrix',
    'User_Access', 'Role_Definitions', 'Department_Access',
    'Access_Audit_Log', 'Dashboard_Security_Audit',
    'BuiltIn_Role_Change_Audit', 'Privileged_Audit_Log',
    'GoLive_Cleanup_Audit'
  ]
});

function glc26Identity_() {
  const signedIn = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const executor = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!signedIn || !executor || signedIn !== executor) {
    throw new Error('Run from the Apps Script editor as the signed-in permanent Admin. Identity is not verifiable.');
  }
  // Read-only authorization: require ACTIVE permanent Admin assignment in User_Access.
  const ss = glc26Spreadsheet_();
  const users = ss.getSheetByName('User_Access');
  if (!users || users.getLastRow() < 2) throw new Error('Cannot verify permanent Admin role in User_Access.');
  const values = users.getDataRange().getValues();
  if (String(values[0][0]) !== 'Email' || String(values[0][2]) !== 'Role' || String(values[0][3]) !== 'Active') {
    throw new Error('User_Access schema mismatch; cleanup blocked.');
  }
  const matches = values.slice(1).filter(r => String(r[0] || '').trim().toLowerCase() === signedIn);
  if (matches.length !== 1 || String(matches[0][2]).trim() !== 'Admin' ||
      !([true, 'true'].includes(matches[0][3]) || String(matches[0][3]).toLowerCase() === 'yes')) {
    throw new Error('A unique, active, permanently assigned Admin is required.');
  }
  return signedIn;
}

function glc26Spreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !ss.getId()) throw new Error('Bind this file to the dashboard spreadsheet project.');
  return ss;
}

function glc26Digest_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}

function glc26Signature_(sheet) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  // Include all populated cells, to prevent cleanup after unnoticed data changes.
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(lastRow && lastCol ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : []),
    Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(digest);
}

function glc26Inventory_(ss) {
  const protectedSet = new Set(GLC26.protected);
  const targets = [];
  const names = ss.getSheets().map(s => s.getName());
  GLC26.targets.forEach(name => {
    if (protectedSet.has(name)) throw new Error('Unsafe config: protected sheet in target list: ' + name);
    const sh = ss.getSheetByName(name);
    if (!sh) { targets.push({name, present:false, rows:0, cols:0, signature:''}); return; }
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr >= 2 && (lc < 1 || !String(sh.getRange(1, 1).getValue() || '').trim())) {
      throw new Error('Header check failed for ' + name + '. Manual inspection required.');
    }
    targets.push({name, present:true, rows:Math.max(0,lr-1), cols:lc, signature:glc26Signature_(sh)});
  });
  const preserved = names.filter(name => GLC26.targets.indexOf(name) < 0);
  return {spreadsheetId:ss.getId(), spreadsheetName:ss.getName(), targets, preserved,
    totalRows:targets.reduce((n,r) => n+r.rows,0)};
}

function glc26PlanDigest_(inventory) {
  return glc26Digest_(JSON.stringify({spreadsheetId:inventory.spreadsheetId,
    targets:inventory.targets, preserved:inventory.preserved.slice().sort()}));
}

function glc26Log_(ss, action, actor, details) {
  let sheet = ss.getSheetByName(GLC26.auditSheet);
  if (!sheet) {
    sheet = ss.insertSheet(GLC26.auditSheet);
    sheet.appendRow(['Timestamp','Action','Admin','Result / Details']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date().toISOString(), action, actor, JSON.stringify(details)]);
}

/** SAFE: Read only. Does NOT create sheets or clear data. */
function goLiveCleanupPreview() {
  const actor = glc26Identity_();
  const plan = glc26Inventory_(glc26Spreadsheet_());
  const result = {
    mode:'DRY RUN — NOTHING DELETED', actor, spreadsheet:plan.spreadsheetName,
    spreadsheetId:plan.spreadsheetId, planDigest:glc26PlanDigest_(plan),
    totalRowsToClear:plan.totalRows, sheetPlan:plan.targets.map(({name,present,rows}) => ({name,present,rowsToClear:rows})),
    preservedSheets:plan.preserved,
    notes:['No users/roles/framework mappings will be reset.',
      'Security and IAM/PAM audit logs stay in the active workbook.',
      'Backup is required before execution.',
      'Drive attachments are NOT deleted. Scheduled triggers are NOT disabled.',
      'Review active IAM/PAM grants and triggers before go-live.']
  };
  Logger.log(JSON.stringify(result,null,2));
  return result;
}

/** Makes a FULL backup in Google Drive; validates it; produces short-lived token. */
function goLiveCleanupPrepare() {
  const actor = glc26Identity_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = glc26Spreadsheet_();
    const plan = glc26Inventory_(ss);
    if (!plan.totalRows) throw new Error('Nothing to clear. Stopping before backup.');
    const suffix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    // The script performs an internal Drive backup; if Drive permission is denied, stop.
    const source = DriveApp.getFileById(ss.getId());
    const backup = source.makeCopy('PREPROD_UAT_BACKUP_' + suffix + '_' + ss.getName());
    const backedUp = SpreadsheetApp.openById(backup.getId());
    if (backedUp.getSheets().length !== ss.getSheets().length) {
      throw new Error('Backup did not contain expected number of sheets. Cleanup blocked.');
    }
    plan.targets.filter(r => r.present).forEach(r => {
      const copySheet = backedUp.getSheetByName(r.name);
      if (!copySheet || glc26Signature_(copySheet) !== r.signature) {
        throw new Error('Backup verification failed for ' + r.name + '. Cleanup blocked.');
      }
    });
    const token = Utilities.getUuid();
    const state = {actor, spreadsheetId:ss.getId(), planDigest:glc26PlanDigest_(plan),
      backupId:backup.getId(), backupUrl:backup.getUrl(), createdAt:Date.now(),
      tokenHash:glc26Digest_(token)};
    PropertiesService.getScriptProperties().setProperty(GLC26.propKey,JSON.stringify(state));
    glc26Log_(ss,'PREPARED',actor,{backupId:state.backupId, planDigest:state.planDigest, rows:plan.totalRows});
    // Log was written on a preserved sheet; its creation changes only preserved list on first use.
    const latest = glc26Inventory_(ss);
    state.planDigest = glc26PlanDigest_(latest);
    PropertiesService.getScriptProperties().setProperty(GLC26.propKey,JSON.stringify(state));
    Logger.log('BACKUP VERIFIED: ' + state.backupUrl);
    Logger.log('CONFIRMATION TOKEN (expires in 20 min): ' + token);
    Logger.log('Required phrase: ' + GLC26.phrase);
    return {prepared:true,backupUrl:state.backupUrl,token,expiresMinutes:GLC26.validityMinutes,
      totalRowsToClear:latest.totalRows, confirmationPhrase:GLC26.phrase};
  } finally { lock.releaseLock(); }
}

/**
 * DESTRUCTIVE — run ONLY after reading the dry-run and verifying the full backup.
 * Paste the token returned by goLiveCleanupPrepare and exact confirmation phrase.
 * E.g.: goLiveCleanupExecute('printed-token', 'ERASE PREPRODUCTION TEST DATA')
 */
function goLiveCleanupExecute(token, typedPhrase) {
  const actor = glc26Identity_();
  if (typedPhrase !== GLC26.phrase || !token || typeof token !== 'string') {
    throw new Error('Confirmation phrase or token missing/incorrect. NOTHING deleted.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let ss;
  try {
    ss = glc26Spreadsheet_();
    const raw = PropertiesService.getScriptProperties().getProperty(GLC26.propKey);
    if (!raw) throw new Error('No prepared cleanup. Run preview and prepare first.');
    const state = JSON.parse(raw);
    if (state.actor !== actor || state.spreadsheetId !== ss.getId() ||
        Date.now()-state.createdAt > GLC26.validityMinutes*60000 ||
        state.tokenHash !== glc26Digest_(token)) {
      throw new Error('Preparation expired or Admin/token mismatch. NOTHING deleted.');
    }
    const plan = glc26Inventory_(ss);
    if (glc26PlanDigest_(plan) !== state.planDigest) {
      throw new Error('Sheet data or structure changed after preparation. NOTHING deleted. Re-run preview/prepare.');
    }
    // Verify backup is still accessible and matches the prepared source.
    const backup = SpreadsheetApp.openById(state.backupId);
    plan.targets.filter(r => r.present).forEach(r => {
      const b = backup.getSheetByName(r.name);
      if (!b || glc26Signature_(b) !== r.signature) {
        throw new Error('Prepared backup could not be verified. NOTHING deleted.');
      }
    });
    // Invalidate token BEFORE first destructive write. Never retry blindly after an error.
    PropertiesService.getScriptProperties().deleteProperty(GLC26.propKey);
    glc26Log_(ss,'EXECUTION_STARTED',actor,{backupId:state.backupId, rows:plan.totalRows});
    const cleared = [];
    plan.targets.filter(r => r.present && r.rows > 0).forEach(r => {
      const sh = ss.getSheetByName(r.name);
      // Clear values only: retains headers, frozen rows, formats, validations.
      sh.getRange(2,1,r.rows,r.cols).clearContent();
      cleared.push({sheet:r.name,rows:r.rows});
      glc26Log_(ss,'SHEET_CLEARED',actor,{sheet:r.name, rows:r.rows});
    });
    const report = {success:true,actor,backupUrl:state.backupUrl,cleared,
      rowsCleared:cleared.reduce((s,r)=>s+r.rows,0),
      preservedSheets:plan.preserved,
      warning:'Reset does not remove Google Drive attachments or disable Apps Script triggers.'};
    glc26Log_(ss,'COMPLETED',actor,{backupId:state.backupId,rowsCleared:report.rowsCleared});
    Logger.log(JSON.stringify(report,null,2));
    return report;
  } catch (error) {
    if (ss) { try { glc26Log_(ss,'FAILED',actor,{message:String(error.message || error)}); } catch(ignore) {} }
    throw error;
  } finally {lock.releaseLock();}
}

// Apps Script editor wrapper. Fill BOTH strings only when ready to execute.
function goLiveCleanupExecuteFromEditor() {
  const CONFIRMATION_TOKEN = '';
  const CONFIRMATION_PHRASE = '';
  return goLiveCleanupExecute(CONFIRMATION_TOKEN, CONFIRMATION_PHRASE);
}
