/**********************************************************************
 * DigiClarity Compliance Dashboard — TC-21 audit archive (V61)
 * Additive module for existing V60 Code.gs. HTML V104 unchanged.
 *
 * SECURITY MODEL
 *  - Daily complete snapshots of configured audit/event sheets in
 *    a user-created, private, script-owner-controlled Drive folder.
 *  - Distinct, newly created ZIP and manifest files (never overwritten).
 *  - SHA-256 archive integrity + manifest hash chain anchored by an
 *    email containing hash/file IDs sent to an INDEPENDENT reviewer.
 *  - No raw log content, credentials, passwords, OTPs, or PHI in emails.
 *  - Fails closed if archive configuration or independent review is absent.
 *
 * LIMITATION: Drive files are NOT immutable/WORM, and a privileged
 * owner can modify source Sheets, archive files and Script Properties.
 * The independent reviewer must retain digest emails outside that
 * owner's control to make subsequent tampering detectable.
 **********************************************************************/
const TC21_ARCHIVE_VERSION_ = 'TC21-V61-2026-10';
const TC21_ARCHIVE_HANDLER_ = 'tc21DailyAuditArchiveJob';
const TC21_ARCHIVE_KEYS_ = {
  folder: 'TC21_ARCHIVE_FOLDER_ID',
  reviewer: 'TC21_INDEPENDENT_REVIEWER_EMAIL',
  trigger: 'TC21_ARCHIVE_TRIGGER_UID',
  lastHash: 'TC21_LAST_MANIFEST_HASH',
  lastId: 'TC21_LAST_MANIFEST_FILE_ID',
  lastDate: 'TC21_LAST_ARCHIVE_DAY',
  lastCounts: 'TC21_LAST_ROW_COUNTS_JSON'
};
// Core audit sheets must exist. Others are copied if available and
// individually shown as MISSING in the manifest; no new sheets are made.
const TC21_REQUIRED_LOG_SHEETS_ = [
  'Dashboard_Security_Audit', 'Access_Audit_Log',
  'Record_History', 'Change_Requests'
];
const TC21_OPTIONAL_LOG_SHEETS_ = [
  'Record_Deletion_Register', 'Privileged_Audit_Log',
  'Emergency_Access_Log', 'Compliance_Reminder_Log',
  'Compliance_Reminder_Runs', 'Compliance_Monitor_Events',
  'Compliance_Monitor_Runs', 'Evidence_Request_Notifications',
  'Audit_Readiness_Exports'
];
const TC21_MAX_UNCOMPRESSED_BYTES_ = 35 * 1024 * 1024;

function tc21HashBytes_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(function(b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); })
    .join('');
}
function tc21HashText_(text) {
  return tc21HashBytes_(Utilities.newBlob(String(text), 'text/plain').getBytes());
}
function tc21Time_() { return new Date().toISOString(); }
function tc21Day_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd');
}
function tc21Actor_() {
  return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
}
function tc21ArchiveConfig_() {
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty(TC21_ARCHIVE_KEYS_.folder) || '').trim();
  const email = String(props.getProperty(TC21_ARCHIVE_KEYS_.reviewer) || '').trim().toLowerCase();
  if (!id) throw new Error('TC21: Set TC21_ARCHIVE_FOLDER_ID in Script Properties first.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('TC21: Configure an independent reviewer mailbox in Script Properties.');
  }
  const actor = tc21Actor_();
  if (!actor || actor === email) {
    throw new Error('TC21: Archive operator and independent reviewer must be different accounts.');
  }
  const folder = DriveApp.getFolderById(id);
  const owner = folder.getOwner();
  if (!owner || String(owner.getEmail() || '').toLowerCase() !== actor) {
    throw new Error('TC21: Archive folder must be owned by the Apps Script execution account.');
  }
  if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    throw new Error('TC21: Folder link sharing must be Restricted/Private.');
  }
  const thirdPartyAccess = folder.getEditors().concat(folder.getViewers())
    .map(function(u) { return String(u.getEmail() || '').toLowerCase(); })
    .filter(function(e) { return e && e !== actor; });
  if (thirdPartyAccess.length) {
    throw new Error('TC21: Remove other direct folder editors/viewers; reviewer receives digest by email.');
  }
  return { folder: folder, reviewer: email, owner: actor, props: props };
}
function tc21Snapshot_() {
  const ss = getSpreadsheet_();
  const names = TC21_REQUIRED_LOG_SHEETS_.concat(TC21_OPTIONAL_LOG_SHEETS_);
  const data = {};
  const counts = {};
  const missing = [];
  names.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      counts[name] = null;
      missing.push(name);
      return;
    }
    // Preserve the COMPLETE sheet (including header and all rows) so that
    // deleting/editing log rows later can be detected against this archive.
    const values = sheet.getDataRange().getValues();
    const normalized = values.map(function(row) {
      return row.map(function(cell) {
        return cell instanceof Date ? cell.toISOString() : cell;
      });
    });
    data[name] = {sheetId: sheet.getSheetId(), values: normalized};
    counts[name] = Math.max(0, normalized.length - 1);
  });
  const missingCore = TC21_REQUIRED_LOG_SHEETS_.filter(function(n) { return missing.indexOf(n) >= 0; });
  if (missingCore.length) throw new Error('TC21: Required logs missing; archive aborted: ' + missingCore.join(', '));
  return {data: data, counts: counts, missing: missing};
}
function tc21ReadManifest_(fileId) {
  const contents = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
  const manifest = JSON.parse(contents);
  if (!manifest || !manifest.core || !manifest.manifestSha256) throw new Error('TC21: Manifest format invalid.');
  const calculated = tc21HashText_(JSON.stringify(manifest.core));
  if (calculated !== manifest.manifestSha256) throw new Error('TC21: Manifest digest mismatch. Possible alteration.');
  return manifest;
}
function tc21ValidateLastAnchor_(cfg) {
  const lastId = cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastId);
  const lastHash = cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastHash);
  if (!!lastId !== !!lastHash) throw new Error('TC21: Broken prior archive anchor configuration.');
  if (!lastId) return '';
  const last = tc21ReadManifest_(lastId);
  if (last.manifestSha256 !== lastHash) {
    throw new Error('TC21: Previous manifest anchor hash does not match. Possible alteration.');
  }
  const archiveFile = DriveApp.getFileById(last.core.archiveFileId);
  const currentDigest = tc21HashBytes_(archiveFile.getBlob().getBytes());
  if (currentDigest !== last.core.archiveSha256) {
    throw new Error('TC21: Previous audit archive changed or corrupted.');
  }
  return lastHash;
}
function tc21MakeArchive_(manual) {
  const cfg = tc21ArchiveConfig_();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const day = tc21Day_();
    if (!manual && cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastDate) === day) {
      return {success: true, skipped: true, message: 'Daily audit archive already completed.'};
    }
    const previousHash = tc21ValidateLastAnchor_(cfg);
    const snapshot = tc21Snapshot_();
    const createdAt = tc21Time_();
    const payload = JSON.stringify({format:TC21_ARCHIVE_VERSION_, generatedAt:createdAt,
      spreadsheetId:getSpreadsheet_().getId(), sheets:snapshot.data});
    const bytes = Utilities.newBlob(payload, 'application/json', 'audit-snapshot.json').getBytes();
    if (bytes.length > TC21_MAX_UNCOMPRESSED_BYTES_) {
      throw new Error('TC21: Snapshot is too large for configured archive threshold; contact IT to plan split exports.');
    }
    const baseName = 'DGS_TC21_' + day + '_' + Utilities.getUuid();
    const zipped = Utilities.zip([Utilities.newBlob(bytes, 'application/json', 'audit-snapshot.json')], baseName + '.zip');
    const archiveBytes = zipped.getBytes();
    const archiveHash = tc21HashBytes_(archiveBytes);
    const archive = cfg.folder.createFile(zipped).setName(baseName + '.zip');
    if (tc21HashBytes_(archive.getBlob().getBytes()) !== archiveHash) {
      throw new Error('TC21: ZIP read-back verification failed.');
    }
    const oldCounts = cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastId)
      ? tc21ReadManifest_(cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastId)).core.rowCounts : {};
    const decreases = Object.keys(snapshot.counts).filter(function(name) {
      return typeof oldCounts[name] === 'number' &&
        typeof snapshot.counts[name] === 'number' &&
        snapshot.counts[name] < oldCounts[name];
    });
    const core = {
      format: TC21_ARCHIVE_VERSION_, timestamp: createdAt, archiveDate: day,
      archiveFileId: archive.getId(), archiveSha256: archiveHash,
      previousManifestSha256: previousHash, rowCounts: snapshot.counts,
      missingOptionalSheets: snapshot.missing,
      decreasedRowCountSheets: decreases
    };
    const manifest = {core: core, manifestSha256: tc21HashText_(JSON.stringify(core))};
    const manifestFile = cfg.folder.createFile(
      Utilities.newBlob(JSON.stringify(manifest, null, 2), 'application/json', baseName + '.manifest.json')
    );
    if (tc21ReadManifest_(manifestFile.getId()).manifestSha256 !== manifest.manifestSha256) {
      throw new Error('TC21: Manifest read-back verification failed.');
    }
    // The email goes to an INDEPENDENT mailbox and becomes an external
    // anchor for the manifest/hash chain. No sensitive log rows are emailed.
    const subject = '[DGS Audit Archive] ' + day + (decreases.length ? ' - REVIEW ROW DECREASE' : ' - Digest');
    const body = [
      'DigiClarity TC-21 independent audit archive receipt',
      'Timestamp (UTC): ' + createdAt,
      'Archive date: ' + day,
      'Archive ZIP file ID: ' + archive.getId(),
      'Archive SHA-256: ' + archiveHash,
      'Manifest file ID: ' + manifestFile.getId(),
      'Manifest SHA-256: ' + manifest.manifestSha256,
      'Previous manifest SHA-256: ' + (previousHash || '(genesis)'),
      'Row counts (not log content): ' + JSON.stringify(snapshot.counts),
      'Missing optional sheets: ' + (snapshot.missing.join(', ') || 'none'),
      'Row-count decreases requiring investigation: ' + (decreases.join(', ') || 'none'),
      '',
      'Retain this email in a reviewer-controlled mailbox with retention and restricted access.',
      'Drive archive files are not immutable; this independent receipt detects later inconsistencies.',
      'Do not reply with PHI, passwords, tokens, or log contents.'
    ].join('\n');
    MailApp.sendEmail({to:cfg.reviewer, subject:subject, body:body, name:'DigiClarity Audit Archive'});
    // Advance the chain only after independent receipt is successfully sent.
    cfg.props.setProperties({
      TC21_LAST_MANIFEST_HASH: manifest.manifestSha256,
      TC21_LAST_MANIFEST_FILE_ID: manifestFile.getId(),
      TC21_LAST_ARCHIVE_DAY: day,
      TC21_LAST_ROW_COUNTS_JSON: JSON.stringify(snapshot.counts)
    });
    return {success:true, date:day, archiveFileId:archive.getId(),
      manifestFileId:manifestFile.getId(), manifestSha256:manifest.manifestSha256,
      counts:snapshot.counts, decreasedRowCountSheets:decreases};
  } finally { lock.releaseLock(); }
}
// Editor-only operations: trailing underscores keep these out of
// google.script.run (dashboardSecureInvoke does not dispatch them).
function tc21ArchiveConfigurationCheck_() {
  const cfg = tc21ArchiveConfig_();
  return {ready:true, owner:cfg.owner, reviewer:cfg.reviewer, folderId:cfg.folder.getId(),
    requiredSheets:TC21_REQUIRED_LOG_SHEETS_};
}
function tc21RunArchiveNow_() { return tc21MakeArchive_(true); }
function tc21VerifyLatestArchive_() {
  const cfg = tc21ArchiveConfig_();
  const lastHash = tc21ValidateLastAnchor_(cfg);
  if (!lastHash) return {success:false, message:'No archived receipts yet.'};
  const manifest = tc21ReadManifest_(cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastId));
  return {success:true, manifestSha256:lastHash,
    archiveFileId:manifest.core.archiveFileId, manifestFileId:cfg.props.getProperty(TC21_ARCHIVE_KEYS_.lastId),
    rowCounts:manifest.core.rowCounts, decreasedRowCountSheets:manifest.core.decreasedRowCountSheets};
}
function tc21EnableDailyArchive_() {
  const cfg = tc21ArchiveConfig_();
  const previous = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === TC21_ARCHIVE_HANDLER_;
  });
  if (previous.length > 1) throw new Error('TC21: Duplicate archive triggers detected. Resolve manually.');
  if (previous.length === 1) return {success:true, alreadyEnabled:true};
  const trigger = ScriptApp.newTrigger(TC21_ARCHIVE_HANDLER_).timeBased().everyDays(1).atHour(2).create();
  cfg.props.setProperty(TC21_ARCHIVE_KEYS_.trigger, trigger.getUniqueId());
  return {success:true, enabled:true, triggerUid:trigger.getUniqueId()};
}
function tc21DisableDailyArchive_() {
  ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === TC21_ARCHIVE_HANDLER_;
  }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().deleteProperty(TC21_ARCHIVE_KEYS_.trigger);
  return {success:true, enabled:false};
}
// Scheduled entry point must be public for Apps Script time-based triggers.
// Fail closed if called from browser/console without the specific trigger UID.
function tc21DailyAuditArchiveJob(event) {
  const expected = PropertiesService.getScriptProperties().getProperty(TC21_ARCHIVE_KEYS_.trigger);
  if (!expected || !event || String(event.triggerUid || '') !== expected) {
    throw new Error('TC21: Unauthorized scheduled archive invocation.');
  }
  return tc21MakeArchive_(false);
}


function tc21CheckSetup() {
  const result = tc21ArchiveConfigurationCheck_();

  Logger.log(
    "TC-21 CONFIGURATION: " +
    JSON.stringify({ ready: result.ready })
  );
}

function tc21FirstBackupTest() {
  const active = String(
    Session.getActiveUser().getEmail() || ""
  ).trim().toLowerCase();

  const owner = String(
    Session.getEffectiveUser().getEmail() || ""
  ).trim().toLowerCase();

  if (!active || active !== owner) {
    throw new Error("TC21: Owner-only manual test.");
  }

  const result = tc21RunArchiveNow_();
  const check = tc21VerifyLatestArchive_();

  if (!result.success ||
      !check.success ||
      check.manifestSha256 !== result.manifestSha256) {
    throw new Error("TC21: Archive verification failed.");
  }

  Logger.log("TC-21 FIRST BACKUP: " + JSON.stringify({
    archiveCreated: true,
    integrityVerified: true,
    manifestMatched: true,
    archiveDate: result.date
  }));
}
