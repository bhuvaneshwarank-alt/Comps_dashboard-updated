/**
 * DigiClarity Compliance Dashboard | TC-25 automated production backups
 * Add as a NEW .gs file in the existing, spreadsheet-bound Apps Script project.
 * Designed for Code.gs V65 and HTML V112. No HTML or deployment changes.
 *
 * Backed up: Entire bound Google Sheets database + every file immediately inside
 *             the pinned production IAM Shared Drive evidence folder.
 * Not backed up: Apps Script source/deployment, triggers, Script Properties,
 *                other Drive folders, or Google Workspace identity configuration.
 *
 * Security:
 *  - All callable routines end in '_' (private to google.script.run).
 *  - Execute configuration/one-shot/schedule routines only from Script editor.
 *  - Explicit pinned source and destination IDs; validated Shared Drive or My Drive backup.
 *  - Fail-closed on bad configuration, unexpected evidence folders, or copy errors.
 *  - No deletion, changing permissions, or modifying source evidence.
 *  - Disabled by default. No trigger is created until explicitly installed.
 *  - The backup includes credential hashes/salts and session records: only use an
 *    IT-approved, restricted backup folder (Shared Drive or approved My Drive).
 *
 * Advanced Google service: Drive API v3 (already used by V65).
 */

var TC25BK_LOG_SHEET_ = 'DR_Backup_Runs';
var TC25BK_LOG_HEADERS_ = [
  'Run ID', 'Started At', 'Completed At', 'Status', 'Backup Snapshot Folder ID',
  'Database Backup File ID', 'Database Tabs', 'Master Records', 'History Events',
  'Evidence Files Expected', 'Evidence Files Copied', 'Evidence Files SHA256 Verified',
  'Elapsed Seconds', 'Failure / Note'
];

/** Only run from the Apps Script editor. Web clients cannot invoke _ functions. */
function tc25bkPreflight_() {
  var config = tc25bkConfig_();
  var pre = tc25bkValidate_(config);
  var result = {
    safeToConfigure: true,
    enabled: config.enabled,
    sourceSpreadsheetId: config.sourceSheetId,
    liveEvidenceFolderId: pre.sourceFolderId,
    backupFolderId: config.backupFolderId,
    backupLocationType: pre.backupLocationType,
    backupSharedDriveId: pre.backupDriveId || null,
    destinationOnDifferentSharedDrive: pre.backupLocationType === 'SHARED_DRIVE'
      ? pre.backupDriveId !== pre.evidenceDriveId : null,
    destinationOutsideLiveSharedDrive: pre.backupLocationType === 'MY_DRIVE'
      ? true : pre.backupDriveId !== pre.evidenceDriveId,
    destinationOwnershipAndSharingVerified: pre.backupLocationType === 'MY_DRIVE',
    sharedDriveMembershipReview: pre.backupLocationType === 'SHARED_DRIVE' ? 'Manual check required' : null,
    managerAccessContinuity: 'Not assessed by this preflight',
    evidenceFilesNow: pre.files.length,
    maxEvidenceFilesPerSnapshot: config.maxFiles,
    frequencyHours: config.frequencyHours,
    alertRecipient: config.alertEmail,
    triggerInstalled: ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'tc25bkDailyTrigger_';
    }),
    installedFrequencyHours: Number(PropertiesService.getScriptProperties()
      .getProperty('TC25BK_INSTALLED_FREQUENCY_HOURS') || '0') || null,
    triggerScheduleMatchesRecordedSetting: ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'tc25bkDailyTrigger_' &&
        t.getUniqueId() === String(PropertiesService.getScriptProperties().getProperty('TC25BK_TRIGGER_UID') || '');
    }) && Number(PropertiesService.getScriptProperties().getProperty('TC25BK_INSTALLED_FREQUENCY_HOURS') || '0') === config.frequencyHours,
    note: 'Read-only preflight: no files copied, no trigger enabled, no live data modified.'
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/** Run AFTER preflight, and only after explicitly setting TC25BK_ENABLED=YES. */
function tc25bkRunOnce_() {
  return tc25bkExecute_('MANUAL');
}

/** Installed time-driven handler. May be invoked only by its own trigger UID. */
function tc25bkDailyTrigger_(event) {
  var props = PropertiesService.getScriptProperties();
  var expectedUid = String(props.getProperty('TC25BK_TRIGGER_UID') || '');
  if (!expectedUid || !event || String(event.triggerUid || '') !== expectedUid) {
    throw new Error('TC25BK_TRIGGER_UNAUTHORIZED: Unexpected trigger invocation.');
  }
  return tc25bkExecute_('SCHEDULED');
}

/** Create a trigger at the requested interval. The trigger executes the private handler only. */
function tc25bkCreateScheduleTrigger_(hours) {
  var builder = ScriptApp.newTrigger('tc25bkDailyTrigger_').timeBased();
  if (hours === 1) return builder.everyHours(1).create();
  if (hours === 12) return builder.everyHours(12).create();
  if (hours === 24) return builder.everyDays(1).atHour(2).inTimezone('Asia/Kolkata').create();
  throw new Error('TC25BK_INVALID_INTERVAL: Only 1, 12 or 24 hours supported.');
}

/**
 * Install or switch the one authorized trigger to match TC25BK_FREQUENCY_HOURS.
 * Apps Script does NOT expose an existing trigger's clock interval. We record our
 * installed interval in Script Properties and conservatively replace a legacy
 * trigger when the installed interval is unknown.
 *
 * Create -> pin new UID -> delete old trigger. While both temporarily exist,
 * the UID check in tc25bkDailyTrigger_ rejects the old one.
 */
function tc25bkInstallSchedule_() {
  var config = tc25bkConfig_();
  if (!config.enabled) {
    throw new Error('TC25BK_DISABLED: Set TC25BK_ENABLED=YES only after approval.');
  }
  tc25bkValidate_(config);
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('TC25BK_LAST_STATUS') !== 'PASS' ||
      props.getProperty('TC25BK_LAST_SUCCESS_TARGET_ID') !== config.backupFolderId) {
    throw new Error('TC25BK_TARGET_MANUAL_TEST_REQUIRED: Complete a PASS backup to the currently configured folder.');
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('TC25BK_SCHEDULE_BUSY: Another backup or schedule update is in progress.');
  }
  try {
    var existing = ScriptApp.getProjectTriggers().filter(function(t) {
      return t.getHandlerFunction() === 'tc25bkDailyTrigger_';
    });
    if (existing.length > 1) {
      throw new Error('TC25BK_MULTIPLE_TRIGGERS: Stop and review duplicates manually.');
    }
    var configuredUid = String(props.getProperty('TC25BK_TRIGGER_UID') || '');
    var installedHours = Number(props.getProperty('TC25BK_INSTALLED_FREQUENCY_HOURS') || '0');
    var canReuse = existing.length === 1 &&
      configuredUid === existing[0].getUniqueId() &&
      installedHours === config.frequencyHours;
    if (canReuse) {
      var reused = {installed: true, reused: true, replaced: false,
        frequencyHours: config.frequencyHours, triggerCount: 1};
      Logger.log(JSON.stringify(reused));
      return reused;
    }

    // Preserve existing settings if creation or pinning the new trigger fails.
    var fresh = tc25bkCreateScheduleTrigger_(config.frequencyHours);
    try {
      props.setProperties({
        TC25BK_TRIGGER_UID: fresh.getUniqueId(),
        TC25BK_INSTALLED_FREQUENCY_HOURS: String(config.frequencyHours)
      });
    } catch (pinError) {
      try { ScriptApp.deleteTrigger(fresh); } catch (cleanupError) {
        Logger.log('TC25BK_NEW_TRIGGER_CLEANUP_FAILED: ' + String(cleanupError));
      }
      throw pinError;
    }

    // Do not count this as a clean installation if an old trigger can't be removed.
    var oldRemoved = 0;
    existing.forEach(function(t) {
      try {
        ScriptApp.deleteTrigger(t);
        oldRemoved++;
      } catch (removeError) {
        throw new Error('TC25BK_OLD_TRIGGER_REMOVAL_FAILED: New UID pinned, but old trigger cleanup needs manual review: ' + String(removeError));
      }
    });
    var installed = {installed: true, reused: false,
      replaced: oldRemoved > 0, removedOldTriggers: oldRemoved,
      frequencyHours: config.frequencyHours, triggerCount: 1};
    Logger.log(JSON.stringify(installed));
    return installed;
  } finally {
    lock.releaseLock();
  }
}

/** Stop scheduling; existing snapshots are never deleted. */
function tc25bkDisableSchedule_() {
  tc25bkAssertOwner_();
  var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'tc25bkDailyTrigger_';
  });
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().deleteProperty('TC25BK_TRIGGER_UID');
  PropertiesService.getScriptProperties().deleteProperty('TC25BK_INSTALLED_FREQUENCY_HOURS');
  PropertiesService.getScriptProperties().setProperty('TC25BK_ENABLED', 'NO');
  return {disabled: true, triggersRemoved: triggers.length};
}

function tc25bkAssertOwner_() {
  var expected = String(PropertiesService.getScriptProperties()
    .getProperty('TC25BK_OWNER_EMAIL') || '').trim().toLowerCase();
  var effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  var active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!expected || !effective || effective !== expected || (active && active !== expected)) {
    throw new Error('TC25BK_OWNER_MISMATCH: Execute as the IT-approved project/trigger owner.');
  }
}

function tc25bkConfig_() {
  tc25bkAssertOwner_();
  var props = PropertiesService.getScriptProperties();
  var numberFiles = Number(props.getProperty('TC25BK_MAX_FILES') || '100');
  var hours = Number(props.getProperty('TC25BK_FREQUENCY_HOURS') || '12');
  var email = String(props.getProperty('TC25BK_ALERT_EMAIL') || '').trim().toLowerCase();
  if (!/^[a-z0-9._%+\-]+@(digiclarity\.in|clarityrcm\.com)$/.test(email)) {
    throw new Error('TC25BK_ALERT_EMAIL: Use an approved company alert recipient.');
  }
  if (!Number.isInteger(numberFiles) || numberFiles < 1 || numberFiles > 500) {
    throw new Error('TC25BK_MAX_FILES: Integer from 1 to 500 required.');
  }
  if (hours !== 1 && hours !== 12 && hours !== 24) {
    throw new Error('TC25BK_FREQUENCY_HOURS: Only 1, 12 or 24 supported.');
  }
  var location = String(props.getProperty('TC25BK_BACKUP_LOCATION') || 'SHARED_DRIVE').trim().toUpperCase();
  if (location !== 'SHARED_DRIVE' && location !== 'MY_DRIVE') {
    throw new Error('TC25BK_BACKUP_LOCATION: Choose SHARED_DRIVE or MY_DRIVE.');
  }
  var enabled = String(props.getProperty('TC25BK_ENABLED') || 'NO').trim().toUpperCase();
  if (enabled !== 'YES' && enabled !== 'NO') {
    throw new Error('TC25BK_ENABLED: Must be YES or NO.');
  }
  return {
    enabled: enabled === 'YES',
    sourceSheetId: String(props.getProperty('TC25BK_SOURCE_SHEET_ID') || '').trim(),
    backupFolderId: String(props.getProperty('TC25BK_BACKUP_FOLDER_ID') || '').trim(),
    backupDriveId: String(props.getProperty('TC25BK_BACKUP_DRIVE_ID') || '').trim(),
    backupLocationType: location,
    backupOwnerEmail: String(props.getProperty('TC25BK_OWNER_EMAIL') || '').trim().toLowerCase(),
    allowedCustodianEmails: String(props.getProperty('TC25BK_ADDITIONAL_CUSTODIAN_EMAILS') || '').trim(),
    alertEmail: email,
    frequencyHours: hours,
    maxFiles: numberFiles
  };
}

function tc25bkDriveMeta_(fileId) {
  if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.get) {
    throw new Error('TC25BK_DRIVE_API_REQUIRED: Enable Advanced Drive API v3.');
  }
  return Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,parents,trashed,size,md5Checksum,owners(emailAddress),ownedByMe'
  });
}

/** Validate one My Drive object: expected ownership and no broad/unapproved grants.
 * On any inaccessible permission metadata, fail closed (no backup execution). */
function tc25bkMyDriveAcl_(fileId, cfg) {
  var meta = tc25bkDriveMeta_(fileId);
  var expected = cfg.backupOwnerEmail;
  var owned = (meta.owners || []).some(function(owner) {
    return String(owner.emailAddress || '').toLowerCase() === expected;
  });
  if (!meta || meta.trashed || meta.driveId || !owned || meta.ownedByMe !== true) {
    throw new Error('TC25BK_MYDRIVE_OWNER_INVALID: Unexpected item owner or location.');
  }
  var allowed = {};
  allowed[expected] = true;
  if (cfg.allowedCustodianEmails) {
    cfg.allowedCustodianEmails.split(/[;,]/).forEach(function(raw) {
      var email = String(raw).trim().toLowerCase();
      if (!email || !/^[a-z0-9._%+\-]+@(digiclarity\.in|clarityrcm\.com)$/.test(email)) {
        throw new Error('TC25BK_CUSTODIAN_ALLOWLIST_INVALID: Only approved corporate addresses are allowed.');
      }
      allowed[email] = true;
    });
  }
  if (typeof Drive === 'undefined' || !Drive.Permissions || !Drive.Permissions.list) {
    throw new Error('TC25BK_DRIVE_PERMISSION_API_REQUIRED: Cannot verify folder sharing.');
  }
  var token = null;
  var count = 0;
  do {
    var query = {
      supportsAllDrives: true,
      fields: 'nextPageToken,permissions(id,type,emailAddress,role)',
      pageSize: 100
    };
    if (token) query.pageToken = token;
    var page = Drive.Permissions.list(fileId, query);
    if (!page || !page.permissions) throw new Error('TC25BK_ACL_UNAVAILABLE: Cannot inspect file permissions.');
    page.permissions.forEach(function(permission) {
      count++;
      var email = String(permission.emailAddress || '').toLowerCase();
      if (permission.type !== 'user' || !allowed[email]) {
        throw new Error('TC25BK_ACL_NOT_RESTRICTED: Unapproved individual, group, domain or public permission.');
      }
    });
    token = page.nextPageToken || null;
    if (count > 500) throw new Error('TC25BK_ACL_TOO_LARGE: Unexpected number of sharing grants.');
  } while (token);
  if (!count) throw new Error('TC25BK_ACL_EMPTY: No ownership permissions could be verified.');
  return meta;
}

function tc25bkValidate_(cfg) {
  if (!cfg.sourceSheetId || !cfg.backupFolderId ||
      (cfg.backupLocationType === 'SHARED_DRIVE' && !cfg.backupDriveId)) {
    throw new Error('TC25BK_CONFIG_MISSING: Pin the production sheet, backup location and backup folder IDs.');
  }
  var spreadsheet = getSpreadsheet_(); // Existing V65 helper; requires bound project.
  if (spreadsheet.getId() !== cfg.sourceSheetId) {
    throw new Error('TC25BK_SOURCE_MISMATCH: Refusing to back up an unexpected spreadsheet.');
  }
  var props = PropertiesService.getScriptProperties();
  if (String(props.getProperty('IAM_EVIDENCE_STORAGE_MODE') || '').toUpperCase() !== 'SHARED_DRIVE') {
    throw new Error('TC25BK_SHARED_DRIVE_REQUIRED: Production evidence must remain on its pinned Shared Drive.');
  }
  var liveId = String(props.getProperty('IAM_EVIDENCE_FOLDER_ID') || '').trim();
  var tc25PinnedId = String(props.getProperty('TC25_TEST_LIVE_FOLDER_ID') || '').trim();
  if (!liveId || liveId !== tc25PinnedId || liveId === cfg.backupFolderId) {
    throw new Error('TC25BK_SOURCE_DESTINATION_MISMATCH: Evidence pin or backup destination invalid.');
  }
  var liveFolder = iamPrivateFolder_(); // Reuse V65 source ownership / Shared Drive checks.
  var liveMeta = tc25bkDriveMeta_(liveId);
  if (!liveMeta || liveMeta.trashed || !liveMeta.driveId) {
    throw new Error('TC25BK_LIVE_FOLDER_INVALID: Cannot verify live evidence folder.');
  }
  var destMeta = tc25bkDriveMeta_(cfg.backupFolderId);
  if (!destMeta || destMeta.trashed ||
      destMeta.mimeType !== 'application/vnd.google-apps.folder') {
    throw new Error('TC25BK_DESTINATION_INVALID: Backup destination must be an active folder.');
  }
  var destinationDriveId = '';
  if (cfg.backupLocationType === 'SHARED_DRIVE') {
    if (!destMeta.driveId || destMeta.driveId !== cfg.backupDriveId) {
      throw new Error('TC25BK_DESTINATION_INVALID: Backup folder must match the pinned Shared Drive.');
    }
    // Preserve the original ancestry check for Shared Drive destinations.
    var parentId = cfg.backupFolderId;
    for (var depth = 0; depth < 30 && parentId; depth++) {
      if (parentId === liveId) {
        throw new Error('TC25BK_DEST_INSIDE_LIVE: Backup folder cannot be inside live evidence.');
      }
      if (parentId === destMeta.driveId) { parentId = ''; break; }
      var parentMeta = tc25bkDriveMeta_(parentId);
      parentId = (parentMeta.parents && parentMeta.parents.length) ? parentMeta.parents[0] : '';
    }
    if (parentId) throw new Error('TC25BK_FOLDER_DEPTH: Cannot establish destination ancestry.');
    destinationDriveId = destMeta.driveId;
  } else if (cfg.backupLocationType === 'MY_DRIVE') {
    // Fail closed if an obsolete Shared Drive ID remains pinned from an earlier setup.
    if (cfg.backupDriveId) {
      throw new Error('TC25BK_MYDRIVE_CONFIG_CONFLICT: Clear TC25BK_BACKUP_DRIVE_ID for My Drive mode.');
    }
    if (destMeta.driveId) {
      throw new Error('TC25BK_MYDRIVE_LOCATION_INVALID: Folder is actually in a Shared Drive.');
    }
    if (destMeta.name !== 'DGS_DR_Automated_Backups' ||
        !destMeta.parents || destMeta.parents.length !== 1) {
      throw new Error('TC25BK_MYDRIVE_FOLDER_INVALID: Unexpected backup folder name or parents.');
    }
    var parentFolder = tc25bkDriveMeta_(destMeta.parents[0]);
    if (!parentFolder || parentFolder.name !== 'DGS Compliance DR Backup' ||
        parentFolder.mimeType !== 'application/vnd.google-apps.folder' ||
        !parentFolder.parents || parentFolder.parents.length !== 1) {
      throw new Error('TC25BK_MYDRIVE_PARENT_INVALID: Backup parent folder not recognized.');
    }
    var rootMeta = Drive.Files.get('root', {fields: 'id', supportsAllDrives: true});
    if (!rootMeta || !rootMeta.id || parentFolder.parents[0] !== rootMeta.id) {
      throw new Error('TC25BK_MYDRIVE_ANCESTRY_INVALID: Destination must be under the owner My Drive root.');
    }
    // Reject unsafe direct AND inherited permissions on both levels.
    tc25bkMyDriveAcl_(parentFolder.id, cfg);
    tc25bkMyDriveAcl_(destMeta.id, cfg);
  } else {
    throw new Error('TC25BK_LOCATION_UNSUPPORTED: Use MY_DRIVE or SHARED_DRIVE.');
  }
  var backupFolder = DriveApp.getFolderById(cfg.backupFolderId);
  if (liveFolder.getFolders().hasNext()) {
    throw new Error('TC25BK_NESTED_EVIDENCE: Live evidence contains subfolders; review backup scope.');
  }
  var files = [];
  var it = liveFolder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    files.push(f);
    if (files.length > cfg.maxFiles) {
      throw new Error('TC25BK_FILE_LIMIT: Evidence count exceeds approved snapshot cap.');
    }
  }
  return {
    spreadsheet: spreadsheet, sourceFolderId: liveId, sourceFolder: liveFolder,
    backupFolder: backupFolder, evidenceDriveId: liveMeta.driveId,
    backupDriveId: destinationDriveId, backupLocationType: cfg.backupLocationType,
    files: files
  };
}

function tc25bkSha256_(bytes) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return digest.map(function(b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function tc25bkSheetInventory_(ss) {
  return ss.getSheets().map(function(sheet) {
    return {name: sheet.getName(), rows: sheet.getLastRow(), columns: sheet.getLastColumn()};
  }).sort(function(a, b) { return a.name.localeCompare(b.name); });
}

function tc25bkCompareSheets_(original, restored) {
  var src = tc25bkSheetInventory_(original);
  var dst = tc25bkSheetInventory_(restored);
  if (JSON.stringify(src) !== JSON.stringify(dst)) {
    throw new Error('TC25BK_SHEET_MISMATCH: Tab names or data dimensions differ in copied database.');
  }
  // Verify full displayed values for critical registers without exposing values
  // in the manifest/log. Google may reformat underlying native cell values.
  ['Compliance_Master_Register', 'Record_History'].forEach(function(name) {
    var a = original.getSheetByName(name);
    var b = restored.getSheetByName(name);
    if (!a || !b) throw new Error('TC25BK_REQUIRED_SHEET_MISSING: ' + name);
    var cells = a.getLastRow() * a.getLastColumn();
    if (cells > 100000) {
      throw new Error('TC25BK_SHEET_LIMIT: Critical tab exceeds current digest threshold.');
    }
    var hashA = tc25bkSha256_(Utilities.newBlob(JSON.stringify(a.getDataRange().getDisplayValues())).getBytes());
    var hashB = tc25bkSha256_(Utilities.newBlob(JSON.stringify(b.getDataRange().getDisplayValues())).getBytes());
    if (hashA !== hashB) {
      throw new Error('TC25BK_DATA_DIGEST_MISMATCH: Copied critical register differs: ' + name);
    }
  });
  return {
    tabs: src.length,
    masterRecords: Math.max(0, restored.getSheetByName('Compliance_Master_Register').getLastRow() - 1),
    historyEvents: Math.max(0, restored.getSheetByName('Record_History').getLastRow() - 1),
    inventory: dst
  };
}

function tc25bkCopyEvidence_(sourceFiles, evidenceDest, cfg) {
  var details = [];
  var verified = 0;
  sourceFiles.forEach(function(src) {
    var copied = src.makeCopy(src.getName(), evidenceDest);
    if (cfg.backupLocationType === 'MY_DRIVE') tc25bkMyDriveAcl_(copied.getId(), cfg);
    var sourceMeta = tc25bkDriveMeta_(src.getId());
    var destMeta = tc25bkDriveMeta_(copied.getId());
    if (destMeta.trashed || sourceMeta.mimeType !== destMeta.mimeType ||
        Number(src.getSize()) !== Number(copied.getSize())) {
      throw new Error('TC25BK_EVIDENCE_META_MISMATCH: ' + src.getId());
    }
    var hashSource, hashCopied, method;
    // SHA-256 for modest files, MD5 Drive-provided checksum for larger binaries.
    if (Number(src.getSize()) <= 20 * 1024 * 1024) {
      hashSource = tc25bkSha256_(src.getBlob().getBytes());
      hashCopied = tc25bkSha256_(copied.getBlob().getBytes());
      method = 'SHA256';
    } else if (sourceMeta.md5Checksum && destMeta.md5Checksum) {
      hashSource = String(sourceMeta.md5Checksum);
      hashCopied = String(destMeta.md5Checksum);
      method = 'DriveMD5';
    } else {
      throw new Error('TC25BK_CHECKSUM_UNAVAILABLE: Cannot validate oversized evidence file.');
    }
    if (hashSource !== hashCopied) {
      throw new Error('TC25BK_EVIDENCE_HASH_MISMATCH: ' + src.getId());
    }
    if (method === 'SHA256') verified++;
    details.push({
      sourceFileId: src.getId(), copiedFileId: copied.getId(),
      byteSize: Number(src.getSize()), verification: method,
      checksum: hashCopied
    });
  });
  return {files: details, sha256Count: verified};
}

function tc25bkLog_(result) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(TC25BK_LOG_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(TC25BK_LOG_SHEET_);
    sh.appendRow(TC25BK_LOG_HEADERS_);
    sh.setFrozenRows(1);
  } else {
    var found = sh.getRange(1, 1, 1, TC25BK_LOG_HEADERS_.length).getDisplayValues()[0];
    if (JSON.stringify(found) !== JSON.stringify(TC25BK_LOG_HEADERS_)) {
      throw new Error('TC25BK_LOG_SCHEMA_MISMATCH: Existing DR_Backup_Runs has different headers.');
    }
  }
  sh.appendRow([
    result.runId, result.startIso, result.endIso, result.status,
    result.snapshotFolderId || '', result.databaseBackupId || '',
    result.tabs || 0, result.master || 0, result.history || 0,
    result.expected || 0, result.copied || 0, result.sha256Count || 0,
    result.seconds || 0, String(result.error || '').slice(0, 350)
  ]);
}

function tc25bkNotify_(cfg, result) {
  var props = PropertiesService.getScriptProperties();
  if (result.status === 'PASS' &&
      String(props.getProperty('TC25BK_NOTIFY_SUCCESS') || 'NO').toUpperCase() !== 'YES') return;
  var subject = '[DGS DR Backup] ' + result.status + ' — ' + result.runId;
  var body = 'Production dashboard backup ' + result.status + '\n' +
    'Run ID: ' + result.runId + '\n' +
    'Started: ' + result.startIso + '\n' +
    'Completed: ' + result.endIso + '\n' +
    'Snapshot folder ID: ' + (result.snapshotFolderId || 'Not created') + '\n' +
    'Expected evidence: ' + result.expected + '; copied: ' + result.copied + '\n' +
    'Issue: ' + (result.error || 'None') + '\n' +
    'Check Apps Script Executions and restricted DR_Backup_Runs for details.\n';
  MailApp.sendEmail({to: cfg.alertEmail, subject: subject, body: body});
}

function tc25bkExecute_(kind) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('TC25BK_ALREADY_RUNNING: Another backup is in progress.');
  var started = new Date();
  var runId = 'DR-' + Utilities.formatDate(started, 'Asia/Kolkata', 'yyyyMMdd-HHmmss') +
    '-' + Utilities.getUuid().slice(0, 8);
  var result = {
    runId: runId, startIso: started.toISOString(), endIso: '',
    status: 'FAILED', expected: 0, copied: 0, sha256Count: 0,
    snapshotFolderId: '', databaseBackupId: '', tabs: 0, master: 0,
    history: 0, seconds: 0, error: ''
  };
  var cfg;
  try {
    cfg = tc25bkConfig_();
    if (!cfg.enabled) throw new Error('TC25BK_DISABLED: Set TC25BK_ENABLED=YES before backup runs.');
    var pre = tc25bkValidate_(cfg);
    result.expected = pre.files.length;

    // Snapshot folder and files are always newly created; sources remain intact.
    var snapshot = pre.backupFolder.createFolder(runId);
    result.snapshotFolderId = snapshot.getId();
    if (cfg.backupLocationType === 'MY_DRIVE') tc25bkMyDriveAcl_(snapshot.getId(), cfg);
    var dbCopy = DriveApp.getFileById(pre.spreadsheet.getId())
      .makeCopy('Compliance_Database_' + runId, snapshot);
    result.databaseBackupId = dbCopy.getId();
    if (cfg.backupLocationType === 'MY_DRIVE') tc25bkMyDriveAcl_(dbCopy.getId(), cfg);
    var copiedSpreadsheet = SpreadsheetApp.openById(dbCopy.getId());
    var verifiedDatabase = tc25bkCompareSheets_(pre.spreadsheet, copiedSpreadsheet);
    result.tabs = verifiedDatabase.tabs;
    result.master = verifiedDatabase.masterRecords;
    result.history = verifiedDatabase.historyEvents;

    var evidenceDest = snapshot.createFolder('Evidence_Files');
    if (cfg.backupLocationType === 'MY_DRIVE') tc25bkMyDriveAcl_(evidenceDest.getId(), cfg);
    var copies = tc25bkCopyEvidence_(pre.files, evidenceDest, cfg);
    result.copied = copies.files.length;
    result.sha256Count = copies.sha256Count;
    if (result.copied !== result.expected) {
      throw new Error('TC25BK_EVIDENCE_COUNT_MISMATCH.');
    }

    var manifest = {
      formatVersion: 1, status: 'PASS', runId: runId,
      createdAt: started.toISOString(), sourceSpreadsheetId: pre.spreadsheet.getId(),
      databaseBackupFileId: dbCopy.getId(),
      evidenceSourceFolderId: pre.sourceFolderId,
      evidenceBackupFolderId: evidenceDest.getId(),
      evidenceCount: result.copied, database: verifiedDatabase,
      evidence: copies.files, backupLocationType: pre.backupLocationType,
      sameSharedDrive: pre.backupLocationType === 'SHARED_DRIVE' &&
        pre.evidenceDriveId === pre.backupDriveId,
      notes: 'Apps Script source/credentials in Script Properties/triggers NOT part of snapshot.'
    };
    var json = JSON.stringify(manifest, null, 2);
    var manifestFile = snapshot.createFile('DR_Manifest_' + runId + '.json', json, MimeType.PLAIN_TEXT);
    if (cfg.backupLocationType === 'MY_DRIVE') tc25bkMyDriveAcl_(manifestFile.getId(), cfg);
    result.status = 'PASS';
  } catch (error) {
    result.error = String(error && error.message || error);
    Logger.log('TC25BK FAILED ' + runId + ': ' + result.error);
  } finally {
    result.endIso = new Date().toISOString();
    result.seconds = Math.round((Date.now() - started.getTime()) / 1000);
    try {
      PropertiesService.getScriptProperties().setProperty('TC25BK_LAST_STATUS', result.status);
      PropertiesService.getScriptProperties().setProperty('TC25BK_LAST_RUN_AT', result.endIso);
      tc25bkLog_(result);
      if (result.status === 'PASS') {
        PropertiesService.getScriptProperties().setProperty('TC25BK_LAST_SUCCESS_AT', result.endIso);
        PropertiesService.getScriptProperties().setProperty('TC25BK_LAST_SUCCESS_RUN', runId);
        PropertiesService.getScriptProperties().setProperty('TC25BK_LAST_SUCCESS_TARGET_ID', cfg.backupFolderId);
      }
    } catch (loggingError) {
      Logger.log('TC25BK_LOG_FAILURE: ' + String(loggingError));
      result.status = 'FAILED';
      result.error += ' | Logging failure: ' + String(loggingError);
    }
    try {
      // A configuration failure must still alert a previously pinned company
      // recipient, if the recipient property itself is valid.
      if (!cfg) {
        var recoveryEmail = String(PropertiesService.getScriptProperties()
          .getProperty('TC25BK_ALERT_EMAIL') || '').trim().toLowerCase();
        if (/^[a-z0-9._%+\-]+@(digiclarity\.in|clarityrcm\.com)$/.test(recoveryEmail)) {
          cfg = {alertEmail: recoveryEmail};
        }
      }
      if (cfg) tc25bkNotify_(cfg, result);
    } catch (mailError) {
      Logger.log('TC25BK_NOTIFICATION_FAILURE: ' + String(mailError));
    }
    lock.releaseLock();
  }
  Logger.log(JSON.stringify(result));
  if (result.status !== 'PASS') throw new Error('TC25BK_FAILED: ' + result.error + ' | Run ' + runId);
  return result;
}
