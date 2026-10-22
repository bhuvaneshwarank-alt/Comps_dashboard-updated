/**
 * TC-25-DR-13 — read-only integrity verification for restored evidence.
 * Production dashboard: add as a NEW .gs file; do not replace Code.gs or DR.gs.
 * This checks the 21-Oct-2026 19:23 IST scheduled backup snapshot.
 * No writes, permissions changes, mail, or triggers.
 */
function tc25drVerifyRestoredEvidence() {
  // Pinned to the verified production backup and user-provided restore folder.
  var RUN_ID = 'DR-20261021-192330-52b459fb';
  var SNAPSHOT_FOLDER_ID = '1TDj2n8-HxYrva18ROQkk6vseA9TRr3UC';
  var RESTORED_FOLDER_ID = '1wEaY1d8UU87q_Ww1ZiDJesREjqMZYfFo';
  var EXPECTED_FILES = 3;
  var MAX_HASH_BYTES = 20 * 1024 * 1024;

  // Use the same configured execution owner as the approved backup module.
  var props = PropertiesService.getScriptProperties();
  var expectedOwner = String(props.getProperty('TC25BK_OWNER_EMAIL') || '').trim().toLowerCase();
  var executor = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!expectedOwner || executor !== expectedOwner) {
    throw new Error('TC25_DR13_OWNER_DENIED: Use the configured backup execution owner.');
  }
  if (RESTORED_FOLDER_ID === String(props.getProperty('IAM_EVIDENCE_FOLDER_ID') || '') ||
      RESTORED_FOLDER_ID === String(props.getProperty('TC25BK_BACKUP_FOLDER_ID') || '') ||
      RESTORED_FOLDER_ID === SNAPSHOT_FOLDER_ID) {
    throw new Error('TC25_DR13_FOLDER_MISMATCH: Restored folder must be isolated.');
  }

  var snapshot = DriveApp.getFolderById(SNAPSHOT_FOLDER_ID);
  var restored = DriveApp.getFolderById(RESTORED_FOLDER_ID);
  if (snapshot.getName() !== RUN_ID || restored.getName() !== 'Evidence_Restored') {
    throw new Error('TC25_DR13_FOLDER_NAME_MISMATCH: Verify snapshot and restored folders.');
  }
  var restoredParent = restored.getParents();
  if (!restoredParent.hasNext() ||
      restoredParent.next().getName() !== 'TC25_DR_Restore_Test_21Oct2026') {
    throw new Error('TC25_DR13_RESTORE_PARENT_MISMATCH: Unexpected restored folder parent.');
  }

  var manifestFiles = snapshot.getFilesByName('DR_Manifest_' + RUN_ID + '.json');
  if (!manifestFiles.hasNext()) throw new Error('TC25_DR13_MANIFEST_MISSING');
  var manifestFile = manifestFiles.next();
  if (manifestFiles.hasNext()) throw new Error('TC25_DR13_MANIFEST_AMBIGUOUS');
  var manifest = JSON.parse(manifestFile.getBlob().getDataAsString('UTF-8'));
  if (manifest.runId !== RUN_ID || manifest.status !== 'PASS' ||
      Number(manifest.evidenceCount) !== EXPECTED_FILES ||
      !Array.isArray(manifest.evidence) || manifest.evidence.length !== EXPECTED_FILES ||
      !manifest.evidenceBackupFolderId) {
    throw new Error('TC25_DR13_MANIFEST_INVALID: Unexpected backup content.');
  }

  var backupEvidenceFolder = DriveApp.getFolderById(manifest.evidenceBackupFolderId);
  var backupFolderParents = backupEvidenceFolder.getParents();
  if (!backupFolderParents.hasNext() || backupFolderParents.next().getId() !== SNAPSHOT_FOLDER_ID) {
    throw new Error('TC25_DR13_BACKUP_FOLDER_MISMATCH');
  }

  function sha256_(file) {
    if (file.getSize() > MAX_HASH_BYTES) {
      throw new Error('TC25_DR13_FILE_TOO_LARGE: Manual validation required (>20MB).');
    }
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      file.getBlob().getBytes()).map(function(b) {
      return ('0' + (b & 255).toString(16)).slice(-2);
    }).join('');
  }
  function key_(hash, size) { return String(size) + ':' + hash.toLowerCase(); }

  // Multiset matching handles duplicate files and Google Drive's "Copy of" renaming.
  var expected = {};
  manifest.evidence.forEach(function(entry) {
    if (entry.verification !== 'SHA256' || !/^[a-f0-9]{64}$/i.test(String(entry.checksum || ''))) {
      throw new Error('TC25_DR13_MANIFEST_CHECKSUM_UNSUPPORTED');
    }
    var backupFile = DriveApp.getFileById(entry.copiedFileId);
    var parents = backupFile.getParents();
    if (!parents.hasNext() || parents.next().getId() !== manifest.evidenceBackupFolderId) {
      throw new Error('TC25_DR13_BACKUP_FILE_PARENT_MISMATCH');
    }
    var actualBackupHash = sha256_(backupFile);
    if (actualBackupHash.toLowerCase() !== String(entry.checksum).toLowerCase() ||
        Number(backupFile.getSize()) !== Number(entry.byteSize)) {
      throw new Error('TC25_DR13_BACKUP_FILE_INTEGRITY_FAIL');
    }
    var k = key_(entry.checksum, entry.byteSize);
    expected[k] = (expected[k] || 0) + 1;
  });

  var count = 0;
  var files = restored.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    count++;
    if (count > EXPECTED_FILES) throw new Error('TC25_DR13_EXTRA_RESTORED_FILES');
    var restoredHash = sha256_(file);
    var k = key_(restoredHash, file.getSize());
    if (!expected[k]) throw new Error('TC25_DR13_HASH_MISMATCH: Unexpected restored file ' + count);
    expected[k]--;
  }
  if (count !== EXPECTED_FILES || Object.keys(expected).some(function(k) { return expected[k] !== 0; })) {
    throw new Error('TC25_DR13_INCOMPLETE_RESTORE: Missing or mismatched evidence.');
  }

  var result = {test:'TC-25-DR-13', status:'PASS', runId:RUN_ID,
    manifestVerified:true, backupFilesSha256Verified:EXPECTED_FILES,
    restoredFilesSha256Verified:count, originalAndBackupFilesModified:false,
    note:'Read-only integrity check; no restore operation or application RTO measured.'};
  Logger.log('TC25_DR13: PASS — restored evidence SHA-256 ' + count + '/' + EXPECTED_FILES);
  Logger.log(JSON.stringify(result));
  return result;
}
