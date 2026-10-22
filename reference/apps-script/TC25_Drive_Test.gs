
function tc25TestSharedDriveAccess() {
  const props = PropertiesService.getScriptProperties();

  const folders = [
    {
      key: 'TC25_TEST_LIVE_FOLDER_ID',
      expected: '01_Live_Evidence'
    },
    {
      key: 'TC25_TEST_BACKUP_FOLDER_ID',
      expected: '02_Evidence_Backup'
    }
  ];

  folders.forEach(function(item) {
    const id = props.getProperty(item.key);

    if (!id || !/^[a-zA-Z0-9_-]{10,}$/.test(id)) {
      throw new Error(
        item.key + ': Invalid or missing folder ID'
      );
    }

    const folder = DriveApp.getFolderById(id);
    const name = folder.getName();

    if (name !== item.expected) {
      throw new Error(
        item.key + ': Unexpected folder name: ' + name
      );
    }

    Logger.log(item.expected + ': ACCESS PASS');
  });

  Logger.log('TC-25: BOTH FOLDERS ACCESSIBLE');
}

function tc25TestSharedDriveWrite() {
  const props = PropertiesService.getScriptProperties();

  const folders = [
    {
      key: 'TC25_TEST_LIVE_FOLDER_ID',
      name: '01_Live_Evidence'
    },
    {
      key: 'TC25_TEST_BACKUP_FOLDER_ID',
      name: '02_Evidence_Backup'
    }
  ];

  folders.forEach(function(item) {
    const id = props.getProperty(item.key);
    const folder = DriveApp.getFolderById(id);

    if (folder.getName() !== item.name) {
      throw new Error('Unexpected folder: ' + item.name);
    }

    const content = 'DGS TC-25 Shared Drive test only';
    const filename = 'TC25_TEST_' +
      new Date().getTime() + '.txt';

    const blob = Utilities.newBlob(
      content,
      'text/plain',
      filename
    );

    const file = folder.createFile(blob);

    if (file.getBlob().getDataAsString() !== content) {
      throw new Error(item.name + ': File verification failed');
    }

    Logger.log(item.name + ': WRITE PASS');
  });

  Logger.log('TC-25: BOTH FOLDERS WRITE VERIFIED');
}

function tc25VerifySharedDriveIdentity() {
  const props = PropertiesService.getScriptProperties();

  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');
  const folderId = props.getProperty('TC25_TEST_LIVE_FOLDER_ID');

  if (!driveId || !folderId) {
    throw new Error('Missing Shared Drive configuration.');
  }

  const folder = Drive.Files.get(folderId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,trashed'
  });

  if (
    folder.trashed ||
    folder.name !== '01_Live_Evidence' ||
    folder.mimeType !== 'application/vnd.google-apps.folder' ||
    folder.driveId !== driveId
  ) {
    throw new Error('TC25: Shared Drive identity mismatch.');
  }

  Logger.log('TC25: SHARED DRIVE IDENTITY PASS');
}

function tc25StagingStorageCheck() {
  const props = PropertiesService.getScriptProperties();

  // Ensure this is the staging spreadsheet.
  const ss = getSpreadsheet_();
  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  // Verify the selected storage mode.
  if (iamEvidenceStorageMode_() !== 'SHARED_DRIVE') {
    throw new Error('TC25: Shared Drive mode not enabled.');
  }

  // Verify staging folder is configured.
  const liveId = props.getProperty(
    'TC25_TEST_LIVE_FOLDER_ID'
  );

  if (!liveId || props.getProperty(
    'IAM_EVIDENCE_FOLDER_ID'
  ) !== liveId) {
    throw new Error('TC25: Evidence folder mismatch.');
  }

  // Run V65's existing folder and owner checks.
  const folder = iamPrivateFolder_();

  Logger.log(
    'TC25: STAGING SHARED DRIVE COMPATIBILITY PASS'
  );
  Logger.log(
    'Verified folder: ' + folder.getName()
  );
}
function tc25CheckFolderProperties() {
  const props = PropertiesService.getScriptProperties();

  const live = String(
    props.getProperty('TC25_TEST_LIVE_FOLDER_ID') || ''
  ).trim();

  const evidence = String(
    props.getProperty('IAM_EVIDENCE_FOLDER_ID') || ''
  ).trim();

  Logger.log('Live folder configured: ' + Boolean(live));
  Logger.log('Evidence folder configured: ' + Boolean(evidence));
  Logger.log('Folder IDs match: ' + (live === evidence && !!live));
}

function tc25CheckStagingOwner() {
  const props = PropertiesService.getScriptProperties();

  const configured = String(
    props.getProperty('IAM_STORAGE_OWNER') || ''
  ).trim().toLowerCase();

  const effective = String(
    Session.getEffectiveUser().getEmail() || ''
  ).trim().toLowerCase();

  Logger.log('Owner configured: ' + Boolean(configured));
  Logger.log('Executing account available: ' + Boolean(effective));
  Logger.log(
    'Owner matches executing account: ' +
    Boolean(configured && effective && configured === effective)
  );
}

function tc25TestEvidenceBackup() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging project.');
  }

  const props = PropertiesService.getScriptProperties();

  const liveId = props.getProperty('TC25_TEST_LIVE_FOLDER_ID');
  const backupId = props.getProperty('TC25_TEST_BACKUP_FOLDER_ID');

  if (!liveId || !backupId || liveId === backupId) {
    throw new Error('TC25: Invalid backup folder configuration.');
  }

  const liveFolder = DriveApp.getFolderById(liveId);
  const backupFolder = DriveApp.getFolderById(backupId);

  if (liveFolder.getName() !== '01_Live_Evidence' ||
      backupFolder.getName() !== '02_Evidence_Backup') {
    throw new Error('TC25: Unexpected folder names.');
  }

  // Select only the harmless TC25 text test file.
  const files = liveFolder.getFiles();
  let source = null;

  while (files.hasNext()) {
    const file = files.next();

    if (/^TC25_TEST_.*\.txt$/i.test(file.getName()) &&
        file.getSize() <= 1024 * 1024) {
      source = file;
      break;
    }
  }

  if (!source) {
    throw new Error('TC25: Dummy TXT file not found.');
  }

  function sha256(blob) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      blob.getBytes()
    ).map(function(byte) {
      return ('0' + (byte & 255).toString(16)).slice(-2);
    }).join('');
  }

  const sourceHash = sha256(source.getBlob());

  const backupName =
    'TC25_BACKUP_' +
    Utilities.formatDate(
      new Date(),
      'Asia/Kolkata',
      'yyyyMMdd_HHmmss'
    ) +
    '_' + Utilities.getUuid().slice(0, 8) + '.txt';

  const backup = source.makeCopy(
    backupName,
    backupFolder
  );

  const backupHash = sha256(backup.getBlob());

  if (sourceHash !== backupHash ||
      source.getId() === backup.getId()) {
    throw new Error('TC25: Backup integrity verification FAILED.');
  }

  Logger.log('TC25: EVIDENCE BACKUP INTEGRITY PASS');
  Logger.log('SHA256 match: true');
  Logger.log('Backup created in: ' + backupFolder.getName());
  Logger.log('Original file preserved: ' + !source.isTrashed());
}

function tc25TestEvidenceRecovery() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();
  const liveId = props.getProperty('TC25_TEST_LIVE_FOLDER_ID');
  const backupId = props.getProperty('TC25_TEST_BACKUP_FOLDER_ID');

  if (!liveId || !backupId || liveId === backupId) {
    throw new Error('TC25: Invalid folder configuration.');
  }

  const live = DriveApp.getFolderById(liveId);
  const backup = DriveApp.getFolderById(backupId);

  if (live.getName() !== '01_Live_Evidence' ||
      backup.getName() !== '02_Evidence_Backup') {
    throw new Error('TC25: Unexpected folder names.');
  }

  // Select the most recent TC25 dummy backup.
  const files = backup.getFiles();
  let source = null;

  while (files.hasNext()) {
    const file = files.next();

    if (/^TC25_BACKUP_.*\.txt$/i.test(file.getName()) &&
        file.getSize() <= 1024 * 1024 &&
        (!source ||
         file.getDateCreated() > source.getDateCreated())) {
      source = file;
    }
  }

  if (!source) {
    throw new Error('TC25: Test backup not found.');
  }

  function sha256(file) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      file.getBlob().getBytes()
    ).map(function(byte) {
      return ('0' + (byte & 255).toString(16)).slice(-2);
    }).join('');
  }

  const originalHash = sha256(source);

  // Create a separate recovery test copy.
  // Never overwrite an existing evidence file.
  const restored = source.makeCopy(
    'TC25_RESTORE_TEST_' + Utilities.getUuid() + '.txt',
    live
  );

  const restoredHash = sha256(restored);

  if (originalHash !== restoredHash) {
    throw new Error('TC25: Recovery integrity FAILED.');
  }

  Logger.log('TC25: EVIDENCE RECOVERY PASS');
  Logger.log('SHA256 match: true');
  Logger.log('Recovered into: ' + live.getName());
  Logger.log('Backup preserved: ' + !source.isTrashed());
}

function tc25EvidenceInventoryCheck() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const sheet = ss.getSheetByName('Evidence_Register');

  if (!sheet) {
    throw new Error('TC25: Evidence_Register not found.');
  }

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0] || [];

  const evidenceCol = headers.indexOf('Evidence ID');
  const recordCol = headers.indexOf('Record ID');
  const fileCol = headers.indexOf('File ID');

  if ([evidenceCol, recordCol, fileCol].includes(-1)) {
    throw new Error('TC25: Required evidence columns missing.');
  }

  let total = 0;
  let complete = 0;
  let incomplete = 0;
  const fileIds = new Set();
  let duplicateReferences = 0;

  data.slice(1).forEach(function(row) {
    if (!row[evidenceCol]) return;
    total++;

    if (row[recordCol] && row[fileCol]) {
      complete++;
    } else {
      incomplete++;
    }

    if (row[fileCol]) {
      if (fileIds.has(row[fileCol])) {
        duplicateReferences++;
      }
      fileIds.add(row[fileCol]);
    }
  });

  Logger.log('TC25: EVIDENCE INVENTORY COMPLETED');
  Logger.log('Total evidence records: ' + total);
  Logger.log('Complete references: ' + complete);
  Logger.log('Incomplete references: ' + incomplete);
  Logger.log('Duplicate file references: ' + duplicateReferences);
  Logger.log('Unique file IDs: ' + fileIds.size);
  Logger.log('No files moved or modified.');
}

function tc25VerifyEvidenceFiles() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();
  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  if (!driveId) {
    throw new Error('TC25: Shared Drive ID missing.');
  }

  const sheet = ss.getSheetByName('Evidence_Register');
  if (!sheet) {
    throw new Error('TC25: Evidence_Register missing.');
  }

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0] || [];
  const idColumn = headers.indexOf('File ID');

  if (idColumn === -1) {
    throw new Error('TC25: File ID column missing.');
  }

  let checked = 0;
  let accessible = 0;
  let sharedDrive = 0;
  let otherDrive = 0;
  let unavailable = 0;

  data.slice(1).forEach(function(row) {
    const fileId = String(row[idColumn] || '').trim();
    if (!fileId) return;

    checked++;

    try {
      const file = Drive.Files.get(fileId, {
        supportsAllDrives: true,
        fields: 'id,driveId,trashed'
      });

      if (!file || file.trashed) {
        unavailable++;
        return;
      }

      accessible++;

      if (file.driveId === driveId) {
        sharedDrive++;
      } else {
        otherDrive++;
      }

    } catch (error) {
      unavailable++;
    }
  });

  Logger.log('TC25: PHYSICAL EVIDENCE VERIFICATION');
  Logger.log('Files checked: ' + checked);
  Logger.log('Accessible: ' + accessible);
  Logger.log('In target Shared Drive: ' + sharedDrive);
  Logger.log('Outside target Shared Drive: ' + otherDrive);
  Logger.log('Unavailable: ' + unavailable);

  Logger.log(
    unavailable === 0
      ? 'TC25: FILE ACCESSIBILITY PASS'
      : 'TC25: FILE ACCESSIBILITY REVIEW REQUIRED'
  );
}

function tc25PrepareMigrationDummy() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('TC25_MIGRATION_DUMMY_FILE_ID')) {
    throw new Error(
      'TC25: Migration dummy already configured.'
    );
  }

  const file = DriveApp.createFile(
    'TC25_MIGRATION_SOURCE_' +
      Utilities.getUuid().slice(0, 8) + '.txt',
    'TC25 dummy migration test. No real evidence.',
    MimeType.PLAIN_TEXT
  );

  props.setProperty(
    'TC25_MIGRATION_DUMMY_FILE_ID',
    file.getId()
  );

  Logger.log('TC25: MIGRATION DUMMY READY');
  Logger.log('Source location: My Drive');
  Logger.log('Dummy file created: true');
  Logger.log('Existing evidence modified: false');
}

function tc25TestDummyMigration() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  const fileId = props.getProperty(
    'TC25_MIGRATION_DUMMY_FILE_ID'
  );
  const folderId = props.getProperty(
    'TC25_TEST_LIVE_FOLDER_ID'
  );
  const driveId = props.getProperty(
    'IAM_SHARED_DRIVE_ID'
  );

  if (!fileId || !folderId || !driveId) {
    throw new Error('TC25: Missing migration configuration.');
  }

  const source = DriveApp.getFileById(fileId);
  const target = DriveApp.getFolderById(folderId);

  const before = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,parents,trashed'
  });

  if (
    before.trashed ||
    before.driveId ||
    !/^TC25_MIGRATION_SOURCE_.*\.txt$/.test(before.name) ||
    before.mimeType !== 'text/plain' ||
    source.getSize() > 4096
  ) {
    throw new Error('TC25: Unsafe or already migrated source.');
  }

  const targetInfo = Drive.Files.get(folderId, {
    supportsAllDrives: true,
    fields: 'id,name,driveId,trashed'
  });

  if (
    targetInfo.trashed ||
    targetInfo.name !== '01_Live_Evidence' ||
    targetInfo.driveId !== driveId
  ) {
    throw new Error('TC25: Target Shared Drive mismatch.');
  }

  function hash(file) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      file.getBlob().getBytes()
    ).map(function(b) {
      return ('0' + (b & 255).toString(16)).slice(-2);
    }).join('');
  }

  const originalHash = hash(source);

  // Preserve a safety copy in My Drive before moving.
  const safetyCopy = source.makeCopy(
    'TC25_MIGRATION_SAFETY_' +
    Utilities.getUuid() + '.txt'
  );

  props.setProperty(
    'TC25_MIGRATION_SAFETY_COPY_ID',
    safetyCopy.getId()
  );

  if (hash(safetyCopy) !== originalHash) {
    throw new Error('TC25: Safety copy integrity failed.');
  }

  // Move only the disposable dummy.
  source.moveTo(target);

  const after = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,driveId,parents,trashed'
  });

  const verified =
    !after.trashed &&
    after.id === fileId &&
    after.driveId === driveId &&
    (after.parents || []).includes(folderId) &&
    hash(DriveApp.getFileById(fileId)) === originalHash;

  if (!verified) {
    throw new Error(
      'TC25: Migration verification failed. Inspect dummy manually.'
    );
  }

  Logger.log('TC25: DUMMY MIGRATION PASS');
  Logger.log('Original file ID preserved: true');
  Logger.log('SHA256 match: true');
  Logger.log('Destination: 01_Live_Evidence');
  Logger.log('Safety copy retained in My Drive: true');
}

function tc25TestDummyRollback() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  const fileId = props.getProperty(
    'TC25_MIGRATION_DUMMY_FILE_ID'
  );
  const safetyId = props.getProperty(
    'TC25_MIGRATION_SAFETY_COPY_ID'
  );
  const folderId = props.getProperty(
    'TC25_TEST_LIVE_FOLDER_ID'
  );
  const driveId = props.getProperty(
    'IAM_SHARED_DRIVE_ID'
  );

  if (!fileId || !safetyId || !folderId || !driveId) {
    throw new Error('TC25: Missing rollback configuration.');
  }

  const file = DriveApp.getFileById(fileId);
  const safety = DriveApp.getFileById(safetyId);

  const before = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,parents,trashed'
  });

  if (
    before.trashed ||
    before.driveId !== driveId ||
    !(before.parents || []).includes(folderId) ||
    !/^TC25_MIGRATION_SOURCE_.*\.txt$/.test(before.name) ||
    before.mimeType !== 'text/plain' ||
    file.getSize() > 4096 ||
    fileId === safetyId
  ) {
    throw new Error('TC25: Unsafe rollback target.');
  }

  function sha256(f) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      f.getBlob().getBytes()
    ).map(function(byte) {
      return ('0' + (byte & 255).toString(16)).slice(-2);
    }).join('');
  }

  const expectedHash = sha256(safety);

  if (sha256(file) !== expectedHash) {
    throw new Error('TC25: Safety copy hash mismatch.');
  }

  // Move only our disposable test file.
  const myDrive = DriveApp.getRootFolder();

  file.moveTo(myDrive);

  const after = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,driveId,parents,trashed'
  });

  if (
    after.trashed ||
    after.id !== fileId ||
    after.driveId ||
    !(after.parents || []).includes(myDrive.getId()) ||
    sha256(DriveApp.getFileById(fileId)) !== expectedHash
  ) {
    throw new Error(
      'TC25: Rollback verification failed. Inspect dummy manually.'
    );
  }

  Logger.log('TC25: DUMMY ROLLBACK PASS');
  Logger.log('Restored to My Drive: true');
  Logger.log('Original file ID preserved: true');
  Logger.log('SHA256 match: true');
  Logger.log('Safety copy retained: ' + !safety.isTrashed());
}

function tc25VerifyBackupFolderIdentity() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();
  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  const folders = [
    {
      key: 'TC25_TEST_DB_BACKUP_FOLDER_ID',
      name: '03_Database_Backup'
    },
    {
      key: 'TC25_TEST_SOURCE_BACKUP_FOLDER_ID',
      name: '04_Source_Code_Backup'
    }
  ];

  if (!driveId) {
    throw new Error('TC25: Shared Drive ID missing.');
  }

  const seen = new Set();

  folders.forEach(function(item) {
    const id = props.getProperty(item.key);

    if (!id || seen.has(id)) {
      throw new Error('TC25: Missing or duplicate folder ID.');
    }

    seen.add(id);

    const folder = Drive.Files.get(id, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,driveId,trashed'
    });

    if (
      folder.trashed ||
      folder.name !== item.name ||
      folder.driveId !== driveId ||
      folder.mimeType !==
        'application/vnd.google-apps.folder'
    ) {
      throw new Error(
        'TC25: Invalid backup folder: ' + item.name
      );
    }

    Logger.log('TC25: ' + item.name + ' VERIFIED');
  });

  Logger.log('TC25: BACKUP FOLDER IDENTITY PASS');
  Logger.log('No files created or modified.');
}

function tc25CreateDummyDatabaseBackup() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('TC25_DUMMY_DB_BACKUP_ID')) {
    throw new Error(
      'TC25: Dummy database backup already registered.'
    );
  }

  const folderId = props.getProperty(
    'TC25_TEST_DB_BACKUP_FOLDER_ID'
  );

  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  if (!folderId || !driveId) {
    throw new Error('TC25: Missing backup configuration.');
  }

  const info = Drive.Files.get(folderId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,trashed'
  });

  if (
    info.trashed ||
    info.name !== '03_Database_Backup' ||
    info.driveId !== driveId ||
    info.mimeType !== 'application/vnd.google-apps.folder'
  ) {
    throw new Error('TC25: Backup destination mismatch.');
  }

  const dummyData = [
    ['Record ID', 'Department', 'Status'],
    ['TC25-TEST-001', 'IT', 'Open'],
    ['TC25-TEST-002', 'HR', 'Closed']
  ];

  const backupName =
    'TC25_DUMMY_DB_BACKUP_' +
    Utilities.formatDate(
      new Date(),
      'Asia/Kolkata',
      'yyyyMMdd_HHmmss'
    );

  // Create a synthetic database: NO real dashboard data.
  const backupSS = SpreadsheetApp.create(backupName);
  const backupId = backupSS.getId();

  // Save the reference before any further operations.
  props.setProperty(
    'TC25_DUMMY_DB_BACKUP_ID',
    backupId
  );

  const sheet = backupSS.getSheets()[0];
  sheet.setName('TC25_Test_Data');

  sheet.getRange(
    1, 1,
    dummyData.length,
    dummyData[0].length
  ).setValues(dummyData);

  SpreadsheetApp.flush();

  // Move the synthetic backup into the approved folder.
  const backupFile = DriveApp.getFileById(backupId);
  backupFile.moveTo(DriveApp.getFolderById(folderId));

  const fileInfo = Drive.Files.get(backupId, {
    supportsAllDrives: true,
    fields: 'id,mimeType,parents,driveId,trashed'
  });

  if (
    fileInfo.trashed ||
    fileInfo.driveId !== driveId ||
    !(fileInfo.parents || []).includes(folderId)
  ) {
    throw new Error(
      'TC25: Backup destination verification failed.'
    );
  }

  const readback = SpreadsheetApp
    .openById(backupId)
    .getSheetByName('TC25_Test_Data')
    .getDataRange()
    .getValues();

  if (
    JSON.stringify(readback) !==
    JSON.stringify(dummyData)
  ) {
    throw new Error(
      'TC25: Database backup data verification failed.'
    );
  }

  Logger.log('TC25: DATABASE BACKUP CREATION PASS');
  Logger.log('Stored in: 03_Database_Backup');
  Logger.log('Dummy records verified: 2');
  Logger.log('Data integrity match: true');
  Logger.log('Production database modified: false');
}

function tc25TestDummyDatabaseRestore() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('TC25_DUMMY_DB_RESTORE_ID')) {
    throw new Error('TC25: Restore test already registered.');
  }

  const backupId = props.getProperty('TC25_DUMMY_DB_BACKUP_ID');
  const folderId = props.getProperty('TC25_TEST_DB_BACKUP_FOLDER_ID');
  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  if (!backupId || !folderId || !driveId) {
    throw new Error('TC25: Missing backup configuration.');
  }

  const folder = DriveApp.getFolderById(folderId);
  const source = DriveApp.getFileById(backupId);

  const info = Drive.Files.get(backupId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,parents,trashed'
  });

  if (
    info.trashed ||
    info.driveId !== driveId ||
    !(info.parents || []).includes(folderId) ||
    !/^TC25_DUMMY_DB_BACKUP_/.test(info.name) ||
    info.mimeType !==
      'application/vnd.google-apps.spreadsheet'
  ) {
    throw new Error('TC25: Unexpected backup source.');
  }

  const expected = [
    ['Record ID', 'Department', 'Status'],
    ['TC25-TEST-001', 'IT', 'Open'],
    ['TC25-TEST-002', 'HR', 'Closed']
  ];

  // Restore to a separate file, never overwrite staging.
  const restored = source.makeCopy(
    'TC25_DUMMY_DB_RESTORED_' + Utilities.getUuid(),
    folder
  );

  props.setProperty('TC25_DUMMY_DB_RESTORE_ID', restored.getId());

  const restoreInfo = Drive.Files.get(restored.getId(), {
    supportsAllDrives: true,
    fields: 'id,driveId,parents,trashed'
  });

  if (
    restoreInfo.trashed ||
    restoreInfo.id === backupId ||
    restoreInfo.driveId !== driveId ||
    !(restoreInfo.parents || []).includes(folderId)
  ) {
    throw new Error('TC25: Restored file location invalid.');
  }

  const restoredSS = SpreadsheetApp.openById(restored.getId());
  const restoredSheet = restoredSS.getSheetByName('TC25_Test_Data');

  if (!restoredSheet) {
    throw new Error('TC25: Restored sheet missing.');
  }

  const actual = restoredSheet.getDataRange().getDisplayValues();

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error('TC25: Restored data integrity FAILED.');
  }

  Logger.log('TC25: DATABASE RESTORATION PASS');
  Logger.log('Restored records: 2');
  Logger.log('Data integrity match: true');
  Logger.log('Original backup preserved: ' + !source.isTrashed());
  Logger.log('Production database modified: false');
}

function tc25TestSourceCodeBackup() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('TC25_SOURCE_BACKUP_FILE_ID')) {
    throw new Error('TC25: Source backup already registered.');
  }

  const folderId = props.getProperty(
    'TC25_TEST_SOURCE_BACKUP_FOLDER_ID'
  );
  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  if (!folderId || !driveId) {
    throw new Error('TC25: Missing backup configuration.');
  }

  const folderInfo = Drive.Files.get(folderId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,trashed'
  });

  if (
    folderInfo.trashed ||
    folderInfo.name !== '04_Source_Code_Backup' ||
    folderInfo.driveId !== driveId ||
    folderInfo.mimeType !==
      'application/vnd.google-apps.folder'
  ) {
    throw new Error('TC25: Source backup folder mismatch.');
  }

  const dummySource = [
    '// TC25 synthetic Apps Script source',
    'function tc25DummySource() {',
    '  return "TC25_SOURCE_TEST_OK";',
    '}'
  ].join('\n');

  function sha256(value) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      Utilities.newBlob(value).getBytes()
    ).map(function(byte) {
      return ('0' + (byte & 255).toString(16)).slice(-2);
    }).join('');
  }

  const folder = DriveApp.getFolderById(folderId);

  const backup = folder.createFile(
    'TC25_DUMMY_SOURCE_BACKUP_' +
      Utilities.getUuid() + '.gs',
    dummySource,
    MimeType.PLAIN_TEXT
  );

  props.setProperty(
    'TC25_SOURCE_BACKUP_FILE_ID',
    backup.getId()
  );

  const fileInfo = Drive.Files.get(backup.getId(), {
    supportsAllDrives: true,
    fields: 'id,driveId,parents,trashed'
  });

  const readback = backup.getBlob().getDataAsString();

  if (
    fileInfo.trashed ||
    fileInfo.driveId !== driveId ||
    !(fileInfo.parents || []).includes(folderId) ||
    sha256(dummySource) !== sha256(readback)
  ) {
    throw new Error('TC25: Source backup verification failed.');
  }

  Logger.log('TC25: SOURCE CODE BACKUP PASS');
  Logger.log('Stored in: 04_Source_Code_Backup');
  Logger.log('SHA256 match: true');
  Logger.log('Actual dashboard source modified: false');
}

function tc25TestSourceCodeRecovery() {
  const ss = getSpreadsheet_();

  if (ss.getName() !== 'DGS_Compliance_TC25_Staging') {
    throw new Error('TC25: Not the staging spreadsheet.');
  }

  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('TC25_SOURCE_RESTORE_FILE_ID')) {
    throw new Error('TC25: Recovery already registered.');
  }

  const sourceId = props.getProperty(
    'TC25_SOURCE_BACKUP_FILE_ID'
  );
  const folderId = props.getProperty(
    'TC25_TEST_SOURCE_BACKUP_FOLDER_ID'
  );
  const driveId = props.getProperty('IAM_SHARED_DRIVE_ID');

  if (!sourceId || !folderId || !driveId) {
    throw new Error('TC25: Missing recovery configuration.');
  }

  const sourceInfo = Drive.Files.get(sourceId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType,driveId,parents,trashed'
  });

  if (
    sourceInfo.trashed ||
    sourceInfo.driveId !== driveId ||
    !(sourceInfo.parents || []).includes(folderId) ||
    !/^TC25_DUMMY_SOURCE_BACKUP_.*\.gs$/.test(sourceInfo.name)
  ) {
    throw new Error('TC25: Invalid source backup.');
  }

  const source = DriveApp.getFileById(sourceId);
  const folder = DriveApp.getFolderById(folderId);

  function sha256(file) {
    return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      file.getBlob().getBytes()
    ).map(function(byte) {
      return ('0' + (byte & 255).toString(16)).slice(-2);
    }).join('');
  }

  const originalHash = sha256(source);

  // Restore as a separate file, without overwriting code.
  const restored = source.makeCopy(
    'TC25_SOURCE_RESTORED_' +
      Utilities.getUuid() + '.gs',
    folder
  );

  props.setProperty(
    'TC25_SOURCE_RESTORE_FILE_ID',
    restored.getId()
  );

  const restoredInfo = Drive.Files.get(restored.getId(), {
    supportsAllDrives: true,
    fields: 'id,driveId,parents,trashed'
  });

  if (
    restoredInfo.trashed ||
    restoredInfo.driveId !== driveId ||
    !(restoredInfo.parents || []).includes(folderId) ||
    restored.getId() === sourceId ||
    sha256(restored) !== originalHash
  ) {
    throw new Error('TC25: Source recovery verification FAILED.');
  }

  Logger.log('TC25: SOURCE CODE RECOVERY PASS');
  Logger.log('SHA256 match: true');
  Logger.log('Recovered source file: verified');
  Logger.log('Original backup preserved: ' + !source.isTrashed());
  Logger.log('Dashboard code modified: false');
}