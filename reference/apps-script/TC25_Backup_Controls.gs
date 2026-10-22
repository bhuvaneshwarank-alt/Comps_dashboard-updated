
function tc25bkPreflight() {
  var owner = String(
    PropertiesService.getScriptProperties()
      .getProperty('TC25BK_OWNER_EMAIL') || ''
  ).trim().toLowerCase();

  var active = String(
    Session.getActiveUser().getEmail() || ''
  ).trim().toLowerCase();

  var effective = String(
    Session.getEffectiveUser().getEmail() || ''
  ).trim().toLowerCase();

  if (!owner || active !== owner || effective !== owner) {
    throw new Error('TC25BK_OWNER_DENIED');
  }

  return tc25bkPreflight_();
}

function tc25bkRunOnce() {
  var enabled = String(
    PropertiesService.getScriptProperties()
      .getProperty('TC25BK_ENABLED') || 'NO'
  ).trim().toUpperCase();

  if (enabled !== 'YES') {
    throw new Error(
      'TC25BK_DISABLED: Backup must be explicitly enabled before execution.'
    );
  }

  return tc25bkRunOnce_();
}

function tc25bkInstallSchedule() {
  var props = PropertiesService.getScriptProperties();

  var owner = String(
    props.getProperty('TC25BK_OWNER_EMAIL') || ''
  ).trim().toLowerCase();

  var active = String(
    Session.getActiveUser().getEmail() || ''
  ).trim().toLowerCase();

  var effective = String(
    Session.getEffectiveUser().getEmail() || ''
  ).trim().toLowerCase();

  if (!owner || active !== owner || effective !== owner) {
    throw new Error('TC25BK_OWNER_DENIED');
  }

  if (props.getProperty('TC25BK_LAST_STATUS') !== 'PASS') {
    throw new Error('TC25BK_MANUAL_BACKUP_REQUIRED');
  }

  return tc25bkInstallSchedule_();
}

function tc25bkLockDiagnostic() {
  var lock = LockService.getScriptLock();
  var acquired = false;

  try {
    acquired = lock.tryLock(5000);

    if (!acquired) {
      Logger.log('TC25_LOCK_TEST: BUSY');
      return;
    }

    Logger.log('TC25_LOCK_TEST: PASS');
    Logger.log('Script lock acquired successfully.');

  } catch (error) {
    Logger.log('TC25_LOCK_TEST: ERROR');
    Logger.log(String(error));
    throw error;

  } finally {
    if (acquired) {
      lock.releaseLock();
    }
  }
}

