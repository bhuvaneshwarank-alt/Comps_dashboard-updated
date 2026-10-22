/**
 * DIGICLARITY — TC-23.6 HIPAA incident manager alert, V63.
 * ADDITIVE SCRIPT MODULE. Code.gs V63 calls tc23NotifyHipaaIncidentCreated_()
 * only AFTER the master record and Record_History are written.
 *
 * This sends an INTERNAL notification, NOT a HIPAA breach determination,
 * legal notification, or clinical incident-management workflow.
 *
 * Script Properties (set from Project Settings, no secrets in code):
 *   TC23_INCIDENT_ALERT_ENABLED   = true
 *   TC23_INCIDENT_MANAGER_EMAIL   = designated manager's company email
 *   TC23_INCIDENT_TEST_MODE       = true  (keep true in prototype)
 *
 * Notification conditions:
 *   Record Type = Incident AND Primary Compliance Framework contains HIPAA.
 * Content is limited to record ID, UTC timestamp, coarse risk category,
 * and a dashboard landing-page link. NO titles/descriptions/PHI/tokens.
 */

const TC23_INCIDENT_ALERT_SHEET_ = 'Incident_Alert_Log';
const TC23_INCIDENT_ALERT_HEADERS_ = [
  'Alert ID', 'Record ID', 'Event Time UTC', 'Record Type', 'Framework',
  'Recipient Email', 'Status', 'Sent At UTC', 'Attempt Count',
  'Test Mode', 'Error Class'
];

function tc23IncidentAlertConfig_() {
  const props = PropertiesService.getScriptProperties();
  const rawEnabled = String(props.getProperty('TC23_INCIDENT_ALERT_ENABLED') || 'false').trim();
  const enabled = /^(true|yes|1)$/i.test(rawEnabled);
  const manager = String(props.getProperty('TC23_INCIDENT_MANAGER_EMAIL') || '').trim().toLowerCase();
  const testMode = !/^(false|no|0)$/i.test(
    String(props.getProperty('TC23_INCIDENT_TEST_MODE') || 'true').trim()
  );
  const validManager = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(manager);
  const owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  return {
    enabled: enabled, manager: manager, validManager: validManager,
    testMode: testMode, owner: owner,
    ready: enabled && validManager && !!owner && owner !== manager
  };
}

function tc23IncidentAlertLog_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(TC23_INCIDENT_ALERT_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(TC23_INCIDENT_ALERT_SHEET_);
    sheet.getRange(1, 1, 1, TC23_INCIDENT_ALERT_HEADERS_.length)
      .setValues([TC23_INCIDENT_ALERT_HEADERS_]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = sheet.getRange(1, 1, 1, TC23_INCIDENT_ALERT_HEADERS_.length).getDisplayValues()[0];
  if (TC23_INCIDENT_ALERT_HEADERS_.some(function(h, i) { return headers[i] !== h; })) {
    throw new Error('TC23_ALERT_LOG_SCHEMA_MISMATCH');
  }
  return sheet;
}

function tc23IncidentAlertStatusWrite_(sheet, row, status, sentAt, errorClass) {
  sheet.getRange(row, 7).setValue(status);
  sheet.getRange(row, 8).setValue(sentAt || '');
  sheet.getRange(row, 11).setValue(errorClass || '');
  SpreadsheetApp.flush();
}

function tc23IncidentAlertSafeRisk_(riskValue) {
  const value = String(riskValue || '').trim();
  return /^(Critical|High|Medium|Low)$/i.test(value)
    ? value.charAt(0).toUpperCase() + value.slice(1).toLowerCase()
    : 'Not assessed';
}

function tc23IncidentAlertEmail_(cfg, recordId, createdAt, riskCategory) {
  const url = APPROVAL_DASHBOARD_URL; // existing trusted, non-token dashboard landing page
  const escapedId = approvalEmailEscape_(recordId);
  const escapedDate = approvalEmailEscape_(createdAt);
  const escapedRisk = approvalEmailEscape_(riskCategory);
  const escapedUrl = approvalEmailEscape_(url);
  const subject = (cfg.testMode ? '[TEST] ' : '') +
    '[Internal Review] HIPAA-related incident created | ' + recordId;
  const body =
    'DIGICLARITY COMPLIANCE DASHBOARD\n\n' +
    (cfg.testMode ? 'PROTOTYPE TEST — NO REAL INCIDENT HAS BEEN REPORTED.\n\n' : '') +
    'An incident record marked with the HIPAA framework requires internal review.\n\n' +
    'Record ID: ' + recordId + '\n' +
    'Created (UTC): ' + createdAt + '\n' +
    'Risk category: ' + riskCategory + '\n\n' +
    'Sign in to the Compliance Dashboard and review the record:\n' + url + '\n\n' +
    'No patient details or record narrative are included in this notification.\n' +
    'This internal alert does not establish whether a reportable HIPAA breach occurred.\n' +
    'Follow the company incident response/escalation procedure; do not send PHI in replies.';
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:auto;color:#172742">' +
    '<div style="background:#142440;color:#fff;padding:24px;border-radius:12px 12px 0 0">' +
    '<div style="font-size:23px;font-weight:700;letter-spacing:1px">DIGICLARITY</div>' +
    '<div style="font-size:12px;margin-top:7px;letter-spacing:1px">COMPLIANCE DASHBOARD</div></div>' +
    '<div style="padding:24px;border:1px solid #e3e9f1;border-radius:0 0 12px 12px">' +
    (cfg.testMode ? '<p style="color:#9b2260;font-weight:700">PROTOTYPE TEST — NO REAL INCIDENT</p>' : '') +
    '<h2 style="font-size:21px">HIPAA-related incident requires review</h2>' +
    '<p>An incident record was created. Please assess it inside the authenticated dashboard.</p>' +
    '<div style="background:#f3f6fa;border-left:4px solid #2469ad;padding:18px">' +
    '<div><strong>Record ID:</strong> ' + escapedId + '</div>' +
    '<div><strong>Created (UTC):</strong> ' + escapedDate + '</div>' +
    '<div><strong>Risk category:</strong> ' + escapedRisk + '</div></div>' +
    '<p style="margin:26px 0"><a href="' + escapedUrl + '" ' +
    'style="background:#2469ad;color:#fff;text-decoration:none;padding:13px 20px;border-radius:6px">' +
    'REVIEW IN DASHBOARD</a></p>' +
    '<p style="font-size:12px;color:#5c6b82">Patient names, descriptions and PHI are intentionally excluded. ' +
    'An internal alert is not a determination of a reportable HIPAA breach. ' +
    'Use approved response and escalation procedures.</p></div></div>';
  return {to:cfg.manager, subject:subject, body:body, htmlBody:html,
    name:'DigiClarity Compliance Dashboard'};
}

function tc23NotifyHipaaIncidentCreated_(context) {
  context = context || {};
  const type = String(context.recordType || '').trim();
  const framework = String(context.framework || '').trim();
  if (type.toLowerCase() !== 'incident' || !/\bHIPAA\b/i.test(framework)) {
    return {applicable:false, status:'NOT_APPLICABLE', sent:false};
  }
  const recordId = String(context.recordId || '').trim();
  // Record ID is generated by the trusted backend, never a free-text title.
  if (!/^[A-Za-z0-9-]{4,50}$/.test(recordId)) throw new Error('TC23_INVALID_RECORD_ID');
  const eventTime = String(context.createdAt || new Date().toISOString());
  const cfg = tc23IncidentAlertConfig_();
  const risk = tc23IncidentAlertSafeRisk_(context.risk);
  const sheet = tc23IncidentAlertLog_();
  const initialStatus = !cfg.enabled ? 'DISABLED' : (!cfg.ready ? 'CONFIG_REQUIRED' : 'PENDING');
  const initialRow = [
    'IA-' + Utilities.getUuid(), recordId, eventTime, 'Incident', 'HIPAA',
    cfg.validManager ? cfg.manager : '', initialStatus, '',
    initialStatus === 'PENDING' ? 1 : 0, cfg.testMode ? 'Yes' : 'No', ''
  ];
  sheet.appendRow(initialRow);
  SpreadsheetApp.flush();
  const rowNum = sheet.getLastRow();
  if (initialStatus !== 'PENDING') {
    return {applicable:true, status:initialStatus, sent:false};
  }
  try {
    MailApp.sendEmail(tc23IncidentAlertEmail_(cfg, recordId, eventTime, risk));
  } catch (error) {
    console.error('TC23 incident alert send failed for ' + recordId + ': ' + String(error));
    try { tc23IncidentAlertStatusWrite_(sheet, rowNum, 'FAILED', '', 'MAIL_SEND_ERROR'); }
    catch (auditError) { console.error('TC23 alert failure-log write failed: ' + String(auditError)); }
    return {applicable:true, status:'FAILED', sent:false};
  }
  try {
    tc23IncidentAlertStatusWrite_(sheet, rowNum, 'SENT', new Date().toISOString(), '');
    return {applicable:true, status:'SENT', sent:true, testMode:cfg.testMode};
  } catch (error) {
    // MailApp returned successfully; do NOT retry automatically: that may duplicate an alert.
    console.error('TC23 incident alert sent but status write failed for ' + recordId + ': ' + String(error));
    return {applicable:true, status:'SENT_LOG_PENDING', sent:true, testMode:cfg.testMode};
  }
}

// Private configuration validation. Temporary editor-only wrapper can call it;
// do NOT leave a public unauthenticated testing endpoint in production.
function tc23IncidentAlertConfigurationCheck_() {
  const cfg = tc23IncidentAlertConfig_();
  return {enabled:cfg.enabled, testMode:cfg.testMode,
    managerConfigured:cfg.validManager, separateFromOwner:!!cfg.owner && cfg.owner !== cfg.manager,
    ready:cfg.ready};
}
