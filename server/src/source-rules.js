// Pure rules ported verbatim from the uploaded Code.gs. Regenerate: npm run audit:source
const DASHBOARD_PASSWORD_MIN_LENGTH = 12;
const DASHBOARD_PASSWORD_MAX_LENGTH = 128;
const RBAC_PERMISSION_KEYS = [
  'manageUsers','assignRoles','manageAccounts',
  'viewDashboard','viewRecords','createRecords','editRecords','deleteRecords',
  'submitEvidence','reviewResponses','reviewChanges','viewApprovalHistory',
  'viewReports','viewIamReports','viewExecutive','viewAnalytics','viewCrossFramework',
  'viewEvidence','uploadEvidence','manageEvidence','manageEvidenceRequests',
  'viewControlTests','manageControlTests','viewMonitoring','manageMonitoring',
  'viewAuditReadiness','exportAuditReadiness'
];
function rbacEmptyPermissions_() {
  const out = {};
  RBAC_PERMISSION_KEYS.forEach(function(key){ out[key] = false; });
  return out;
}

function rbacNormalizePermissions_(value) {
  let source = value || {};
  if (typeof source === 'string') {
    try { source = JSON.parse(source || '{}'); } catch (error) { source = {}; }
  }
  const out = rbacEmptyPermissions_();
  RBAC_PERMISSION_KEYS.forEach(function(key){ out[key] = source[key] === true || String(source[key]).toLowerCase() === 'true'; });
  return out;
}

function rbacBuiltInRoleSeeds_() {
  function permissions(keys) {
    const out = rbacEmptyPermissions_();
    keys.forEach(function(key){ out[key] = true; });
    return out;
  }
  return [
    {
      name:'Admin', description:'Full administrative dashboard access. Core IAM permissions are protected to prevent lockout.', builtIn:true, active:true,
      scopeType:'ANY', allowedDepartments:[], allowedUsers:[],
      permissions:permissions(RBAC_PERMISSION_KEYS)
    },
    {
      name:'Compliance Manager', description:'Operational compliance access with controlled record changes, review and reporting.', builtIn:true, active:true,
      scopeType:'ANY', allowedDepartments:[], allowedUsers:[],
      permissions:permissions(['viewDashboard','viewRecords','createRecords','editRecords','deleteRecords','reviewResponses','reviewChanges','viewApprovalHistory','viewReports','viewIamReports','viewExecutive','viewAnalytics','viewCrossFramework'])
    },
    {
      name:'Auditor', description:'Read-only assurance access to compliance information and approval history.', builtIn:true, active:true,
      scopeType:'ANY', allowedDepartments:[], allowedUsers:[],
      permissions:permissions(['viewDashboard','viewRecords','viewApprovalHistory','viewReports','viewExecutive','viewAnalytics','viewCrossFramework'])
    },
    {
      name:'Department User', description:'Department-scoped access to view assigned records and submit responses and evidence.', builtIn:true, active:true,
      scopeType:'DEPARTMENTS', allowedDepartments:[], allowedUsers:[],
      permissions:permissions(['viewDashboard','viewRecords','submitEvidence','viewReports','viewAnalytics'])
    }
  ];
}

function rbacSafeJsonArray_(value) {
  if (Array.isArray(value)) return value.map(function(item){ return String(item || '').trim(); }).filter(Boolean);
  const raw = String(value || '').trim();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(function(item){ return String(item || '').trim(); }).filter(Boolean) : [];
  } catch (error) {
    return raw.split(',').map(function(item){ return item.trim(); }).filter(Boolean);
  }
}

function rbacNormalizeRole_(role) {
  role = role || {};
  const scope = String(role.scopeType || 'ANY').trim().toUpperCase();
  // Upgrade existing built-in role rows in memory, without overwriting Role_Definitions.
  // Explicit false is respected for Compliance Manager; custom roles default to denied.
  const roleName = String(role.name || '').trim();
  let storedPermissions = role.permissions || {};
  if (typeof storedPermissions === 'string') {
    try { storedPermissions = JSON.parse(storedPermissions || '{}'); } catch (error) { storedPermissions = {}; }
  }
  const normalizedPermissions = rbacNormalizePermissions_(role.permissions);
  // V53: backfill only legacy built-in role rows which predate the evidence permissions.
  // A persisted explicit false is always authoritative.
  const evidenceDefaults = roleName === 'Admin' || roleName === 'Compliance Manager'
    ? ['viewEvidence','uploadEvidence','manageEvidence','manageEvidenceRequests','viewControlTests','manageControlTests','viewMonitoring','manageMonitoring','viewAuditReadiness','exportAuditReadiness']
    : (roleName === 'Auditor' || roleName === 'Department User'
      ? (roleName === 'Auditor' ? ['viewEvidence','uploadEvidence','viewControlTests','viewMonitoring','viewAuditReadiness'] : ['viewEvidence','uploadEvidence']) : []);
  evidenceDefaults.forEach(function(key){
    if (!Object.prototype.hasOwnProperty.call(storedPermissions,key)) normalizedPermissions[key] = true;
  });
  // Evidence upload uses the existing secure response submission flow as well.
  if (['Auditor','Compliance Manager'].indexOf(roleName) !== -1 && !Object.prototype.hasOwnProperty.call(storedPermissions,'uploadEvidence')) {
    normalizedPermissions.submitEvidence = true;
  }
  if (roleName === 'Admin' && (role.builtIn === true || String(role.builtIn).toLowerCase() === 'true')) {
    // Admin is a fixed, built-in full-access role. Legacy stored false flags
    // cannot silently remove its permissions. This does not alter approver identity checks.
    RBAC_PERMISSION_KEYS.forEach(function(key){ normalizedPermissions[key] = true; });
  }
  else if (roleName === 'Compliance Manager' &&
           !Object.prototype.hasOwnProperty.call(storedPermissions, 'viewIamReports')) {
    normalizedPermissions.viewIamReports = true;
  }
  return {
    name:String(role.name || '').trim(),
    description:String(role.description || '').trim(),
    builtIn:role.builtIn === true || String(role.builtIn).toLowerCase() === 'true',
    active:role.active !== false && String(role.active).toLowerCase() !== 'false',
    scopeType:['ANY','DEPARTMENTS','USERS'].indexOf(scope) === -1 ? 'ANY' : scope,
    allowedDepartments:rbacSafeJsonArray_(role.allowedDepartments),
    allowedUsers:rbacSafeJsonArray_(role.allowedUsers).map(function(email){ return email.toLowerCase(); }),
    permissions:normalizedPermissions,
    updatedBy:String(role.updatedBy || ''),
    updatedAt:String(role.updatedAt || '')
  };
}

function iamParseDepartments_(value) {
  if (Array.isArray(value)) return value.map(function(v){ return String(v).trim(); }).filter(Boolean);
  if (!value) return [];
  try { const parsed = JSON.parse(String(value)); return Array.isArray(parsed) ? iamParseDepartments_(parsed) : []; }
  catch (ignore) { return []; } // Invalid scope never becomes global access.
}

function iamDepartmentKey_(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g,' '); }

function iamCanReadRecord_(access, record) {
  if (!access.departmentScoped) return true;
  const key = iamDepartmentKey_(record && record.Department);
  return !!key && access.departments.some(function(dept){ return iamDepartmentKey_(dept) === key; });
}

function dashboardPasswordNormalizeForBlocklist_(value) {
  return String(value || '').toLowerCase()
    .replace(/[@4]/g,'a').replace(/[0]/g,'o').replace(/[1!|]/g,'i')
    .replace(/[3]/g,'e').replace(/[5$]/g,'s').replace(/[7+]/g,'t')
    .replace(/[^a-z0-9]/g,'');
}

function dashboardPasswordIdentityTokens_(access) {
  access = access || {};
  const tokens = [];
  function add(value) {
    String(value || '').toLowerCase().split(/[^a-z0-9]+/).forEach(function(part){
      if (part.length >= 3 && tokens.indexOf(part) === -1) tokens.push(part);
    });
  }
  add(access.name);
  const email = String(access.email || '');
  add(email.replace(/@.*$/,''));
  return tokens;
}

function dashboardPasswordHasObviousSequence_(password) {
  const p = String(password || '').toLowerCase();
  if (/(.)\1{3,}/.test(p)) return true;
  const sources = [
    'abcdefghijklmnopqrstuvwxyz', '0123456789',
    'qwertyuiopasdfghjklzxcvbnm'
  ];
  for (let s = 0; s < sources.length; s++) {
    const source = sources[s];
    const reverse = source.split('').reverse().join('');
    for (let i = 0; i <= source.length - 4; i++) {
      const seq = source.slice(i, i + 4);
      if (p.indexOf(seq) !== -1 || p.indexOf(reverse.slice(i, i + 4)) !== -1) return true;
    }
  }
  return false;
}

function dashboardPasswordHasDateLikePattern_(password) {
  const p = String(password || '');
  if (/(?:^|\D)(?:19|20)\d{2}[-/.](?:0?[1-9]|1[0-2])[-/.](?:0?[1-9]|[12]\d|3[01])(?:\D|$)/.test(p)) return true;
  if (/(?:^|\D)(?:0?[1-9]|[12]\d|3[01])[-/.](?:0?[1-9]|1[0-2])[-/.](?:(?:19|20)?\d{2})(?:\D|$)/.test(p)) return true;
  if (/(?:^|\D)(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])(?:\D|$)/.test(p)) return true;
  if (/(?:^|\D)(?:0[1-9]|[12]\d|3[01])(?:0[1-9]|1[0-2])(?:19|20)\d{2}(?:\D|$)/.test(p)) return true;
  return false;
}

function dashboardPasswordPolicy_(password, access) {
  password = String(password || '');
  if (password.length < DASHBOARD_PASSWORD_MIN_LENGTH) throw new Error('Dashboard password must contain at least ' + DASHBOARD_PASSWORD_MIN_LENGTH + ' characters.');
  if (password.length > DASHBOARD_PASSWORD_MAX_LENGTH) throw new Error('Dashboard password must be ' + DASHBOARD_PASSWORD_MAX_LENGTH + ' characters or fewer.');
  if (!/[a-z]/.test(password)) throw new Error('Dashboard password must include at least one lowercase letter.');
  if (!/[A-Z]/.test(password)) throw new Error('Dashboard password must include at least one uppercase letter.');
  if (!/[0-9]/.test(password)) throw new Error('Dashboard password must include at least one number.');
  if (!/[^A-Za-z0-9\s]/.test(password)) throw new Error('Dashboard password must include at least one special character.');
  if (/\s/.test(password)) throw new Error('Dashboard password cannot contain spaces.');
  if (dashboardPasswordHasObviousSequence_(password)) throw new Error('Dashboard password cannot contain obvious sequences or repeated characters such as 1234, abcd or qwerty.');
  if (dashboardPasswordHasDateLikePattern_(password)) throw new Error('Dashboard password cannot contain a birthday-like or calendar-date pattern.');

  const normalized = dashboardPasswordNormalizeForBlocklist_(password);
  const commonWords = [
    'password','passcode','welcome','letmein','qwerty','admin','administrator','user','login',
    'dashboard','compliance','digiclarity','clarityrcm','company','office','security',
    'january','february','march','april','may','june','july','august','september','october','november','december',
    'summer','winter','spring','autumn','india','chennai','mumbai','delhi','newdelhi','bengaluru','bangalore',
    'hyderabad','kolkata','pune','coimbatore','madurai','trichy','kochi','ernakulam'
  ];
  for (let i = 0; i < commonWords.length; i++) {
    if (normalized.indexOf(dashboardPasswordNormalizeForBlocklist_(commonWords[i])) !== -1) {
      throw new Error('Dashboard password cannot contain common words, company terms, month/season names or common location names.');
    }
  }

  const identityTokens = dashboardPasswordIdentityTokens_(access);
  for (let j = 0; j < identityTokens.length; j++) {
    if (normalized.indexOf(dashboardPasswordNormalizeForBlocklist_(identityTokens[j])) !== -1) {
      throw new Error('Dashboard password cannot contain your name or email identity.');
    }
  }
  return password;
}

function phase3Value_(value) {
  const text = String(value == null ? '' : value).trim();
  return /^(n\/?a|not applicable|—|-)$/i.test(text) ? '' : text;
}

function phase3Risk_(record) {
  const raw = phase3Value_(record['Residual Risk Rating']) || phase3Value_(record['Inherent Risk Rating']);
  const standard = ['Critical','High','Medium','Low','Informational'];
  return standard.find(function(label) { return label.toLowerCase() === raw.toLowerCase(); }) || raw;
}

function phase3Status_(record) { return phase3Value_(record['Record Status']).toLowerCase(); }

function phase3IsTerminal_(record) { return ['closed','cancelled','canceled'].indexOf(phase3Status_(record)) !== -1; }

function phase3IsOpen_(record) { return !!phase3Status_(record) && !phase3IsTerminal_(record); }

function phase3Date_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : new Date(value.getTime());
  const text = phase3Value_(value); if (!text) return null;
  let y, m, d, match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:$|[T ])/);
  if (match) { y=Number(match[1]);m=Number(match[2]);d=Number(match[3]); }
  else {
    match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (match) { d=Number(match[1]);m=Number(match[2]);y=Number(match[3]); }
    else {
      match = text.match(/^(\d{1,2})-([a-z]{3})-(\d{4})$/i);
      if (match) { d=Number(match[1]);m=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(match[2].toLowerCase())+1;y=Number(match[3]); }
      else { const parsed = new Date(text); return isNaN(parsed.getTime()) ? null : parsed; }
    }
  }
  const date = new Date(y,m-1,d);
  return date.getFullYear() === y && date.getMonth() === m-1 && date.getDate() === d ? date : null;
}

function phase3DayKey_(value) {
  const date = phase3Date_(value); if (!date) return '';
  return String(date.getFullYear()).padStart(4,'0') + '-' + String(date.getMonth()+1).padStart(2,'0') + '-' + String(date.getDate()).padStart(2,'0');
}

function phase3Overdue_(record, todayKey) {
  if (!phase3IsOpen_(record)) return false;
  const action = phase3Value_(record['Action Status']).toLowerCase();
  if (['closed','completed','cancelled','canceled'].indexOf(action) !== -1) return false;
  const target = phase3DayKey_(record['Target Completion Date']);
  return !!target && !!todayKey && target < todayKey;
}

function phase3PendingEvidence_(record) {
  if (!phase3IsOpen_(record)) return false;
  return phase3Status_(record) === 'pending evidence' ||
    (phase3Value_(record['Evidence Required?']).toLowerCase() === 'yes' &&
     phase3Value_(record['Evidence Validated?']).toLowerCase() !== 'yes');
}

function phase3FrameworkKey_(value) {
  const text = phase3Value_(value).toLowerCase().replace(/\s+/g,' ');
  const key = text.replace(/[^a-z0-9]/g,'');
  if (['soc2','soc2type2','soc2typeii','soc2trustservicescriteria'].indexOf(key) !== -1) return 'soc 2';
  if (['hipaa','hipaasecurityrule','hipaaprivacyrule'].indexOf(key) !== -1) return 'hipaa';
  if (['iso27k','iso27k2022','iso27001','iso270012022','isoiec27001','isoiec270012022'].indexOf(key) !== -1) return 'iso 27001:2022';
  return text;
}

function phase3FrameworkLabel_(value) {
  const key = phase3FrameworkKey_(value);
  return {'soc 2':'SOC 2',hipaa:'HIPAA','iso 27001:2022':'ISO 27001:2022'}[key] || phase3Value_(value);
}
export { RBAC_PERMISSION_KEYS, DASHBOARD_PASSWORD_MIN_LENGTH, DASHBOARD_PASSWORD_MAX_LENGTH, rbacEmptyPermissions_, rbacNormalizePermissions_, rbacBuiltInRoleSeeds_, rbacSafeJsonArray_, rbacNormalizeRole_, iamParseDepartments_, iamDepartmentKey_, iamCanReadRecord_, dashboardPasswordNormalizeForBlocklist_, dashboardPasswordIdentityTokens_, dashboardPasswordHasObviousSequence_, dashboardPasswordHasDateLikePattern_, dashboardPasswordPolicy_, phase3Value_, phase3Risk_, phase3Status_, phase3IsTerminal_, phase3IsOpen_, phase3Date_, phase3DayKey_, phase3Overdue_, phase3PendingEvidence_, phase3FrameworkKey_, phase3FrameworkLabel_ };
