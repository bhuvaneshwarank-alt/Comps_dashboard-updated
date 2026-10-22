/***************************************************************
 * DIGICLARITY COMPLIANCE DASHBOARD
 * Google Apps Script Backend
 *
 * Sheets:
 *   Compliance_Master_Register
 *   Lookup_Master
 *
 * Master schema:
 *   148 columns
 *
 * Compatible with the current Index.html frontend.
 ***************************************************************/

const MASTER_SHEET = 'Compliance_Master_Register';
const LOOKUP_SHEET = 'Lookup_Master';
const FRAMEWORK_CONTROL_SHEET = 'Framework_Control_Master';
const RECORD_HISTORY_SHEET = 'Record_History';
const DRAFT_SHEET = 'Compliance_Draft_Register';
const RECORD_DELETION_SHEET = 'Record_Deletion_Register';
const CHANGE_REQUEST_SHEET = 'Change_Requests';
const APPROVAL_MATRIX_SHEET = 'Approval_Matrix';

// Phase 8.2: dashboard-local credential store. This password is separate from
// the employee's Google Workspace password and is used only by this dashboard.
const DASHBOARD_CREDENTIAL_SHEET = 'Dashboard_Credentials';
const DASHBOARD_CREDENTIAL_HEADERS = [
  'Email', 'Salt', 'Password Hash', 'Updated By', 'Updated At',
  'Password History JSON', 'Password Expires At', 'Failed Attempts', 'Locked Until', 'Last Login At'
];
const DASHBOARD_PASSWORD_PEPPER_KEY = 'DGS_DASHBOARD_PASSWORD_PEPPER_V1';
const DASHBOARD_PASSWORD_MIN_LENGTH = 12;
const DASHBOARD_PASSWORD_MAX_LENGTH = 128;
const DASHBOARD_PASSWORD_HISTORY_DEPTH = 5;
const DASHBOARD_PASSWORD_MAX_AGE_DAYS = 90;
const DASHBOARD_PASSWORD_MAX_FAILED_ATTEMPTS = 5;
const DASHBOARD_PASSWORD_LOCK_MINUTES = 15;

// Phase 8.3: persistent dashboard sessions + email OTP password recovery.
const DASHBOARD_SESSION_SHEET = 'Dashboard_Sessions';
const DASHBOARD_SESSION_HEADERS = [
  'Session ID', 'Email', 'Token Hash', 'Created At', 'Last Seen At',
  'Revoked At', 'Revoke Reason', 'Credential Updated At',
  'Display Name', 'Role', 'Browser', 'Platform', 'Device', 'User Agent'
];
const DASHBOARD_SESSION_PEPPER_KEY = 'DGS_DASHBOARD_SESSION_PEPPER_V1';

const DASHBOARD_RESET_SHEET = 'Dashboard_Password_Reset';
const DASHBOARD_RESET_HEADERS = [
  'Email', 'OTP Salt', 'OTP Hash', 'OTP Expires At', 'OTP Attempts',
  'Window Started At', 'Window Send Count', 'Last Sent At',
  'Reset Token Hash', 'Reset Token Expires At', 'Verified At', 'Completed At'
];
const DASHBOARD_RESET_PEPPER_KEY = 'DGS_DASHBOARD_RESET_PEPPER_V1';
const DASHBOARD_RESET_OTP_MINUTES = 10;
const DASHBOARD_RESET_TOKEN_MINUTES = 15;
const DASHBOARD_RESET_RESEND_SECONDS = 60;
const DASHBOARD_RESET_MAX_SENDS_PER_HOUR = 5;
const DASHBOARD_RESET_MAX_OTP_ATTEMPTS = 5;

// Phase 8.3 security center: dedicated authentication/session audit trail.
const DASHBOARD_SECURITY_AUDIT_SHEET = 'Dashboard_Security_Audit';
const DASHBOARD_SECURITY_AUDIT_HEADERS = [
  'Event ID', 'Timestamp', 'Actor Email', 'Target Email', 'Event Type', 'Result',
  'Session ID', 'Role', 'Browser', 'Platform', 'Device', 'Details JSON'
];
const DASHBOARD_SECURITY_AUDIT_LIMIT = 300;

const LOGO_FILE_ID = '15boXwaxGYWpXb7T0I5fLz9ere-ldUILx';

const EXPECTED_MASTER_COLUMNS = 148;
const DEFAULT_RECORD_LIMIT = 250;
const WATCHLIST_LIMIT = 15;


/***************************************************************
 * WEB APP
 ***************************************************************/
function doGet() {
  // Phase 8.1: the static authentication gateway may render before RBAC.
  // All compliance data/actions remain protected by server-side permission checks.
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Digiclarity Compliance Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/***************************************************************
 * BASIC HELPERS
 ***************************************************************/
function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Unable to access the active spreadsheet. Make sure this Apps Script project is bound to the dashboard spreadsheet.'
    );
  }

  return ss;
}

function cleanString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function isBlank_(value) {
  return cleanString_(value) === '';
}

function isYes_(value) {
  return cleanString_(value).toLowerCase() === 'yes';
}

function uniqueSorted_(values) {
  const seen = {};

  (values || []).forEach(function(value) {
    const cleaned = cleanString_(value);
    if (cleaned) seen[cleaned] = true;
  });

  return Object.keys(seen).sort(function(a, b) {
    return a.localeCompare(b);
  });
}


/***************************************************************
 * MASTER SHEET
 ***************************************************************/
function getMasterSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(MASTER_SHEET);

  if (!sheet) {
    throw new Error('Master sheet not found: ' + MASTER_SHEET);
  }

  return sheet;
}


/***************************************************************
 * ONE-TIME MASTER SCHEMA MIGRATION
 *
 * Adds Sub Department immediately after Department and initializes
 * the new column for existing records as N/A.
 ***************************************************************/
function migrateMasterForSubDepartment_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const sheet = getMasterSheet_();
  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(cleanString_);

  const existingIndex = headers.indexOf('Sub Department');

  if (existingIndex !== -1) {
    return {
      success: true,
      changed: false,
      message: 'Sub Department column already exists.',
      column: existingIndex + 1,
      currentColumns: lastColumn
    };
  }

  const departmentIndex = headers.indexOf('Department');

  if (departmentIndex === -1) {
    throw new Error('Department column not found in Compliance_Master_Register.');
  }

  sheet.insertColumnAfter(departmentIndex + 1);
  sheet.getRange(1, departmentIndex + 2).setValue('Sub Department');

  const rowCount = sheet.getLastRow() - 1;

  if (rowCount > 0) {
    const defaults = Array.from(
      { length: rowCount },
      function() { return ['N/A']; }
    );

    sheet
      .getRange(2, departmentIndex + 2, rowCount, 1)
      .setValues(defaults);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    changed: true,
    message: 'Sub Department column added successfully.',
    column: departmentIndex + 2,
    currentColumns: sheet.getLastColumn(),
    initializedExistingRecords: rowCount
  };
}


/***************************************************************
 * MASTER SCHEMA
 ***************************************************************/
function validateMasterSchema_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn !== EXPECTED_MASTER_COLUMNS) {
    throw new Error(
      'Master sheet schema mismatch. Expected ' +
      EXPECTED_MASTER_COLUMNS +
      ' columns, found ' +
      lastColumn +
      '.'
    );
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(cleanString_);

  const requiredHeaders = [
    'Record ID',
    'Record Type',
    'Date Identified',
    'Date Logged',
    'Record Title',
    'Record Description',
    'Record Status',
    'Primary Compliance Framework',
    'Department',
    'Sub Department',
    'Likelihood (1-5)',
    'Impact (1-5)',
    'Inherent Risk Score',
    'Inherent Risk Rating',
    'Residual Likelihood (1-5)',
    'Residual Impact (1-5)',
    'Residual Risk Score',
    'Residual Risk Rating',
    'Target Completion Date',
    'Action Status',
    'Logged By',
    'Last Updated By',
    'Last Updated Date'
  ];

  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      throw new Error('Required master header missing: ' + header);
    }
  });

  const seen = {};

  headers.forEach(function(header) {
    if (!header) return;

    if (seen[header]) {
      throw new Error('Duplicate master header found: ' + header);
    }

    seen[header] = true;
  });

  return {
    headers: headers,
    requiredHeaders: requiredHeaders
  };
}


/***************************************************************
 * READ MASTER RECORDS
 ***************************************************************/
function readMasterRecords_() {
  const sheet = getMasterSheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const headers = lastColumn
    ? sheet
        .getRange(1, 1, 1, lastColumn)
        .getDisplayValues()[0]
        .map(cleanString_)
    : [];

  if (headers.length === 0 || lastRow < 2) {
    return {
      headers: headers,
      records: []
    };
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();

  const deletedMap = getDeletedRecordIdMap_();

  const records = values
    .filter(function(row) {
      return row.some(function(cell) {
        return !isBlank_(cell);
      });
    })
    .map(function(row) {
      const record = {};

      headers.forEach(function(header, index) {
        if (header) {
          record[header] = row[index] !== undefined ? row[index] : '';
        }
      });

      return record;
    })
    .filter(function(record) {
      const recordId = cleanString_(record['Record ID']);
      return !!phase3Value_(recordId) && !deletedMap[recordId];
    });

  return {
    headers: headers,
    records: iamScopeRecords_(records)
  };
}


/***************************************************************
 * GET RECORDS
 *
 * Compatible filter names from Index.html:
 *   primaryFramework
 *   recordType
 *   recordStatus
 *   riskRating
 *   department
 ***************************************************************/
function getRecords_(filters) {
  rbacRequireAnyPermission_(['viewDashboard','viewRecords','viewReports','viewAnalytics','viewExecutive','reviewChanges','viewApprovalHistory','submitEvidence','reviewResponses']);

  filters = filters || {};

  const master = readMasterRecords_();
  let records = master.records.slice();

  const primaryFramework = cleanString_(
    filters.primaryFramework || filters.framework || filters['Primary Compliance Framework']
  );

  const recordType = cleanString_(
    filters.recordType || filters['Record Type']
  );

  const recordStatus = cleanString_(
    filters.recordStatus || filters['Record Status']
  );

  const riskRating = cleanString_(
    filters.riskRating ||
    filters['Risk Rating'] ||
    filters['Residual Risk Rating'] ||
    filters['Inherent Risk Rating']
  );

  const department = cleanString_(
    filters.department || filters['Department']
  );

  if (primaryFramework && primaryFramework.toLowerCase() !== 'all') {
    records = records.filter(function(record) {
      return phase3FrameworkKey_(record['Primary Compliance Framework']) === phase3FrameworkKey_(primaryFramework);
    });
  }

  if (recordType && recordType.toLowerCase() !== 'all') {
    records = records.filter(function(record) {
      return phase3Value_(record['Record Type']).toLowerCase() === recordType.toLowerCase();
    });
  }

  if (recordStatus && recordStatus.toLowerCase() !== 'all') {
    records = records.filter(function(record) {
      return phase3Value_(record['Record Status']).toLowerCase() === recordStatus.toLowerCase();
    });
  }

  if (riskRating && riskRating.toLowerCase() !== 'all') {
    records = records.filter(function(record) {
      return phase3Risk_(record).toLowerCase() === riskRating.toLowerCase();
    });
  }

  if (department && department.toLowerCase() !== 'all') {
    records = records.filter(function(record) {
      return phase3Value_(record['Department']).toLowerCase() === department.toLowerCase();
    });
  }

  return records; // All matching records are needed for accurate totals and reports.
}


/***************************************************************
 * LOOKUP MASTER
 ***************************************************************/
function getFrameworkControlSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(FRAMEWORK_CONTROL_SHEET);

  if (!sheet) {
    throw new Error(
      'Framework control sheet not found: ' + FRAMEWORK_CONTROL_SHEET
    );
  }

  return sheet;
}


function getLookupSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(LOOKUP_SHEET);

  if (!sheet) {
    throw new Error('Lookup sheet not found: ' + LOOKUP_SHEET);
  }

  return sheet;
}

function getLookupValues_() {
  rbacRequire_('read');

  const sheet = getLookupSheet_();
  const data = sheet.getDataRange().getDisplayValues();

  if (!data.length) return {};

  const headers = data[0].map(cleanString_);
  const categoryIndex = headers.indexOf('Lookup Category');
  const valueIndex = headers.indexOf('Value');

  if (categoryIndex === -1 || valueIndex === -1) {
    throw new Error(
      'Lookup_Master must contain "Lookup Category" and "Value" columns.'
    );
  }

  const result = {};

  for (let i = 1; i < data.length; i++) {
    const category = cleanString_(data[i][categoryIndex]);
    const value = cleanString_(data[i][valueIndex]);

    if (!category || !value) continue;

    if (!result[category]) result[category] = [];

    if (result[category].indexOf(value) === -1) {
      result[category].push(value);
    }
  }

  const frameworkCategory = 'Primary Compliance Framework';
  if (!result[frameworkCategory]) result[frameworkCategory] = [];
  readFrameworkControlMaster_().controls.forEach(function(control) {
    if (!result[frameworkCategory].some(function(value) {
      return frameworkControlAliasKeyV18_(value) === frameworkControlAliasKeyV18_(control.framework);
    })) result[frameworkCategory].push(control.framework);
  });
  Object.keys(result).forEach(function(category) {
    result[category].sort(function(a, b) {
      return a.localeCompare(b);
    });
  });

  const access = rbacRequire_('read');
  if (access.departmentScoped) result.Department = access.departments.slice();
  return result;
}

function getDependentLookupMap_() {
  const sheet = getLookupSheet_();
  const data = sheet.getDataRange().getDisplayValues();

  if (!data.length) return {};

  const headers = data[0].map(cleanString_);
  const categoryIndex = headers.indexOf('Lookup Category');
  const parentIndex = headers.indexOf('Parent Value');
  const valueIndex = headers.indexOf('Value');
  const activeIndex = headers.indexOf('Active');

  if (
    categoryIndex === -1 ||
    parentIndex === -1 ||
    valueIndex === -1
  ) {
    return {};
  }

  const result = {};

  for (let i = 1; i < data.length; i++) {
    const category = cleanString_(data[i][categoryIndex]);
    const parent = cleanString_(data[i][parentIndex]);
    const value = cleanString_(data[i][valueIndex]);
    const active = activeIndex === -1
      ? true
      : cleanString_(data[i][activeIndex]).toLowerCase() !== 'no';

    if (!category || !parent || !value || !active) continue;

    if (!result[category]) result[category] = {};
    if (!result[category][parent]) result[category][parent] = [];

    if (result[category][parent].indexOf(value) === -1) {
      result[category][parent].push(value);
    }
  }

  Object.keys(result).forEach(function(category) {
    Object.keys(result[category]).forEach(function(parent) {
      result[category][parent].sort(function(a, b) {
        return a.localeCompare(b);
      });
    });
  });

  return result;
}


function getDependentLookupValues_(category, parentValue) {
  rbacRequire_('read');

  const targetCategory = cleanString_(category);
  const targetParent = cleanString_(parentValue);

  if (!targetCategory || !targetParent) return [];

  const map = getDependentLookupMap_();

  return (
    map[targetCategory] &&
    map[targetCategory][targetParent]
  ) || [];
}


function buildLookupMap_() {
  const values = getLookupValues_();
  const map = {};

  Object.keys(values).forEach(function(category) {
    map[category] = {};

    values[category].forEach(function(value) {
      map[category][cleanString_(value)] = true;
    });
  });

  return {
    values: values,
    map: map
  };
}



/***************************************************************
 * FRAMEWORK CONTROL MASTER
 *
 * Sheet:
 *   Framework_Control_Master
 *
 * Expected columns (recommended):
 *   Framework
 *   Control ID
 *   Control Name
 *   Control Description
 *   Domain
 *   Common Control Area (optional)
 *   SOC 2 CC Mapping (optional)
 *   HIPAA Mapping (optional)
 *   Active
 *
 * The reader also accepts common header aliases so the sheet
 * remains easy to maintain.
 ***************************************************************/
function getFrameworkControlHeaderMap_(headers) {
  const normalized = headers.map(function(header) {
    return cleanString_(header).toLowerCase();
  });

  function findIndex(aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const index = normalized.indexOf(aliases[i].toLowerCase());
      if (index !== -1) return index;
    }
    return -1;
  }

  return {
    framework: findIndex([
      'framework',
      'primary compliance framework',
      'compliance framework'
    ]),
    controlId: findIndex([
      'control id',
      'control / requirement id',
      'control/requirement id',
      'requirement / control id',
      'requirement/control id'
    ]),
    controlName: findIndex([
      'control name',
      'control title',
      'name'
    ]),
    controlDescription: findIndex([
      'control description',
      'description',
      'requirement description'
    ]),
    domain: findIndex([
      'domain',
      'framework domain',
      'control domain'
    ]),
    commonControlArea: findIndex([
      'common control area',
      'common control',
      'control area'
    ]),
    soc2CCMapping: findIndex([
      'soc 2 cc mapping',
      'soc2 cc mapping',
      'soc 2 mapping',
      'soc2 mapping'
    ]),
    hipaaMapping: findIndex([
      'hipaa mapping',
      'hipaa control mapping',
      'hipaa crosswalk'
    ]),
    iso27001Mapping: findIndex(['ISO 27001 Mapping', 'ISO 27001:2022 Mapping', 'ISO/IEC 27001:2022 Mapping', 'ISO Mapping']),
    active: findIndex([
      'active',
      'status',
      'enabled'
    ])
  };
}


function isFrameworkControlActive_(value) {
  const text = cleanString_(value).toLowerCase();

  if (!text) return true;

  return [
    'yes',
    'active',
    'true',
    '1',
    'enabled'
  ].indexOf(text) !== -1;
}


function readFrameworkControlMaster_() {
  const sheet = getFrameworkControlSheet_();
  const data = sheet.getDataRange().getDisplayValues();

  if (!data.length) {
    return {
      headers: [],
      controls: []
    };
  }

  const headerRow = data.findIndex(function(row) {
    const m = getFrameworkControlHeaderMap_(row);
    return m.framework >= 0 && m.controlId >= 0 && m.controlName >= 0;
  });
  if (headerRow < 0) throw new Error('Framework_Control_Master headers were not found.');
  const headers = data[headerRow].map(cleanString_);
  let map = getFrameworkControlHeaderMap_(headers);

  if (map.framework === -1) {
    throw new Error(
      'Framework_Control_Master is missing a Framework column.'
    );
  }

  if (map.controlId === -1) {
    throw new Error(
      'Framework_Control_Master is missing a Control ID column.'
    );
  }

  if (map.controlName === -1) {
    throw new Error(
      'Framework_Control_Master is missing a Control Name column.'
    );
  }

  if (map.controlDescription === -1) {
    throw new Error(
      'Framework_Control_Master is missing a Control Description column.'
    );
  }

  const controls = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const sectionMap = getFrameworkControlHeaderMap_(row);
    if (sectionMap.framework >= 0 && sectionMap.controlId >= 0 && sectionMap.controlName >= 0) {
      if (sectionMap.controlDescription < 0) throw new Error('Control Description header is missing on row ' + (i + 1));
      map = sectionMap;
      continue;
    }

    const framework = cleanString_(row[map.framework]);
    const controlId = cleanString_(row[map.controlId]);
    const controlName = cleanString_(row[map.controlName]);
    const controlDescription = cleanString_(row[map.controlDescription]);
    const domain = map.domain !== -1
      ? cleanString_(row[map.domain])
      : '';
    const commonControlArea = map.commonControlArea !== -1
      ? cleanString_(row[map.commonControlArea])
      : '';
    const soc2CCMapping = map.soc2CCMapping !== -1
      ? cleanString_(row[map.soc2CCMapping])
      : '';
    const hipaaMapping = map.hipaaMapping !== -1
      ? cleanString_(row[map.hipaaMapping])
      : '';
    const active = map.active !== -1
      ? cleanString_(row[map.active])
      : 'Yes';

    if (!framework || !controlId || !controlName || !controlDescription) {
      continue;
    }

    if (!isFrameworkControlActive_(active)) {
      continue;
    }

    controls.push({
      framework: framework,
      controlId: controlId,
      controlName: controlName,
      controlDescription: controlDescription,
      domain: domain,
      commonControlArea: commonControlArea,
      soc2CCMapping: soc2CCMapping,
      hipaaMapping: hipaaMapping,
      iso27001Mapping: map.iso27001Mapping >= 0 ? cleanString_(row[map.iso27001Mapping]) : '',
      active: true
    });
  }

  controls.sort(function(a, b) {
    const frameworkCompare =
      a.framework.localeCompare(b.framework);

    if (frameworkCompare !== 0) return frameworkCompare;

    return a.controlId.localeCompare(b.controlId, undefined, {
      numeric: true,
      sensitivity: 'base'
    });
  });

  return {
    headers: headers,
    controls: controls
  };
}


function getFrameworkControls_(framework) {
  rbacRequire_('read');

  const target = cleanString_(framework);

  if (!target || target === 'All') {
    return {
      success: true,
      framework: '',
      controls: readFrameworkControlMaster_().controls
    };
  }

  const all = readFrameworkControlMaster_().controls;

  const controls = all.filter(function(control) {
    return frameworkControlAliasKeyV18_(control.framework) === frameworkControlAliasKeyV18_(target);
  });

  return {
    success: true,
    framework: target,
    controls: controls
  };
}



/**
 * Returns all active common control areas from Framework_Control_Master.
 */
function getCommonControlAreas_() {
  rbacRequire_('read');

  const controls = readFrameworkControlMaster_().controls;

  return {
    success: true,
    areas: uniqueSorted_(
      controls.map(function(control) {
        return control.commonControlArea;
      })
    )
  };
}


/**
 * Returns all active framework controls linked to a Common Control Area.
 * The result is intentionally cross-framework so the frontend can build
 * the consolidated crosswalk view in one call.
 */
function getControlsByCommonControlArea_(commonControlArea) {
  rbacRequire_('read');

  const target = cleanString_(commonControlArea);

  if (!target) {
    return {
      success: true,
      commonControlArea: '',
      controls: []
    };
  }

  const controls = readFrameworkControlMaster_().controls.filter(function(control) {
    return cleanString_(control.commonControlArea).toLowerCase() ===
      target.toLowerCase();
  });

  return {
    success: true,
    commonControlArea: target,
    controls: controls
  };
}


function getFrameworkControlById_(framework, controlId) {
  rbacRequire_('read');

  const targetFramework = cleanString_(framework);
  const targetControlId = cleanString_(controlId);

  if (!targetFramework || !targetControlId) {
    return {
      success: false,
      message: 'Framework and Control ID are required.'
    };
  }

  const controls = readFrameworkControlMaster_().controls;

  const control = controls.find(function(item) {
    return (
      frameworkControlAliasKeyV18_(item.framework) === frameworkControlAliasKeyV18_(targetFramework) &&
      item.controlId.toLowerCase() === targetControlId.toLowerCase()
    );
  });

  if (!control) {
    return {
      success: false,
      message:
        'Control ' + targetControlId +
        ' was not found for framework ' + targetFramework + '.'
    };
  }

  return {
    success: true,
    control: control
  };
}


/***************************************************************
 * FILTER OPTIONS
 *
 * The current Index.html expects:
 *   primaryFramework
 *   recordType
 *   recordStatus
 *   riskRating
 *   department
 ***************************************************************/
function getFilterOptions_() {
  rbacRequireAnyPermission_(['viewDashboard','viewRecords','viewReports','viewAnalytics','viewExecutive','createRecords','editRecords']);

  const records = readMasterRecords_().records;

  const primaryFramework = uniqueSorted_(
    records.map(function(record) {
      return phase3FrameworkLabel_(record['Primary Compliance Framework']);
    })
  );

  const recordType = uniqueSorted_(
    records.map(function(record) {
      return record['Record Type'];
    })
  );

  const recordStatus = uniqueSorted_(
    records.map(function(record) {
      return record['Record Status'];
    })
  );

  const riskRating = uniqueSorted_(
    records
      .map(function(record) {
        return phase3Risk_(record);
      })
  );

  const lookup = getLookupValues_();

  const department =
    (lookup['Department'] && lookup['Department'].length)
      ? uniqueSorted_(lookup['Department'])
      : uniqueSorted_(
          records.map(function(record) {
            return record['Department'];
          })
        );

  return {
    primaryFramework: primaryFramework,
    recordType: recordType,
    recordStatus: recordStatus,
    riskRating: riskRating,
    department: department,

    // Backward-friendly aliases
    frameworks: primaryFramework,
    recordTypes: recordType,
    recordStatuses: recordStatus,
    riskRatings: riskRating,
    departments: department
  };
}


/***************************************************************
 * DATE HELPERS
 ***************************************************************/
function parseRecordDate_(value) {
  return phase3Date_(value);
}

function dateKey_(date) {
  if (!date) return '';

  return Utilities.formatDate(
    date,
    getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/***************************************************************
 * DASHBOARD
 ***************************************************************/
function countByField_(records, fieldName) {
  const counts = {};

  records.forEach(function(record) {
    const value = phase3Value_(record[fieldName]);

    if (!value) return;

    counts[value] = (counts[value] || 0) + 1;
  });

  return Object.keys(counts)
    .filter(function(label) {
      return counts[label] > 0;
    })
    .sort(function(a, b) {
      if (counts[b] !== counts[a]) {
        return counts[b] - counts[a];
      }

      return a.localeCompare(b);
    })
    .map(function(label) {
      return {
        label: label,
        value: counts[label]
      };
    });
}

function countOpenDepartments_(records) {
  const counts = {};

  records.forEach(function(record) {
    const status = cleanString_(record['Record Status']);

    if (!phase3IsOpen_(record)) {
      return;
    }

    const department =
      phase3Value_(record['Department']) || 'Unassigned';

    counts[department] = (counts[department] || 0) + 1;
  });

  return Object.keys(counts)
    .sort(function(a, b) {
      if (counts[b] !== counts[a]) {
        return counts[b] - counts[a];
      }

      return a.localeCompare(b);
    })
    .map(function(label) {
      return {
        label: label,
        value: counts[label]
      };
    });
}

function getDashboardData_(filters) {
  rbacRequirePermission_('viewDashboard');

  const records = getRecords_(filters || {});

  const timeZone = getAnalyticsTimeZone_();
  const todayKey = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const openRecords = records.filter(phase3IsOpen_).length;
  const highCriticalRisks = records.filter(function(record) { return ['High','Critical'].indexOf(phase3Risk_(record)) !== -1; }).length;
  const overdueActions = records.filter(function(record) { return phase3Overdue_(record, todayKey); }).length;
  const pendingEvidence = records.filter(phase3PendingEvidence_).length;
  const closedRecords = records.filter(function(record) { return phase3Status_(record) === 'closed'; }).length;
  const chartRecords = records.map(function(record) {
    return Object.assign({}, record, {'Primary Compliance Framework':phase3FrameworkLabel_(record['Primary Compliance Framework']),
      'Residual Risk Rating':phase3Risk_(record)});
  });

  const charts = {
    frameworks:
      countByField_(chartRecords, 'Primary Compliance Framework'),

    statuses:
      countByField_(records, 'Record Status'),

    risks:
      countByField_(chartRecords, 'Residual Risk Rating'),

    findings:
      countByField_(records, 'Finding Classification'),

    actions:
      countByField_(records, 'Action Status'),

    departments:
      countOpenDepartments_(records)
  };

  return {
    success: true,
    timeZone: timeZone,
    todayKey: todayKey,

    kpis: {
      totalRecords: records.length,
      openRecords: openRecords,
      highCriticalRisks: highCriticalRisks,
      overdueActions: overdueActions,
      pendingEvidence: pendingEvidence,
      closedRecords: closedRecords
    },

    charts: charts,

    watchlist:
      buildWatchlist_(records, todayKey),

    recordCount:
      records.length,

    generatedAt:
      formatGeneratedAt_(),

    filters:
      filters || {}
  };
}

function formatGeneratedAt_() {
  const tz =
    getSpreadsheet_().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone();

  return Utilities.formatDate(
    new Date(),
    tz,
    'dd-MMM-yyyy HH:mm:ss'
  );
}


/***************************************************************
 * WATCHLIST
 *
 * Returns RAW RECORD OBJECTS because the current Index.html
 * reads fields directly from each watchlist record.
 ***************************************************************/
function buildWatchlist_(records, today) {
  const items = [];
  const todayKey = typeof today === 'string' ? today : phase3DayKey_(today);

  records.forEach(function(record) {
    if (!phase3IsOpen_(record)) return;
    const status = phase3Status_(record);
    const risk = phase3Risk_(record);

    const priority =
      cleanString_(record['Action Priority']);

    const actionStatus =
      cleanString_(record['Action Status']);

    const targetDate =
      parseRecordDate_(record['Target Completion Date']);

    let score = 0;

    if (risk === 'Critical') {
      score += 100;
    } else if (risk === 'High') {
      score += 80;
    } else if (risk === 'Medium') {
      score += 40;
    }

    if (
      priority === 'Critical'
    ) {
      score += 60;
    } else if (
      priority === 'High'
    ) {
      score += 45;
    } else if (
      priority === 'Medium'
    ) {
      score += 20;
    }

    if (phase3Overdue_(record, todayKey)) {
      score += 70;
    }

    if (phase3PendingEvidence_(record)) {
      score += 50;
    }

    if (
      status === 'under investigation' ||
      status === 'remediation in progress'
    ) {
      score += 15;
    }

    if (score > 0) {
      items.push({
        record: record,
        score: score
      });
    }
  });

  items.sort(function(a, b) {
    if (b.score !== a.score) {
      return b.score - a.score;
    }

    return cleanString_(
      a.record['Record ID']
    ).localeCompare(
      cleanString_(b.record['Record ID'])
    );
  });

  return items
    .slice(0, WATCHLIST_LIMIT)
    .map(function(item) {
      return item.record;
    });
}


/***************************************************************
 * RECORD DETAIL
 *
 * The current Index.html expects:
 * {
 *   success: true,
 *   record: {...}
 * }
 ***************************************************************/
function getRecoveredRecordTypeFromHistory_(recordId) {
  const target = cleanString_(recordId);
  if (!target) return '';

  try {
    const history = getRecordHistory_(target) || [];
    if (!history.length) return '';

    const supportedTypes = Object.keys(RECORD_TYPE_GROUPS);
    const isSupportedType = function(value) {
      const text = cleanString_(value);
      return supportedTypes.some(function(type) {
        return type.toLowerCase() === text.toLowerCase();
      });
    };

    /* getRecordHistory_() returns newest first. Reconstruct the last
     * known valid Record Type from oldest to newest. This specifically
     * repairs records affected by the earlier edit bug that changed
     * Record Type to N/A. */
    let recovered = '';

    history.slice().reverse().forEach(function(item) {
      if (cleanString_(item['Field Name']) !== 'Record Type') return;

      const oldValue = cleanString_(item['Old Value']);
      const newValue = cleanString_(item['New Value']);

      if (isSupportedType(newValue)) {
        recovered = supportedTypes.find(function(type) {
          return type.toLowerCase() === newValue.toLowerCase();
        }) || newValue;
      } else if (!recovered && isSupportedType(oldValue)) {
        recovered = supportedTypes.find(function(type) {
          return type.toLowerCase() === oldValue.toLowerCase();
        }) || oldValue;
      }
    });

    return recovered;
  } catch (error) {
    console.error('Record type recovery failed:', error);
    return '';
  }
}

function getRecordById_(recordId) {
  rbacRequireAnyPermission_(['viewRecords','createRecords','editRecords','deleteRecords','submitEvidence','reviewResponses','reviewChanges','viewApprovalHistory']);

  const target = cleanString_(recordId);

  if (!target) {
    return {
      success: false,
      message: 'Record ID is required.'
    };
  }

  const records =
    readMasterRecords_().records;

  const deletedMap = getDeletedRecordIdMap_();

  const record =
    records.find(function(item) {
      return cleanString_(item['Record ID']) === target;
    });

  if (!record) {
    return {
      success: false,
      message: 'Record not found.'
    };
  }

  const currentType = cleanString_(record['Record Type']);
  if (!normalizeRecordTypeKey_(currentType)) {
    const recoveredType = getRecoveredRecordTypeFromHistory_(target);
    if (recoveredType) {
      record['Record Type'] = recoveredType;
    }
  }

  return {
    success: true,
    record: record,
    isDeleted: !!deletedMap[target]
  };
}



/***************************************************************
 * PHASE 2 SUPPORT SHEETS
 *
 * These sheets are created automatically when the feature is
 * first used. The Compliance Master Register schema remains
 * unchanged at 148 columns.
 ***************************************************************/
function getOrCreateSheetWithHeaders_(sheetName, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentLastColumn = sheet.getLastColumn();
  const currentHeaders = currentLastColumn > 0
    ? sheet.getRange(1, 1, 1, currentLastColumn).getDisplayValues()[0].map(cleanString_)
    : [];

  headers.forEach(function(header, index) {
    const existingIndex = currentHeaders.indexOf(header);
    if (existingIndex === -1) {
      sheet.getRange(1, currentHeaders.length + 1).setValue(header);
      currentHeaders.push(header);
    }
  });

  return sheet;
}

function getRecordDeletionSheet_() {
  return getOrCreateSheetWithHeaders_(
    RECORD_DELETION_SHEET,
    [
      'Record ID',
      'Record Type',
      'Record Title',
      'Deletion Status',
      'Deleted By',
      'Deleted Date',
      'Deletion Reason',
      'Restored By',
      'Restored Date',
      'Restoration Reason'
    ]
  );
}

function getDeletedRecordIdMap_() {
  const sheet = getSpreadsheet_().getSheetByName(RECORD_DELETION_SHEET);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return {};

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(cleanString_);

  const idIndex = headers.indexOf('Record ID');
  const statusIndex = headers.indexOf('Deletion Status');
  if (idIndex === -1) return {};

  const map = {};
  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  values.forEach(function(row) {
    const recordId = cleanString_(row[idIndex]);
    if (!recordId) return;

    const status = statusIndex === -1
      ? 'Deleted'
      : cleanString_(row[statusIndex]).toLowerCase();

    if (status === 'deleted') {
      map[recordId] = true;
    } else if (status === 'restored') {
      delete map[recordId];
    }
  });

  return map;
}

function deleteComplianceRecord_(recordId, deletionReason) {
  rbacRequirePermission_('deleteRecords');

  const target = cleanString_(recordId);
  const reason = cleanString_(deletionReason);

  if (!target) throw new Error('Record ID is required.');
  if (!reason) throw new Error('Deletion reason is required.');

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const sheet = getMasterSheet_();
    const schema = validateMasterSchema_(sheet);
    const headers = schema.headers;
    const rowNumber = findRecordRowById_(sheet, target, headers);

    if (rowNumber === -1) {
      throw new Error('Record not found: ' + target);
    }

    const deletedMap = getDeletedRecordIdMap_();
    if (deletedMap[target]) {
      throw new Error('Record ' + target + ' is already deleted.');
    }

    // TC-24: a legal hold blocks even the existing soft-delete route.
    tc24AssertDeletionAllowed_(target);

    const displayRow = sheet
      .getRange(rowNumber, 1, 1, headers.length)
      .getDisplayValues()[0];

    const record = {};
    headers.forEach(function(header, index) {
      if (header) record[header] = displayRow[index] !== undefined ? displayRow[index] : '';
    });

    const deletedBy = getCurrentUserEmail_();
    const deletedDate = new Date();
    const deletionSheet = getRecordDeletionSheet_();

    deletionSheet.appendRow([
      target,
      cleanString_(record['Record Type']),
      cleanString_(record['Record Title']),
      'Deleted',
      deletedBy,
      deletedDate,
      reason
    ]);

    SpreadsheetApp.flush();

    let historyLogged = true;
    try {
      appendRecordHistoryEntries_([
        {
          timestamp: deletedDate.toISOString(),
          recordId: target,
          action: 'DELETE',
          fieldName: 'Record',
          oldValue: 'Active',
          newValue: 'Deleted',
          changedBy: deletedBy,
          changeReason: reason
        }
      ]);
    } catch (historyError) {
      historyLogged = false;
      console.error('Record deletion history logging failed:', historyError);
    }

    return {
      success: true,
      recordId: target,
      deletedBy: deletedBy,
      deletedDate: deletedDate.toISOString(),
      deletionReason: reason,
      historyLogged: historyLogged,
      message: 'Compliance record soft-deleted successfully.'
    };
  } catch (error) {
    console.error('deleteComplianceRecord error:', error);
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to delete compliance record.'
    );
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


function getDeletedRecordsForRecovery_() {
  rbacRequirePermission_('deleteRecords');

  const deletionSheet = getSpreadsheet_().getSheetByName(RECORD_DELETION_SHEET);
  if (!deletionSheet || deletionSheet.getLastRow() < 2) return [];

  const deletionHeaders = deletionSheet
    .getRange(1, 1, 1, deletionSheet.getLastColumn())
    .getDisplayValues()[0]
    .map(cleanString_);

  const deletionValues = deletionSheet
    .getRange(2, 1, deletionSheet.getLastRow() - 1, deletionSheet.getLastColumn())
    .getDisplayValues();

  const statusIndex = deletionHeaders.indexOf('Deletion Status');
  const idIndex = deletionHeaders.indexOf('Record ID');

  if (statusIndex === -1 || idIndex === -1) return [];

  const master = readMasterRecords_().records;
  const masterById = {};
  master.forEach(function(record) {
    const id = cleanString_(record['Record ID']);
    if (id) masterById[id] = record;
  });

  const latestDeleted = {};

  deletionValues.forEach(function(row) {
    const id = cleanString_(row[idIndex]);
    const status = cleanString_(row[statusIndex]).toLowerCase();
    if (!id) return;

    const item = {};
    deletionHeaders.forEach(function(header, index) {
      if (header) item[header] = row[index] !== undefined ? row[index] : '';
    });

    if (status === 'deleted') {
      latestDeleted[id] = item;
    } else if (status === 'restored') {
      delete latestDeleted[id];
    }
  });

  return Object.keys(latestDeleted).map(function(id) {
    const deletion = latestDeleted[id];
    const masterRecord = masterById[id] || {};

    return {
      'Record ID': id,
      'Record Type': cleanString_(masterRecord['Record Type']) || cleanString_(deletion['Record Type']),
      'Record Title': cleanString_(masterRecord['Record Title']) || cleanString_(deletion['Record Title']),
      'Deleted By': cleanString_(deletion['Deleted By']),
      'Deleted Date': cleanString_(deletion['Deleted Date']),
      'Deletion Reason': cleanString_(deletion['Deletion Reason'])
    };
  }).sort(function(a, b) {
    return String(b['Deleted Date'] || '').localeCompare(String(a['Deleted Date'] || ''));
  });
}


function restoreComplianceRecord_(recordId, restorationReason) {
  rbacRequirePermission_('deleteRecords');

  const target = cleanString_(recordId);
  const reason = cleanString_(restorationReason);

  if (!target) throw new Error('Record ID is required.');
  if (!reason) throw new Error('Restoration reason is required.');

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const masterSheet = getMasterSheet_();
    const schema = validateMasterSchema_(masterSheet);
    const headers = schema.headers;
    const rowNumber = findRecordRowById_(masterSheet, target, headers);

    if (rowNumber === -1) {
      throw new Error('Record not found: ' + target);
    }

    const deletedMap = getDeletedRecordIdMap_();
    if (!deletedMap[target]) {
      throw new Error('Record ' + target + ' is not currently deleted.');
    }

    const deletionSheet = getRecordDeletionSheet_();
    const deletionHeaders = deletionSheet
      .getRange(1, 1, 1, deletionSheet.getLastColumn())
      .getDisplayValues()[0]
      .map(cleanString_);

    const idIndex = deletionHeaders.indexOf('Record ID');
    const statusIndex = deletionHeaders.indexOf('Deletion Status');

    let restoreRow = -1;
    for (let row = deletionSheet.getLastRow(); row >= 2; row--) {
      const rowValues = deletionSheet
        .getRange(row, 1, 1, deletionHeaders.length)
        .getDisplayValues()[0];

      if (
        cleanString_(rowValues[idIndex]) === target &&
        cleanString_(rowValues[statusIndex]).toLowerCase() === 'deleted'
      ) {
        restoreRow = row;
        break;
      }
    }

    if (restoreRow === -1) {
      throw new Error('Deletion record entry not found for ' + target + '.');
    }

    const restoredBy = getCurrentUserEmail_();
    const restoredDate = new Date();

    function setCell(header, value) {
      const index = deletionHeaders.indexOf(header);
      if (index !== -1) {
        deletionSheet.getRange(restoreRow, index + 1).setValue(value);
      }
    }

    setCell('Deletion Status', 'Restored');
    setCell('Restored By', restoredBy);
    setCell('Restored Date', restoredDate);
    setCell('Restoration Reason', reason);

    SpreadsheetApp.flush();

    let historyLogged = true;
    try {
      appendRecordHistoryEntries_([
        {
          timestamp: restoredDate.toISOString(),
          recordId: target,
          action: 'RESTORE',
          fieldName: 'Record',
          oldValue: 'Deleted',
          newValue: 'Active',
          changedBy: restoredBy,
          changeReason: reason
        }
      ]);
    } catch (historyError) {
      historyLogged = false;
      console.error('Record restoration history logging failed:', historyError);
    }

    return {
      success: true,
      recordId: target,
      restoredBy: restoredBy,
      restoredDate: restoredDate.toISOString(),
      restorationReason: reason,
      historyLogged: historyLogged,
      message: 'Compliance record restored successfully.'
    };

  } catch (error) {
    console.error('restoreComplianceRecord error:', error);
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to restore compliance record.'
    );
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


function getRecordHistorySheet_() {
  return getOrCreateSheetWithHeaders_(
    RECORD_HISTORY_SHEET,
    [
      'Timestamp',
      'Record ID',
      'Action',
      'Field Name',
      'Old Value',
      'New Value',
      'Changed By',
      'Change Reason'
    ]
  );
}

function getDraftSheet_() {
  return getOrCreateSheetWithHeaders_(
    DRAFT_SHEET,
    [
      'Draft ID',
      'Record Type',
      'Record Title',
      'Current Step',
      'Draft Status',
      'Payload JSON',
      'Created By',
      'Created Date',
      'Last Updated By',
      'Last Updated Date',
      'Submitted Record ID'
    ]
  );
}

function appendRecordHistoryEntries_(entries) {
  if (!entries || !entries.length) return;

  const sheet = getRecordHistorySheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
  const rows = entries.map(function(entry) {
    const row = Array(headers.length).fill('');
    const set = function(header, value) {
      const index = headers.indexOf(header);
      if (index !== -1) row[index] = value === undefined || value === null ? '' : String(value);
    };

    set('Timestamp', entry.timestamp || new Date().toISOString());
    set('Record ID', entry.recordId);
    set('Action', entry.action);
    set('Field Name', entry.fieldName);
    set('Old Value', entry.oldValue);
    set('New Value', entry.newValue);
    set('Changed By', entry.changedBy || getCurrentUserEmail_());
    set('Change Reason', entry.changeReason || '');
    return row;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();
}

function getRecordHistory_(recordId) {
  rbacRequireAnyPermission_(['viewRecords','editRecords','reviewChanges','viewApprovalHistory']);

  const target = cleanString_(recordId);
  if (!target) return [];
  iamRequireRecord_(target);

  const sheet = getRecordHistorySheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(cleanString_);
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();

  const records = values.map(function(row) {
    const item = {};
    headers.forEach(function(header, index) {
      if (header) item[header] = row[index] !== undefined ? row[index] : '';
    });
    return item;
  }).filter(function(item) {
    return cleanString_(item['Record ID']) === target;
  });

  records.sort(function(a, b) {
    return String(b['Timestamp'] || '').localeCompare(String(a['Timestamp'] || ''));
  });

  return records;
}

function findRecordRowById_(sheet, recordId, headers) {
  const target = cleanString_(recordId);
  const index = headers.indexOf('Record ID');
  if (index === -1 || !target) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, index + 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (cleanString_(values[i][0]) === target) return i + 2;
  }

  return -1;
}

function buildMergedUpdatePayload_(existingRecord, payload, applicableFields, headers) {
  const merged = {};
  const source = payload && typeof payload === 'object' ? payload : {};

  /*
   * Start with the complete existing record so every value is preserved
   * by default. We then overlay only fields that the edit form actually
   * submitted and that belong to the current Record Type workflow.
   *
   * This makes repeated edits safe: a second edit cannot accidentally
   * replace fields from an earlier edit with blanks or N/A values.
   */
  headers.forEach(function(header) {
    if (!header || SYSTEM_MANAGED_FIELDS.indexOf(header) !== -1) return;
    merged[header] = existingRecord[header] !== undefined
      ? existingRecord[header]
      : '';
  });

  applicableFields.forEach(function(field) {
    if (SYSTEM_MANAGED_FIELDS.indexOf(field) !== -1) return;
    if (Object.prototype.hasOwnProperty.call(source, field)) {
      merged[field] = source[field];
    }
  });

  const requestedType = Object.prototype.hasOwnProperty.call(source, 'Record Type')
    ? source['Record Type'] : existingRecord['Record Type'];
  const targetType = normalizeRecordTypeKey_(requestedType);
  if (!targetType) throw new Error('Select a supported Record Type.');
  merged['Record Type'] = targetType;
  if (targetType !== normalizeRecordTypeKey_(existingRecord['Record Type'])) {
    const previousFields = getApplicableFieldsForType_(existingRecord['Record Type']);
    applicableFields.forEach(function(field) {
      if (SYSTEM_MANAGED_FIELDS.indexOf(field) === -1 && previousFields.indexOf(field) === -1 && cleanString_(merged[field]).toUpperCase() === 'N/A') {
        merged[field] = '';
      }
    });
  }
  return merged;
}

function updateComplianceRecord_(recordId, payload, changeReason) {
  const target = cleanString_(recordId);
  if (!target) throw new Error('Record ID is required.');
  if (!payload || typeof payload !== 'object') throw new Error('Invalid record data received.');

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const sheet = getMasterSheet_();
    const schema = validateMasterSchema_(sheet);
    const headers = schema.headers;
    const rowNumber = findRecordRowById_(sheet, target, headers);

    if (rowNumber === -1) throw new Error('Record not found: ' + target);

    const deletedMap = getDeletedRecordIdMap_();
    if (deletedMap[target]) {
      throw new Error(
        'Record ' + target + ' is deleted. Restore the record from Recovery before editing or updating it.'
      );
    }

    const oldRawRow = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const oldDisplayRow = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
    const existingRecord = {};
    headers.forEach(function(header, index) {
      if (header) existingRecord[header] = oldDisplayRow[index] !== undefined ? oldDisplayRow[index] : '';
    });

    let existingType = normalizeRecordTypeKey_(existingRecord['Record Type']);

    /*
     * Recover records that were damaged by the earlier edit implementation
     * when Record Type was accidentally written as N/A. The recovered type
     * is also written back to the master row during this edit so the record
     * becomes clean again.
     */
    if (!existingType) {
      const recoveredType = getRecoveredRecordTypeFromHistory_(target);
      if (recoveredType) {
        existingRecord['Record Type'] = recoveredType;
        existingType = normalizeRecordTypeKey_(recoveredType);
      }
    }

    // Legacy N/A types can be corrected by an explicitly approved supported type.
    const targetType = normalizeRecordTypeKey_(Object.prototype.hasOwnProperty.call(payload, 'Record Type')
      ? payload['Record Type'] : existingType);
    if (!targetType) throw new Error('Select a supported Record Type.');
    const applicableFields = getApplicableFieldsForType_(targetType);
    const merged = buildMergedUpdatePayload_(existingRecord, payload, applicableFields, headers);

    getCreateRecordRequiredFields_().forEach(function(field) {
      if (cleanString_(merged[field]) === '') throw new Error(field + ' is required.');
    });

    /*
     * Existing records may contain legacy lookup values that are no
     * longer present in Lookup_Master (for example, historical
     * department names). Do not block an edit merely because an
     * unchanged legacy value is not in today's lookup list. Any
     * lookup-controlled value that is actually changed is still
     * validated against the current lookup master.
     */
    validateUpdateRecordLookups_(existingRecord, merged);

    /* Backend-enforced conditional field dependencies. */
    applyConditionalFieldDependencies_(merged);

    const selectedFramework = cleanString_(merged['Primary Compliance Framework']);
    const selectedControlId = cleanString_(merged['Requirement / Control ID']);
    let selectedFrameworkControl = null;

    if (selectedControlId) {
      const controlResult = getFrameworkControlById_(selectedFramework, selectedControlId);
      if (!controlResult.success) throw new Error(controlResult.message);
      selectedFrameworkControl = controlResult.control;
    }

    const now = new Date();
    const currentUser = getCurrentUserEmail_();
    const newRow = headers.map(function(header, index) {
      if (header === 'Record ID') return oldRawRow[index];
      if (header === 'Record Type') return targetType;
      if (header === 'Date Logged') return oldRawRow[index];
      if (header === 'Logged By') return oldRawRow[index];
      if (header === 'Last Updated By') return currentUser;
      if (header === 'Last Updated Date') return now;

      if (header === 'Inherent Risk Score') {
        if (applicableFields.indexOf(header) === -1) return oldRawRow[index];
        return calculateRiskScore_(merged['Likelihood (1-5)'], merged['Impact (1-5)']);
      }

      if (header === 'Residual Risk Score') {
        if (applicableFields.indexOf(header) === -1) return oldRawRow[index];
        return calculateRiskScore_(merged['Residual Likelihood (1-5)'], merged['Residual Impact (1-5)']);
      }

      /*
       * EDIT MODE: preserve existing values for fields that are outside
       * the current Record Type matrix. Do not turn legacy/non-applicable
       * fields into N/A just because the user is editing the record.
       */
      if (applicableFields.indexOf(header) === -1) {
        return oldRawRow[index];
      }

      if (header === 'Requirement Description' && selectedFrameworkControl) {
        return selectedFrameworkControl.controlDescription;
      }

      if (header === 'Framework Domain' && selectedFrameworkControl) {
        return selectedFrameworkControl.domain || '';
      }

      if (Object.prototype.hasOwnProperty.call(merged, header)) {
        return normalizeRecordFieldValue_(header, merged[header]);
      }

      return oldRawRow[index];
    });

    if (newRow.length !== EXPECTED_MASTER_COLUMNS) {
      throw new Error('Generated update row contains ' + newRow.length + ' columns. Expected ' + EXPECTED_MASTER_COLUMNS + '.');
    }

    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([newRow]);
    SpreadsheetApp.flush();

    const newDisplayRow = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
    const timestamp = new Date().toISOString();
    const reason = cleanString_(changeReason);
    const historyEntries = [];

    headers.forEach(function(header, index) {
      if (!header || SYSTEM_MANAGED_FIELDS.indexOf(header) !== -1) return;

      const oldValue = oldDisplayRow[index] !== undefined ? String(oldDisplayRow[index]) : '';
      const newValue = newDisplayRow[index] !== undefined ? String(newDisplayRow[index]) : '';

      if (oldValue !== newValue) {
        historyEntries.push({
          timestamp: timestamp,
          recordId: target,
          action: 'UPDATE',
          fieldName: header,
          oldValue: oldValue,
          newValue: newValue,
          changedBy: currentUser,
          changeReason: reason
        });
      }
    });

    let historyLogged = true;
    try {
      appendRecordHistoryEntries_(historyEntries);
    } catch (historyError) {
      historyLogged = false;
      console.error('Record history logging failed:', historyError);
    }

    const updatedRecord = {};
    headers.forEach(function(header, index) {
      if (header) {
        updatedRecord[header] =
          newDisplayRow[index] !== undefined
            ? newDisplayRow[index]
            : '';
      }
    });

    return {
      success: true,
      recordId: target,
      changedFieldCount: historyEntries.length,
      historyLogged: historyLogged,
      updatedBy: currentUser,
      timestamp: timestamp,
      record: updatedRecord,
      message: historyEntries.length
        ? 'Compliance record updated successfully.'
        : 'No field changes were detected.'
    };

  } catch (error) {
    console.error('updateComplianceRecord error:', error);
    throw new Error(error && error.message ? error.message : 'Unable to update compliance record.');
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


/***************************************************************
 * COMPLIANCE DRAFT REGISTER
 ***************************************************************/
function generateNextDraftId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return newDraftId_(getCurrentYear_(), 1);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
  const idIndex = headers.indexOf('Draft ID');
  if (idIndex === -1) throw new Error('Draft ID column not found.');

  const ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getDisplayValues().map(function(row) { return cleanString_(row[0]); });
  const year = String(getCurrentYear_());
  let highest = 0;

  ids.forEach(function(id) {
    const match = id.match(new RegExp('^DRAFT-' + year + '-(\\d+)$'));
    if (!match) return;
    const number = parseInt(match[1], 10);
    if (!isNaN(number) && number > highest) highest = number;
  });

  return newDraftId_(year, highest + 1);
}

function newDraftId_(year, number) {
  return 'DRAFT-' + String(year) + '-' + String(number).padStart(4, '0');
}

function findDraftRowById_(sheet, draftId, headers) {
  const index = headers.indexOf('Draft ID');
  if (index === -1) return -1;
  const target = cleanString_(draftId);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !target) return -1;

  const values = sheet.getRange(2, index + 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (cleanString_(values[i][0]) === target) return i + 2;
  }
  return -1;
}

function parseDraftRow_(headers, row) {
  const item = {};
  headers.forEach(function(header, index) {
    if (header) item[header] = row[index] !== undefined ? row[index] : '';
  });

  let payload = {};
  try {
    payload = item['Payload JSON'] ? JSON.parse(item['Payload JSON']) : {};
  } catch (error) {
    payload = {};
  }
  item.payload = payload;
  return item;
}

function saveComplianceDraft_(draftId, payload, currentStep) {
  rbacRequireAnyPermission_(['createRecords','editRecords']);

  if (!payload || typeof payload !== 'object') throw new Error('Invalid draft data received.');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const sheet = getDraftSheet_();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
    const now = new Date();
    const currentUser = getCurrentUserEmail_();
    const step = Math.min(Math.max(parseInt(currentStep, 10) || 1, 1), 6);
    const safeDraftId = cleanString_(draftId) || generateNextDraftId_(sheet);
    const rowNumber = findDraftRowById_(sheet, safeDraftId, headers);
    const title = cleanString_(payload['Record Title']) || 'Untitled Draft';
    const recordType = cleanString_(payload['Record Type']);
    const payloadJson = JSON.stringify(payload);

    if (rowNumber === -1) {
      const rowObject = {
        'Draft ID': safeDraftId,
        'Record Type': recordType,
        'Record Title': title,
        'Current Step': step,
        'Draft Status': 'Saved',
        'Payload JSON': payloadJson,
        'Created By': currentUser,
        'Created Date': now,
        'Last Updated By': currentUser,
        'Last Updated Date': now
      };
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([headers.map(function(header) { return rowObject[header] !== undefined ? rowObject[header] : ''; })]);
    } else {
      const oldRow = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      const rowObject = {
        'Draft ID': safeDraftId,
        'Record Type': recordType,
        'Record Title': title,
        'Current Step': step,
        'Draft Status': 'Saved',
        'Payload JSON': payloadJson,
        'Created By': oldRow[headers.indexOf('Created By')],
        'Created Date': oldRow[headers.indexOf('Created Date')],
        'Last Updated By': currentUser,
        'Last Updated Date': now
      };
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([headers.map(function(header) { return rowObject[header] !== undefined ? rowObject[header] : ''; })]);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      draftId: safeDraftId,
      recordType: recordType,
      recordTitle: title,
      currentStep: step,
      updatedBy: currentUser,
      timestamp: now.toISOString(),
      message: 'Draft saved successfully.'
    };

  } catch (error) {
    console.error('saveComplianceDraft error:', error);
    throw new Error(error && error.message ? error.message : 'Unable to save draft.');
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getComplianceDrafts_() {
  rbacRequireAnyPermission_(['createRecords','editRecords']);

  const sheet = getDraftSheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(cleanString_);
  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues()
    .map(function(row) { return parseDraftRow_(headers, row); })
    .filter(function(item) { return cleanString_(item['Draft Status']) === 'Saved'; })
    .sort(function(a, b) { return String(b['Last Updated Date'] || '').localeCompare(String(a['Last Updated Date'] || '')); })
    .map(function(item) {
      return {
        'Draft ID': item['Draft ID'],
        'Record Type': item['Record Type'],
        'Record Title': item['Record Title'],
        'Current Step': item['Current Step'],
        'Draft Status': item['Draft Status'],
        'Created By': item['Created By'],
        'Created Date': item['Created Date'],
        'Last Updated By': item['Last Updated By'],
        'Last Updated Date': item['Last Updated Date']
      };
    });
}

function getComplianceDraftById_(draftId) {
  rbacRequireAnyPermission_(['createRecords','editRecords']);

  const target = cleanString_(draftId);
  if (!target) return { success: false, message: 'Draft ID is required.' };

  const sheet = getDraftSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
  const rowNumber = findDraftRowById_(sheet, target, headers);
  if (rowNumber === -1) return { success: false, message: 'Draft not found.' };

  const row = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  const item = parseDraftRow_(headers, row);

  return {
    success: cleanString_(item['Draft Status']) === 'Saved',
    message: cleanString_(item['Draft Status']) === 'Saved' ? 'Draft loaded.' : 'Draft is no longer available.',
    draft: {
      'Draft ID': item['Draft ID'],
      'Record Type': item['Record Type'],
      'Record Title': item['Record Title'],
      'Current Step': Number(item['Current Step']) || 1,
      'Draft Status': item['Draft Status'],
      'Created By': item['Created By'],
      'Created Date': item['Created Date'],
      'Last Updated By': item['Last Updated By'],
      'Last Updated Date': item['Last Updated Date'],
      payload: item.payload || {}
    }
  };
}

function deleteComplianceDraft_(draftId) {
  rbacRequireAnyPermission_(['createRecords','editRecords']);

  const target = cleanString_(draftId);
  if (!target) throw new Error('Draft ID is required.');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getDraftSheet_();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
    const rowNumber = findDraftRowById_(sheet, target, headers);
    if (rowNumber === -1) throw new Error('Draft not found.');

    const statusIndex = headers.indexOf('Draft Status');
    const updatedByIndex = headers.indexOf('Last Updated By');
    const updatedDateIndex = headers.indexOf('Last Updated Date');
    const now = new Date();
    if (statusIndex !== -1) sheet.getRange(rowNumber, statusIndex + 1).setValue('Deleted');
    if (updatedByIndex !== -1) sheet.getRange(rowNumber, updatedByIndex + 1).setValue(getCurrentUserEmail_());
    if (updatedDateIndex !== -1) sheet.getRange(rowNumber, updatedDateIndex + 1).setValue(now);
    SpreadsheetApp.flush();

    return { success: true, draftId: target, message: 'Draft deleted.' };
  } catch (error) {
    console.error('deleteComplianceDraft error:', error);
    throw new Error(error && error.message ? error.message : 'Unable to delete draft.');
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function markComplianceDraftSubmitted_(draftId, recordId) {
  rbacRequireAnyPermission_(['createRecords','editRecords']);

  const target = cleanString_(draftId);
  if (!target) return { success: false, message: 'No draft ID supplied.' };

  const sheet = getDraftSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(cleanString_);
  const rowNumber = findDraftRowById_(sheet, target, headers);
  if (rowNumber === -1) return { success: false, message: 'Draft not found.' };

  const setValue = function(header, value) {
    const index = headers.indexOf(header);
    if (index !== -1) sheet.getRange(rowNumber, index + 1).setValue(value);
  };

  const now = new Date();
  setValue('Draft Status', 'Submitted');
  setValue('Last Updated By', getCurrentUserEmail_());
  setValue('Last Updated Date', now);
  setValue('Submitted Record ID', cleanString_(recordId));

  SpreadsheetApp.flush();
  return { success: true, draftId: target, recordId: cleanString_(recordId), message: 'Draft marked as submitted.' };
}


/***************************************************************
 * LOGO
 ***************************************************************/
function getLogoDataUrl_() {
  rbacRequire_('read');

  try {
    const file = DriveApp.getFileById(LOGO_FILE_ID);
    const blob = file.getBlob();
    const contentType = blob.getContentType() || 'image/png';

    const base64 =
      Utilities.base64Encode(blob.getBytes());

    const dataUrl =
      'data:' +
      contentType +
      ';base64,' +
      base64;

    return dataUrl;

  } catch (error) {
    console.error('Logo loading failed:', error);
    return '';
  }
}


/***************************************************************
 * CONFIG
 ***************************************************************/
function getConfig_() {
  rbacRequire_('read');

  return {
    masterSheet: MASTER_SHEET,
    lookupSheet: LOOKUP_SHEET,
    frameworkControlSheet: FRAMEWORK_CONTROL_SHEET,
    expectedMasterColumns: EXPECTED_MASTER_COLUMNS,
    defaultRecordLimit: DEFAULT_RECORD_LIMIT,
    watchlistLimit: WATCHLIST_LIMIT,
    currentYear: new Date().getFullYear(),
    timestamp: new Date().toISOString()
  };
}



/***************************************************************
 * RECORD-TYPE APPLICABILITY ENGINE
 *
 * One master register / many case types.
 * The UI activates only applicable fields for the selected
 * Record Type. All non-applicable master columns are written
 * as N/A by the backend at record creation.
 ***************************************************************/

const RECORD_TYPE_GROUPS = {
  'Audit Finding': [
    'common', 'compliance', 'finding', 'risk', 'audit',
    'action', 'closure', 'retention'
  ],
  'Incident': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'impact', 'incident', 'notification', 'action', 'closure',
    'retention'
  ],
  'Control Deficiency': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'action', 'closure', 'retention'
  ],
  'Compliance Issue': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'impact', 'notification', 'action', 'closure', 'retention'
  ],
  'Risk': [
    'common', 'context', 'compliance', 'risk', 'impact',
    'action', 'closure', 'retention'
  ],
  'Policy Exception': [
    'common', 'context', 'compliance', 'risk', 'impact',
    'action', 'closure', 'retention'
  ],
  'Observation': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'action', 'closure', 'retention'
  ],
  'CAPA': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'action', 'closure', 'retention'
  ],
  'Client Finding': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'impact', 'notification', 'action', 'closure', 'retention'
  ],
  'Regulatory Finding': [
    'common', 'context', 'compliance', 'finding', 'risk',
    'impact', 'notification', 'action', 'closure', 'retention'
  ]
};

const RECORD_GROUP_DEFINITIONS = {
  common: {
    title: 'Record & Scope',
    fields: [
      'Record Sub-Type',
      'Source / Trigger',
      'Date Identified',
      'Record Title',
      'Record Description',
      'Business Unit',
      'Department',
      'Sub Department',
      'Location',
      'Record Status',
      'Primary Compliance Framework',
      'Secondary Framework(s)'
    ]
  },

  context: {
    title: 'Practice / Process / System',
    fields: [
      'Practice / Client',
      'Process / Function',
      'System / Application'
    ]
  },

  compliance: {
    title: 'Compliance & Control',
    fields: [
      'Requirement / Control ID',
      'Framework Domain',
      'Requirement Description',
      'Control Owner',
      'Control Frequency',
      'Control Type',
      'Control Effectiveness',
      'Policy / Procedure Violated',
      'Regulatory / Contractual Obligation'
    ]
  },

  finding: {
    title: 'Issue & Finding',
    fields: [
      'Issue Category',
      'Issue Type',
      'Nature of Issue',
      'Detection Method',
      'Detection Source',
      'Cause Category',
      'Root Cause',
      'Control Failure Identified',
      'Control Failure Type',
      'Repeat Finding?',
      'Previous Record ID',
      'Auditor / Management Observation',
      'Detailed Finding / Incident Summary'
    ]
  },

  risk: {
    title: 'Risk Assessment',
    fields: [
      'Risk Category',
      'Likelihood (1-5)',
      'Impact (1-5)',
      'Inherent Risk Score',
      'Inherent Risk Rating',
      'Existing Controls',
      'Residual Likelihood (1-5)',
      'Residual Impact (1-5)',
      'Residual Risk Score',
      'Residual Risk Rating',
      'Risk Treatment',
      'Risk Owner',
      'Risk Acceptance Required?',
      'Risk Acceptance Authority',
      'Risk Acceptance Date'
    ]
  },

  impact: {
    title: 'Business & Impact Assessment',
    fields: [
      'Business Impact',
      'Compliance Impact',
      'Security Impact',
      'Privacy Impact',
      'Financial Impact',
      'Operational Impact',
      'Reputational Impact',
      'Potential Financial Exposure',
      'Potential Regulatory Exposure',
      'Records / Population Affected',
      'Estimated Records Affected',
      'Third Parties Impacted',
      'Data Breach Potential',
      'Systems Compromised',
      'Customer / Client Impact',
      'Employee Impact'
    ]
  },

  incident: {
    title: 'Incident & Investigation',
    fields: [
      'Incident ID',
      'Reported By',
      'Mail Subject',
      'Threat Type',
      'Threat Actor',
      'Intent',
      'Forensics Required?',
      'Evidence Collected',
      'Containment Actions',
      'Investigation Lead',
      'Investigation Start Date',
      'Investigation Status',
      'Access Revocation / Modification Completed?',
      'Incident Closure Date'
    ]
  },

  audit: {
    title: 'Audit Details',
    fields: [
      'Audit ID',
      'Audit Type',
      'Audit Name',
      'Audit Period',
      'Audit Start Date',
      'Audit End Date',
      'Auditor / Audit Firm',
      'Audit Scope',
      'Testing Area',
      'Sample Size',
      'Sample Exceptions',
      'Testing Result',
      'Finding Classification',
      'Audit Finding Reference',
      'Management Response'
    ]
  },

  notification: {
    title: 'Notification & Legal',
    fields: [
      'Regulatory Notification Required?',
      'Regulatory Body',
      'Regulatory Notification Date',
      'Client Notification Required?',
      'Client Notification Date',
      'Legal Counsel Engaged?',
      'Legal Review Required?',
      'Breach Notification Required?',
      'Notification Decision',
      'Notification Decision Date',
      'Notification Owner',
      'Data Destruction Certification Required?'
    ]
  },

  action: {
    title: 'Corrective Action & Remediation',
    fields: [
      'Corrective Action Required?',
      'Corrective Action ID',
      'Corrective Action Description',
      'Root Cause Action',
      'Preventive Action',
      'Action Owner',
      'Action Owner Department',
      'Action Priority',
      'Target Completion Date',
      'Actual Completion Date',
      'Action Status',
      'Evidence Required?',
      'Evidence Submitted Date',
      'Evidence Validated?',
      'Validation Date',
      'Validation Owner',
      'Effectiveness Check Required?',
      'Effectiveness Check Date',
      'Effectiveness Result'
    ]
  },

  closure: {
    title: 'Closure & Governance',
    fields: [
      'Business Owner',
      'Management Approver',
      'Compliance Sign-off',
      'Compliance Sign-off Date',
      'Closure Criteria Met?',
      'Closure Approved By',
      'Closure Date',
      'Final Outcome',
      'Lessons Learned',
      'Residual Risk Accepted?'
    ]
  },

  retention: {
    title: 'Retention & Review',
    fields: [
      'Retention Period',
      'Record Retention Expiry Date',
      'Review Frequency',
      'Next Review Date'
    ]
  }
};

const RECORD_FIELD_DEFINITIONS = {
  'Record Sub-Type': { type: 'text' },
  'Source / Trigger': { type: 'text' },
  'Date Identified': { type: 'date' },
  'Record Title': { type: 'text', required: true },
  'Record Description': { type: 'textarea', required: true },
  'Business Unit': { type: 'text' },
  'Department': { type: 'department' },
  'Sub Department': { type: 'subdepartment' },
  'Location': { type: 'text' },
  'Record Status': { type: 'lookup', lookup: 'Record Status', required: true },
  'Primary Compliance Framework': { type: 'lookup', lookup: 'Primary Compliance Framework', required: true },
  'Secondary Framework(s)': { type: 'multilookup', lookup: 'Primary Compliance Framework' },
  'Framework Domain': { type: 'text' },
  'Practice / Client': { type: 'text' },
  'Process / Function': { type: 'text' },
  'System / Application': { type: 'text' },

  'Requirement / Control ID': { type: 'text' },
  'Requirement Description': { type: 'textarea' },
  'Control Owner': { type: 'text' },
  'Control Frequency': { type: 'lookup', lookup: 'Control Frequency' },
  'Control Type': { type: 'lookup', lookup: 'Control Type' },
  'Control Effectiveness': { type: 'lookup', lookup: 'Control Effectiveness' },
  'Policy / Procedure Violated': { type: 'lookup', lookup: 'Policy / Procedure Violated' },
  'Regulatory / Contractual Obligation': { type: 'lookup', lookup: 'Regulatory / Contractual Obligation' },

  'Issue Category': { type: 'text' },
  'Issue Type': { type: 'text' },
  'Nature of Issue': { type: 'text' },
  'Detection Method': { type: 'text' },
  'Detection Source': { type: 'text' },
  'Cause Category': { type: 'text' },
  'Root Cause': { type: 'textarea' },
  'Control Failure Identified': { type: 'lookup', lookup: 'Yes/No' },
  'Control Failure Type': { type: 'text' },
  'Repeat Finding?': { type: 'lookup', lookup: 'Repeat Finding?' },
  'Previous Record ID': { type: 'text' },
  'Auditor / Management Observation': { type: 'textarea' },
  'Detailed Finding / Incident Summary': { type: 'textarea' },

  'Risk Category': { type: 'lookup', lookup: 'Risk Category' },
  'Likelihood (1-5)': { type: 'number', min: 1, max: 5 },
  'Impact (1-5)': { type: 'number', min: 1, max: 5 },
  'Inherent Risk Score': { type: 'calculated' },
  'Inherent Risk Rating': { type: 'lookup', lookup: 'Risk Rating' },
  'Existing Controls': { type: 'textarea' },
  'Residual Likelihood (1-5)': { type: 'number', min: 1, max: 5 },
  'Residual Impact (1-5)': { type: 'number', min: 1, max: 5 },
  'Residual Risk Score': { type: 'calculated' },
  'Residual Risk Rating': { type: 'lookup', lookup: 'Risk Rating' },
  'Risk Treatment': { type: 'lookup', lookup: 'Risk Treatment' },
  'Risk Owner': { type: 'text' },
  'Risk Acceptance Required?': { type: 'lookup', lookup: 'Risk Acceptance Required?' },
  'Risk Acceptance Authority': { type: 'text' },
  'Risk Acceptance Date': { type: 'date' },

  'Business Impact': { type: 'textarea' },
  'Compliance Impact': { type: 'textarea' },
  'Security Impact': { type: 'textarea' },
  'Privacy Impact': { type: 'textarea' },
  'Financial Impact': { type: 'textarea' },
  'Operational Impact': { type: 'textarea' },
  'Reputational Impact': { type: 'textarea' },
  'Potential Financial Exposure': { type: 'text' },
  'Potential Regulatory Exposure': { type: 'text' },
  'Records / Population Affected': { type: 'text' },
  'Estimated Records Affected': { type: 'number', min: 0 },
  'Third Parties Impacted': { type: 'text' },
  'Data Breach Potential': { type: 'text' },
  'Systems Compromised': { type: 'text' },
  'Customer / Client Impact': { type: 'textarea' },
  'Employee Impact': { type: 'textarea' },

  'Incident ID': { type: 'text' },
  'Reported By': { type: 'text' },
  'Mail Subject': { type: 'text' },
  'Threat Type': { type: 'text' },
  'Threat Actor': { type: 'text' },
  'Intent': { type: 'text' },
  'Forensics Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Evidence Collected': { type: 'textarea' },
  'Containment Actions': { type: 'textarea' },
  'Investigation Lead': { type: 'text' },
  'Investigation Start Date': { type: 'date' },
  'Investigation Status': { type: 'lookup', lookup: 'Investigation Status' },
  'Access Revocation / Modification Completed?': { type: 'lookup', lookup: 'Yes/No' },
  'Incident Closure Date': { type: 'date' },

  'Audit ID': { type: 'text' },
  'Audit Type': { type: 'lookup', lookup: 'Audit Type' },
  'Audit Name': { type: 'text' },
  'Audit Period': { type: 'text' },
  'Audit Start Date': { type: 'date' },
  'Audit End Date': { type: 'date' },
  'Auditor / Audit Firm': { type: 'text' },
  'Audit Scope': { type: 'textarea' },
  'Testing Area': { type: 'text' },
  'Sample Size': { type: 'number', min: 0 },
  'Sample Exceptions': { type: 'number', min: 0 },
  'Testing Result': { type: 'lookup', lookup: 'Testing Result' },
  'Finding Classification': { type: 'lookup', lookup: 'Finding Classification' },
  'Audit Finding Reference': { type: 'text' },
  'Management Response': { type: 'textarea' },

  'Regulatory Notification Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Regulatory Body': { type: 'text' },
  'Regulatory Notification Date': { type: 'date' },
  'Client Notification Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Client Notification Date': { type: 'date' },
  'Legal Counsel Engaged?': { type: 'lookup', lookup: 'Yes/No' },
  'Legal Review Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Breach Notification Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Notification Decision': { type: 'text' },
  'Notification Decision Date': { type: 'date' },
  'Notification Owner': { type: 'text' },
  'Data Destruction Certification Required?': { type: 'lookup', lookup: 'Yes/No' },

  'Corrective Action Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Corrective Action ID': { type: 'text' },
  'Corrective Action Description': { type: 'textarea' },
  'Root Cause Action': { type: 'textarea' },
  'Preventive Action': { type: 'textarea' },
  'Action Owner': { type: 'text' },
  'Action Owner Department': { type: 'department' },
  'Action Priority': { type: 'lookup', lookup: 'Action Priority' },
  'Target Completion Date': { type: 'date' },
  'Actual Completion Date': { type: 'date' },
  'Action Status': { type: 'lookup', lookup: 'Action Status' },
  'Evidence Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Evidence Submitted Date': { type: 'date' },
  'Evidence Validated?': { type: 'lookup', lookup: 'Yes/No' },
  'Validation Date': { type: 'date' },
  'Validation Owner': { type: 'text' },
  'Effectiveness Check Required?': { type: 'lookup', lookup: 'Yes/No' },
  'Effectiveness Check Date': { type: 'date' },
  'Effectiveness Result': { type: 'lookup', lookup: 'Effectiveness Result' },

  'Business Owner': { type: 'text' },
  'Management Approver': { type: 'text' },
  'Compliance Sign-off': { type: 'text' },
  'Compliance Sign-off Date': { type: 'date' },
  'Closure Criteria Met?': { type: 'lookup', lookup: 'Yes/No' },
  'Closure Approved By': { type: 'text' },
  'Closure Date': { type: 'date' },
  'Final Outcome': { type: 'textarea' },
  'Lessons Learned': { type: 'textarea' },
  'Residual Risk Accepted?': { type: 'lookup', lookup: 'Yes/No' },

  'Retention Period': { type: 'text' },
  'Record Retention Expiry Date': { type: 'date' },
  'Review Frequency': { type: 'lookup', lookup: 'Review Frequency' },
  'Next Review Date': { type: 'date' }
};

const SYSTEM_MANAGED_FIELDS = [
  'Record ID',
  'Date Logged',
  'Logged By',
  'Last Updated By',
  'Last Updated Date'
];

function normalizeRecordTypeKey_(value) {
  const text = cleanString_(value).toLowerCase();
  return Object.keys(RECORD_TYPE_GROUPS).find(function(key) {
    return key.toLowerCase() === text;
  }) || '';
}

function getSupportedRecordTypes_() {
  const lookup = getLookupValues_();
  const lookupTypes = lookup['Record Type'] || [];

  const configured = Object.keys(RECORD_TYPE_GROUPS);

  const result = configured.filter(function(type) {
    return lookupTypes.some(function(value) {
      return cleanString_(value).toLowerCase() === type.toLowerCase();
    });
  });

  return result.length ? result : configured;
}

function getApplicableFieldsForType_(recordType) {
  const key = normalizeRecordTypeKey_(recordType);

  if (!key) return [];

  const groupNames = RECORD_TYPE_GROUPS[key] || [];
  const fields = [];

  groupNames.forEach(function(groupName) {
    const group = RECORD_GROUP_DEFINITIONS[groupName];
    if (!group) return;

    group.fields.forEach(function(field) {
      if (fields.indexOf(field) === -1) {
        fields.push(field);
      }
    });
  });

  return fields;
}

function getNonApplicableFields_(recordType, headers) {
  const applicable = getApplicableFieldsForType_(recordType);

  return headers.filter(function(header) {
    if (!header) return false;
    if (SYSTEM_MANAGED_FIELDS.indexOf(header) !== -1) return false;
    return applicable.indexOf(header) === -1;
  });
}

function getRecordTypeConfig_() {
  rbacRequire_('read');

  const lookup = getLookupValues_();

  const supportedRecordTypes = getSupportedRecordTypes_();

  const groups = {};
  Object.keys(RECORD_GROUP_DEFINITIONS).forEach(function(groupName) {
    groups[groupName] = RECORD_GROUP_DEFINITIONS[groupName];
  });

  return {
    success: true,
    recordTypes: supportedRecordTypes,
    groups: groups,
    fieldDefinitions: RECORD_FIELD_DEFINITIONS,
    systemManagedFields: SYSTEM_MANAGED_FIELDS,
    defaultRecordType: supportedRecordTypes.indexOf('Audit Finding') !== -1
      ? 'Audit Finding'
      : (supportedRecordTypes[0] || ''),
    lookup: lookup,
    dependentLookups: getDependentLookupMap_()
  };
}

/***************************************************************
 * NEW RECORD CREATION
 ***************************************************************/

/*
 * Required fields for initial record creation.
 *
 * We keep this intentionally small.
 * The eventual Index.html wizard will collect more fields
 * based on Record Type.
 */
function getCreateRecordRequiredFields_() {
  return [
    'Record Type',
    'Date Identified',
    'Record Title',
    'Record Description',
    'Record Status',
    'Primary Compliance Framework'
  ];
}


/*
 * Lookup categories used during record creation.
 *
 * Only categories actually present in Lookup_Master are
 * enforced. This allows the Lookup_Master to be expanded later.
 */
function getCreateRecordLookupMap_() {
  return {
    'Record Type': 'Record Type',
    'Record Status': 'Record Status',
    'Primary Compliance Framework': 'Primary Compliance Framework',
    'Department': 'Department',
    'Sub Department': 'Sub Department',
    'Risk Category': 'Risk Category',
    'Control Frequency': 'Control Frequency',
    'Policy / Procedure Violated': 'Policy / Procedure Violated',
    'Regulatory / Contractual Obligation': 'Regulatory / Contractual Obligation',
    'Action Owner Department': 'Department',
    'Control Type': 'Control Type',
    'Control Effectiveness': 'Control Effectiveness',
    'Repeat Finding?': 'Repeat Finding?',
    'Risk Treatment': 'Risk Treatment',
    'Risk Acceptance Required?': 'Risk Acceptance Required?',
    'Inherent Risk Rating': 'Risk Rating',
    'Residual Risk Rating': 'Risk Rating',
    'Action Status': 'Action Status',
    'Action Priority': 'Action Priority',
    'Audit Type': 'Audit Type',
    'Finding Classification': 'Finding Classification',
    'Testing Result': 'Testing Result',
    'Investigation Status': 'Investigation Status',
    'Effectiveness Result': 'Effectiveness Result',
    'Review Frequency': 'Review Frequency'
  };
}


function normalizeLookupComparisonValues_(value) {
  let values;

  if (Array.isArray(value)) {
    values = value;
  } else {
    values = String(value === null || value === undefined ? '' : value)
      .split(',')
      .map(function(item) { return item.trim(); })
      .filter(function(item) { return item !== ''; });
  }

  return values
    .map(function(item) { return cleanString_(item).toLowerCase(); })
    .filter(function(item) { return item !== ''; })
    .sort();
}

function lookupValuesEqual_(oldValue, newValue) {
  const oldValues = normalizeLookupComparisonValues_(oldValue);
  const newValues = normalizeLookupComparisonValues_(newValue);

  if (oldValues.length !== newValues.length) return false;

  for (let i = 0; i < oldValues.length; i++) {
    if (oldValues[i] !== newValues[i]) return false;
  }

  return true;
}

/**
 * Returns true when a field is intentionally disabled by the
 * conditional dependency engine for the current payload.
 *
 * Disabled dependent fields are system-normalized to N/A and must
 * not be validated against Lookup_Master as if N/A were a user choice.
 */
function isConditionallyDisabledField_(field, payload) {
  if (!field || !payload || typeof payload !== 'object') return false;

  return CONDITIONAL_DEPENDENCY_RULES.some(function(rule) {
    if (!rule || !Array.isArray(rule.dependentFields)) return false;
    if (rule.dependentFields.indexOf(field) === -1) return false;

    const controllerValue = cleanString_(payload[rule.controller]).toLowerCase();
    return controllerValue !== 'yes';
  });
}

function validateLookupValueAgainstCurrentMaster_(field, value, lookupData, fieldCategoryMap) {
  const cleaned = cleanString_(value);
  if (!cleaned) return;

  const category = fieldCategoryMap[field];

  // Do not enforce categories that are not configured in Lookup_Master.
  if (!category || !lookupData.values[category] || !lookupData.map[category]) {
    return;
  }

  const submittedValues = Array.isArray(value)
    ? value
    : String(value)
        .split(',')
        .map(function(item) { return item.trim(); })
        .filter(function(item) { return item !== ''; });

  submittedValues.forEach(function(item) {
    const itemClean = cleanString_(item);
    if (!itemClean) return;

    const exists = Object.keys(lookupData.map[category]).some(function(key) {
      return field === 'Primary Compliance Framework'
        ? frameworkControlAliasKeyV18_(key) === frameworkControlAliasKeyV18_(itemClean)
        : cleanString_(key).toLowerCase() === itemClean.toLowerCase();
    });

    if (!exists) {
      throw new Error(
        'Invalid value for "' +
        field +
        '": "' +
        itemClean +
        '". Please select a valid lookup value.'
      );
    }
  });
}

function validateUpdateRecordLookups_(existingRecord, payload) {
  const lookupData = buildLookupMap_();
  const fieldCategoryMap = getCreateRecordLookupMap_();
  const current = existingRecord || {};
  const submitted = payload || {};

  Object.keys(fieldCategoryMap).forEach(function(field) {
    if (!Object.prototype.hasOwnProperty.call(submitted, field)) return;

    /*
     * A dependency-disabled field is intentionally stored as N/A.
     * N/A is a system sentinel, not a Lookup_Master selection.
     */
    if (isConditionallyDisabledField_(field, submitted)) return;

    const newValue = submitted[field];
    const oldValue = current[field] !== undefined ? current[field] : '';

    /*
     * Legacy value protection:
     * unchanged historical values are allowed to remain as-is.
     */
    if (lookupValuesEqual_(oldValue, newValue)) return;

    validateLookupValueAgainstCurrentMaster_(
      field,
      newValue,
      lookupData,
      fieldCategoryMap
    );
  });

  /* Enforce Department → Sub Department relationship when the
     department or sub department has actually changed. */
  const department = cleanString_(submitted['Department']);
  const subDepartment = cleanString_(submitted['Sub Department']);
  const oldDepartment = cleanString_(current['Department']);
  const oldSubDepartment = cleanString_(current['Sub Department']);

  const departmentChanged =
    !lookupValuesEqual_(oldDepartment, department);
  const subDepartmentChanged =
    !lookupValuesEqual_(oldSubDepartment, subDepartment);

  if (departmentChanged || subDepartmentChanged) {
    if (subDepartment) {
      if (!department) {
        throw new Error(
          'Department must be selected before selecting Sub Department.'
        );
      }

      const allowedSubDepartments =
        getDependentLookupValues_('Sub Department', department);

      if (
        !allowedSubDepartments.some(function(value) {
          return cleanString_(value).toLowerCase() ===
            subDepartment.toLowerCase();
        })
      ) {
        throw new Error(
          'Invalid Sub Department "' +
          subDepartment +
          '" for Department "' +
          department +
          '".'
        );
      }
    }
  }
}


function validateCreateRecordLookups_(payload) {
  const lookupData = buildLookupMap_();
  const values = lookupData.values;
  const map = lookupData.map;
  const fieldCategoryMap = getCreateRecordLookupMap_();

  Object.keys(fieldCategoryMap).forEach(function(field) {
    if (!Object.prototype.hasOwnProperty.call(payload, field)) {
      return;
    }

    /*
     * Dependency-disabled fields are written as N/A by the system.
     * Skip lookup validation because N/A is not a selectable lookup value.
     */
    if (isConditionallyDisabledField_(field, payload)) {
      return;
    }

    const rawValue = payload[field];

    if (
      rawValue === null ||
      rawValue === undefined ||
      cleanString_(rawValue) === ''
    ) {
      return;
    }

    const category = fieldCategoryMap[field];

    // Do not reject a field if the category has not yet
    // been created in Lookup_Master.
    if (!values[category] || !map[category]) {
      return;
    }

    let submittedValues;

    if (Array.isArray(rawValue)) {
      submittedValues = rawValue;
    } else {
      submittedValues =
        String(rawValue)
          .split(',')
          .map(function(item) {
            return item.trim();
          })
          .filter(function(item) {
            return item !== '';
          });
    }

    submittedValues.forEach(function(value) {
      const cleaned = cleanString_(value);

      if (!cleaned) return;

      if (!map[category][cleaned]) {
        throw new Error(
          'Invalid value for "' +
          field +
          '": "' +
          cleaned +
          '". Please select a valid lookup value.'
        );
      }
    });
  });

  /* Enforce Department → Sub Department relationship. */
  const department = cleanString_(payload['Department']);
  const subDepartment = cleanString_(payload['Sub Department']);

  if (subDepartment) {
    if (!department) {
      throw new Error(
        'Department must be selected before selecting Sub Department.'
      );
    }

    const allowedSubDepartments =
      getDependentLookupValues_('Sub Department', department);

    if (
      !allowedSubDepartments.some(function(value) {
        return cleanString_(value).toLowerCase() ===
          subDepartment.toLowerCase();
      })
    ) {
      throw new Error(
        'Invalid Sub Department "' +
        subDepartment +
        '" for Department "' +
        department +
        '".'
      );
    }
  }
}


/***************************************************************
 * CURRENT USER
 ***************************************************************/
function getCurrentUserEmail_() {
  let email = '';

  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (error) {
    email = '';
  }

  return email || 'System User';
}


/***************************************************************
 * TC-23 ENHANCEMENT - TYPE-SPECIFIC, IMMUTABLE RECORD IDS
 * New records: AUD-01-2026 / INC-01-2026 / RSK-01-2026 / SEC-01-2026
 * Existing IDs (such as 2026-0019) are NEVER rewritten.
 * A change of Record Type AFTER creation does NOT re-key references;
 * the original Record ID remains stable for Maker-Checker/history/evidence.
 * The final ID is allocated inside createComplianceRecord_'s script lock.
 ***************************************************************/
function recordIdPrefixForType_(recordType) {
  const type = normalizeRecordTypeKey_(recordType);
  const prefixes = {
    'Audit Finding': 'AUD',
    'Incident': 'INC',
    'Risk': 'RSK',
    'Security Finding': 'SEC', // reserved for a future configured type
    'Control Deficiency': 'CDF',
    'Compliance Issue': 'CIS',
    'Policy Exception': 'PEX',
    'Observation': 'OBS',
    'CAPA': 'CAP',
    'Client Finding': 'CLF',
    'Regulatory Finding': 'RGF'
  };
  const prefix = prefixes[type];
  if (!prefix) throw new Error('Unsupported Record Type for ID generation.');
  return prefix;
}

function recordIdYear_() {
  // Respect the Apps Script project timezone rather than a browser timezone.
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
}

function generateNextRecordId_(sheet, headers, recordType) {
  const index = headers.indexOf('Record ID');
  if (index === -1) throw new Error('Record ID column not found.');
  const prefix = recordIdPrefixForType_(recordType);
  const year = recordIdYear_();
  const matcher = new RegExp('^' + prefix + '-(\\d+)-' + year + '$', 'i');
  let highest = 0;

  // The master sheet contains active and soft-deleted records.
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, index + 1, lastRow - 1, 1).getDisplayValues();
    ids.forEach(function(row) {
      const match = String(row[0] || '').trim().match(matcher);
      if (!match) return;
      const seq = Number(match[1]);
      if (Number.isSafeInteger(seq) && seq > highest) highest = seq;
    });
  }

  // High-water marks prevent ID reuse if a record is later physically removed.
  // Writes happen ONLY after the new master row is flushed successfully.
  const propertyKey = 'DGS_RECORD_ID_SEQ_' + prefix + '_' + year;
  const stored = Number(PropertiesService.getScriptProperties().getProperty(propertyKey) || 0);
  if (Number.isSafeInteger(stored) && stored > highest) highest = stored;
  const next = highest + 1;
  return prefix + '-' + String(next).padStart(2, '0') + '-' + year;
}

function recordIdRememberCommitted_(recordId) {
  const match = String(recordId || '').match(/^(AUD|INC|RSK|SEC|CDF|CIS|PEX|OBS|CAP|CLF|RGF)-(\d+)-(\d{4})$/i);
  if (!match) return;
  const prefix = match[1].toUpperCase();
  const value = Number(match[2]);
  const year = match[3];
  const key = 'DGS_RECORD_ID_SEQ_' + prefix + '_' + year;
  const props = PropertiesService.getScriptProperties();
  const previous = Number(props.getProperty(key) || 0);
  if (Number.isSafeInteger(value) && value > previous) props.setProperty(key, String(value));
}

// Read-only preview: never reserves an ID, and may change before Save.
// The create operation allocates and commits the authoritative ID.
function getNextRecordIdPreview_(recordType) {
  rbacRequirePermission_('createRecords');
  const sheet = getMasterSheet_();
  const schema = validateMasterSchema_(sheet);
  return {
    recordId: generateNextRecordId_(sheet, schema.headers, recordType),
    estimate: true
  };
}

function getCurrentYear_() {
  return new Date().getFullYear();
}


/***************************************************************
 * RISK SCORE
 *
 * Score = Likelihood × Impact
 ***************************************************************/
function calculateRiskScore_(likelihood, impact) {
  const l = Number(likelihood);
  const i = Number(impact);

  if (!isFinite(l) || !isFinite(i)) {
    return '';
  }

  if (
    l < 1 ||
    l > 5 ||
    i < 1 ||
    i > 5
  ) {
    return '';
  }

  return l * i;
}


/***************************************************************
 * FIELD NORMALIZATION
 ***************************************************************/
function normalizeRecordFieldValue_(header, value) {
  if (
    value === null ||
    value === undefined ||
    cleanString_(value) === ''
  ) {
    return '';
  }

  const dateFields = {
    'Date Identified': true,
    'Risk Acceptance Date': true,
    'Investigation Start Date': true,
    'Incident Closure Date': true,
    'Audit Start Date': true,
    'Audit End Date': true,
    'Regulatory Notification Date': true,
    'Client Notification Date': true,
    'Notification Decision Date': true,
    'Target Completion Date': true,
    'Actual Completion Date': true,
    'Evidence Submitted Date': true,
    'Validation Date': true,
    'Effectiveness Check Date': true,
    'Compliance Sign-off Date': true,
    'Closure Date': true,
    'Record Retention Expiry Date': true,
    'Next Review Date': true
  };

  if (dateFields[header]) {
    const date = parseRecordDate_(value);
    return date || cleanString_(value);
  }

  const numericFields = {
    'Likelihood (1-5)': true,
    'Impact (1-5)': true,
    'Residual Likelihood (1-5)': true,
    'Residual Impact (1-5)': true,
    'Estimated Records Affected': true,
    'Sample Size': true,
    'Sample Exceptions': true
  };

  if (numericFields[header]) {
    const number = Number(value);
    return isFinite(number)
      ? number
      : cleanString_(value);
  }

  if (Array.isArray(value)) {
    return value.join(', ');
  }

  if (typeof value === 'object') {
    return JSON.stringify(value);
  }

  return cleanString_(value);
}


/***************************************************************
 * CONDITIONAL FIELD DEPENDENCY ENGINE
 *
 * The frontend controls visibility/editability, while the backend
 * remains authoritative and writes N/A for any dependent field
 * whose controller is not Yes.
 ***************************************************************/
const CONDITIONAL_DEPENDENCY_RULES = [
  {
    controller: 'Corrective Action Required?',
    dependentFields: [
      'Corrective Action ID',
      'Corrective Action Description',
      'Root Cause Action',
      'Preventive Action',
      'Action Owner',
      'Action Owner Department',
      'Action Priority',
      'Target Completion Date',
      'Actual Completion Date',
      'Action Status',
      'Evidence Required?'
    ]
  },
  {
    controller: 'Evidence Required?',
    dependentFields: [
      'Evidence Submitted Date',
      'Evidence Validated?',
      'Validation Date',
      'Validation Owner',
      'Effectiveness Check Required?',
      'Effectiveness Check Date',
      'Effectiveness Result'
    ]
  },
  {
    controller: 'Effectiveness Check Required?',
    dependentFields: [
      'Effectiveness Check Date',
      'Effectiveness Result'
    ]
  },
  {
    controller: 'Risk Acceptance Required?',
    dependentFields: [
      'Risk Acceptance Authority',
      'Risk Acceptance Date'
    ]
  },
  {
    controller: 'Forensics Required?',
    dependentFields: [
      'Evidence Collected',
      'Investigation Lead',
      'Investigation Start Date',
      'Investigation Status'
    ]
  },
  {
    controller: 'Regulatory Notification Required?',
    dependentFields: [
      'Regulatory Body',
      'Regulatory Notification Date'
    ]
  },
  {
    controller: 'Client Notification Required?',
    dependentFields: [
      'Client Notification Date'
    ]
  },
  {
    controller: 'Breach Notification Required?',
    dependentFields: [
      'Notification Decision',
      'Notification Decision Date',
      'Notification Owner'
    ]
  },
  {
    controller: 'Control Failure Identified',
    dependentFields: [
      'Control Failure Type'
    ]
  },
  {
    controller: 'Repeat Finding?',
    dependentFields: [
      'Previous Record ID'
    ]
  },
  {
    controller: 'Closure Criteria Met?',
    dependentFields: [
      'Closure Approved By',
      'Closure Date'
    ]
  }
];

function applyConditionalFieldDependencies_(payload) {
  if (!payload || typeof payload !== 'object') return;

  CONDITIONAL_DEPENDENCY_RULES.forEach(function(rule) {
    const controllerValue = cleanString_(payload[rule.controller]).toLowerCase();
    if (controllerValue === 'yes') return;

    rule.dependentFields.forEach(function(field) {
      payload[field] = 'N/A';
    });
  });
}

/* Backward-compatible wrapper retained for existing callers. */
function applyEvidenceRequiredDependency_(payload) {
  applyConditionalFieldDependencies_(payload);
}

/***************************************************************
 * CREATE COMPLIANCE RECORD
 ***************************************************************/
function createComplianceRecord_(payload) {
  rbacRequirePermission_('createRecords');

  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid record data received.');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const sheet = getMasterSheet_();
    const schema = validateMasterSchema_(sheet);
    const headers = schema.headers;

    const recordType = cleanString_(payload['Record Type']);
    const normalizedType = normalizeRecordTypeKey_(recordType);

    if (!normalizedType) {
      throw new Error(
        'Unsupported Record Type: ' + recordType +
        '. Please select a supported Record Type.'
      );
    }

    const applicableFields =
      getApplicableFieldsForType_(normalizedType);

    const requiredFields = getCreateRecordRequiredFields_();

    requiredFields.forEach(function(field) {
      if (
        !Object.prototype.hasOwnProperty.call(payload, field) ||
        cleanString_(payload[field]) === ''
      ) {
        throw new Error(field + ' is required.');
      }
    });

    /*
     * Record Type is mandatory and becomes immutable in the
     * lifecycle of the record. The backend is the final authority.
     */
    if (cleanString_(payload['Record Type']).toLowerCase() !== normalizedType.toLowerCase()) {
      throw new Error('Invalid Record Type.');
    }

    validateCreateRecordLookups_(payload);

    /* V55: for new Risk records, operational context is applicable only
       when Operations and an actual sub-department are selected. This is
       enforced server-side even if a client submits hidden/stale values. */
    if (normalizedType.toLowerCase() === 'risk' &&
        !(cleanString_(payload['Department']).toLowerCase() === 'operations' &&
          cleanString_(payload['Sub Department']) &&
          cleanString_(payload['Sub Department']).toUpperCase() !== 'N/A')) {
      ['Practice / Client', 'Process / Function', 'System / Application'].forEach(function(field) {
        payload[field] = 'N/A';
      });
    }

    /* Backend-enforced conditional field dependencies. */
    applyConditionalFieldDependencies_(payload);

    /*
     * Framework Control Master validation and enrichment.
     * If a control is selected, the backend is authoritative:
     * - validates that the control belongs to the selected framework
     * - writes the approved control description
     * - writes the approved framework domain
     */
    let selectedFrameworkControl = null;

    const selectedFramework =
      cleanString_(payload['Primary Compliance Framework']);

    const selectedControlId =
      cleanString_(payload['Requirement / Control ID']);

    if (selectedControlId) {
      const controlResult = getFrameworkControlById_(
        selectedFramework,
        selectedControlId
      );

      if (!controlResult.success) {
        throw new Error(controlResult.message);
      }

      selectedFrameworkControl = controlResult.control;
    }

    const recordId = generateNextRecordId_(sheet, headers, normalizedType);
    const now = new Date();
    const currentUser = getCurrentUserEmail_();

    const row = headers.map(function(header) {
      /* System-managed columns */
      if (header === 'Record ID') return recordId;
      if (header === 'Record Type') return normalizedType;
      if (header === 'Date Logged') return now;
      if (header === 'Logged By') return currentUser;
      if (header === 'Last Updated By') return currentUser;
      if (header === 'Last Updated Date') return now;

      /* Risk scores are always calculated by the backend. */
      if (header === 'Inherent Risk Score') {
        if (applicableFields.indexOf(header) === -1) return 'N/A';

        return calculateRiskScore_(
          payload['Likelihood (1-5)'],
          payload['Impact (1-5)']
        );
      }

      if (header === 'Residual Risk Score') {
        if (applicableFields.indexOf(header) === -1) return 'N/A';

        return calculateRiskScore_(
          payload['Residual Likelihood (1-5)'],
          payload['Residual Impact (1-5)']
        );
      }

      /*
       * Non-applicable column handling:
       * every field outside the selected Record Type's matrix
       * receives N/A so the database clearly records intentional
       * non-applicability rather than an ambiguous blank.
       */
      if (applicableFields.indexOf(header) === -1) {
        return 'N/A';
      }

      /* Framework-control values are system-controlled. */
      if (
        header === 'Requirement Description' &&
        selectedFrameworkControl
      ) {
        return selectedFrameworkControl.controlDescription;
      }

      if (
        header === 'Framework Domain' &&
        selectedFrameworkControl
      ) {
        return selectedFrameworkControl.domain || '';
      }

      if (Object.prototype.hasOwnProperty.call(payload, header)) {
        return normalizeRecordFieldValue_(header, payload[header]);
      }

      return '';
    });

    if (row.length !== EXPECTED_MASTER_COLUMNS) {
      throw new Error(
        'Generated row contains ' +
        row.length +
        ' columns. Expected ' +
        EXPECTED_MASTER_COLUMNS +
        '.'
      );
    }

    /* Final collision check */
    const existing = getRecordById_(recordId);

    if (existing && existing.success) {
      throw new Error(
        'Record ID collision detected: ' +
        recordId +
        '. Please retry.'
      );
    }

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    // Persist the high-water mark after the master row is durable.
    // A property-write failure must not falsely report the already-saved
    // record as unsuccessful: the next allocation scans the master anyway.
    try {
      recordIdRememberCommitted_(recordId);
    } catch (counterError) {
      console.error('Record ID high-water mark update failed:', counterError);
    }

    let historyLogged = true;
    try {
      appendRecordHistoryEntries_([
        {
          timestamp: now.toISOString(),
          recordId: recordId,
          action: 'CREATE',
          fieldName: 'Record',
          oldValue: '',
          newValue: 'Record created',
          changedBy: currentUser,
          changeReason: 'Initial record creation'
        },
        {timestamp:now.toISOString(), recordId:recordId, action:'CREATE',
          fieldName:'Record Type', oldValue:'', newValue:normalizedType,
          changedBy:currentUser, changeReason:'Initial record creation'}
      ]);
    } catch (historyError) {
      historyLogged = false;
      console.error('Initial record history logging failed:', historyError);
    }

    // TC-23.6: notify the designated manager of new HIPAA-related Incidents.
    // Record and history have already been persisted. Mail failures must NOT
    // turn a successfully saved record into an apparent failed submission.
    let incidentAlert = { applicable: false, status: 'NOT_APPLICABLE', sent: false };
    if (normalizedType === 'Incident' && /\bHIPAA\b/i.test(
      selectedFramework + ' ' + cleanString_(payload['Secondary Framework(s)'])
    )) {
      try {
        incidentAlert = tc23NotifyHipaaIncidentCreated_({
          recordId: recordId,
          recordType: normalizedType,
          framework: selectedFramework + ' ' + cleanString_(payload['Secondary Framework(s)']),
          risk: payload['Residual Risk Rating'] || payload['Inherent Risk Rating'],
          createdAt: now.toISOString()
        });
      } catch (alertError) {
        console.error('TC23 incident manager alert could not complete:', alertError);
        incidentAlert = { applicable: true, status: 'FAILED', sent: false };
      }
    }

    return {
      success: true,
      recordId: recordId,
      incidentAlert: incidentAlert,
      recordType: normalizedType,
      applicableFieldCount: applicableFields.length,
      nonApplicableFieldCount:
        Math.max(headers.length - SYSTEM_MANAGED_FIELDS.length - applicableFields.length, 0),
      timestamp: now.toISOString(),
      loggedBy: currentUser,
      historyLogged: historyLogged,
      message:
        'Compliance record created successfully.'
    };

  } catch (error) {
    console.error(
      'createComplianceRecord error:',
      error
    );

    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to create compliance record.'
    );

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/***************************************************************
 * NON-DESTRUCTIVE API TEST
 ***************************************************************/
function testCreateRecordAPI_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  return {
    success:
      typeof createComplianceRecord_ === 'function',

    functionAvailable:
      typeof createComplianceRecord_ === 'function',

    message:
      'Create record API is available. No record was created.'
  };
}


/***************************************************************
 * DESTRUCTIVE INSERT TEST
 *
 * DO NOT RUN unless you intentionally want a test row added.
 ***************************************************************/
function testCreateComplianceRecord_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const payload = {
    'Record Type': 'Observation',
    'Record Sub-Type': 'System Test',
    'Source / Trigger': 'Apps Script Test',
    'Date Identified': new Date(),
    'Record Title': 'System Test Record',
    'Record Description':
      'Temporary backend insertion test record.',
    'Business Unit': 'Test',
    'Department': 'Compliance',
    'Location': 'Chennai',
    'Practice / Client': 'Internal',
    'Process / Function': 'Compliance',
    'Record Status': 'Open',
    'Primary Compliance Framework': 'SOC 2',
    'Detailed Finding / Incident Summary':
      'Temporary backend insertion test record.',
    'Likelihood (1-5)': 2,
    'Impact (1-5)': 2,
    'Residual Likelihood (1-5)': 1,
    'Residual Impact (1-5)': 1
  };

  return createComplianceRecord_(payload);
}


/***************************************************************
 * SIMPLE TESTS
 ***************************************************************/
function testMasterSheet_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const sheet =
    getMasterSheet_();

  const schema =
    validateMasterSchema_(sheet);

  return {
    success: true,
    sheetName: sheet.getName(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn(),
    headerCount: schema.headers.length,
    recordCount:
      Math.max(sheet.getLastRow() - 1, 0)
  };
}

function testLookupSheet_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const sheet =
    getLookupSheet_();

  const values =
    getLookupValues_();

  const categories =
    Object.keys(values);

  return {
    success: true,
    sheetName: sheet.getName(),
    categories: categories.length,
    values: categories.reduce(function(total, category) {
      return total + values[category].length;
    }, 0),
    categoryNames: categories
  };
}

function testDashboardData_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  return getDashboardData_({});
}

function testFilterOptions_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  return getFilterOptions_();
}

function testLogo_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const dataUrl =
    getLogoDataUrl_();

  return {
    success: !!dataUrl,
    dataUrl: dataUrl
  };
}

function testRiskCalculation_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const inherent =
    calculateRiskScore_(4, 4);

  const residual =
    calculateRiskScore_(3, 3);

  return {
    success:
      inherent === 16 &&
      residual === 9,

    inherentTest:
      '4 x 4 = ' + inherent,

    residualTest:
      '3 x 3 = ' + residual
  };
}



/***************************************************************
 * TEST: FRAMEWORK CONTROL MASTER
 *
 * Non-destructive. Verifies the control library exists and
 * that the required fields are readable.
 ***************************************************************/
function testFrameworkControlMaster_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const sheet = getFrameworkControlSheet_();
  const master = readFrameworkControlMaster_();

  const frameworks = uniqueSorted_(
    master.controls.map(function(control) {
      return control.framework;
    })
  );

  const result = {
    success: master.controls.length > 0,
    sheetName: sheet.getName(),
    controlCount: master.controls.length,
    frameworkCount: frameworks.length,
    frameworks: frameworks,
    sample: master.controls.slice(0, 5)
  };

  console.log(JSON.stringify(result, null, 2));

  return result;
}


/***************************************************************
 * FINAL SYSTEM TEST
 *
 * 30 non-destructive checks.
 *
 * This does NOT create a record.
 ***************************************************************/
function testDynamicLookupIntegration_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const lookup = getLookupValues_();
  const config = getRecordTypeConfig_();
  const requiredCategories = [
    'Risk Category',
    'Control Frequency',
    'Policy / Procedure Violated',
    'Regulatory / Contractual Obligation'
  ];

  const checks = requiredCategories.map(function(category) {
    return {
      category: category,
      lookupValues: Array.isArray(lookup[category]) ? lookup[category] : [],
      fieldDefinition: config.fieldDefinitions[category] || null
    };
  });

  return {
    success: checks.every(function(item) {
      return item.lookupValues.length > 0 &&
        item.fieldDefinition &&
        item.fieldDefinition.type === 'lookup' &&
        item.fieldDefinition.lookup === item.category;
    }),
    checks: checks
  };
}


/***************************************************************
 * FINAL SYSTEM TEST
 ***************************************************************/
function runFinalSystemTest_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const results = [];
  let passed = 0;
  let failed = 0;

  function check(name, fn) {
    try {
      const value = fn();

      const success =
        typeof value === 'boolean'
          ? value
          : value &&
            typeof value.success === 'boolean'
            ? value.success
            : true;

      if (success) {
        passed++;

        results.push({
          check: name,
          status: 'PASSED'
        });

      } else {
        failed++;

        results.push({
          check: name,
          status: 'FAILED',
          details: value
        });
      }

    } catch (error) {
      failed++;

      results.push({
        check: name,
        status: 'FAILED',
        error:
          error && error.message
            ? error.message
            : String(error)
      });
    }
  }

  check(
    'Master sheet exists',
    function() {
      return !!getMasterSheet_();
    }
  );

  check(
    'Master sheet has exactly 148 headers',
    function() {
      return (
        getMasterSheet_().getLastColumn() ===
        EXPECTED_MASTER_COLUMNS
      );
    }
  );

  check(
    'Master data records available',
    function() {
      return (
        readMasterRecords_().records.length > 0
      );
    }
  );

  check(
    'Critical master headers present',
    function() {
      validateMasterSchema_(
        getMasterSheet_()
      );

      return true;
    }
  );

  check(
    'Sub Department header present',
    function() {
      const headers =
        readMasterRecords_().headers;

      return headers.indexOf('Sub Department') !== -1;
    }
  );

  check(
    'No duplicate master headers',
    function() {
      const headers =
        readMasterRecords_().headers
          .filter(function(header) {
            return !!header;
          });

      return (
        new Set(headers).size === headers.length
      );
    }
  );

  check(
    'Lookup sheet exists',
    function() {
      return !!getLookupSheet_();
    }
  );

  check(
    'Lookup headers present',
    function() {
      const sheet =
        getLookupSheet_();

      const headers =
        sheet
          .getRange(
            1,
            1,
            1,
            sheet.getLastColumn()
          )
          .getDisplayValues()[0]
          .map(cleanString_);

      return (
        headers.indexOf('Lookup Category') !== -1 &&
        headers.indexOf('Value') !== -1
      );
    }
  );

  check(
    'Record Type lookup values available',
    function() {
      const lookup =
        getLookupValues_();

      return (
        lookup['Record Type'] &&
        lookup['Record Type'].length > 0
      );
    }
  );

  check(
    'Record Status lookup values available',
    function() {
      const lookup =
        getLookupValues_();

      return (
        lookup['Record Status'] &&
        lookup['Record Status'].length > 0
      );
    }
  );

  check(
    'Compliance Framework lookup values available',
    function() {
      const lookup =
        getLookupValues_();

      return (
        lookup['Primary Compliance Framework'] &&
        lookup['Primary Compliance Framework'].length > 0
      );
    }
  );

  check(
    'Risk Rating lookup values available',
    function() {
      const lookup =
        getLookupValues_();

      return (
        lookup['Risk Rating'] &&
        lookup['Risk Rating'].length > 0
      );
    }
  );

  check(
    'Dashboard API works',
    function() {
      const result =
        getDashboardData_({});

      return (
        result &&
        result.success === true
      );
    }
  );

  check(
    'Dashboard record count available',
    function() {
      const result =
        getDashboardData_({});

      return (
        typeof result.recordCount === 'number'
      );
    }
  );

  check(
    'Dashboard KPI data available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.kpis &&
        typeof result.kpis.totalRecords === 'number' &&
        typeof result.kpis.openRecords === 'number' &&
        typeof result.kpis.highCriticalRisks === 'number' &&
        typeof result.kpis.overdueActions === 'number' &&
        typeof result.kpis.pendingEvidence === 'number' &&
        typeof result.kpis.closedRecords === 'number'
      );
    }
  );

  check(
    'Framework chart dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.frameworks)
      );
    }
  );

  check(
    'Status chart dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.statuses)
      );
    }
  );

  check(
    'Residual risk chart dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.risks)
      );
    }
  );

  check(
    'Finding classification dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.findings)
      );
    }
  );

  check(
    'Action status chart dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.actions)
      );
    }
  );

  check(
    'Open department chart dataset available',
    function() {
      const result =
        getDashboardData_({});

      return (
        result.charts &&
        Array.isArray(result.charts.departments)
      );
    }
  );

  check(
    'Watchlist available',
    function() {
      const result =
        getDashboardData_({});

      return Array.isArray(result.watchlist);
    }
  );

  check(
    'Dashboard generated timestamp available',
    function() {
      const result =
        getDashboardData_({});

      return !!result.generatedAt;
    }
  );

  check(
    'SOC 2 filter works',
    function() {
      return Array.isArray(
        getRecords_({
          primaryFramework: 'SOC 2'
        })
      );
    }
  );

  check(
    'Combined SOC 2 + High risk filter works',
    function() {
      return Array.isArray(
        getRecords_({
          primaryFramework: 'SOC 2',
          riskRating: 'High'
        })
      );
    }
  );

  check(
    'Record detail lookup works',
    function() {
      const records =
        readMasterRecords_().records;

      if (!records.length) {
        return true;
      }

      const id =
        cleanString_(
          records[0]['Record ID']
        );

      if (!id) {
        return false;
      }

      const result =
        getRecordById_(id);

      return (
        result &&
        result.success === true &&
        result.record &&
        cleanString_(
          result.record['Record ID']
        ) === id
      );
    }
  );

  check(
    'Filter options work',
    function() {
      const result =
        getFilterOptions_();

      return (
        result &&
        Array.isArray(result.primaryFramework) &&
        Array.isArray(result.recordType) &&
        Array.isArray(result.recordStatus) &&
        Array.isArray(result.riskRating) &&
        Array.isArray(result.department)
      );
    }
  );

  check(
    'Digiclarity logo works',
    function() {
      const result =
        getLogoDataUrl_();

      return !!result;
    }
  );

  check(
    'Configuration works',
    function() {
      const result =
        getConfig_();

      return (
        result.masterSheet === MASTER_SHEET &&
        result.lookupSheet === LOOKUP_SHEET &&
        result.frameworkControlSheet === FRAMEWORK_CONTROL_SHEET &&
        result.expectedMasterColumns ===
          EXPECTED_MASTER_COLUMNS
      );
    }
  );

  check(
    'Create record API + risk engine + record-type matrix works',
    function() {
      const config = getRecordTypeConfig_();
      return (
        typeof createComplianceRecord_ === 'function' &&
        calculateRiskScore_(4, 4) === 16 &&
        calculateRiskScore_(3, 3) === 9 &&
        config &&
        config.success === true &&
        Array.isArray(config.recordTypes) &&
        config.recordTypes.length >= 1 &&
        !!config.groups &&
        !!config.fieldDefinitions
      );
    }
  );

  const finalResult = {
    success: failed === 0,
    passed: passed,
    failed: failed,
    total: passed + failed,
    results: results,
    timestamp: new Date().toISOString()
  };

  console.log(
    '=============================================='
  );

  console.log(
    'DIGICLARITY COMPLIANCE DASHBOARD SYSTEM TEST'
  );

  console.log(
    '=============================================='
  );

  results.forEach(function(result, index) {
    console.log(
      (index + 1) +
      '. ' +
      result.status +
      ' - ' +
      result.check
    );
  });

  console.log(
    '----------------------------------------------'
  );

  console.log(
    'PASSED: ' + passed
  );

  console.log(
    'FAILED: ' + failed
  );

  console.log(
    'TOTAL: ' + (passed + failed)
  );

  if (failed === 0) {
    console.log(
      '🎉 FINAL RESULT: ALL SYSTEMS PASSED'
    );
  } else {
    console.log(
      '⚠️ FINAL RESULT: CHECK FAILED TESTS'
    );
  }

  console.log(
    '=============================================='
  );

  return finalResult;
}


/***************************************************************
 * TEST: CROSS-FRAMEWORK CONTROL MAPPING
 *
 * Non-destructive test.
 * Verifies that the Common Control Area engine can return
 * controls from multiple frameworks for the same area.
 ***************************************************************/
function testCrossFrameworkControlMapping_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const areasResult = getCommonControlAreas_();

  if (!areasResult || areasResult.success !== true) {
    throw new Error('Unable to load Common Control Areas.');
  }

  const targetArea = 'Access Control';
  const result = getControlsByCommonControlArea_(targetArea);

  if (!result || result.success !== true) {
    throw new Error(
      'Unable to load controls for Common Control Area: ' + targetArea
    );
  }

  const controls = Array.isArray(result.controls)
    ? result.controls
    : [];

  const frameworks = uniqueSorted_(
    controls.map(function(control) {
      return control.framework;
    })
  );

  const hasHIPAA = frameworks.some(function(framework) {
    return String(framework).toLowerCase() === 'hipaa';
  });

  const hasSOC2 = frameworks.some(function(framework) {
    return String(framework).toLowerCase() === 'soc 2';
  });

  const finalResult = {
    success:
      areasResult.areas.length > 0 &&
      controls.length > 0 &&
      hasHIPAA &&
      hasSOC2,
    commonControlAreaCount: areasResult.areas.length,
    targetArea: targetArea,
    matchedControlCount: controls.length,
    frameworks: frameworks,
    controls: controls.map(function(control) {
      return {
        framework: control.framework,
        controlId: control.controlId,
        controlName: control.controlName,
        domain: control.domain || '',
        commonControlArea: control.commonControlArea || '',
        soc2CCMapping: control.soc2CCMapping || '',
        hipaaMapping: control.hipaaMapping || ''
      };
    }),
    timestamp: new Date().toISOString()
  };

  console.log(
    '=============================================='
  );
  console.log(
    'CROSS-FRAMEWORK CONTROL MAPPING TEST'
  );
  console.log(
    '=============================================='
  );
  console.log(JSON.stringify(finalResult, null, 2));

  if (!finalResult.success) {
    console.log(
      '⚠️ CROSS-FRAMEWORK TEST FAILED'
    );
  } else {
    console.log(
      '🎉 CROSS-FRAMEWORK TEST PASSED'
    );
  }

  console.log(
    '=============================================='
  );

  return finalResult;
}
/***************************************************************
 * MAKER-CHECKER V1
 *
 * Existing record edits do NOT update Compliance_Master_Register
 * immediately. They are written to Change_Requests and require an
 * assigned checker to approve them.
 *
 * V1 scope:
 *   - Existing-record changes only
 *   - One pending request per Record ID
 *   - Maker cannot approve/reject own request
 *   - Assigned approver / backup approver only
 *   - Email notification to approver
 *   - Concurrency check before approval
 *   - Reuses updateComplianceRecord_() only AFTER approval
 ***************************************************************/

const CHANGE_REQUEST_HEADERS = [
  'Request ID',
  'Record ID',
  'Record Type',
  'Record Title',
  'Request Type',
  'Requested By',
  'Requested Date',
  'Change Reason',
  'Changed Field Count',
  'Old Values JSON',
  'Proposed Values JSON',
  'Status',
  'Approver Email',
  'Backup Approver Email',
  'Reviewed By',
  'Reviewed Date',
  'Reviewer Comment',
  'Applied Date',
  'Approval Token'
];

const APPROVAL_MATRIX_HEADERS = [
  'Rule ID',
  'Department',
  'Request Type',
  'Approver Email',
  'Backup Approver Email',
  'Active'
];

function setupMakerCheckerSheets_() {
  rbacRequire_('admin');

  const changeSheet = getChangeRequestSheet_();
  const matrixSheet = getApprovalMatrixSheet_();

  if (matrixSheet.getLastRow() < 2) {
    matrixSheet.appendRow([
      'DEFAULT-RECORD-CHANGE',
      'ALL',
      'RECORD_CHANGE',
      'ENTER_APPROVER_EMAIL',
      '',
      'Yes'
    ]);
  }

  changeSheet.setFrozenRows(1);
  matrixSheet.setFrozenRows(1);

  try {
    changeSheet.autoResizeColumns(1, CHANGE_REQUEST_HEADERS.length);
    matrixSheet.autoResizeColumns(1, APPROVAL_MATRIX_HEADERS.length);
  } catch (ignore) {}

  return {
    success: true,
    changeRequestSheet: CHANGE_REQUEST_SHEET,
    approvalMatrixSheet: APPROVAL_MATRIX_SHEET,
    message:
      'Maker-Checker sheets are ready. Open Approval_Matrix and replace ENTER_APPROVER_EMAIL with the checker email before submitting an edit.'
  };
}

function getChangeRequestSheet_() {
  return getOrCreateSheetWithHeaders_(
    CHANGE_REQUEST_SHEET,
    CHANGE_REQUEST_HEADERS
  );
}

function getApprovalMatrixSheet_() {
  return getOrCreateSheetWithHeaders_(
    APPROVAL_MATRIX_SHEET,
    APPROVAL_MATRIX_HEADERS
  );
}

function generateNextChangeRequestId_(sheet) {
  const year = String(getCurrentYear_());
  const lastRow = sheet.getLastRow();
  let highest = 0;

  if (lastRow >= 2) {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0]
      .map(cleanString_);

    const idIndex = headers.indexOf('Request ID');

    if (idIndex !== -1) {
      const values = sheet
        .getRange(2, idIndex + 1, lastRow - 1, 1)
        .getDisplayValues();

      values.forEach(function(row) {
        const id = cleanString_(row[0]);
        const match = id.match(
          new RegExp('^CR-' + year + '-(\\d+)$')
        );

        if (!match) return;

        const number = parseInt(match[1], 10);
        if (!isNaN(number) && number > highest) {
          highest = number;
        }
      });
    }
  }

  return 'CR-' + year + '-' + String(highest + 1).padStart(4, '0');
}

function getActualCurrentUserEmail_() {
  let email = '';

  try {
    email = cleanString_(Session.getActiveUser().getEmail());
  } catch (ignore) {
    email = '';
  }

  return email;
}

function normalizeEmail_(value) {
  return cleanString_(value).toLowerCase();
}

function isActiveApprovalRule_(value) {
  const normalized = cleanString_(value).toLowerCase();
  return normalized === 'yes' ||
    normalized === 'true' ||
    normalized === 'active' ||
    normalized === '1';
}

function getApprovalRuleForRecord_(record) {
  const sheet = getApprovalMatrixSheet_();

  if (sheet.getLastRow() < 2) {
    setupMakerCheckerSheets_();
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'Approval_Matrix is empty. Run setupMakerCheckerSheets_() and configure an approver.'
    );
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(cleanString_);

  const rows = sheet
    .getRange(2, 1, lastRow - 1, headers.length)
    .getDisplayValues();

  const department = cleanString_(record['Department']);
  const requestType = 'RECORD_CHANGE';

  let exactMatch = null;
  let defaultMatch = null;

  rows.forEach(function(row) {
    const item = {};
    headers.forEach(function(header, index) {
      item[header] = row[index] !== undefined ? row[index] : '';
    });

    if (!isActiveApprovalRule_(item['Active'])) return;

    const ruleRequestType = cleanString_(item['Request Type']).toUpperCase();
    if (ruleRequestType && ruleRequestType !== requestType) return;

    const ruleDepartment = cleanString_(item['Department']);

    if (
      department &&
      ruleDepartment &&
      ruleDepartment.toLowerCase() === department.toLowerCase()
    ) {
      if (!exactMatch) exactMatch = item;
      return;
    }

    if (
      !ruleDepartment ||
      ruleDepartment.toUpperCase() === 'ALL' ||
      ruleDepartment === '*'
    ) {
      if (!defaultMatch) defaultMatch = item;
    }
  });

  const selected = exactMatch || defaultMatch;

  if (!selected) {
    throw new Error(
      'No active RECORD_CHANGE approval rule is configured for Department "' +
      (department || 'N/A') +
      '". Update Approval_Matrix.'
    );
  }

  const approverEmail = cleanString_(selected['Approver Email']);
  const backupApproverEmail = cleanString_(selected['Backup Approver Email']);

  if (
    !approverEmail ||
    approverEmail.toUpperCase() === 'ENTER_APPROVER_EMAIL' ||
    approverEmail.indexOf('@') === -1
  ) {
    throw new Error(
      'Approval_Matrix is not configured. Replace ENTER_APPROVER_EMAIL with the manager/checker email.'
    );
  }

  return {
    ruleId: cleanString_(selected['Rule ID']),
    approverEmail: approverEmail,
    backupApproverEmail: backupApproverEmail
  };
}

function getPendingChangeRequestForRecord_(recordId) {
  const sheet = getSpreadsheet_().getSheetByName(CHANGE_REQUEST_SHEET);

  if (!sheet || sheet.getLastRow() < 2) return null;

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(cleanString_);

  const recordIndex = headers.indexOf('Record ID');
  const statusIndex = headers.indexOf('Status');

  if (recordIndex === -1 || statusIndex === -1) return null;

  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, headers.length)
    .getDisplayValues();

  for (let i = 0; i < rows.length; i++) {
    const rowRecordId = cleanString_(rows[i][recordIndex]);
    const status = cleanString_(rows[i][statusIndex]).toLowerCase();

    if (
      rowRecordId === cleanString_(recordId) &&
      status.indexOf('pending') !== -1
    ) {
      const item = {};
      headers.forEach(function(header, index) {
        item[header] = rows[i][index] !== undefined ? rows[i][index] : '';
      });
      item._rowNumber = i + 2;
      return item;
    }
  }

  return null;
}


/**
 * Normalize values before Maker-Checker comparison.
 * Prevents formatting-only differences (especially dates) from
 * being counted as actual record changes.
 */
function normalizeMakerCheckerComparisonValue_(field, value) {
  if (value === null || value === undefined) return '';

  const definition =
    typeof RECORD_FIELD_DEFINITIONS !== 'undefined' &&
    RECORD_FIELD_DEFINITIONS[field]
      ? RECORD_FIELD_DEFINITIONS[field]
      : null;

  const type = definition && definition.type
    ? cleanString_(definition.type).toLowerCase()
    : '';

  const raw = cleanString_(value);

  if (!raw) return '';

  if (type === 'date') {
    return normalizeMakerCheckerDate_(value);
  }

  if (type === 'number' || type === 'calculated') {
    const numberValue = Number(raw);
    return isNaN(numberValue) ? raw : String(numberValue);
  }

  if (type === 'multilookup') {
    return raw
      .split(',')
      .map(function(item) { return cleanString_(item).toLowerCase(); })
      .filter(function(item) { return item !== ''; })
      .sort()
      .join(',');
  }

  return raw;
}


/**
 * Converts common dashboard / Google Sheets date representations
 * to YYYY-MM-DD exclusively for equality comparison.
 */
function normalizeMakerCheckerDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const raw = cleanString_(value);
  if (!raw || raw.toUpperCase() === 'N/A') return raw;

  // Already ISO date: 2026-09-05 or 2026-09-05T...
  let match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);
  if (match) {
    return [
      match[1],
      String(match[2]).padStart(2, '0'),
      String(match[3]).padStart(2, '0')
    ].join('-');
  }

  // dd-MMM-yyyy, e.g. 05-Sep-2026
  match = raw.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (match) {
    const months = {
      jan:'01', feb:'02', mar:'03', apr:'04',
      may:'05', jun:'06', jul:'07', aug:'08',
      sep:'09', oct:'10', nov:'11', dec:'12'
    };
    const month = months[String(match[2]).toLowerCase()];
    if (month) {
      return [
        match[3],
        month,
        String(match[1]).padStart(2, '0')
      ].join('-');
    }
  }

  // dd/MM/yyyy
  match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    return [
      match[3],
      String(match[2]).padStart(2, '0'),
      String(match[1]).padStart(2, '0')
    ].join('-');
  }

  return raw;
}


function makerCheckerValuesEqual_(field, oldValue, newValue) {
  return normalizeMakerCheckerComparisonValue_(field, oldValue) ===
    normalizeMakerCheckerComparisonValue_(field, newValue);
}


function buildChangeRequestComparison_(recordId, payload) {
  const result = getRecordById_(recordId);

  if (!result || !result.success || !result.record) {
    throw new Error(
      result && result.message
        ? result.message
        : 'Unable to load the existing record.'
    );
  }

  if (result.isDeleted) {
    throw new Error(
      'Record ' + recordId +
      ' is deleted. Restore it from Recovery before requesting a change.'
    );
  }

  const existingRecord = result.record;
  const existingType = normalizeRecordTypeKey_(existingRecord['Record Type']);

  // Keep the original invalid type in the review snapshot; never guess its replacement.
  const targetType = normalizeRecordTypeKey_(Object.prototype.hasOwnProperty.call(payload, 'Record Type')
    ? payload['Record Type'] : existingType);
  if (!targetType) throw new Error('Select a supported Record Type.');
  const applicableFields = getApplicableFieldsForType_(targetType);
  const merged = buildMergedUpdatePayload_(
    existingRecord,
    payload,
    applicableFields,
    Object.keys(existingRecord)
  );

  getCreateRecordRequiredFields_().forEach(function(field) {
    if (cleanString_(merged[field]) === '') {
      throw new Error(field + ' is required.');
    }
  });

  validateUpdateRecordLookups_(existingRecord, merged);
  applyConditionalFieldDependencies_(merged);

  const selectedFramework = cleanString_(
    merged['Primary Compliance Framework']
  );
  const selectedControlId = cleanString_(
    merged['Requirement / Control ID']
  );

  if (selectedFramework && selectedControlId) {
    const controlResult = getFrameworkControlById_(
      selectedFramework,
      selectedControlId
    );

    if (
      !controlResult ||
      !controlResult.success ||
      !controlResult.control
    ) {
      throw new Error(
        controlResult && controlResult.message
          ? controlResult.message
          : 'Requirement / Control ID "' + selectedControlId +
            '" is not valid for framework "' + selectedFramework + '".'
      );
    }

    const frameworkControl = controlResult.control;

    merged['Requirement Description'] =
      frameworkControl.controlDescription || '';

    merged['Framework Domain'] =
      frameworkControl.domain || '';
  }

  // Always snapshot the source type to detect a reclassification while review is pending.
  const oldValues = {'Record Type': existingRecord['Record Type']};
  const proposedValues = {};
  const changes = [];

  ['Record Type'].concat(applicableFields.filter(function(field) { return field !== 'Record Type'; })).forEach(function(field) {
    if (SYSTEM_MANAGED_FIELDS.indexOf(field) !== -1) return;

    const oldValue =
      existingRecord[field] === null ||
      existingRecord[field] === undefined
        ? ''
        : String(existingRecord[field]);

    const newValue =
      merged[field] === null ||
      merged[field] === undefined
        ? ''
        : String(merged[field]);

    if (!makerCheckerValuesEqual_(field, oldValue, newValue)) {
      oldValues[field] = oldValue;
      proposedValues[field] = newValue;

      changes.push({
        field: field,
        oldValue: oldValue,
        newValue: newValue
      });
    }
  });

  return {
    existingRecord: existingRecord,
    mergedPayload: merged,
    oldValues: oldValues,
    proposedValues: proposedValues,
    changes: changes
  };
}

function submitRecordChangeRequest_(recordId, payload, changeReason) {
  rbacRequirePermission_('editRecords');

  const target = cleanString_(recordId);
  const reason = cleanString_(changeReason);

  if (!target) throw new Error('Record ID is required.');

  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid record data received.');
  }

  if (!reason) {
    throw new Error(
      'Change reason is required before submitting a record change for approval.'
    );
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const existingPending = getPendingChangeRequestForRecord_(target);

    if (existingPending) {
      throw new Error(
        'Record ' + target +
        ' already has a pending change request: ' +
        cleanString_(existingPending['Request ID']) +
        '. Approve or send back the existing request before submitting another change.'
      );
    }

    const comparison = buildChangeRequestComparison_(target, payload);

    if (!comparison.changes.length) {
      throw new Error(
        'No field changes were detected. Change at least one field before submitting for approval.'
      );
    }

    const makerEmail = getActualCurrentUserEmail_();

    if (!makerEmail) {
      throw new Error(
        'Unable to identify the logged-in user email. Maker-Checker requires authenticated user identity.'
      );
    }

    const approvalRule =
      getApprovalRuleForRecord_(comparison.existingRecord);

    if (
      normalizeEmail_(makerEmail) ===
      normalizeEmail_(approvalRule.approverEmail)
    ) {
      throw new Error(
        'Maker and checker cannot be the same user. Configure a different approver in Approval_Matrix.'
      );
    }

    const sheet = getChangeRequestSheet_();
    const requestId = generateNextChangeRequestId_(sheet);
    const requestedDate = new Date().toISOString();
    const approvalToken = Utilities.getUuid();

    const rowObject = {
      'Request ID': requestId,
      'Record ID': target,
      'Record Type': cleanString_(
        comparison.mergedPayload['Record Type']
      ),
      'Record Title': cleanString_(
        comparison.existingRecord['Record Title']
      ),
      'Request Type': 'RECORD_CHANGE',
      'Requested By': makerEmail,
      'Requested Date': requestedDate,
      'Change Reason': reason,
      'Changed Field Count': comparison.changes.length,
      'Old Values JSON': JSON.stringify(comparison.oldValues),
      'Proposed Values JSON': JSON.stringify(comparison.proposedValues),
      'Status': 'Pending Approval',
      'Approver Email': approvalRule.approverEmail,
      'Backup Approver Email': approvalRule.backupApproverEmail,
      'Reviewed By': '',
      'Reviewed Date': '',
      'Reviewer Comment': '',
      'Applied Date': '',
      'Approval Token': approvalToken
    };

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0]
      .map(cleanString_);

    const row = headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(rowObject, header)
        ? rowObject[header]
        : '';
    });

    sheet.appendRow(row);

    try {
      appendRecordHistoryEntries_([{
        timestamp: requestedDate,
        recordId: target,
        action: 'CHANGE REQUESTED',
        fieldName: 'Maker-Checker Request',
        oldValue: '',
        newValue: requestId,
        changedBy: makerEmail,
        changeReason: reason
      }]);
    } catch (historyError) {
      console.error(
        'Unable to write CHANGE REQUESTED history:',
        historyError
      );
    }

    let emailSent = false;
    let emailWarning = '';

    try {
      sendChangeRequestApprovalEmail_(rowObject, comparison.changes);
      emailSent = true;
    } catch (emailError) {
      emailWarning =
        emailError && emailError.message
          ? emailError.message
          : String(emailError);

      console.error(
        'Maker-Checker approval email failed:',
        emailError
      );
    }

    return {
      success: true,
      requestId: requestId,
      recordId: target,
      status: 'Pending Approval',
      changedFieldCount: comparison.changes.length,
      approverEmail: approvalRule.approverEmail,
      emailSent: emailSent,
      emailWarning: emailWarning,
      message:
        'Change request ' + requestId +
        ' submitted for approval. The master record has not been changed.'
    };

  } catch (error) {
    console.error('submitRecordChangeRequest error:', error);
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to submit the change request.'
    );
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Approval SLA is informational: no automatic expiry or rejection.
const APPROVAL_DASHBOARD_URL = 'https://sites.google.com/digiclarity.in/dgs-compliance-dashboard/dgs-cd';
const APPROVAL_REVIEW_HOURS = 48;

function approvalEmailEscape_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(c) {
    return {'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c];
  });
}

function buildChangeRequestApprovalEmail_(request, changes) {
  const submitted = new Date(request['Requested Date']);
  if (isNaN(submitted.getTime())) throw new Error('Invalid change request submission date.');
  const deadline = new Date(submitted.getTime() + APPROVAL_REVIEW_HOURS * 60 * 60 * 1000);
  const format = function(date) {
    return Utilities.formatDate(date, 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a') + ' IST';
  };
  const compact = function(value, limit) {
    const text = cleanString_(value).replace(/\s+/g, ' ');
    return text.length > limit ? text.slice(0, limit - 1) + '…' : text;
  };
  const e = approvalEmailEscape_;
  const visibleChanges = (changes || []).filter(function(change) {
    return cleanString_(change.field || change.fieldName);
  });
  const snapshots = visibleChanges.slice(0, 3).map(function(change) {
    return {
      field: compact(change.field || change.fieldName, 48),
      before: compact(change.oldValue, 90) || 'Blank',
      after: compact(change.newValue, 90) || 'Blank'
    };
  });
  const remaining = Math.max(visibleChanges.length - snapshots.length, 0);
  const summary = snapshots.length
    ? snapshots.map(function(item) {return item.field + ': ' + item.before + ' → ' + item.after;}).join('\n') +
      (remaining ? '\n+' + remaining + ' more change' + (remaining === 1 ? '' : 's') + ' in the dashboard.' : '')
    : 'Review the proposed changes in the dashboard.';
  const snapshotHtml = snapshots.length
    ? snapshots.map(function(item) {
        return '<div style="margin:12px 0;font-size:13px;line-height:20px;overflow-wrap:anywhere">' +
          '<strong>' + e(item.field) + '</strong><br><span style="color:#647089">' +
          e(item.before) + '</span> &rarr; <span style="color:#142440">' + e(item.after) + '</span></div>';
      }).join('') + (remaining ? '<div style="font-size:12px;color:#647089">+' + remaining + ' more changes in the dashboard.</div>' : '')
    : '<p>Review the proposed changes in the dashboard.</p>';
  const title = compact(request['Record Title'] || request['Record ID'] || 'Compliance record', 140);
  const submittedText = format(submitted);
  const deadlineText = format(deadline);
  const body = 'DIGICLARITY | Change Approval\n\nAction required: A record change is awaiting your approval.\n\n' +
    'Record: ' + title + '\n' + summary + '\nSubmitted on: ' + submittedText +
    '\nApproval deadline: ' + deadlineText + ' (48 hours after submission)\n\n' +
    'Review the full comparison in Dashboard > Approvals.\n' + APPROVAL_DASHBOARD_URL;
  const html = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:#f2f4f8;font-family:Arial,Helvetica,sans-serif;color:#202a40">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f2f4f8"><tr><td align="center" style="padding:28px 12px">' +
    '<table role="presentation" width="600" cellspacing="0" cellpadding="0" style="width:100%;max-width:600px;background:#ffffff;border:1px solid #e4e8f0;border-radius:16px;overflow:hidden">' +
    '<tr><td style="background:#142440;padding:30px 32px;text-align:center">' +
    '<div style="font-size:26px;font-weight:bold;letter-spacing:2px;color:#ffffff">DIGI<span style="color:#f05b9c">CLARITY</span></div>' +
    '<div style="margin-top:9px;color:#c7d4eb;font-size:12px;letter-spacing:2px">COMPLIANCE DASHBOARD</div></td></tr>' +
    '<tr><td style="padding:30px 32px">' +
    '<span style="background:#fce8f1;color:#a92161;font-size:11px;font-weight:bold;letter-spacing:1px;padding:7px 10px;border-radius:4px">ACTION REQUIRED</span>' +
    '<h1 style="font-size:24px;line-height:32px;margin:22px 0 10px;color:#142440">A change is ready for review</h1>' +
    '<p style="font-size:14px;line-height:23px;color:#647089;margin:0 0 24px">Please review the proposed record change and submit your decision in the dashboard.</p>' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr><td style="padding:20px;background:#f7f9fc;border-left:4px solid #297ac8">' +
    '<div style="font-size:11px;font-weight:bold;letter-spacing:1px;color:#297ac8;margin-bottom:10px">CHANGE SNAPSHOT</div>' +
    '<div style="font-size:16px;line-height:24px;font-weight:bold;overflow-wrap:anywhere">' + e(title) + '</div>' +
    snapshotHtml +
    '<div style="font-size:12px;line-height:20px;color:#647089"><strong>Submitted on</strong><br>' + e(submittedText) + '</div></td></tr></table>' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:18px"><tr><td style="padding:17px 20px;background:#fff0f6;border-left:4px solid #dc3983">' +
    '<div style="font-size:12px;font-weight:bold;color:#a92161">APPROVAL DEADLINE · 48 HOURS</div>' +
    '<div style="font-size:15px;font-weight:bold;margin-top:7px;color:#142440">' + e(deadlineText) + '</div></td></tr></table>' +
    '<table role="presentation" cellspacing="0" cellpadding="0" align="center" style="margin:28px auto 14px"><tr><td bgcolor="#2469ad" style="border-radius:7px;text-align:center">' +
    '<a href="' + e(APPROVAL_DASHBOARD_URL) + '" target="_blank" style="display:inline-block;border:16px solid #2469ad;border-left-width:24px;border-right-width:24px;border-radius:7px;color:#ffffff;font-size:13px;font-weight:bold;text-decoration:none">REVIEW CHANGE REQUEST</a>' +
    '</td></tr></table><p style="text-align:center;font-size:12px;line-height:19px;color:#647089;margin:0">Open the Approvals tab to view the full comparison.</p>' +
    '<p style="font-size:12px;line-height:20px;color:#647089;margin:28px 0 0">Thank you,<br><strong style="color:#142440">Global Compliance</strong></p></td></tr>' +
    '</table></td></tr></table></body></html>';
  return {subject: 'Approval Required | ' + cleanString_(request['Request ID']) + ' | ' + cleanString_(request['Record ID']), body: body, htmlBody: html};
}

function sendChangeRequestApprovalEmail_(request, changes) {
  const approver = cleanString_(request['Approver Email']);
  if (!approver) return;
  const message = buildChangeRequestApprovalEmail_(request, changes);
  message.to = approver;
  const backup = cleanString_(request['Backup Approver Email']);
  if (backup) message.cc = backup;
  message.name = 'Digiclarity Compliance Dashboard';
  MailApp.sendEmail(message);
}

function readChangeRequestRows_() {
  const sheet = getSpreadsheet_().getSheetByName(CHANGE_REQUEST_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      sheet: sheet,
      headers: CHANGE_REQUEST_HEADERS.slice(),
      rows: []
    };
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(cleanString_);

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, headers.length)
    .getDisplayValues();

  const rows = values.map(function(row, index) {
    const item = {};

    headers.forEach(function(header, columnIndex) {
      item[header] =
        row[columnIndex] !== undefined
          ? row[columnIndex]
          : '';
    });

    item._rowNumber = index + 2;
    return item;
  });

  return {
    sheet: sheet,
    headers: headers,
    rows: rows
  };
}

function getChangeRequestById_(requestId) {
  const target = cleanString_(requestId);

  if (!target) return null;

  const data = readChangeRequestRows_();

  for (let i = 0; i < data.rows.length; i++) {
    if (cleanString_(data.rows[i]['Request ID']) === target) {
      return {
        sheet: data.sheet,
        headers: data.headers,
        item: data.rows[i]
      };
    }
  }

  return null;
}

function canCurrentUserReviewRequest_(request, currentEmail) {
  const current = normalizeEmail_(currentEmail);
  const primary = normalizeEmail_(request['Approver Email']);
  const backup = normalizeEmail_(request['Backup Approver Email']);
  const access = rbacRequire_('read');
  return Boolean(access.permissions.reviewChanges && current &&
    (current === primary || (backup && current === backup)));
}

function getChangeRequestsForCurrentUser_() {
  rbacRequireAnyPermission_(['reviewChanges','viewApprovalHistory']);

  const data = readChangeRequestRows_();
  const currentEmail = getActualCurrentUserEmail_();

  if (!currentEmail) {
    throw new Error(
      'Unable to identify the logged-in user email. Approval access requires authenticated user identity.'
    );
  }

  const current = normalizeEmail_(currentEmail);

  const requests = data.rows
    .sort(function(a, b) {
      return String(b['Requested Date'] || '')
        .localeCompare(String(a['Requested Date'] || ''));
    })
    .map(function(item) {
      const copy = Object.assign({}, item);
      delete copy._rowNumber;
      delete copy['Approval Token'];
      return copy;
    });

  return {
    success: true,
    currentUser: currentEmail,
    requests: requests
  };
}

function parseJsonObjectSafe_(value) {
  try {
    const parsed = value ? JSON.parse(value) : {};
    return parsed && typeof parsed === 'object'
      ? parsed
      : {};
  } catch (ignore) {
    return {};
  }
}

function getChangeRequestDetails_(requestId) {
  rbacRequireAnyPermission_(['reviewChanges','viewApprovalHistory']);

  const found = getChangeRequestById_(requestId);

  if (!found) {
    throw new Error('Change request not found: ' + requestId);
  }

  const request = found.item;
  const currentEmail = getActualCurrentUserEmail_();

  if (!currentEmail) {
    throw new Error(
      'Unable to identify the logged-in user email.'
    );
  }

  const current = normalizeEmail_(currentEmail);

  const isMaker =
    normalizeEmail_(request['Requested By']) === current;

  const isChecker =
    canCurrentUserReviewRequest_(request, currentEmail);


  const oldValues =
    parseJsonObjectSafe_(request['Old Values JSON']);

  const proposedValues =
    parseJsonObjectSafe_(request['Proposed Values JSON']);

  const fieldNames = {};

  Object.keys(oldValues).forEach(function(field) {
    fieldNames[field] = true;
  });

  Object.keys(proposedValues).forEach(function(field) {
    fieldNames[field] = true;
  });

  const changes = Object.keys(fieldNames).map(function(field) {
    return {
      field: field,
      oldValue:
        oldValues[field] === undefined
          ? ''
          : oldValues[field],
      newValue:
        proposedValues[field] === undefined
          ? ''
          : proposedValues[field]
    };
  });

  const response = Object.assign({}, request);
  delete response._rowNumber;
  delete response['Approval Token'];
  response.changes = changes;
  const access = rbacRequire_('read');
  const pending = cleanString_(request.Status).toLowerCase() === 'pending approval';
  response.canReview = Boolean(pending && isChecker && !isMaker && access.permissions.reviewChanges);
  response.canRevise = Boolean(access.permissions.editRecords && isMaker && request.Status === 'Returned');
  response.canReassign = Boolean(pending && access.canManageUsers);
  response.reviewAccessMessage = !pending ? 'This request is ' + request.Status + '.' :
    !access.permissions.reviewChanges ? 'Your role does not include approval review permission.' :
    isMaker ? 'You submitted this request. Another assigned checker must review it.' :
    !isChecker ? 'You are signed in as ' + currentEmail + '. This request is assigned to ' +
      request['Approver Email'] + (request['Backup Approver Email'] ? ' (backup: ' + request['Backup Approver Email'] + ')' : '') +
      '. Ask an Admin to correct the assignment if needed.' : '';

  return response;
}

function updateChangeRequestRow_(found, updates) {
  const sheet = found.sheet;
  const headers = found.headers;
  const rowNumber = found.item._rowNumber;

  Object.keys(updates).forEach(function(header) {
    const index = headers.indexOf(header);
    if (index === -1) return;
    sheet.getRange(rowNumber, index + 1).setValue(updates[header]);
  });

  SpreadsheetApp.flush();
}

function assertChangeRequestSnapshotIsCurrent_(request) {
  const recordResult = getRecordById_(request['Record ID']);

  if (
    !recordResult ||
    !recordResult.success ||
    !recordResult.record
  ) {
    throw new Error(
      'Unable to load the current master record before approval.'
    );
  }

  if (recordResult.isDeleted) {
    throw new Error(
      'The record was deleted after this request was submitted.'
    );
  }

  const oldValues =
    parseJsonObjectSafe_(request['Old Values JSON']);

  const conflicts = [];

  Object.keys(oldValues).forEach(function(field) {
    const expected =
      oldValues[field] === null ||
      oldValues[field] === undefined
        ? ''
        : String(oldValues[field]);

    const actual =
      recordResult.record[field] === null ||
      recordResult.record[field] === undefined
        ? ''
        : String(recordResult.record[field]);

    if (expected !== actual) {
      conflicts.push(field);
    }
  });

  if (conflicts.length) {
    throw new Error(
      'Approval blocked because the master record changed after this request was submitted. Conflicting field(s): ' +
      conflicts.join(', ') +
      '. Reject this request and submit a new change request.'
    );
  }

  return recordResult.record;
}

function approveRecordChangeRequest_(requestId, reviewerComment) {
  rbacRequirePermission_('reviewChanges');

  const requestTarget = cleanString_(requestId);
  if (!requestTarget) {
    throw new Error('Request ID is required.');
  }

  let requestSnapshot = null;
  let checkerEmail = '';
  let proposedValues = {};
  let processingMarked = false;

  /*
   * Stage 1:
   * Lock the approval request, validate checker permissions and the
   * original master snapshot, then mark the request as Processing
   * Approval. We intentionally release this lock BEFORE calling
   * updateComplianceRecord_(), because that existing function uses its
   * own ScriptLock.
   */
  const stageOneLock = LockService.getScriptLock();

  try {
    stageOneLock.waitLock(20000);

    const found = getChangeRequestById_(requestTarget);

    if (!found) {
      throw new Error('Change request not found: ' + requestTarget);
    }

    const request = found.item;

    if (
      cleanString_(request['Status']).toLowerCase() !==
      'pending approval'
    ) {
      throw new Error(
        'Only Pending Approval requests can be approved.'
      );
    }

    checkerEmail = getActualCurrentUserEmail_();

    if (!checkerEmail) {
      throw new Error(
        'Unable to identify the logged-in checker email.'
      );
    }

    if (
      normalizeEmail_(request['Requested By']) ===
      normalizeEmail_(checkerEmail)
    ) {
      throw new Error(
        'Maker cannot approve their own change request.'
      );
    }

    if (!canCurrentUserReviewRequest_(request, checkerEmail)) {
      throw new Error(
        'You are not the assigned checker for this change request.'
      );
    }

    assertChangeRequestSnapshotIsCurrent_(request);

    proposedValues =
      parseJsonObjectSafe_(request['Proposed Values JSON']);

    if (!Object.prototype.hasOwnProperty.call(proposedValues, 'Record Type')) {
      proposedValues['Record Type'] = cleanString_(request['Record Type']);
    }

    requestSnapshot = Object.assign({}, request);

    updateChangeRequestRow_(found, {
      'Status': 'Processing Approval',
      'Reviewed By': checkerEmail,
      'Reviewed Date': new Date().toISOString(),
      'Reviewer Comment': cleanString_(reviewerComment)
    });

    processingMarked = true;

  } catch (error) {
    console.error(
      'approveRecordChangeRequest validation error:',
      error
    );
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to validate the change request.'
    );
  } finally {
    try { stageOneLock.releaseLock(); } catch (ignore) {}
  }

  /*
   * Stage 2:
   * Apply the approved values using the dashboard's existing update
   * engine. That function performs its own locking and final backend
   * validations.
   */
  let updateResult;

  try {
    updateResult = updateComplianceRecord_(
      requestSnapshot['Record ID'],
      proposedValues,
      'Approved via ' +
        requestSnapshot['Request ID'] +
        '. Original reason: ' +
        cleanString_(requestSnapshot['Change Reason'])
    );
  } catch (updateError) {
    /*
     * Return the request to Pending Approval when the master update
     * fails so the checker can retry after the underlying issue is
     * corrected.
     */
    if (processingMarked) {
      const rollbackLock = LockService.getScriptLock();

      try {
        rollbackLock.waitLock(15000);
        const latest = getChangeRequestById_(requestTarget);

        if (
          latest &&
          cleanString_(latest.item['Status']).toLowerCase() ===
            'processing approval'
        ) {
          updateChangeRequestRow_(latest, {
            'Status': 'Pending Approval',
            'Reviewed By': '',
            'Reviewed Date': '',
            'Reviewer Comment': '',
            'Applied Date': ''
          });
        }
      } catch (rollbackError) {
        console.error(
          'Unable to restore approval request after update failure:',
          rollbackError
        );
      } finally {
        try { rollbackLock.releaseLock(); } catch (ignore) {}
      }
    }

    console.error(
      'approveRecordChangeRequest update error:',
      updateError
    );

    throw new Error(
      updateError && updateError.message
        ? updateError.message
        : 'Unable to apply the approved record change.'
    );
  }

  /*
   * Stage 3:
   * Finalize the approval request and write approval history.
   */
  const finalLock = LockService.getScriptLock();

  try {
    finalLock.waitLock(15000);

    const latest = getChangeRequestById_(requestTarget);

    if (!latest) {
      throw new Error(
        'The master record was updated, but the approval request could not be reloaded for finalization.'
      );
    }

    const reviewedDate = new Date().toISOString();

    updateChangeRequestRow_(latest, {
      'Status': 'Approved',
      'Reviewed By': checkerEmail,
      'Reviewed Date': reviewedDate,
      'Reviewer Comment': cleanString_(reviewerComment),
      'Applied Date': reviewedDate
    });

    try {
      appendRecordHistoryEntries_([{
        timestamp: reviewedDate,
        recordId: requestSnapshot['Record ID'],
        action: 'CHANGE APPROVED',
        fieldName: 'Maker-Checker Request',
        oldValue: 'Pending Approval',
        newValue: 'Approved',
        changedBy: checkerEmail,
        changeReason:
          cleanString_(reviewerComment) ||
          ('Approved request ' + requestSnapshot['Request ID'])
      }]);
    } catch (historyError) {
      console.error(
        'Unable to write CHANGE APPROVED history:',
        historyError
      );
    }

    try {
      sendChangeRequestDecisionEmail_(
        requestSnapshot,
        'Approved',
        checkerEmail,
        reviewerComment
      );
    } catch (emailError) {
      console.error(
        'Approval decision email failed:',
        emailError
      );
    }

    return {
      success: true,
      requestId: requestSnapshot['Request ID'],
      recordId: requestSnapshot['Record ID'],
      status: 'Approved',
      reviewedBy: checkerEmail,
      changedFieldCount:
        updateResult && updateResult.changedFieldCount
          ? updateResult.changedFieldCount
          : 0,
      record:
        updateResult && updateResult.record
          ? updateResult.record
          : null,
      message:
        'Change request approved and the master record was updated.'
    };

  } catch (error) {
    console.error(
      'approveRecordChangeRequest finalization error:',
      error
    );
    throw new Error(
      error && error.message
        ? error.message
        : 'The record update completed, but approval finalization failed.'
    );
  } finally {
    try { finalLock.releaseLock(); } catch (ignore) {}
  }
}
function rejectRecordChangeRequest_(requestId, reviewerComment) {
  rbacRequirePermission_('reviewChanges');

  const comment = cleanString_(reviewerComment);

  if (!comment) {
    throw new Error(
      'Reviewer comment is required when rejecting a change request.'
    );
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const found = getChangeRequestById_(requestId);

    if (!found) {
      throw new Error('Change request not found: ' + requestId);
    }

    const request = found.item;

    if (
      cleanString_(request['Status']).toLowerCase() !==
      'pending approval'
    ) {
      throw new Error(
        'Only Pending Approval requests can be rejected.'
      );
    }

    const checkerEmail = getActualCurrentUserEmail_();

    if (!checkerEmail) {
      throw new Error(
        'Unable to identify the logged-in checker email.'
      );
    }

    if (
      normalizeEmail_(request['Requested By']) ===
      normalizeEmail_(checkerEmail)
    ) {
      throw new Error(
        'Maker cannot reject their own change request.'
      );
    }

    if (
      !canCurrentUserReviewRequest_(request, checkerEmail)
    ) {
      throw new Error(
        'You are not the assigned checker for this change request.'
      );
    }

    const reviewedDate = new Date().toISOString();

    updateChangeRequestRow_(found, {
      'Status': 'Rejected',
      'Reviewed By': checkerEmail,
      'Reviewed Date': reviewedDate,
      'Reviewer Comment': comment,
      'Applied Date': ''
    });

    try {
      appendRecordHistoryEntries_([{
        timestamp: reviewedDate,
        recordId: request['Record ID'],
        action: 'CHANGE REJECTED',
        fieldName: 'Maker-Checker Request',
        oldValue: 'Pending Approval',
        newValue: 'Rejected',
        changedBy: checkerEmail,
        changeReason: comment
      }]);
    } catch (historyError) {
      console.error(
        'Unable to write CHANGE REJECTED history:',
        historyError
      );
    }

    try {
      sendChangeRequestDecisionEmail_(
        request,
        'Rejected',
        checkerEmail,
        comment
      );
    } catch (emailError) {
      console.error(
        'Rejection decision email failed:',
        emailError
      );
    }

    return {
      success: true,
      requestId: request['Request ID'],
      recordId: request['Record ID'],
      status: 'Rejected',
      reviewedBy: checkerEmail,
      message:
        'Change request rejected. The master record remains unchanged.'
    };

  } catch (error) {
    console.error('rejectRecordChangeRequest error:', error);
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to reject the change request.'
    );
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function returnRecordChangeRequest_(requestId, reviewerComment) {
  rbacRequirePermission_('reviewChanges');

  const comment = cleanString_(reviewerComment);

  if (!comment) {
    throw new Error(
      'Reviewer comment is required when sending back a change request.'
    );
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const found = getChangeRequestById_(requestId);

    if (!found) {
      throw new Error('Change request not found: ' + requestId);
    }

    const request = found.item;

    if (
      cleanString_(request['Status']).toLowerCase() !==
      'pending approval'
    ) {
      throw new Error(
        'Only Pending Approval requests can be sent back.'
      );
    }

    const checkerEmail = getActualCurrentUserEmail_();

    if (!checkerEmail) {
      throw new Error(
        'Unable to identify the logged-in checker email.'
      );
    }

    if (
      normalizeEmail_(request['Requested By']) ===
      normalizeEmail_(checkerEmail)
    ) {
      throw new Error(
        'Maker cannot send back their own change request.'
      );
    }

    if (
      !canCurrentUserReviewRequest_(request, checkerEmail)
    ) {
      throw new Error(
        'You are not the assigned checker for this change request.'
      );
    }

    const reviewedDate = new Date().toISOString();

    updateChangeRequestRow_(found, {
      'Status': 'Returned',
      'Reviewed By': checkerEmail,
      'Reviewed Date': reviewedDate,
      'Reviewer Comment': comment,
      'Applied Date': ''
    });

    try {
      appendRecordHistoryEntries_([{
        timestamp: reviewedDate,
        recordId: request['Record ID'],
        action: 'CHANGE RETURNED',
        fieldName: 'Maker-Checker Request',
        oldValue: 'Pending Approval',
        newValue: 'Returned',
        changedBy: checkerEmail,
        changeReason: comment
      }]);
    } catch (historyError) {
      console.error(
        'Unable to write CHANGE RETURNED history:',
        historyError
      );
    }

    try {
      sendChangeRequestDecisionEmail_(
        request,
        'Sent Back',
        checkerEmail,
        comment
      );
    } catch (emailError) {
      console.error(
        'Send-back decision email failed:',
        emailError
      );
    }

    return {
      success: true,
      requestId: request['Request ID'],
      recordId: request['Record ID'],
      status: 'Returned',
      reviewedBy: checkerEmail,
      message:
        'Change request sent back. The master record remains unchanged.'
    };

  } catch (error) {
    console.error('returnRecordChangeRequest error:', error);
    throw new Error(
      error && error.message
        ? error.message
        : 'Unable to reject the change request.'
    );
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


function sendChangeRequestDecisionEmail_(
  request,
  decision,
  checkerEmail,
  reviewerComment
) {
  const makerEmail = cleanString_(request['Requested By']);

  if (!makerEmail) return;

  const subject =
    request['Request ID'] +
    ' ' +
    decision +
    ' | ' +
    request['Record ID'];

  let body =
    'Your compliance record change request has been ' +
    decision.toLowerCase() +
    '.\n\n' +
    'Request ID: ' + request['Request ID'] + '\n' +
    'Record ID: ' + request['Record ID'] + '\n' +
    'Reviewed By: ' + checkerEmail + '\n' +
    'Decision: ' + decision + '\n';

  if (cleanString_(reviewerComment)) {
    body +=
      'Reviewer Comment: ' +
      cleanString_(reviewerComment) +
      '\n';
  }

  body +=
    decision === 'Approved'
      ? '\nThe approved changes have been applied to the master compliance record.'
      : '\nThe master compliance record was not changed.';

  MailApp.sendEmail({
    to: makerEmail,
    subject: subject,
    body: body,
    name: 'Digiclarity Compliance Dashboard'
  });
}


/***************************************************************
 * MAKER-CHECKER DIAGNOSTIC
 ***************************************************************/
function runMakerCheckerSystemTest_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const results = [];

  function pushTest(name, passed, details) {
    results.push({
      test: name,
      passed: !!passed,
      details: details || ''
    });
  }

  try {
    const master = getMasterSheet_();
    pushTest(
      'Master Register',
      !!master,
      master ? master.getName() : 'Missing'
    );
  } catch (error) {
    pushTest(
      'Master Register',
      false,
      error.message
    );
  }

  try {
    const changeSheet = getChangeRequestSheet_();
    pushTest(
      'Change_Requests',
      !!changeSheet,
      changeSheet ? changeSheet.getName() : 'Missing'
    );
  } catch (error) {
    pushTest(
      'Change_Requests',
      false,
      error.message
    );
  }

  try {
    const matrixSheet = getApprovalMatrixSheet_();
    pushTest(
      'Approval_Matrix',
      !!matrixSheet,
      matrixSheet ? matrixSheet.getName() : 'Missing'
    );

    let configured = false;

    if (matrixSheet.getLastRow() >= 2) {
      const values = matrixSheet
        .getRange(
          2,
          1,
          matrixSheet.getLastRow() - 1,
          matrixSheet.getLastColumn()
        )
        .getDisplayValues();

      configured = values.some(function(row) {
        const email = cleanString_(row[3]);
        return (
          email &&
          email.toUpperCase() !== 'ENTER_APPROVER_EMAIL' &&
          email.indexOf('@') !== -1
        );
      });
    }

    pushTest(
      'Approver Configured',
      configured,
      configured
        ? 'At least one approver is configured.'
        : 'Replace ENTER_APPROVER_EMAIL in Approval_Matrix.'
    );

  } catch (error) {
    pushTest(
      'Approval_Matrix',
      false,
      error.message
    );
  }

  try {
    const currentEmail = getActualCurrentUserEmail_();
    pushTest(
      'Authenticated User Email',
      !!currentEmail,
      currentEmail ||
        'Session.getActiveUser().getEmail() returned blank.'
    );
  } catch (error) {
    pushTest(
      'Authenticated User Email',
      false,
      error.message
    );
  }

  const passed =
    results.filter(function(item) {
      return item.passed;
    }).length;

  return {
    success: passed === results.length,
    passed: passed,
    total: results.length,
    results: results
  };
}

/***************************************************************
 * RBAC — validated server-side on every public entry point.
 * Bootstrap is editor-only (trailing underscore), never callable
 * using google.script.run. No effective-user fallback for visitors.
 ***************************************************************/
const RBAC_INITIAL_ADMIN = 'bhuvaneshwaran.k@digiclarity.in';
const RBAC_ALLOWED_DOMAINS = ['digiclarity.in', 'clarityrcm.com'];

function rbacAllowedEmail_(email) {
  const value = String(email || '').trim().toLowerCase();
  return value.length <= 254 && /^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+$/.test(value) &&
    RBAC_ALLOWED_DOMAINS.indexOf(value.split('@')[1]) !== -1;
}
const RBAC_HEADERS = ['Email', 'Display Name', 'Role', 'Active', 'Updated By', 'Updated At', 'Departments'];
const RBAC_ROLE_HEADERS = ['Role Name', 'Description', 'Built In', 'Active', 'Scope Type', 'Allowed Departments', 'Allowed Users', 'Permissions JSON', 'Updated By', 'Updated At'];
const RBAC_BUILTIN_ROLES = ['Admin', 'Compliance Manager', 'Auditor', 'Department User'];
// Kept for compatibility with older helper code. Custom roles are loaded from Role_Definitions.
const RBAC_ROLES = RBAC_BUILTIN_ROLES.slice();
const RBAC_PERMISSION_KEYS = [
  'manageUsers','assignRoles','manageAccounts',
  'viewDashboard','viewRecords','createRecords','editRecords','deleteRecords',
  'submitEvidence','reviewResponses','reviewChanges','viewApprovalHistory',
  'viewReports','viewIamReports','viewExecutive','viewAnalytics','viewCrossFramework',
  'viewEvidence','uploadEvidence','manageEvidence','manageEvidenceRequests',
  'viewControlTests','manageControlTests','viewMonitoring','manageMonitoring',
  'viewAuditReadiness','exportAuditReadiness'
];

function rbacIdentity_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!rbacAllowedEmail_(email)) {
    throw new Error('Access denied: sign in with your assigned digiclarity.in or clarityrcm.com account. The deployment must expose your signed-in identity.');
  }
  return email;
}

function rbacRows_() {
  const sheet = getSpreadsheet_().getSheetByName('User_Access');
  if (!sheet) throw new Error('Access control is not initialized. Ask the dashboard Admin to complete setup.');
  const values = sheet.getDataRange().getValues();
  if (!values.length || RBAC_HEADERS.slice(0,6).some(function(h, i) { return values[0][i] !== h; }) || (values[0].length > 6 && values[0][6] !== 'Departments')) {
    throw new Error('User_Access headers are invalid. Access has been blocked.');
  }
  return {sheet: sheet, rows: values.slice(1).map(function(row, index) {
    return {email: String(row[0] || '').trim().toLowerCase(), name: String(row[1] || ''),
      role: String(row[2] || '').trim(), active: row[3] === true || String(row[3]).toLowerCase() === 'true',
      departments: iamParseDepartments_(row[6]), updatedBy: String(row[4] || ''), updatedAt: row[5] instanceof Date ? row[5].toISOString() : String(row[5] || ''), row: index + 2};
  }).filter(function(row) { return !!row.email; })};
}

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

function rbacReadRoleDefinitions_() {
  const sheet = getSpreadsheet_().getSheetByName('Role_Definitions');
  if (!sheet) return rbacBuiltInRoleSeeds_().map(rbacNormalizeRole_);
  const values = sheet.getDataRange().getValues();
  if (!values.length || RBAC_ROLE_HEADERS.some(function(header, index){ return String((values[0] || [])[index] || '') !== header; })) {
    throw new Error('Role_Definitions headers are invalid. Access has been blocked until the role configuration is repaired.');
  }
  return values.slice(1).map(function(row){
    return rbacNormalizeRole_({
      name:row[0], description:row[1], builtIn:row[2], active:row[3], scopeType:row[4],
      allowedDepartments:row[5], allowedUsers:row[6], permissions:row[7], updatedBy:row[8],
      updatedAt:row[9] instanceof Date ? row[9].toISOString() : row[9]
    });
  }).filter(function(role){ return !!role.name; });
}

function rbacGetRoleDefinition_(name) {
  const key = String(name || '').trim().toLowerCase();
  return rbacReadRoleDefinitions_().find(function(role){ return role.name.toLowerCase() === key; }) || null;
}

function rbacEnsureRoleDefinitionsSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('Role_Definitions');
  if (sheet) {
    const header = sheet.getRange(1,1,1,RBAC_ROLE_HEADERS.length).getDisplayValues()[0];
    if (RBAC_ROLE_HEADERS.some(function(value,index){ return header[index] !== value; })) throw new Error('Role_Definitions headers are invalid.');
    return sheet;
  }
  sheet = ss.insertSheet('Role_Definitions');
  sheet.getRange(1,1,1,RBAC_ROLE_HEADERS.length).setValues([RBAC_ROLE_HEADERS]);
  const actor = rbacIdentity_();
  const now = new Date().toISOString();
  const rows = rbacBuiltInRoleSeeds_().map(function(role){
    return [role.name, role.description, true, true, role.scopeType, JSON.stringify(role.allowedDepartments), JSON.stringify(role.allowedUsers), JSON.stringify(role.permissions), actor, now];
  });
  if (rows.length) sheet.getRange(2,1,rows.length,RBAC_ROLE_HEADERS.length).setValues(rows);
  sheet.setFrozenRows(1);
  return sheet;
}

function rbacCoreAdminPermissions_(permissions) {
  permissions = rbacNormalizePermissions_(permissions);
  RBAC_PERMISSION_KEYS.forEach(function(key){ permissions[key] = true; });
  return permissions;
}

function rbacPermissionAllowed_(access, key) {
  return !!(access && access.permissions && access.permissions[key]);
}

function rbacAccessForEmail_(email) {
  // Phase 8.7: do NOT run the temporary-access lifecycle from every RBAC read.
  // The time-based worker and explicit access-request paths own lifecycle processing.
  // Running it here can recursively re-enter RBAC while Privileged Access is loading.
  email = String(email || '').trim().toLowerCase();
  phase87AssertExpirySafe_(email);
  if (!rbacAllowedEmail_(email)) throw new Error('Access denied: sign in with your assigned company account.');
  const matches = rbacRows_().rows.filter(function(row) { return row.email === email; });
  if (matches.length !== 1 || !matches[0].active) {
    throw new Error('Access denied: your account has no unique, active dashboard role. Contact the Admin.');
  }
  const user = matches[0];
  const role = rbacGetRoleDefinition_(user.role);
  if (!role || !role.active) throw new Error('Access denied: the assigned role is missing or inactive. Contact the Admin.');
  const permissions = rbacNormalizePermissions_(role.permissions);
  const departmentScoped = role.scopeType === 'DEPARTMENTS';
  if (departmentScoped) {
    if (!user.departments.length || user.departments.some(function(dept){ return /^(all|\*)$/i.test(dept); })) throw new Error('Access denied: Admin must assign specific departments for this role.');
    iamAssertStorageOwner_();
  }
  const canWrite = permissions.createRecords || permissions.editRecords || permissions.deleteRecords;
  const canRespond = (permissions.submitEvidence && permissions.uploadEvidence) || permissions.reviewResponses || canWrite;
  const isSystemAdmin = role.builtIn && role.name === 'Admin';
  return {
    email:email,name:user.name,role:user.role,roleDescription:role.description,departments:user.departments,
    permissions:permissions,scopeType:role.scopeType,allowedDepartments:role.allowedDepartments,allowedUsers:role.allowedUsers,
    canWrite:canWrite,canRespond:canRespond,canManageUsers:permissions.manageUsers,canAssignRoles:permissions.assignRoles,
    canManageAccounts:permissions.manageAccounts,departmentScoped:departmentScoped,isSystemAdmin:isSystemAdmin
  };
}

function rbacRequire_(permission) {
  const access = rbacAccessForEmail_(rbacIdentity_());
  if ((permission === 'admin' && !access.isSystemAdmin) ||
      (permission === 'write' && !access.canWrite) ||
      (permission === 'global' && access.departmentScoped) ||
      (permission === 'respond' && !access.canRespond)) {
    throw new Error('Access denied: your ' + access.role + ' role cannot perform this action.');
  }
  return access;
}

function rbacRequirePermission_(key) {
  if (RBAC_PERMISSION_KEYS.indexOf(key) === -1) throw new Error('Unknown access permission: ' + key);
  const access = rbacRequire_('read');
  if (!rbacPermissionAllowed_(access, key)) throw new Error('Access denied: your ' + access.role + ' role does not include ' + key + '.');
  return access;
}

function rbacRequireAnyPermission_(keys) {
  const access = rbacRequire_('read');
  const ok = (keys || []).some(function(key){ return rbacPermissionAllowed_(access,key); });
  if (!ok) throw new Error('Access denied: your ' + access.role + ' role does not include the required permission.');
  return access;
}

function getAccessContext_() {
  return rbacRequire_('read');
}



/***************************************************************
 * PHASE 8.2 — DASHBOARD-LOCAL PASSWORD
 * The Google Workspace identity remains the primary identity.
 * This separate password belongs only to the Compliance Dashboard.
 *
 * Policy:
 * - minimum 12 characters
 * - uppercase + lowercase + number + special character
 * - blocks whitespace, obvious sequences/repeats, common words,
 *   identity terms, common location terms and birthday-like dates
 * - expires every 90 days
 * - last five passwords cannot be reused
 * - five failed attempts lock dashboard-password verification for 15 minutes
 ***************************************************************/
function dashboardCredentialSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(DASHBOARD_CREDENTIAL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_CREDENTIAL_SHEET);
    sheet.getRange(1,1,1,DASHBOARD_CREDENTIAL_HEADERS.length).setValues([DASHBOARD_CREDENTIAL_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Backward-compatible migration from the original 5-column credential sheet.
  const legacyHeaders = ['Email', 'Salt', 'Password Hash', 'Updated By', 'Updated At'];
  const width = Math.max(sheet.getLastColumn(), legacyHeaders.length);
  const current = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  if (legacyHeaders.some(function(value,index){ return current[index] !== value; })) {
    throw new Error('Dashboard_Credentials headers are invalid. Contact the Compliance Administrator.');
  }
  for (let i = legacyHeaders.length; i < DASHBOARD_CREDENTIAL_HEADERS.length; i++) {
    if (current[i] && current[i] !== DASHBOARD_CREDENTIAL_HEADERS[i]) {
      throw new Error('Dashboard_Credentials contains an unexpected column at position ' + (i + 1) + '. Contact the Compliance Administrator.');
    }
  }
  sheet.getRange(1,1,1,DASHBOARD_CREDENTIAL_HEADERS.length).setValues([DASHBOARD_CREDENTIAL_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function dashboardParseHistory_(value) {
  if (!value) return [];
  try {
    const parsed = typeof value === 'string' ? JSON.parse(value) : value;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(function(entry){
      return entry && typeof entry === 'object' && entry.salt && entry.hash;
    }).map(function(entry){
      return {salt:String(entry.salt), hash:String(entry.hash), changedAt:String(entry.changedAt || '')};
    }).slice(0, DASHBOARD_PASSWORD_HISTORY_DEPTH);
  } catch (error) {
    return [];
  }
}

function dashboardCredentialRows_() {
  const sheet = dashboardCredentialSheet_();
  if (sheet.getLastRow() < 2) return {sheet:sheet, rows:[]};
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,DASHBOARD_CREDENTIAL_HEADERS.length).getValues();
  return {sheet:sheet, rows:values.map(function(row,index){
    return {
      email:String(row[0]||'').trim().toLowerCase(),
      salt:String(row[1]||''),
      hash:String(row[2]||''),
      updatedBy:String(row[3]||''),
      updatedAt:row[4] instanceof Date ? row[4].toISOString() : String(row[4]||''),
      history:dashboardParseHistory_(row[5]),
      expiresAt:row[6] instanceof Date ? row[6].toISOString() : String(row[6]||''),
      failedAttempts:Number(row[7]||0) || 0,
      lockedUntil:row[8] instanceof Date ? row[8].toISOString() : String(row[8]||''),
      lastLoginAt:row[9] instanceof Date ? row[9].toISOString() : String(row[9]||''),
      row:index+2
    };
  }).filter(function(row){ return !!row.email; })};
}

function dashboardCredentialForEmail_(email) {
  email = String(email || '').trim().toLowerCase();
  const matches = dashboardCredentialRows_().rows.filter(function(row){ return row.email === email; });
  if (matches.length > 1) throw new Error('Duplicate dashboard credential entries exist for this account. Contact the Compliance Administrator.');
  return matches[0] || null;
}

function dashboardPasswordPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty(DASHBOARD_PASSWORD_PEPPER_KEY);
  if (!pepper) {
    pepper = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(DASHBOARD_PASSWORD_PEPPER_KEY, pepper);
  }
  return pepper;
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

function dashboardPasswordHash_(password, salt) {
  let value = String(salt || '') + '\n' + String(password || '');
  const pepper = dashboardPasswordPepper_();
  // Repeated HMAC plus a server-only pepper keeps the Sheet alone insufficient
  // for offline password verification. Plaintext passwords are never stored.
  for (let i = 0; i < 180; i++) {
    value = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value, pepper));
  }
  return value;
}

function dashboardPasswordEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function dashboardPasswordExpiryIso_(fromDate) {
  const d = fromDate instanceof Date ? new Date(fromDate.getTime()) : new Date(fromDate || Date.now());
  if (isNaN(d.getTime())) d.setTime(Date.now());
  d.setDate(d.getDate() + DASHBOARD_PASSWORD_MAX_AGE_DAYS);
  return d.toISOString();
}

function dashboardCredentialExpiry_(credential) {
  if (!credential) return '';
  if (credential.expiresAt) return credential.expiresAt;
  return dashboardPasswordExpiryIso_(credential.updatedAt || Date.now());
}

function dashboardCredentialExpired_(credential) {
  const expiry = new Date(dashboardCredentialExpiry_(credential));
  return !isNaN(expiry.getTime()) && expiry.getTime() <= Date.now();
}

function dashboardPasswordMatchesStored_(password, entry) {
  if (!entry || !entry.salt || !entry.hash) return false;
  return dashboardPasswordEqual_(dashboardPasswordHash_(password, entry.salt), entry.hash);
}

function dashboardAssertPasswordNotReused_(password, credential) {
  if (!credential) return;
  if (dashboardPasswordMatchesStored_(password, credential)) {
    throw new Error('You cannot reuse your current dashboard password.');
  }
  const history = Array.isArray(credential.history) ? credential.history : [];
  for (let i = 0; i < history.length; i++) {
    if (dashboardPasswordMatchesStored_(password, history[i])) {
      throw new Error('You cannot reuse any of your last ' + DASHBOARD_PASSWORD_HISTORY_DEPTH + ' dashboard passwords.');
    }
  }
}

function dashboardWriteCredential_(actor, email, password, existing) {
  const data = dashboardCredentialRows_();
  const salt = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  const hash = dashboardPasswordHash_(password, salt);
  const now = new Date().toISOString();
  let history = existing && Array.isArray(existing.history) ? existing.history.slice() : [];
  if (existing && existing.salt && existing.hash) {
    history.unshift({salt:existing.salt, hash:existing.hash, changedAt:existing.updatedAt || now});
  }
  history = history.slice(0, DASHBOARD_PASSWORD_HISTORY_DEPTH);
  const row = [
    email, salt, hash, actor, now, JSON.stringify(history), dashboardPasswordExpiryIso_(now),
    0, '', existing ? String(existing.lastLoginAt || '') : ''
  ];
  if (existing) data.sheet.getRange(existing.row,1,1,DASHBOARD_CREDENTIAL_HEADERS.length).setValues([row]);
  else data.sheet.appendRow(row);
  return {email:email, updatedAt:now, expiresAt:row[6]};
}

function dashboardCredentialSecurityUpdate_(credential, failedAttempts, lockedUntil, lastLoginAt) {
  if (!credential || !credential.row) return;
  const sheet = dashboardCredentialSheet_();
  sheet.getRange(credential.row, 8, 1, 3).setValues([[
    Number(failedAttempts || 0), String(lockedUntil || ''), String(lastLoginAt || credential.lastLoginAt || '')
  ]]);
}

function dashboardCredentialLocked_(credential) {
  if (!credential || !credential.lockedUntil) return false;
  const d = new Date(credential.lockedUntil);
  return !isNaN(d.getTime()) && d.getTime() > Date.now();
}

function dashboardCredentialLockMessage_(credential) {
  const until = new Date(credential.lockedUntil);
  const minutes = isNaN(until.getTime()) ? DASHBOARD_PASSWORD_LOCK_MINUTES : Math.max(1, Math.ceil((until.getTime() - Date.now()) / 60000));
  return 'Too many incorrect dashboard-password attempts. Try again in about ' + minutes + ' minute' + (minutes === 1 ? '' : 's') + '.';
}

function createDashboardPassword(input) {
  input = input || {};
  const email = rbacIdentity_();
  const access = rbacAccessForEmail_(email);
  const current = dashboardCredentialForEmail_(email);
  if (current) throw new Error('A dashboard password already exists. Use Change Dashboard Password from My Account.');
  const password = dashboardPasswordPolicy_(input.newPassword, access);
  if (password !== String(input.confirmPassword || '')) throw new Error('The new passwords do not match.');
  const result = dashboardWriteCredential_(email, email, password, null);
  try { rbacAudit_(email,email,null,{action:'DASHBOARD PASSWORD CREATED',updatedAt:result.updatedAt,expiresAt:result.expiresAt}); } catch (error) {}
  const session = dashboardCreateSession_(email, input.clientMeta);
  dashboardSecurityAudit_('PASSWORD_CREATED','SUCCESS',email,email,session.sessionId,access.role,input.clientMeta,{});
  return {success:true,message:'Dashboard password created. It will expire in ' + DASHBOARD_PASSWORD_MAX_AGE_DAYS + ' days.',access:access,expiresAt:result.expiresAt,sessionToken:session.token};
}

function verifyDashboardPassword(input) {
  input = input || {};
  const email = rbacIdentity_();
  const access = rbacAccessForEmail_(email);
  const credential = dashboardCredentialForEmail_(email);
  if (!credential) return {success:false,code:'PASSWORD_SETUP_REQUIRED',message:'Create your dashboard password before continuing.'};

  if (dashboardCredentialLocked_(credential)) {
    return {success:false,code:'ACCOUNT_LOCKED',message:dashboardCredentialLockMessage_(credential)};
  }

  const password = String(input.password || '');
  const candidate = dashboardPasswordHash_(password, credential.salt);
  if (!dashboardPasswordEqual_(candidate, credential.hash)) {
    const failures = Number(credential.failedAttempts || 0) + 1;
    if (failures >= DASHBOARD_PASSWORD_MAX_FAILED_ATTEMPTS) {
      const until = new Date(Date.now() + DASHBOARD_PASSWORD_LOCK_MINUTES * 60000).toISOString();
      dashboardCredentialSecurityUpdate_(credential, failures, until, credential.lastLoginAt);
      return {success:false,code:'ACCOUNT_LOCKED',message:'Too many incorrect attempts. Dashboard password verification is locked for ' + DASHBOARD_PASSWORD_LOCK_MINUTES + ' minutes.'};
    }
    dashboardCredentialSecurityUpdate_(credential, failures, '', credential.lastLoginAt);
    return {success:false,code:'INVALID_PASSWORD',message:'Incorrect dashboard password. ' + (DASHBOARD_PASSWORD_MAX_FAILED_ATTEMPTS - failures) + ' attempt(s) remaining before a temporary lock.'};
  }

  const now = new Date().toISOString();
  dashboardCredentialSecurityUpdate_(credential, 0, '', now);

  if (dashboardCredentialExpired_(credential)) {
    return {
      success:false, code:'PASSWORD_EXPIRED', passwordChangeRequired:true,
      email:access.email, name:access.name, role:access.role,
      message:'Your Compliance Dashboard password has expired after ' + DASHBOARD_PASSWORD_MAX_AGE_DAYS + ' days. Change it to continue.'
    };
  }

  // Existing passwords created before this policy are checked at login and must
  // be upgraded if they no longer satisfy the current policy.
  try {
    dashboardPasswordPolicy_(password, access);
  } catch (policyError) {
    return {
      success:false, code:'PASSWORD_POLICY_UPGRADE_REQUIRED', passwordChangeRequired:true,
      email:access.email, name:access.name, role:access.role,
      message:'Your existing dashboard password does not meet the current security policy. Change it to continue.'
    };
  }

  const expiry = new Date(dashboardCredentialExpiry_(credential));
  const daysRemaining = isNaN(expiry.getTime()) ? null : Math.max(0, Math.ceil((expiry.getTime() - Date.now()) / 86400000));
  return {success:true,code:'AUTHORIZED',message:'Dashboard password verified.',access:access,expiresAt:dashboardCredentialExpiry_(credential),daysRemaining:daysRemaining};
}

function changeDashboardPassword_(input) {
  input = input || {};
  const access = rbacRequire_('read');
  const email = access.email;
  const current = dashboardCredentialForEmail_(email);
  if (!current) throw new Error('No dashboard password exists yet. Sign out and create one from the login screen.');
  if (dashboardCredentialLocked_(current)) throw new Error(dashboardCredentialLockMessage_(current));

  const currentCandidate = dashboardPasswordHash_(String(input.currentPassword || ''), current.salt);
  if (!dashboardPasswordEqual_(currentCandidate, current.hash)) throw new Error('Current dashboard password is incorrect.');

  const next = dashboardPasswordPolicy_(input.newPassword, access);
  if (next !== String(input.confirmPassword || '')) throw new Error('The new passwords do not match.');
  dashboardAssertPasswordNotReused_(next, current);

  const data = dashboardCredentialRows_();
  const oldValues = data.sheet.getRange(current.row,1,1,DASHBOARD_CREDENTIAL_HEADERS.length).getValues();
  const result = dashboardWriteCredential_(email,email,next,current);
  try {
    rbacAudit_(email,email,{action:'DASHBOARD PASSWORD ACTIVE'},{action:'DASHBOARD PASSWORD CHANGED',updatedAt:result.updatedAt,expiresAt:result.expiresAt});
  } catch (error) {
    data.sheet.getRange(current.row,1,1,DASHBOARD_CREDENTIAL_HEADERS.length).setValues(oldValues);
    throw new Error('Password change was rolled back because the access audit log could not be written.');
  }
  dashboardRevokeAllSessionsForEmail_(email,'PASSWORD_CHANGED');
  const session = dashboardCreateSession_(email, input.clientMeta);
  dashboardSecurityAudit_('PASSWORD_CHANGED','SUCCESS',email,email,session.sessionId,access.role,input.clientMeta,{});
  return {success:true,message:'Dashboard password changed successfully. The new password will expire in ' + DASHBOARD_PASSWORD_MAX_AGE_DAYS + ' days.',access:access,expiresAt:result.expiresAt,sessionToken:session.token};
}

// Admin recovery endpoint for User Management. Reset removes the local dashboard
// credential and history so the user must create a new password at next sign-in.
function adminResetDashboardPassword_(email) {
  const actor = rbacRequirePermission_('manageAccounts');
  const target = String(email || '').trim().toLowerCase();
  if (!rbacAllowedEmail_(target)) throw new Error('Enter a valid dashboard user email.');
  const data = dashboardCredentialRows_();
  const matches = data.rows.filter(function(row){ return row.email === target; });
  if (matches.length > 1) throw new Error('Duplicate credential entries exist for this account.');
  if (!matches.length) return {success:true,message:'No dashboard password was configured for this user.'};
  data.sheet.deleteRow(matches[0].row);
  dashboardRevokeAllSessionsForEmail_(target,'ADMIN_PASSWORD_RESET');
  rbacAudit_(actor.email,target,{action:'DASHBOARD PASSWORD ACTIVE'},{action:'DASHBOARD PASSWORD RESET',resetBy:actor.email});
  return {success:true,message:'Dashboard password reset. The user must create a new password at next sign-in.'};
}

/**
 * Phase 8.1 authentication gateway.
 * Returns only the currently signed-in user's own authorization state.
 * It intentionally does not expose the User_Access directory.
 */
function getLoginGatewayContext() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();

  if (!email) {
    return {
      ok: false,
      code: 'IDENTITY_UNAVAILABLE',
      email: '',
      message: 'A Google Workspace identity could not be detected. Sign in with your assigned company account and try again.'
    };
  }

  if (!rbacAllowedEmail_(email)) {
    return {
      ok: false,
      code: 'DOMAIN_NOT_ALLOWED',
      email: email,
      message: 'This Google account is not an approved DigiClarity company account. Use your @digiclarity.in or @clarityrcm.com account.'
    };
  }

  try {
    const data = rbacRows_();
    const matches = data.rows.filter(function(row) { return row.email === email; });

    if (matches.length === 0) {
      return {
        ok: false,
        code: 'ACCESS_NOT_ASSIGNED',
        email: email,
        message: 'Your company account is recognized, but dashboard access has not been assigned. Contact the Compliance Administrator.'
      };
    }

    if (matches.length > 1) {
      return {
        ok: false,
        code: 'DUPLICATE_ACCESS',
        email: email,
        message: 'Duplicate access entries exist for this account. Contact the Compliance Administrator before continuing.'
      };
    }

    if (!matches[0].active) {
      return {
        ok: false,
        code: 'ACCOUNT_DISABLED',
        email: email,
        message: 'Your dashboard account is currently disabled. Contact the Compliance Administrator if access is required.'
      };
    }

    const role = rbacGetRoleDefinition_(matches[0].role);
    if (!role || !role.active) {
      return {
        ok: false,
        code: 'ROLE_UNAVAILABLE',
        email: email,
        message: 'Your assigned dashboard role is unavailable or inactive. Contact the Compliance Administrator.'
      };
    }

    const access = rbacAccessForEmail_(email);
    const credential = dashboardCredentialForEmail_(email);
    if (!credential) {
      return {
        ok: true,
        code: 'PASSWORD_SETUP_REQUIRED',
        email: access.email,
        name: access.name,
        role: access.role,
        requiresPasswordSetup: true,
        message: 'Company account verified. Create a Compliance Dashboard password to continue.'
      };
    }
    return {
      ok: true,
      code: 'PASSWORD_REQUIRED',
      email: access.email,
      name: access.name,
      role: access.role,
      requiresDashboardPassword: true,
      message: 'Company account verified. Enter your Compliance Dashboard password to continue.'
    };
  } catch (error) {
    return {
      ok: false,
      code: 'ACCESS_CONFIGURATION_ERROR',
      email: email,
      message: error && error.message ? error.message : 'Dashboard access could not be verified. Contact the Compliance Administrator.'
    };
  }
}


function rbacAudit_(actor, email, before, after) {
  const sheet = getSpreadsheet_().getSheetByName('Access_Audit_Log');
  if (!sheet) throw new Error('Access_Audit_Log is missing. Restore the audit sheet before changing roles.');
  sheet.appendRow([new Date().toISOString(), actor, email, JSON.stringify(before || null), JSON.stringify(after)]);
}

// Run once from the Apps Script editor, signed in as RBAC_INITIAL_ADMIN.
function setupDashboardRbac_() {
  const actor = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (actor !== RBAC_INITIAL_ADMIN) throw new Error('Only the configured initial Admin can initialize RBAC.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    if (ss.getSheetByName('User_Access')) throw new Error('RBAC already exists. Use User Management; setup never resets existing roles.');
    let audit = ss.getSheetByName('Access_Audit_Log');
    if (!audit) {
      audit = ss.insertSheet('Access_Audit_Log');
      audit.appendRow(['Timestamp', 'Changed By', 'User Email', 'Before JSON', 'After JSON']);
      audit.setFrozenRows(1);
    }
    const sheet = ss.insertSheet('User_Access');
    sheet.appendRow(RBAC_HEADERS);
    sheet.appendRow([actor, 'Bhuvaneshwaran K', 'Admin', true, actor, new Date().toISOString(), '[]']);
    sheet.setFrozenRows(1);
    rbacEnsureRoleDefinitionsSheet_();
    rbacAudit_(actor, actor, null, {role: 'Admin', active: true, action: 'Initial setup'});
    return 'RBAC ready. Initial Admin: ' + actor;
  } finally { lock.releaseLock(); }
}

function listDashboardUsers_() {
  rbacRequirePermission_('manageUsers');
  return rbacRows_().rows.map(function(row) {
    const user = Object.assign({}, row); delete user.row; return user;
  });
}

function listRoleDefinitions_() {
  rbacRequirePermission_('manageUsers');
  rbacEnsureRoleDefinitionsSheet_();
  return rbacReadRoleDefinitions_();
}

function rbacValidateRoleScope_(role, email, departments) {
  role = rbacNormalizeRole_(role);
  if (role.scopeType === 'USERS') {
    if (!role.allowedUsers.length) throw new Error('This role has no allowed users configured.');
    if (role.allowedUsers.indexOf(String(email || '').toLowerCase()) === -1) throw new Error('This role is restricted to specific users and cannot be assigned to this email.');
    return [];
  }
  if (role.scopeType === 'DEPARTMENTS') {
    const selected = iamValidateDepartments_(departments);
    if (role.allowedDepartments.length) {
      const allowedKeys = role.allowedDepartments.map(iamDepartmentKey_);
      selected.forEach(function(dept){
        if (allowedKeys.indexOf(iamDepartmentKey_(dept)) === -1) throw new Error('This role can only be assigned within its configured departments.');
      });
    }
    return selected;
  }
  return [];
}

function saveDashboardUser_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const actor = rbacRequirePermission_('assignRoles');
    if (!actor.permissions.manageAccounts) throw new Error('Access denied: account management permission is also required.');
    if (actor.role === 'Compliance Manager' && !actor.isSystemAdmin) {
      const requested = String((input||{}).role||'');
      const current = rbacRows_().rows.find(function(u){return u.email === String((input||{}).email||'').trim().toLowerCase();});
      if (requested === 'Admin' || (current && current.role === 'Admin')) throw new Error('Compliance Managers cannot directly create, edit, or reassign Admin accounts. Use the privileged access workflow.');
      if (String((input||{}).email||'').trim().toLowerCase() === actor.email && current && current.role !== requested) throw new Error('You cannot change your own role. Submit an independent access request.');
    }
    input = input || {};
    const email = String(input.email || '').trim().toLowerCase();
    const name = String(input.name || '').trim();
    const roleName = String(input.role || '').trim();
    if (!rbacAllowedEmail_(email)) throw new Error('Enter a valid @digiclarity.in or @clarityrcm.com email.');
    if (name.length > 120 || /^[=+@-]/.test(name) || /^[=+@-]/.test(email)) throw new Error('Invalid display name or email.');
    if (typeof input.active !== 'boolean') throw new Error('Invalid account status.');
    const role = rbacGetRoleDefinition_(roleName);
    if (!role || !role.active) throw new Error('Select an active role from Role Permissions.');
    const data = rbacRows_();
    const matches = data.rows.filter(function(row) { return row.email === email; });
    if (matches.length > 1) throw new Error('Duplicate user entries must be resolved before editing.');
    const before = matches[0];
    if (before && before.role === 'Admin' && before.active && (roleName !== 'Admin' || !input.active) &&
        !data.rows.some(function(row) { return row.email !== email && row.role === 'Admin' && row.active; })) {
      throw new Error('The last active Admin cannot be removed or disabled.');
    }
    if (email === actor.email && actor.isSystemAdmin && (roleName !== 'Admin' || !input.active)) throw new Error('You cannot remove your own Admin access. Another Admin must make that change.');
    if (data.sheet.getLastColumn() < RBAC_HEADERS.length) throw new Error('Admin must run initializeIamPhase before changing access.');
    const departments = rbacValidateRoleScope_(role, email, input.departments || []);
    phase87GuardDirectUser_(before,roleName,input.active,departments);
    const after = {email: email, name: name, role: roleName, active: input.active, departments:departments};
    if (!getSpreadsheet_().getSheetByName('Access_Audit_Log')) throw new Error('Access_Audit_Log is missing. No access change was made.');
    const row = [email, name, roleName, input.active, actor.email, new Date().toISOString(), JSON.stringify(departments)];
    const targetRow = before ? before.row : data.sheet.getLastRow() + 1;
    const oldRow = before ? data.sheet.getRange(targetRow, 1, 1, RBAC_HEADERS.length).getValues() : null;
    data.sheet.getRange(targetRow, 1, 1, RBAC_HEADERS.length).setValues([row]);
    try { rbacAudit_(actor.email, email, before || null, after); }
    catch (error) {
      if (oldRow) data.sheet.getRange(targetRow, 1, 1, RBAC_HEADERS.length).setValues(oldRow);
      else data.sheet.getRange(targetRow, 1, 1, RBAC_HEADERS.length).clearContent();
      throw new Error('Access change rolled back because the audit log could not be written.');
    }
    return {success: true, message: 'User access saved.'};
  } finally { lock.releaseLock(); }
}

function rbacRoleRow_(role, actorEmail) {
  role = rbacNormalizeRole_(role);
  return [role.name, role.description, role.builtIn, role.active, role.scopeType,
    JSON.stringify(role.allowedDepartments), JSON.stringify(role.allowedUsers), JSON.stringify(role.permissions), actorEmail, new Date().toISOString()];
}

function saveRoleDefinition_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const actor = rbacRequirePermission_('assignRoles');
    input = input || {};
    const name = String(input.name || '').trim();
    const description = String(input.description || '').trim();
    if (name.length < 2 || name.length > 60 || /^[=+@-]/.test(name) || /[<>]/.test(name)) throw new Error('Role name must be 2–60 safe characters.');
    if (description.length > 300 || /^[=+@-]/.test(description)) throw new Error('Role description is invalid or too long.');
    const sheet = rbacEnsureRoleDefinitionsSheet_();
    const roles = rbacReadRoleDefinitions_();
    const existing = roles.find(function(role){ return role.name.toLowerCase() === name.toLowerCase(); }) || null;
    const createOnly = input.createOnly === true;
    if (createOnly && existing) throw new Error('A role with this name already exists. Choose a different role name.');
    if (!createOnly && !existing) throw new Error('Role not found. Refresh Role Permissions and try again.');
    if (existing && existing.builtIn && existing.name === 'Admin') {
      throw new Error('Built-in Admin has permanent full access. Its permissions cannot be customized.');
    }

    const scopeType = String(input.scopeType || (existing && existing.scopeType) || 'ANY').toUpperCase();
    if (['ANY','DEPARTMENTS','USERS'].indexOf(scopeType) === -1) throw new Error('Choose a valid assignment restriction.');
    let allowedDepartments = rbacSafeJsonArray_(input.allowedDepartments);
    let allowedUsers = rbacSafeJsonArray_(input.allowedUsers).map(function(email){ return email.toLowerCase(); });
    if (scopeType === 'DEPARTMENTS') {
      if (allowedDepartments.length) {
        const catalog = iamDepartmentCatalog_();
        allowedDepartments = allowedDepartments.map(function(value){
          const match = catalog.find(function(dept){ return iamDepartmentKey_(dept) === iamDepartmentKey_(value); });
          if (!match) throw new Error('Choose role departments from the current department catalog.');
          return match;
        }).filter(function(value,index,array){ return array.indexOf(value) === index; });
      }
      allowedUsers = [];
    } else if (scopeType === 'USERS') {
      if (!allowedUsers.length) throw new Error('Add at least one allowed user email for a user-restricted role.');
      allowedUsers.forEach(function(email){ if (!rbacAllowedEmail_(email)) throw new Error('Allowed users must use an approved company email domain.'); });
      allowedDepartments = [];
    } else {
      allowedDepartments = [];
      allowedUsers = [];
    }

    let permissions = rbacNormalizePermissions_(input.permissions || (existing && existing.permissions));
    const builtIn = existing ? existing.builtIn : false;
    if (builtIn && existing.name === 'Admin') permissions = rbacCoreAdminPermissions_(permissions);
    const role = rbacNormalizeRole_({
      name:name, description:description || (existing && existing.description) || '', builtIn:builtIn, active:input.active !== false,
      scopeType:scopeType, allowedDepartments:allowedDepartments, allowedUsers:allowedUsers, permissions:permissions,
      updatedBy:actor.email, updatedAt:new Date().toISOString()
    });

    // V46: documented permanent-Admin authorization for built-in Compliance Manager policy.
    const v46BuiltInChange = !!(existing && existing.builtIn && ['Compliance Manager','Auditor','Department User'].indexOf(existing.name) !== -1 &&
      JSON.stringify([existing.permissions,existing.scopeType,existing.allowedUsers,existing.allowedDepartments,existing.active]) !==
      JSON.stringify([role.permissions,role.scopeType,role.allowedUsers,role.allowedDepartments,role.active]));
    if (v46BuiltInChange) {
      if (!phase87PermanentAdminEmail_(actor.email)) throw new Error('Only a permanent Admin may authorize a built-in Compliance Manager role change.');
      const reason = String(input.changeReason || '').trim();
      if (reason.length < 12 || reason.length > 1000) throw new Error('Enter a change reason of 12–1000 characters.');
      if (role.scopeType !== existing.scopeType || JSON.stringify(role.allowedUsers)!==JSON.stringify(existing.allowedUsers) ||
          JSON.stringify(role.allowedDepartments)!==JSON.stringify(existing.allowedDepartments) || role.active!==existing.active)
        throw new Error('This authorization allows role permission changes only; use the governed workflow for scope or status changes.');
      if (existing.name !== 'Compliance Manager') {
        const keys = RBAC_PERMISSION_KEYS.filter(function(key){ return ['viewEvidence','uploadEvidence','manageEvidence','manageEvidenceRequests',
          'viewControlTests','manageControlTests','viewMonitoring','manageMonitoring','viewAuditReadiness','exportAuditReadiness'].indexOf(key) === -1; });
        if (keys.some(function(key){return !!role.permissions[key] !== !!existing.permissions[key];}))
          throw new Error('For Auditor and Department User built-in roles, direct authorization is limited to Evidence Management permissions. Use the governed workflow for other access changes.');
      }
      if (accessTempGrantRows_().rows.some(function(g){return ['REVERTED','CANCELLED'].indexOf(g.status)===-1&&(g.originalRole===existing.name||g.grantedRole===existing.name);}))
        throw new Error('Resolve open temporary grants referencing this role first.');
      input.changeReason = reason;
    } else {
      phase87GuardRole_(actor,existing,role);
    }
    const values = sheet.getDataRange().getValues();
    let rowNumber = 0;
    for (let i=1;i<values.length;i++) if (String(values[i][0] || '').trim().toLowerCase() === name.toLowerCase()) { rowNumber = i + 1; break; }
    const before = existing ? Object.assign({}, existing) : null;
    const targetRow = rowNumber || (sheet.getLastRow() + 1);
    const oldRow = rowNumber ? sheet.getRange(rowNumber,1,1,RBAC_ROLE_HEADERS.length).getValues() : null;
    sheet.getRange(targetRow,1,1,RBAC_ROLE_HEADERS.length).setValues([rbacRoleRow_(role,actor.email)]);
    try {
      rbacAudit_(actor.email, 'ROLE::' + name, before, role);
      if (v46BuiltInChange) {
        // Dedicated, mandatory governance audit with before/after and justification.
        const audit = accessRequestEnsureSheet_('BuiltIn_Role_Change_Audit', ['At','Actor','Role','Reason','Before JSON','After JSON','Authorization']);
        audit.appendRow([new Date().toISOString(),actor.email,name,input.changeReason,JSON.stringify(before),JSON.stringify(role),'PERMANENT_ADMIN_DIRECT']);
      }
    }
    catch (error) {
      if (oldRow) sheet.getRange(targetRow,1,1,RBAC_ROLE_HEADERS.length).setValues(oldRow);
      else sheet.getRange(targetRow,1,1,RBAC_ROLE_HEADERS.length).clearContent();
      throw new Error('Role change was rolled back because a required audit log could not be written.');
    }
    return {success:true, message:(existing ? 'Role permissions updated: ' : 'New role created: ') + name, role:role};
  } finally { lock.releaseLock(); }
}

function initializeCustomRoleRbac_() {
  const actor = rbacRequire_('admin');
  const sheet = rbacEnsureRoleDefinitionsSheet_();
  return {success:true, message:'Role_Definitions is ready.', roles:rbacReadRoleDefinitions_(), sheet:sheet.getName(), actor:actor.email};
}

/***************************************************************
 * SELECTABLE ONE-TIME RBAC INITIALIZER
 * Select initializeRbacOnce in the Apps Script editor and Run.
 * Only the configured initial Admin can call this entry point.
 * Existing role assignments are never reset by this function.
 ***************************************************************/
function initializeRbacOnce_() {
  const email = rbacIdentity_();
  if (email !== RBAC_INITIAL_ADMIN) {
    throw new Error('Only the configured initial Admin can run setup.');
  }
  if (getSpreadsheet_().getSheetByName('User_Access')) {
    rbacRequire_('admin');
    const existingMessage = 'RBAC is already initialized. Existing roles are unchanged. Use User Management in the dashboard.';
    console.log(existingMessage);
    return existingMessage;
  }
  const result = setupDashboardRbac_();
  console.log(result);
  return result;
}

// Admin correction applies only to this pending request, not Approval_Matrix.
function reassignPendingChangeRequest_(requestId, approverEmail) {
  const actor = rbacRequirePermission_('manageUsers');
  const email = normalizeEmail_(approverEmail);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const found = getChangeRequestById_(requestId);
    if (!found || found.item.Status !== 'Pending Approval') throw new Error('Only pending requests can be reassigned.');
    const request = found.item;
    const users = rbacRows_().rows.filter(function(u) { return u.email === email; });
    const candidateRole = users.length === 1 ? rbacGetRoleDefinition_(users[0].role) : null;
    if (!rbacAllowedEmail_(email) || users.length !== 1 || !users[0].active || !candidateRole || !candidateRole.permissions.reviewChanges) throw new Error('Select an active user whose role includes approval review permission.');
    if (email === normalizeEmail_(request['Requested By'])) throw new Error('The maker cannot be assigned as checker.');
    const previous = request['Approver Email'];
    if (normalizeEmail_(previous) === email) return {success:true, message:'This checker is already assigned.'};
    updateChangeRequestRow_(found, {'Approver Email':email});
    try {
      appendRecordHistoryEntries_([{timestamp:new Date().toISOString(), recordId:request['Record ID'],
        action:'CHECKER REASSIGNED', fieldName:'Approver Email', oldValue:previous, newValue:email,
        changedBy:actor.email, changeReason:'Checker corrected for request ' + request['Request ID']}]);
    } catch (error) {
      updateChangeRequestRow_(found, {'Approver Email':previous});
      throw new Error('Assignment reverted because its audit history could not be recorded.');
    }
    return {success:true, message:'Checker updated. Ask the manager to refresh Approvals. Update Approval_Matrix separately for future requests.'};
  } finally { lock.releaseLock(); }
}

function getReturnedRequestForRevision_(requestId) {
  const actor = rbacRequirePermission_('editRecords');
  const found = getChangeRequestById_(requestId);
  if (!found || found.item.Status !== 'Returned') throw new Error('Only returned requests can be reopened.');
  const request = found.item;
  if (normalizeEmail_(request['Requested By']) !== actor.email) throw new Error('Only the original maker can reopen this request.');
  if (getPendingChangeRequestForRecord_(request['Record ID'])) throw new Error('This record already has a pending request. Review it before resubmitting.');
  const result = getRecordById_(request['Record ID']);
  if (result && result.isDeleted) throw new Error('Restore the deleted record before revising this request.');
  const record = result && result.record ? result.record : result;
  if (!record || !record['Record ID']) throw new Error('The original record is unavailable.');
  const proposed = parseJsonObjectSafe_(request['Proposed Values JSON']);
  const old = parseJsonObjectSafe_(request['Old Values JSON']);
  const prefill = Object.assign({}, record);
  // Replay only proposed fields, never stale unrelated master fields.
  Object.keys(proposed).forEach(function(field) {
    if (!makerCheckerValuesEqual_(field, old[field], proposed[field])) prefill[field] = proposed[field];
  });
  prefill['Record ID'] = record['Record ID'];
  prefill['Record Type'] = normalizeRecordTypeKey_(prefill['Record Type']) || record['Record Type'];
  return {recordId:record['Record ID'], recordType:prefill['Record Type'], originalRecordType:record['Record Type'], prefillData:prefill,
    sourceRequestId:request['Request ID'], reviewerComment:request['Reviewer Comment'] || ''};
}

/**
 * Permanently removes a dashboard identity from User_Access.
 * Admin only. A deletion reason is mandatory and recorded in Access_Audit_Log.
 */
function deleteDashboardUser_(email, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const actor = rbacRequirePermission_('manageUsers');
    const targetEmail = String(email || '').trim().toLowerCase();
    const deletionReason = String(reason || '').trim();

    if (!rbacAllowedEmail_(targetEmail)) {
      throw new Error('Enter a valid dashboard user email.');
    }
    if (deletionReason.length < 5) {
      throw new Error('A deletion reason of at least 5 characters is required.');
    }
    if (deletionReason.length > 500) {
      throw new Error('Deletion reason must be 500 characters or fewer.');
    }
    if (targetEmail === actor.email) {
      throw new Error('You cannot delete your own Admin access. Another Admin must perform this action.');
    }

    const data = rbacRows_();
    const matches = data.rows.filter(function(row) {
      return row.email === targetEmail;
    });

    if (matches.length !== 1) {
      throw new Error(
        matches.length === 0
          ? 'Dashboard user not found.'
          : 'Duplicate user entries must be resolved before deletion.'
      );
    }

    const target = matches[0];
    phase87AssertNoGrant_(targetEmail);
    if(phase87RolePrivileged_(target.role))throw new Error('Privileged accounts must be deactivated through dual-approved Access Requests; direct deletion is blocked.');

    if (
      target.role === 'Admin' &&
      target.active &&
      !data.rows.some(function(row) {
        return row.email !== targetEmail && row.role === 'Admin' && row.active;
      })
    ) {
      throw new Error('The last active Admin cannot be deleted.');
    }

    if (!getSpreadsheet_().getSheetByName('Access_Audit_Log')) {
      throw new Error('Access_Audit_Log is missing. No user was deleted.');
    }

    const rowNumber = target.row;
    const oldValues = data.sheet
      .getRange(rowNumber, 1, 1, RBAC_HEADERS.length)
      .getValues();

    data.sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();

    try {
      rbacAudit_(
        actor.email,
        targetEmail,
        target,
        {
          action: 'USER DELETED',
          deleted: true,
          deletionReason: deletionReason,
          deletedBy: actor.email,
          deletedAt: new Date().toISOString()
        }
      );
    } catch (auditError) {
      // Restore the deleted row if the mandatory audit entry cannot be written.
      const lastRow = data.sheet.getLastRow();
      if (rowNumber <= lastRow) {
        data.sheet.insertRowBefore(rowNumber);
      } else {
        data.sheet.insertRowAfter(lastRow);
      }
      data.sheet
        .getRange(rowNumber, 1, 1, RBAC_HEADERS.length)
        .setValues(oldValues);
      SpreadsheetApp.flush();

      throw new Error(
        'User deletion was rolled back because the access audit log could not be written.'
      );
    }

    return {
      success: true,
      message: 'Dashboard access deleted for ' + targetEmail + '.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * V9 dashboard-user deletion endpoint.
 * Admin-only, reason-required, audited, and intentionally fail-fast.
 * Payload: { email: string, reason: string }
 */
function deleteDashboardUserV9_(input) {
  input = input || {};
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    throw new Error('User access is busy with another update. Wait a few seconds and try again.');
  }

  try {
    const actor = rbacRequirePermission_('manageUsers');
    const targetEmail = String(input.email || '').trim().toLowerCase();
    const deletionReason = String(input.reason || '').trim();

    if (!rbacAllowedEmail_(targetEmail)) {
      throw new Error('Enter a valid @digiclarity.in or @clarityrcm.com dashboard user email.');
    }
    if (deletionReason.length < 5) {
      throw new Error('A deletion reason of at least 5 characters is required.');
    }
    if (deletionReason.length > 500) {
      throw new Error('Deletion reason must be 500 characters or fewer.');
    }
    if (targetEmail === actor.email) {
      throw new Error('You cannot delete your own Admin access. Another Admin must perform this action.');
    }

    const data = rbacRows_();
    const matches = data.rows.filter(function(row) {
      return row.email === targetEmail;
    });

    if (matches.length === 0) {
      throw new Error('Dashboard user not found. Refresh User Management and try again.');
    }
    if (matches.length > 1) {
      throw new Error('Duplicate User_Access entries exist for this email. Resolve the duplicates before deletion.');
    }

    const target = matches[0];
    phase87AssertNoGrant_(targetEmail);
    if(phase87RolePrivileged_(target.role))throw new Error('Privileged accounts must be deactivated through dual-approved Access Requests; direct deletion is blocked.');
    if (
      target.role === 'Admin' &&
      target.active &&
      !data.rows.some(function(row) {
        return row.email !== targetEmail && row.role === 'Admin' && row.active;
      })
    ) {
      throw new Error('The last active Admin cannot be deleted.');
    }

    const ss = getSpreadsheet_();
    const auditSheet = ss.getSheetByName('Access_Audit_Log');
    if (!auditSheet) {
      throw new Error('Access_Audit_Log is missing. No user was deleted.');
    }

    const rowNumber = target.row;
    const oldValues = data.sheet
      .getRange(rowNumber, 1, 1, RBAC_HEADERS.length)
      .getValues();

    data.sheet.deleteRow(rowNumber);

    try {
      rbacAudit_(
        actor.email,
        targetEmail,
        target,
        {
          action: 'USER DELETED',
          deleted: true,
          deletionReason: deletionReason,
          deletedBy: actor.email,
          deletedAt: new Date().toISOString()
        }
      );
    } catch (auditError) {
      // Restore User_Access if mandatory auditing fails.
      const lastRow = data.sheet.getLastRow();
      if (lastRow < 1) {
        throw new Error('User deletion failed and the User_Access sheet could not be restored safely.');
      }
      if (rowNumber <= lastRow) {
        data.sheet.insertRowBefore(rowNumber);
      } else {
        data.sheet.insertRowAfter(lastRow);
      }
      data.sheet
        .getRange(rowNumber, 1, 1, RBAC_HEADERS.length)
        .setValues(oldValues);

      throw new Error(
        'User deletion was rolled back because Access_Audit_Log could not be written: ' +
        (auditError && auditError.message ? auditError.message : auditError)
      );
    }

    return {
      success: true,
      email: targetEmail,
      message: 'Dashboard access deleted for ' + targetEmail + '.'
    };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/***************************************************************
 * V16 — CROSS-FRAMEWORK INTELLIGENCE BACKEND
 *
 * Adds the normalized Cross_Framework_Matrix engine expected by
 * the V11+ Cross-Framework Explorer frontend.
 *
 * Required existing helpers from the dashboard backend:
 *   getSpreadsheet_()
 *   cleanString_()
 *   uniqueSorted_()
 *   rbacRequire_()
 *   readFrameworkControlMaster_()
 *   isFrameworkControlActive_()
 *
 * Optional Google Sheet:
 *   Cross_Framework_Matrix
 *
 * If the normalized matrix is missing or unusable, the engine
 * safely falls back to Common Control Area relationships from
 * Framework_Control_Master.
 ***************************************************************/

const CROSS_FRAMEWORK_MAPPING_SHEET = 'Cross_Framework_Matrix';

function normalizeCrossFrameworkName_(value) {
  return frameworkControlAliasKeyV18_(value);
}

function crossFrameworkKey_(framework, controlId) {
  return normalizeCrossFrameworkName_(framework) + '||' + cleanString_(controlId).toLowerCase();
}

function crossFrameworkHeaderKey_(value) {
  return cleanString_(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function findCrossFrameworkMatrixHeaderRow_(data) {
  const required = [
    'mappingid',
    'sourceframework',
    'sourcecontrolid',
    'targetframework',
    'targetcontrolid'
  ];

  for (let r = 0; r < Math.min(data.length, 20); r++) {
    const keys = data[r].map(crossFrameworkHeaderKey_);
    const ok = required.every(function(key) {
      return keys.indexOf(key) !== -1;
    });
    if (ok) return r;
  }

  return -1;
}

function getCrossFrameworkMatrixHeaderMap_(headers) {
  const map = {};

  headers.forEach(function(header, index) {
    map[crossFrameworkHeaderKey_(header)] = index;
  });

  return {
    mappingId: map.mappingid,
    sourceFramework: map.sourceframework,
    sourceControlId: map.sourcecontrolid,
    sourceControlName: map.sourcecontrolname,
    targetFramework: map.targetframework,
    targetControlId: map.targetcontrolid,
    targetControlName: map.targetcontrolname,
    targetReference: map.targetreference,
    commonControlArea: map.commoncontrolarea,
    mappingStrength: map.mappingstrength,
    mappingRationale: map.mappingrationale,
    status: map.status
  };
}

function readCrossFrameworkMatrix_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CROSS_FRAMEWORK_MAPPING_SHEET);

  if (!sheet) {
    return {
      available: false,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: ''
    };
  }

  const data = sheet.getDataRange().getDisplayValues();

  if (!data.length) {
    return {
      available: true,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: ''
    };
  }

  const headerRow = findCrossFrameworkMatrixHeaderRow_(data);

  if (headerRow === -1) {
    return {
      available: true,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: 'Required mapping headers were not found. Expected Mapping ID, Source Framework, Source Control ID, Target Framework and Target Control ID.'
    };
  }

  const headers = data[headerRow].map(cleanString_);
  const map = getCrossFrameworkMatrixHeaderMap_(headers);

  const requiredIndexes = [
    map.mappingId,
    map.sourceFramework,
    map.sourceControlId,
    map.targetFramework,
    map.targetControlId
  ];

  if (requiredIndexes.some(function(index) {
    return index === undefined || index === -1;
  })) {
    return {
      available: true,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: 'Cross_Framework_Matrix is missing one or more required columns.'
    };
  }

  const rows = [];

  for (let r = headerRow + 1; r < data.length; r++) {
    const row = data[r];

    const sourceFramework = cleanString_(row[map.sourceFramework]);
    const sourceControlId = cleanString_(row[map.sourceControlId]);
    const targetFramework = cleanString_(row[map.targetFramework]);
    const targetControlId = cleanString_(row[map.targetControlId]);

    if (!sourceFramework || !sourceControlId || !targetFramework || !targetControlId) {
      continue;
    }

    const status = map.status !== undefined
      ? cleanString_(row[map.status])
      : 'Active';

    if (status && !isFrameworkControlActive_(status)) {
      continue;
    }

    rows.push({
      mappingId: cleanString_(row[map.mappingId]),
      sourceFramework: sourceFramework,
      sourceControlId: sourceControlId,
      sourceControlName: map.sourceControlName !== undefined
        ? cleanString_(row[map.sourceControlName])
        : '',
      targetFramework: targetFramework,
      targetControlId: targetControlId,
      targetControlName: map.targetControlName !== undefined
        ? cleanString_(row[map.targetControlName])
        : '',
      targetReference: map.targetReference !== undefined
        ? cleanString_(row[map.targetReference])
        : '',
      commonControlArea: map.commonControlArea !== undefined
        ? cleanString_(row[map.commonControlArea])
        : '',
      mappingStrength: map.mappingStrength !== undefined
        ? cleanString_(row[map.mappingStrength])
        : '',
      mappingRationale: map.mappingRationale !== undefined
        ? cleanString_(row[map.mappingRationale])
        : ''
    });
  }

  return {
    available: true,
    sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
    headerRow: headerRow + 1,
    headers: headers,
    rows: rows,
    error: ''
  };
}

function buildFrameworkControlIndex_(controls) {
  const index = {};

  (controls || []).forEach(function(control) {
    index[crossFrameworkKey_(control.framework, control.controlId)] = control;
  });

  return index;
}

function crossFrameworkReferenceForControl_(control) {
  if (!control) return '';

  const framework = normalizeCrossFrameworkName_(control.framework);

  if (framework === 'soc 2') {
    return cleanString_(control.soc2CCMapping);
  }

  if (framework === 'hipaa') {
    return cleanString_(control.hipaaMapping) || cleanString_(control.controlId);
  }

  return '';
}

function makeMappedControl_(framework, controlId, fallbackName, relationship, controlIndex, direction) {
  const key = crossFrameworkKey_(framework, controlId);
  const master = controlIndex[key] || null;

  return {
    framework: master ? master.framework : cleanString_(framework),
    controlId: master ? master.controlId : cleanString_(controlId),
    controlName: master
      ? master.controlName
      : cleanString_(fallbackName) || cleanString_(controlId),
    controlDescription: master ? master.controlDescription : '',
    domain: master ? master.domain : '',
    commonControlArea:
      cleanString_(relationship.commonControlArea) ||
      (master ? cleanString_(master.commonControlArea) : ''),
    soc2CCMapping: master ? cleanString_(master.soc2CCMapping) : '',
    hipaaMapping: master ? cleanString_(master.hipaaMapping) : '',
    targetReference:
      direction === 'forward'
        ? cleanString_(relationship.targetReference) || crossFrameworkReferenceForControl_(master)
        : crossFrameworkReferenceForControl_(master),
    mappingStrength: cleanString_(relationship.mappingStrength) || 'Mapped',
    mappingRationale: cleanString_(relationship.mappingRationale),
    mappingId: cleanString_(relationship.mappingId),
    mappingDirection: direction
  };
}

function getCrossFrameworkExplorerBootstrap_() {
  rbacRequirePermission_('viewCrossFramework');

  const controls = readFrameworkControlMaster_().controls;
  const frameworkCounts = {};
  const frameworkLabels = {};

  controls.forEach(function(control) {
    const key = normalizeCrossFrameworkName_(control.framework);
    if (!key) return;
    frameworkCounts[key] = (frameworkCounts[key] || 0) + 1;
    if (!frameworkLabels[key]) frameworkLabels[key] = control.framework;
  });

  const preferredOrder = {
    'iso 27001:2022': 1,
    'soc 2': 2,
    'hipaa': 3,
    'gdpr': 4,
    'dpdpa': 5,
    'pci dss': 6
  };

  const frameworks = Object.keys(frameworkCounts)
    .map(function(key) {
      return {
        framework: frameworkLabels[key],
        controlCount: frameworkCounts[key]
      };
    })
    .sort(function(a, b) {
      const ak = normalizeCrossFrameworkName_(a.framework);
      const bk = normalizeCrossFrameworkName_(b.framework);
      const ao = preferredOrder[ak] || 99;
      const bo = preferredOrder[bk] || 99;
      if (ao !== bo) return ao - bo;
      return a.framework.localeCompare(b.framework);
    });

  let matrix;
  try {
    matrix = readCrossFrameworkMatrix_();
  } catch (error) {
    matrix = {
      available: true,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: error && error.message ? error.message : String(error)
    };
  }

  return {
    success: true,
    frameworks: frameworks,
    plannedFrameworks: ['GDPR', 'DPDPA', 'PCI DSS'],
    mappingSheetAvailable: matrix.available === true,
    mappingSheet: matrix.sheetName || CROSS_FRAMEWORK_MAPPING_SHEET,
    mappingRowCount: Array.isArray(matrix.rows) ? matrix.rows.length : 0,
    mappingError: cleanString_(matrix.error)
  };
}

function getCrossFrameworkMappings_(sourceFramework, sourceControlId) {
  rbacRequirePermission_('viewCrossFramework');

  const requestedFramework = cleanString_(sourceFramework);
  const requestedControlId = cleanString_(sourceControlId);

  if (!requestedFramework || !requestedControlId) {
    throw new Error('Framework and Control ID are required for cross-framework mapping.');
  }

  const masterControls = readFrameworkControlMaster_().controls;
  const controlIndex = buildFrameworkControlIndex_(masterControls);
  const sourceKey = crossFrameworkKey_(requestedFramework, requestedControlId);
  const sourceControl = controlIndex[sourceKey] || null;

  if (!sourceControl) {
    throw new Error(
      'Control ' + requestedControlId +
      ' was not found for framework ' + requestedFramework + '.'
    );
  }

  let matrix;
  try {
    matrix = readCrossFrameworkMatrix_();
  } catch (error) {
    matrix = {
      available: true,
      sheetName: CROSS_FRAMEWORK_MAPPING_SHEET,
      rows: [],
      error: error && error.message ? error.message : String(error)
    };
  }

  const exact = [];
  const seenExact = {};

  if (matrix.available && !matrix.error) {
    (matrix.rows || []).forEach(function(relationship) {
      const sourceMatches =
        crossFrameworkKey_(relationship.sourceFramework, relationship.sourceControlId) === sourceKey;

      const targetMatches =
        crossFrameworkKey_(relationship.targetFramework, relationship.targetControlId) === sourceKey;

      let mapped = null;

      if (sourceMatches) {
        mapped = makeMappedControl_(
          relationship.targetFramework,
          relationship.targetControlId,
          relationship.targetControlName,
          relationship,
          controlIndex,
          'forward'
        );
      } else if (targetMatches) {
        mapped = makeMappedControl_(
          relationship.sourceFramework,
          relationship.sourceControlId,
          relationship.sourceControlName,
          relationship,
          controlIndex,
          'reverse'
        );
      }

      if (!mapped) return;

      const mappedKey = crossFrameworkKey_(mapped.framework, mapped.controlId);
      if (!mappedKey || mappedKey === sourceKey || seenExact[mappedKey]) return;

      seenExact[mappedKey] = true;
      exact.push(mapped);
    });
  }

  exact.sort(function(a, b) {
    const frameworkCompare = a.framework.localeCompare(b.framework);
    if (frameworkCompare !== 0) return frameworkCompare;
    return a.controlId.localeCompare(b.controlId, undefined, {
      numeric: true,
      sensitivity: 'base'
    });
  });

  if (exact.length) {
    return {
      success: true,
      sourceFramework: sourceControl.framework,
      sourceControlId: sourceControl.controlId,
      sourceControl: sourceControl,
      mappingMode: 'exact',
      mappingSheetAvailable: matrix.available === true,
      mappingSheet: matrix.sheetName || CROSS_FRAMEWORK_MAPPING_SHEET,
      mappingError: cleanString_(matrix.error),
      controls: exact
    };
  }

  const area = cleanString_(sourceControl.commonControlArea);
  const fallback = [];
  const seenFallback = {};

  if (area) {
    masterControls.forEach(function(control) {
      const sameArea =
        cleanString_(control.commonControlArea).toLowerCase() === area.toLowerCase();

      const key = crossFrameworkKey_(control.framework, control.controlId);

      if (!sameArea || key === sourceKey || seenFallback[key]) return;

      seenFallback[key] = true;
      fallback.push({
        framework: control.framework,
        controlId: control.controlId,
        controlName: control.controlName,
        controlDescription: control.controlDescription,
        domain: control.domain,
        commonControlArea: control.commonControlArea,
        soc2CCMapping: control.soc2CCMapping,
        hipaaMapping: control.hipaaMapping,
        targetReference: crossFrameworkReferenceForControl_(control),
        mappingStrength: 'Common-area related',
        mappingRationale:
          'No exact normalized matrix relationship was available. This control shares the Common Control Area "' +
          area + '" with the selected source control.',
        mappingId: '',
        mappingDirection: 'fallback'
      });
    });
  }

  fallback.sort(function(a, b) {
    const frameworkCompare = a.framework.localeCompare(b.framework);
    if (frameworkCompare !== 0) return frameworkCompare;
    return a.controlId.localeCompare(b.controlId, undefined, {
      numeric: true,
      sensitivity: 'base'
    });
  });

  return {
    success: true,
    sourceFramework: sourceControl.framework,
    sourceControlId: sourceControl.controlId,
    sourceControl: sourceControl,
    commonControlArea: area,
    mappingMode: 'fallback',
    mappingSheetAvailable: matrix.available === true,
    mappingSheet: matrix.sheetName || CROSS_FRAMEWORK_MAPPING_SHEET,
    mappingError: cleanString_(matrix.error),
    controls: fallback
  };
}

/***************************************************************
 * Optional admin-only diagnostic.
 ***************************************************************/
function testCrossFrameworkExplorerV16_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const bootstrap = getCrossFrameworkExplorerBootstrap_();
  const controls = readFrameworkControlMaster_().controls;

  if (!controls.length) {
    throw new Error('Framework_Control_Master has no active controls.');
  }

  const preferred = controls.find(function(control) {
    return normalizeCrossFrameworkName_(control.framework) === 'iso 27001:2022';
  }) || controls[0];

  const mapping = getCrossFrameworkMappings_(
    preferred.framework,
    preferred.controlId
  );

  return {
    success: true,
    version: 'V16',
    frameworkCount: bootstrap.frameworks.length,
    mappingSheetAvailable: bootstrap.mappingSheetAvailable,
    mappingRowCount: bootstrap.mappingRowCount,
    testControl: preferred.framework + ' / ' + preferred.controlId,
    mappingMode: mapping.mappingMode,
    mappedControls: mapping.controls.length
  };
}

/***************************************************************
 * V18 — FRAMEWORK CONTROL ALIAS-SAFE LOOKUP
 *
 * Allows Lookup_Master / UI framework labels to differ slightly
 * from Framework_Control_Master without returning zero controls.
 * Examples treated as the same framework:
 *   ISO 27001 / ISO 27001:2022 / ISO/IEC 27001:2022
 *   SOC2 / SOC 2 / SOC 2 Type 2
 *   HIPAA / HIPAA Security Rule
 *
 * This is intentionally a separate endpoint so the patch can be
 * appended safely without replacing the existing function.
 ***************************************************************/
function frameworkControlAliasKeyV18_(value) {
  const text = cleanString_(value).toLowerCase().replace(/\s+/g, ' ').trim();
  const compact = text.replace(/[^a-z0-9]/g, '');
  if (['soc2','soc2type2','soc2typeii','soc2trustservicescriteria'].indexOf(compact) !== -1) return 'soc 2';
  if (['hipaa','hipaasecurityrule','hipaaprivacyrule'].indexOf(compact) !== -1) return 'hipaa';
  if (['iso27k','iso27k2022','iso27001','iso270012022','isoiec27001','isoiec270012022'].indexOf(compact) !== -1) return 'iso 27001:2022';
  // Keep explicitly named older editions and custom frameworks distinct.
  return text;
}

function getFrameworkControlsAliasSafe_(framework) {
  rbacRequire_('read');

  const target = cleanString_(framework);
  const all = readFrameworkControlMaster_().controls;

  if (!target || target.toLowerCase() === 'all') {
    return {
      success: true,
      requestedFramework: target,
      resolvedFramework: '',
      controls: all
    };
  }

  const targetKey = frameworkControlAliasKeyV18_(target);

  const controls = all.filter(function(control) {
    return frameworkControlAliasKeyV18_(control.framework) === targetKey;
  });

  const resolvedFramework = controls.length
    ? cleanString_(controls[0].framework)
    : target;

  return {
    success: true,
    requestedFramework: target,
    resolvedFramework: resolvedFramework,
    frameworkKey: targetKey,
    controls: controls,
    count: controls.length
  };
}

/***************************************************************
 * V20 — SINGLE-MASTER CROSS-FRAMEWORK ENGINE
 *
 * Framework_Control_Master is the only source of truth.
 * Explicit relationships are derived from:
 *   - SOC 2 CC Mapping
 *   - HIPAA Mapping
 * Common Control Area is used only when no explicit reference
 * mapping is available for the selected control.
 *
 * Safe to append to the V18 backend. The V20 frontend calls only
 * the new V20 endpoints below, so older endpoints are untouched.
 ***************************************************************/

function splitFrameworkMappingRefsV20_(value) {
  const text = cleanString_(value);
  if (!text) return [];

  const seen = {};
  const refs = [];

  text
    .split(/[;,\n]+/)
    .map(function(item) { return cleanString_(item); })
    .filter(Boolean)
    .forEach(function(item) {
      const key = normalizeFrameworkMappingRefV20_(item);
      if (!key || seen[key]) return;
      seen[key] = true;
      refs.push({ key: key, label: item });
    });

  return refs;
}

function normalizeFrameworkMappingRefV20_(value) {
  return cleanString_(value)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[–—]/g, '-');
}

function frameworkReferenceBundleV20_(control) {
  const framework = normalizeCrossFrameworkName_(control && control.framework);
  const soc2 = splitFrameworkMappingRefsV20_(control && control.soc2CCMapping);
  const hipaa = splitFrameworkMappingRefsV20_(control && control.hipaaMapping);
  const iso = splitFrameworkMappingRefsV20_(control && control.iso27001Mapping);
  if (framework === 'iso 27001:2022' && control && control.controlId) {
    const label = cleanString_(control.controlId);
    const key = normalizeFrameworkMappingRefV20_(label);
    if (!iso.some(function(item) { return item.key === key; })) iso.push({key:key,label:label});
  }

  // HIPAA rows often carry the regulation citation in Control ID itself.
  if (framework === 'hipaa' && control && control.controlId) {
    const label = cleanString_(control.controlId);
    const key = normalizeFrameworkMappingRefV20_(label);
    if (key && !hipaa.some(function(item) { return item.key === key; })) {
      hipaa.push({ key: key, label: label });
    }
  }

  return {
    soc2: soc2,
    hipaa: hipaa,
    iso: iso
  };
}

function intersectFrameworkRefsV20_(left, right) {
  const rightMap = {};
  (right || []).forEach(function(item) {
    rightMap[item.key] = item.label;
  });

  const matches = [];
  const seen = {};
  (left || []).forEach(function(item) {
    if (!rightMap[item.key] || seen[item.key]) return;
    seen[item.key] = true;
    matches.push(item.label || rightMap[item.key]);
  });
  return matches;
}

function frameworkMappedReferenceV20_(control, basis, sharedRefs) {
  if (sharedRefs && sharedRefs.length) return sharedRefs.join(', ');

  const framework = normalizeCrossFrameworkName_(control && control.framework);
  if (framework === 'soc 2') return cleanString_(control.soc2CCMapping);
  if (framework === 'hipaa') {
    return cleanString_(control.hipaaMapping) || cleanString_(control.controlId);
  }

  if (basis === 'SOC 2 CC Mapping') return cleanString_(control && control.soc2CCMapping);
  if (basis === 'HIPAA Mapping') return cleanString_(control && control.hipaaMapping);
  return '';
}

function mappedControlFromMasterV20_(target, basis, sharedRefs, sourceArea, mode) {
  const area = cleanString_(target.commonControlArea) || cleanString_(sourceArea);
  const isArea = mode === 'area';

  return {
    framework: target.framework,
    controlId: target.controlId,
    controlName: target.controlName,
    controlDescription: target.controlDescription,
    domain: target.domain,
    commonControlArea: target.commonControlArea,
    soc2CCMapping: target.soc2CCMapping,
    hipaaMapping: target.hipaaMapping,
    targetReference: frameworkMappedReferenceV20_(target, basis, sharedRefs),
    mappingStrength: isArea ? 'Common-area related' : 'Reference overlap · Review required',
    mappingRationale: isArea
      ? 'No explicit SOC 2 or HIPAA reference mapping was available. This control shares the Common Control Area "' + area + '" with the selected source control.'
      : 'Reference overlap requires review and does not establish equivalent compliance. Linked through ' + basis + ' in Framework_Control_Master' +
        (sharedRefs && sharedRefs.length ? ': ' + sharedRefs.join(', ') + '.' : '.'),
    mappingId: '',
    mappingDirection: 'master',
    mappingBasis: basis || (isArea ? 'Common Control Area' : '')
  };
}

function getCrossFrameworkExplorerBootstrapV20_() {
  rbacRequirePermission_('viewCrossFramework');

  const controls = readFrameworkControlMaster_().controls;
  const frameworkCounts = {};
  const frameworkLabels = {};

  controls.forEach(function(control) {
    const key = normalizeCrossFrameworkName_(control.framework);
    if (!key) return;
    frameworkCounts[key] = (frameworkCounts[key] || 0) + 1;
    if (!frameworkLabels[key]) frameworkLabels[key] = control.framework;
  });

  const preferredOrder = {
    'iso 27001:2022': 1,
    'soc 2': 2,
    'hipaa': 3,
    'gdpr': 4,
    'dpdpa': 5,
    'pci dss': 6
  };

  const frameworks = Object.keys(frameworkCounts)
    .map(function(key) {
      return {
        framework: frameworkLabels[key],
        controlCount: frameworkCounts[key]
      };
    })
    .sort(function(a, b) {
      const ak = normalizeCrossFrameworkName_(a.framework);
      const bk = normalizeCrossFrameworkName_(b.framework);
      const ao = preferredOrder[ak] || 99;
      const bo = preferredOrder[bk] || 99;
      if (ao !== bo) return ao - bo;
      return a.framework.localeCompare(b.framework);
    });

  const mappedSoc2Rows = controls.filter(function(control) {
    return splitFrameworkMappingRefsV20_(control.soc2CCMapping).length > 0;
  }).length;

  const mappedHipaaRows = controls.filter(function(control) {
    return splitFrameworkMappingRefsV20_(control.hipaaMapping).length > 0;
  }).length;

  return {
    success: true,
    version: 'V20',
    frameworks: frameworks,
    plannedFrameworks: ['GDPR', 'DPDPA', 'PCI DSS'],
    mappingSource: 'Framework_Control_Master',
    mappingMode: 'master',
    mappedSoc2Rows: mappedSoc2Rows,
    mappedHipaaRows: mappedHipaaRows,
    mappingError: ''
  };
}

function getCrossFrameworkMappingsV20_(sourceFramework, sourceControlId) {
  rbacRequirePermission_('viewCrossFramework');

  const requestedFramework = cleanString_(sourceFramework);
  const requestedControlId = cleanString_(sourceControlId);

  if (!requestedFramework || !requestedControlId) {
    throw new Error('Framework and Control ID are required for cross-framework mapping.');
  }

  const controls = readFrameworkControlMaster_().controls;
  const sourceKey = crossFrameworkKey_(requestedFramework, requestedControlId);
  const source = controls.find(function(control) {
    return crossFrameworkKey_(control.framework, control.controlId) === sourceKey;
  });

  if (!source) {
    throw new Error(
      'Control ' + requestedControlId +
      ' was not found for framework ' + requestedFramework + '.'
    );
  }

  const sourceFrameworkKey = normalizeCrossFrameworkName_(source.framework);
  const sourceRefs = frameworkReferenceBundleV20_(source);
  const explicit = [];
  const seen = {};

  controls.forEach(function(target) {
    const targetKey = crossFrameworkKey_(target.framework, target.controlId);
    const targetFrameworkKey = normalizeCrossFrameworkName_(target.framework);

    // Cross-framework means we do not return siblings from the same framework.
    if (!targetKey || targetKey === sourceKey || targetFrameworkKey === sourceFrameworkKey) return;

    const targetRefs = frameworkReferenceBundleV20_(target);
    const socMatches = intersectFrameworkRefsV20_(sourceRefs.soc2, targetRefs.soc2);
    const hipaaMatches = intersectFrameworkRefsV20_(sourceRefs.hipaa, targetRefs.hipaa);

    const isoMatches = intersectFrameworkRefsV20_(sourceRefs.iso, targetRefs.iso);

    let basis = '';
    let shared = [];

    if (socMatches.length) {
      basis = 'SOC 2 CC Mapping';
      shared = socMatches;
    }

    if (hipaaMatches.length) {
      basis = basis ? 'SOC 2 CC Mapping + HIPAA Mapping' : 'HIPAA Mapping';
      shared = shared.concat(hipaaMatches.filter(function(ref) {
        return shared.indexOf(ref) === -1;
      }));
    }

    if (isoMatches.length) {
      basis = basis ? basis + ' + ISO 27001 Mapping' : 'ISO 27001 Mapping';
      shared = shared.concat(isoMatches.filter(function(ref) { return shared.indexOf(ref) === -1; }));
    }
    if (!basis || seen[targetKey]) return;
    seen[targetKey] = true;
    explicit.push(mappedControlFromMasterV20_(target, basis, shared, source.commonControlArea, 'master'));
  });

  explicit.sort(function(a, b) {
    const frameworkCompare = a.framework.localeCompare(b.framework);
    if (frameworkCompare !== 0) return frameworkCompare;
    return a.controlId.localeCompare(b.controlId, undefined, {
      numeric: true,
      sensitivity: 'base'
    });
  });

  if (explicit.length) {
    return {
      success: true,
      sourceFramework: source.framework,
      sourceControlId: source.controlId,
      sourceControl: source,
      commonControlArea: source.commonControlArea,
      mappingMode: 'master',
      mappingSource: 'Framework_Control_Master',
      controls: explicit
    };
  }

  // Only use Common Control Area when explicit mapping columns yield nothing.
  const area = cleanString_(source.commonControlArea);
  const fallback = [];
  const seenArea = {};

  if (area) {
    controls.forEach(function(target) {
      const targetKey = crossFrameworkKey_(target.framework, target.controlId);
      const targetFrameworkKey = normalizeCrossFrameworkName_(target.framework);
      if (!targetKey || targetKey === sourceKey || targetFrameworkKey === sourceFrameworkKey) return;

      const sameArea = cleanString_(target.commonControlArea).toLowerCase() === area.toLowerCase();
      if (!sameArea || seenArea[targetKey]) return;

      seenArea[targetKey] = true;
      fallback.push(mappedControlFromMasterV20_(target, 'Common Control Area', [], area, 'area'));
    });
  }

  fallback.sort(function(a, b) {
    const frameworkCompare = a.framework.localeCompare(b.framework);
    if (frameworkCompare !== 0) return frameworkCompare;
    return a.controlId.localeCompare(b.controlId, undefined, {
      numeric: true,
      sensitivity: 'base'
    });
  });

  return {
    success: true,
    sourceFramework: source.framework,
    sourceControlId: source.controlId,
    sourceControl: source,
    commonControlArea: area,
    mappingMode: 'area',
    mappingSource: 'Framework_Control_Master',
    controls: fallback
  };
}

function testCrossFrameworkExplorerV20_() {
  rbacRequire_('admin');
  rbacRequire_('admin');

  const bootstrap = getCrossFrameworkExplorerBootstrapV20_();
  const controls = readFrameworkControlMaster_().controls;

  if (!controls.length) {
    throw new Error('Framework_Control_Master has no active controls.');
  }

  const preferred = controls.find(function(control) {
    return normalizeCrossFrameworkName_(control.framework) === 'iso 27001:2022';
  }) || controls[0];

  const mapping = getCrossFrameworkMappingsV20_(preferred.framework, preferred.controlId);

  return {
    success: true,
    version: 'V20',
    frameworkCount: bootstrap.frameworks.length,
    mappingSource: bootstrap.mappingSource,
    testControl: preferred.framework + ' / ' + preferred.controlId,
    mappingMode: mapping.mappingMode,
    mappedControls: mapping.controls.length
  };
}

/** Read-only recurrence picker: closed records are eligible; deleted rows are excluded by the master reader. */
function getPreviousRecordOptions_(currentRecordId) {
  rbacRequireAnyPermission_(['createRecords','editRecords','viewRecords']);
  const current = cleanString_(currentRecordId).toLowerCase();
  const seen = {};
  return readMasterRecords_().records.filter(function(record) {
    const id = cleanString_(record['Record ID']).toLowerCase();
    if (!id || id === current || seen[id]) return false;
    seen[id] = true;
    return true;
  }).map(function(record) {
    return {'Record ID':cleanString_(record['Record ID']),
      'Record Type':cleanString_(record['Record Type']),
      'Record Title':cleanString_(record['Record Title'])};
  }).sort(function(a,b) {return a['Record ID'].localeCompare(b['Record ID'],undefined,{numeric:true});});
}

// Shared dashboard/report semantics. These helpers never change master records.
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

function getAnalyticsTimeZone_() {
  return getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Kolkata';
}
function getAnalyticsRecords_(filters) {
  rbacRequirePermission_('viewAnalytics');
  return {records:getRecords_(filters || {}), timeZone:getAnalyticsTimeZone_()};
}

/***************************************************************
 * IAM — department scope, responses and private evidence
 * Run initializeIamPhase once as an existing Admin.
 * Department accounts require an owner-executed web app with a
 * verified active user identity, and no direct master Sheet access.
 ***************************************************************/
const IAM_RESPONSE_SHEET = 'IAM_Responses';
const IAM_REVIEW_SHEET = 'IAM_Response_Reviews';
const IAM_RESPONSE_HEADERS = ['Submission ID','Record ID','Department','Submitted By','Display Name','Response','Evidence JSON','Submitted At','Status','Reviewed By','Reviewer Name','Reviewed At','Review Comment','Client Token'];
const IAM_REVIEW_HEADERS = ['Event ID','Submission ID','Record ID','Decision','Reviewer','Reviewer Name','Comment','Timestamp'];
const IAM_MAX_EVIDENCE_BYTES = 5 * 1024 * 1024;

function iamParseDepartments_(value) {
  if (Array.isArray(value)) return value.map(function(v){ return String(v).trim(); }).filter(Boolean);
  if (!value) return [];
  try { const parsed = JSON.parse(String(value)); return Array.isArray(parsed) ? iamParseDepartments_(parsed) : []; }
  catch (ignore) { return []; } // Invalid scope never becomes global access.
}
function iamDepartmentKey_(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g,' '); }
function iamDepartmentOffered_(value) {
  return ['quality','security operations','security ops','human resources','human resource','information security','info sec','infosec','compliance'].indexOf(iamDepartmentKey_(value)) === -1;
}
function iamCanReadRecord_(access, record) {
  if (!access.departmentScoped) return true;
  const key = iamDepartmentKey_(record && record.Department);
  return !!key && access.departments.some(function(dept){ return iamDepartmentKey_(dept) === key; });
}
function iamScopeRecords_(records, access) {
  access = access || rbacRequire_('read');
  return records.filter(function(record){ return iamCanReadRecord_(access, record); });
}
function iamRequireRecord_(recordId, access) {
  access = access || rbacRequire_('read');
  const id = String(recordId || '').trim();
  const matches = readMasterRecords_().records.filter(function(record){ return String(record['Record ID']).trim() === id && iamCanReadRecord_(access,record); });
  if (!id || matches.length !== 1) throw new Error('Record is unavailable or outside your permitted departments.');
  return matches[0];
}
function iamAssertStorageOwner_() {
  const owner = PropertiesService.getScriptProperties().getProperty('IAM_STORAGE_OWNER');
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!owner || effective !== owner) throw new Error('IAM requires the web app to execute as its configured owner. Ask Admin to check deployment; do not grant department users master Sheet access.');
}
function iamDepartmentCatalog_() {
  const lookup = getLookupSheet_().getDataRange().getDisplayValues();
  const headers = (lookup[0] || []).map(cleanString_);
  const category = headers.indexOf('Lookup Category'), value = headers.indexOf('Value');
  if (category < 0 || value < 0) throw new Error('Lookup_Master department headers are missing.');
  const departments = lookup.slice(1).filter(function(row){ return row[category] === 'Department'; }).map(function(row){ return cleanString_(row[value]); });
  return uniqueSorted_(departments.concat(readMasterRecords_().records.map(function(record){ return cleanString_(record.Department); }))).filter(function(name){ return phase3Value_(name) && !/^(all|\*)$/i.test(name) && iamDepartmentOffered_(name); });
}
function getIAMDepartmentOptions_() { rbacRequirePermission_('assignRoles'); return iamDepartmentCatalog_(); }
function iamValidateDepartments_(values) {
  if (!Array.isArray(values) || !values.length || values.length > 20) throw new Error('Assign at least one department (maximum 20).');
  const catalog = iamDepartmentCatalog_();
  const result = [];
  values.forEach(function(value){
    const match = catalog.find(function(dept){ return iamDepartmentKey_(dept) === iamDepartmentKey_(value); });
    if (!match) throw new Error('Choose departments from the current department catalog.');
    if (result.indexOf(match) === -1) result.push(match);
  });
  return result;
}
function iamSheet_(name, headers) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet || sheet.getLastColumn() !== headers.length || sheet.getRange(1,1,1,headers.length).getDisplayValues()[0].some(function(h,i){ return h !== headers[i]; })) throw new Error('IAM is not initialized or its headers are invalid. Admin must run initializeIamPhase.');
  return sheet;
}
// TC-25 V65: opt-in Shared Drive evidence storage. The OWNER_PRIVATE mode
// remains the default and preserves every existing folder/permission check.
// SHARED_DRIVE mode requires the Advanced Drive service (Drive API v3) and a
// pinned Shared Drive ID. Never infer authorization from a folder name alone.
function iamEvidenceStorageMode_() {
  const value = String(PropertiesService.getScriptProperties()
    .getProperty('IAM_EVIDENCE_STORAGE_MODE') || 'OWNER_PRIVATE').trim().toUpperCase();
  if (value !== 'OWNER_PRIVATE' && value !== 'SHARED_DRIVE') {
    throw new Error('TC25_STORAGE_MODE_INVALID: Choose OWNER_PRIVATE or SHARED_DRIVE.');
  }
  return value;
}
function iamPrivateFolder_() {
  iamAssertStorageOwner_();
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty('IAM_EVIDENCE_FOLDER_ID') || '').trim();
  if (!id) throw new Error('IAM evidence storage is not initialized.');
  const folder = DriveApp.getFolderById(id);
  if (iamEvidenceStorageMode_() === 'SHARED_DRIVE') {
    const expectedDriveId = String(props.getProperty('IAM_SHARED_DRIVE_ID') || '').trim();
    const expectedLiveId = String(props.getProperty('TC25_TEST_LIVE_FOLDER_ID') || '').trim();
    if (!expectedDriveId || !expectedLiveId || id !== expectedLiveId) {
      throw new Error('TC25_SHARED_DRIVE_CONFIG: Pin the Shared Drive ID and approved live evidence folder.');
    }
    if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.get) {
      throw new Error('TC25_DRIVE_API_REQUIRED: Enable Advanced Google service Drive API v3.');
    }
    // The metadata check fails closed if the destination is not the intended
    // organization-owned Shared Drive. Existing session/RBAC checks remain.
    const meta = Drive.Files.get(id, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,driveId,trashed'
    });
    if (!meta || meta.trashed || meta.mimeType !== 'application/vnd.google-apps.folder' ||
        meta.driveId !== expectedDriveId || meta.id !== expectedLiveId ||
        meta.name !== '01_Live_Evidence') {
      throw new Error('TC25_SHARED_DRIVE_MISMATCH: Evidence destination is not the approved Shared Drive folder.');
    }
    return folder;
  }
  // Original private-folder protection — unchanged for existing deployments.
  const owner = props.getProperty('IAM_STORAGE_OWNER');
  const folderOwner = folder.getOwner();
  const otherMembers = folder.getEditors().concat(folder.getViewers()).filter(function(user){
    return String(user.getEmail() || '').trim().toLowerCase() !== owner;
  });
  if (!folderOwner || String(folderOwner.getEmail() || '').trim().toLowerCase() !== owner ||
      folder.getSharingAccess() !== DriveApp.Access.PRIVATE || otherMembers.length) {
    throw new Error('Evidence folder must remain private to the app owner. Remove broad or direct sharing before uploading evidence.');
  }
  return folder;
}
function initializeIamPhase_() {
  const actor = rbacRequire_('admin');
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (effective !== actor.email) throw new Error('Run initializeIamPhase from the Apps Script editor as the Admin who will own the deployment.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const oldOwner = props.getProperty('IAM_STORAGE_OWNER');
    if (oldOwner && oldOwner !== effective) throw new Error('IAM already belongs to another deployment owner. Its files and assignments were not changed.');
    const data = rbacRows_();
    if (data.sheet.getLastColumn() < RBAC_HEADERS.length) data.sheet.getRange(1,RBAC_HEADERS.length).setValue('Departments');
    [ [IAM_RESPONSE_SHEET,IAM_RESPONSE_HEADERS], [IAM_REVIEW_SHEET,IAM_REVIEW_HEADERS] ].forEach(function(pair){
      let sheet = getSpreadsheet_().getSheetByName(pair[0]);
      if (!sheet) { sheet = getSpreadsheet_().insertSheet(pair[0]); sheet.appendRow(pair[1]); sheet.setFrozenRows(1); }
      iamSheet_(pair[0],pair[1]);
    });
    props.setProperty('IAM_STORAGE_OWNER',effective);
    if (!props.getProperty('IAM_EVIDENCE_FOLDER_ID')) {
      const folder = DriveApp.createFolder('Digiclarity · IAM Evidence · Private');
      folder.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.VIEW);
      props.setProperty('IAM_EVIDENCE_FOLDER_ID',folder.getId());
    }
    iamPrivateFolder_();
    rbacAudit_(actor.email,actor.email,null,{action:'IAM INITIALIZED',owner:effective});
    return 'IAM ready. Existing access is preserved. Use Department User + assigned departments for the pilot; keep master Sheet and evidence folder private.';
  } finally { lock.releaseLock(); }
}
function iamResponseRows_() {
  const sheet = iamSheet_(IAM_RESPONSE_SHEET,IAM_RESPONSE_HEADERS);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1).map(function(row,index){
    const item = {_row:index+2}; IAM_RESPONSE_HEADERS.forEach(function(h,i){ item[h] = row[i] || ''; }); return item;
  }).filter(function(item){ return !!item['Submission ID']; });
}
function iamVisibleResponse_(response, record, access) {
  return !!record && iamCanReadRecord_(access,record) &&
    (!access.departmentScoped || iamCanReadRecord_(access,{Department:response.Department}));
}
function iamPublicResponse_(item) {
  let files = [];
  try { files = JSON.parse(item['Evidence JSON'] || '[]'); } catch (ignore) {}
  return {id:item['Submission ID'],recordId:item['Record ID'],department:item.Department,submittedBy:item['Submitted By'],name:item['Display Name'] || 'Team member',response:item.Response,
    files:Array.isArray(files) ? files.map(function(file){ return {id:file.id,name:file.name,size:file.size,mime:file.mime}; }) : [],
    submittedAt:item['Submitted At'],status:item.Status,reviewedBy:item['Reviewed By'],reviewerName:item['Reviewer Name'],reviewedAt:item['Reviewed At'],reviewComment:item['Review Comment']};
}
function getIAMWorkspace_() {
  const access = rbacRequireAnyPermission_(['viewRecords','submitEvidence','reviewResponses']);
  const records = readMasterRecords_().records;
  const byId = {}; records.forEach(function(record){ byId[record['Record ID']] = record; });
  const responses = iamResponseRows_().filter(function(item){ return iamVisibleResponse_(item,byId[item['Record ID']],access); }).sort(function(a,b){ return b['Submitted At'].localeCompare(a['Submitted At']); });
  return {access:access,timeZone:getAnalyticsTimeZone_(),generatedAt:formatGeneratedAt_(),
    records:records.map(function(record){ return {'Record ID':record['Record ID'],'Record Title':record['Record Title'],'Record Type':record['Record Type'],Department:record.Department,'Record Status':record['Record Status'],'Target Completion Date':record['Target Completion Date'],'Residual Risk Rating':phase3Risk_(record)}; }),
    responses:responses.map(iamPublicResponse_)};
}
function iamSafeCell_(value) { const text = String(value || ''); return /^[=+@-]/.test(text) ? "'" + text : text; }
function iamDecodeEvidence_(file) {
  const formats = {pdf:'application/pdf',png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',txt:'text/plain',csv:'text/csv',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'};
  const name = String(file && file.name || '').trim();
  const ext = name.split('.').pop().toLowerCase();
  if (!name || name.length > 180 || /[\x00-\x1f\\/]/.test(name) || !formats[ext]) throw new Error('Allowed evidence: PDF, PNG, JPG, TXT, CSV, DOCX or XLSX; filename must be under 180 characters.');
  const encoded = String(file.base64 || '');
  if (!encoded || encoded.length > Math.ceil(IAM_MAX_EVIDENCE_BYTES/3)*4 || !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded) || encoded.length % 4 !== 0) throw new Error('Invalid or oversized evidence file.');
  const bytes = Utilities.base64Decode(encoded);
  if (!bytes.length || bytes.length > IAM_MAX_EVIDENCE_BYTES) throw new Error('Evidence files must be nonempty, with a combined maximum of 5 MB.');
  const b = bytes.slice(0,8).map(function(v){ return (v+256)%256; });
  const text = b.map(function(v){ return String.fromCharCode(v); }).join('');
  if ((ext==='pdf' && text.indexOf('%PDF-')!==0) || (ext==='png' && b.join(',')!=='137,80,78,71,13,10,26,10') ||
      (['jpg','jpeg'].indexOf(ext)!==-1 && !(b[0]===255 && b[1]===216 && b[2]===255)) ||
      (['docx','xlsx'].indexOf(ext)!==-1 && !(b[0]===80 && b[1]===75 && b[2]===3 && b[3]===4)) ||
      (['txt','csv'].indexOf(ext)!==-1 && bytes.some(function(v){ return v===0; }))) throw new Error('Evidence contents do not match the file extension.');
  return {name:name,mime:formats[ext],bytes:bytes};
}
function submitIAMResponse_(input) {
  const access = rbacRequirePermission_('submitEvidence');
  input = input || {};
  const response = String(input.response || '').trim();
  const clientToken = String(input.clientToken || '');
  if (response.length < 5 || response.length > 4000) throw new Error('Enter a response between 5 and 4,000 characters.');
  if (!/^[a-zA-Z0-9_-]{16,100}$/.test(clientToken)) throw new Error('Submission token is missing. Reload the response form.');
  if (!Array.isArray(input.files) || input.files.length > 3) throw new Error('Attach up to three evidence files.');
  const decoded = input.files.map(iamDecodeEvidence_);
  if (decoded.reduce(function(total,file){ return total+file.bytes.length; },0) > IAM_MAX_EVIDENCE_BYTES) throw new Error('Combined evidence limit is 5 MB.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  const created = [];
  try {
    const record = iamRequireRecord_(input.recordId,access);
    if (phase3IsTerminal_(record)) throw new Error('Closed or cancelled records cannot receive new responses.');
    const duplicate = iamResponseRows_().find(function(item){ return item['Client Token']===clientToken && item['Submitted By']===access.email; });
    const evidenceRequest = e92PrepareSubmission_(input,record,access,duplicate);
    if (duplicate) {
      if (duplicate['Record ID']!==record['Record ID']) throw new Error('Submission token belongs to another record. Open a new response form.');
      return {success:true,id:duplicate['Submission ID'],message:'This response was already submitted.'};
    }
    iamAssertStorageOwner_();
    const sheet = iamSheet_(IAM_RESPONSE_SHEET,IAM_RESPONSE_HEADERS);
    const folder = decoded.length ? iamPrivateFolder_() : null;
    const id = (evidenceRequest ? 'ERS-' : 'IAM-') + Utilities.getUuid();
    const files = decoded.map(function(file){
      const blob = Utilities.newBlob(file.bytes,file.mime,id+'_'+file.name);
      const stored = folder.createFile(blob); created.push(stored);
      // On Shared Drives, access is inherited from the Shared Drive. Do not
      // call setSharing(PRIVATE), which cannot revoke inherited member access.
      // The approved Shared Drive membership must be audited separately.
      if (iamEvidenceStorageMode_() === 'OWNER_PRIVATE') {
        stored.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.VIEW);
      }
      return {id:stored.getId(),name:file.name,mime:file.mime,size:file.bytes.length};
    });
    const now = new Date().toISOString();
    e92SaveSubmission_(sheet,[id,record['Record ID'],iamSafeCell_(record.Department),access.email,iamSafeCell_(access.name),iamSafeCell_(response),JSON.stringify(files),now,'Submitted','','','','',clientToken],evidenceRequest);
    return {success:true,id:id,message:'Response submitted to Compliance for review.'};
  } catch (error) {
    created.forEach(function(file){ try { file.setTrashed(true); } catch (ignore) {} });
    throw error;
  } finally { lock.releaseLock(); }
}
function reviewIAMResponse_(input) {
  const access = rbacRequirePermission_('reviewResponses');
  input = input || {};
  if (['Accepted','More information required'].indexOf(input.decision)===-1) throw new Error('Choose Accept or Request more information.');
  const comment = String(input.comment || '').trim();
  if (comment.length > 2000 || (input.decision==='More information required' && comment.length < 5)) throw new Error('Explain the additional information needed (5–2,000 characters).');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const item = iamResponseRows_().find(function(row){ return row['Submission ID']===String(input.id || ''); });
    if (!item) throw new Error('Submission not found.');
    iamRequireRecord_(item['Record ID'],access);
    if (item['Submitted By']===access.email) throw new Error('Another Compliance reviewer must review your own submission.');
    if (item.Status!=='Submitted') {
      if (item.Status===input.decision && item['Reviewed By']===access.email && item['Review Comment']===comment) return {success:true,message:'This review was already saved.'};
      throw new Error('This submission was already reviewed. Refresh the workspace.');
    }
    e91ValidateReview_(item,input.decision,access);
    e92ValidateReview_(item,access);
    const reviews = iamSheet_(IAM_REVIEW_SHEET,IAM_REVIEW_HEADERS);
    const sheet = iamSheet_(IAM_RESPONSE_SHEET,IAM_RESPONSE_HEADERS);
    const before = sheet.getRange(item._row,1,1,IAM_RESPONSE_HEADERS.length).getValues();
    const after = before[0].slice();
    const now = new Date().toISOString();
    [ ['Status',input.decision],['Reviewed By',access.email],['Reviewer Name',iamSafeCell_(access.name)],['Reviewed At',now],['Review Comment',iamSafeCell_(comment)] ].forEach(function(pair){ after[IAM_RESPONSE_HEADERS.indexOf(pair[0])] = pair[1]; });
    sheet.getRange(item._row,1,1,IAM_RESPONSE_HEADERS.length).setValues([after]);
    try { reviews.appendRow(['IAM-REVIEW-'+Utilities.getUuid(),item['Submission ID'],item['Record ID'],input.decision,access.email,iamSafeCell_(access.name),iamSafeCell_(comment),now]); }
    catch (error) { sheet.getRange(item._row,1,1,IAM_RESPONSE_HEADERS.length).setValues(before); throw new Error('Review was rolled back because its audit entry could not be saved.'); }
    return {success:true,message:'Response review saved. Main record and approval workflow are unchanged.'};
  } finally { lock.releaseLock(); }
}
function getIAMEvidence_(submissionId,fileId) {
  const access = rbacRequireAnyPermission_(['viewRecords','submitEvidence','reviewResponses']);
  const item = iamResponseRows_().find(function(row){ return row['Submission ID']===String(submissionId || ''); });
  if (!item) throw new Error('Evidence is unavailable.');
  const record = iamRequireRecord_(item['Record ID'],access);
  if (!iamVisibleResponse_(item,record,access)) throw new Error('Evidence is outside your department access.');
  const metadata = iamPublicResponse_(item).files.find(function(file){ return file.id===String(fileId || ''); });
  if (!metadata) throw new Error('Evidence is unavailable.');
  iamAssertStorageOwner_();
  const file = DriveApp.getFileById(metadata.id);
  if (file.isTrashed()) throw new Error('Evidence is unavailable.');
  const bytes = file.getBlob().getBytes();
  if (bytes.length > IAM_MAX_EVIDENCE_BYTES) throw new Error('Evidence exceeds the download limit.');
  return {name:metadata.name,mime:metadata.mime,base64:Utilities.base64Encode(bytes)};
}

/***************************************************************
 * PHASE 8.3 — EMAIL/PASSWORD LOGIN, PERSISTENT SESSIONS & OTP RESET
 *
 * The typed email must match the Google Workspace identity already
 * signed into the Sites/App Script container. This keeps the existing
 * server-side RBAC identity model intact while adding a dashboard-local
 * credential layer and persistent dashboard session.
 ***************************************************************/
function dashboardScriptSecret_(key) {
  const props = PropertiesService.getScriptProperties();
  let value = props.getProperty(key);
  if (!value) {
    value = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(key, value);
  }
  return value;
}

function dashboardHmac_(value, key) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(String(value || ''), dashboardScriptSecret_(key))
  );
}

function dashboardNormalizeLoginEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function dashboardValidateLoginIdentity_(email) {
  const typed = dashboardNormalizeLoginEmail_(email);
  if (!rbacAllowedEmail_(typed)) {
    throw new Error('Enter your assigned @digiclarity.in or @clarityrcm.com email address.');
  }
  const googleEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!googleEmail) {
    throw new Error('Your Google Workspace identity could not be detected. Open the dashboard while signed in with your assigned company Google account.');
  }
  if (typed !== googleEmail) {
    throw new Error('The dashboard email must match the Google Workspace account currently signed in to this site (' + googleEmail + ').');
  }
  return rbacAccessForEmail_(typed);
}

function getDashboardLoginIdentity() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email || !rbacAllowedEmail_(email)) return {email:'',name:''};
  try {
    const access = rbacAccessForEmail_(email);
    return {email:access.email,name:access.name || '',role:access.role || ''};
  } catch (error) {
    return {email:email,name:'',role:'',message:error && error.message ? error.message : ''};
  }
}

/* ---------- Persistent dashboard sessions ---------- */
function dashboardSessionSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(DASHBOARD_SESSION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_SESSION_SHEET);
    sheet.getRange(1,1,1,DASHBOARD_SESSION_HEADERS.length).setValues([DASHBOARD_SESSION_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = Math.max(sheet.getLastColumn(), DASHBOARD_SESSION_HEADERS.length);
  const row = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  for (let i=0;i<DASHBOARD_SESSION_HEADERS.length;i++) {
    if (row[i] && row[i] !== DASHBOARD_SESSION_HEADERS[i]) {
      throw new Error('Dashboard_Sessions headers are invalid. Contact the Compliance Administrator.');
    }
  }
  sheet.getRange(1,1,1,DASHBOARD_SESSION_HEADERS.length).setValues([DASHBOARD_SESSION_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function dashboardSessionRows_() {
  const sheet = dashboardSessionSheet_();
  if (sheet.getLastRow() < 2) return {sheet:sheet,rows:[]};
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,DASHBOARD_SESSION_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(row,index){
    return {
      sessionId:String(row[0]||''), email:String(row[1]||'').trim().toLowerCase(), tokenHash:String(row[2]||''),
      createdAt:row[3] instanceof Date ? row[3].toISOString() : String(row[3]||''),
      lastSeenAt:row[4] instanceof Date ? row[4].toISOString() : String(row[4]||''),
      revokedAt:row[5] instanceof Date ? row[5].toISOString() : String(row[5]||''),
      revokeReason:String(row[6]||''), credentialUpdatedAt:row[7] instanceof Date ? row[7].toISOString() : String(row[7]||''),
      displayName:String(row[8]||''), role:String(row[9]||''), browser:String(row[10]||''), platform:String(row[11]||''),
      device:String(row[12]||''), userAgent:String(row[13]||''), row:index+2
    };
  }).filter(function(row){return !!row.sessionId;})};
}

function dashboardSessionTokenHash_(sessionId, secret, email) {
  return dashboardHmac_(String(sessionId||'')+'\n'+String(secret||'')+'\n'+String(email||'').toLowerCase(), DASHBOARD_SESSION_PEPPER_KEY);
}

function dashboardCreateSession_(email, clientMeta) {
  email = dashboardNormalizeLoginEmail_(email);
  const credential = dashboardCredentialForEmail_(email);
  if (!credential) throw new Error('Dashboard password is not configured for this account.');
  const access = rbacAccessForEmail_(email);
  const meta = dashboardClientMeta_(clientMeta);
  const sessionId = Utilities.getUuid();
  const secret = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  const tokenHash = dashboardSessionTokenHash_(sessionId, secret, email);
  const now = new Date().toISOString();
  dashboardSessionSheet_().appendRow([
    sessionId,email,tokenHash,now,now,'','',String(credential.updatedAt||''),
    String(access.name||''),String(access.role||''),meta.browser,meta.platform,meta.device,meta.userAgent
  ]);
  return {token:sessionId+'.'+secret,sessionId:sessionId,createdAt:now};
}

function dashboardParseSessionToken_(token) {
  const value = String(token || '');
  const pos = value.indexOf('.');
  if (pos <= 0) return null;
  return {sessionId:value.slice(0,pos),secret:value.slice(pos+1)};
}

function dashboardRevokeAllSessionsForEmail_(email, reason) {
  email = dashboardNormalizeLoginEmail_(email);
  const data = dashboardSessionRows_();
  const now = new Date().toISOString();
  data.rows.forEach(function(row){
    if (row.email === email && !row.revokedAt) {
      data.sheet.getRange(row.row,6,1,2).setValues([[now,String(reason||'REVOKED')]]);
    }
  });
}

function resumeDashboardSession(token, clientMeta) {
  const parsed = dashboardParseSessionToken_(token);
  if (!parsed) return {success:false,code:'SESSION_INVALID',message:'Dashboard session is invalid.'};
  let email = '';
  try { email = rbacIdentity_(); } catch (error) {
    return {success:false,code:'IDENTITY_UNAVAILABLE',message:error && error.message ? error.message : 'Company identity could not be verified.'};
  }
  const data = dashboardSessionRows_();
  const matches = data.rows.filter(function(row){return row.sessionId === parsed.sessionId;});
  if (matches.length !== 1) return {success:false,code:'SESSION_INVALID',message:'Dashboard session was not found.'};
  const row = matches[0];
  if (row.email !== email || row.revokedAt) return {success:false,code:'SESSION_REVOKED',message:'Dashboard session is no longer active.'};
  const lastActivityMs = new Date(row.lastSeenAt || row.createdAt).getTime();
  if (!Number.isFinite(lastActivityMs) || Date.now() - lastActivityMs >= 300000) {
    data.sheet.getRange(row.row,6,1,2).setValues([[new Date().toISOString(),'IDLE_TIMEOUT_5_MIN']]);
    dashboardSecurityAudit_('SESSION_IDLE_TIMEOUT','SUCCESS',email,email,row.sessionId,row.role,clientMeta,{idleMinutes:5});
    return {success:false,code:'SESSION_IDLE_TIMEOUT',message:'Dashboard locked after 5 minutes of inactivity. Sign in again.'};
  }
  const expected = dashboardSessionTokenHash_(row.sessionId, parsed.secret, row.email);
  if (!dashboardPasswordEqual_(expected,row.tokenHash)) return {success:false,code:'SESSION_INVALID',message:'Dashboard session is invalid.'};

  let access;
  try { access = rbacAccessForEmail_(email); }
  catch (error) {
    const now = new Date().toISOString();
    data.sheet.getRange(row.row,6,1,2).setValues([[now,'ACCESS_REMOVED']]);
    dashboardSecurityAudit_('SESSION_REVOKED','DENIED',email,email,row.sessionId,row.role,clientMeta,{reason:'ACCESS_REMOVED'});
    return {success:false,code:'ACCESS_REMOVED',message:error && error.message ? error.message : 'Dashboard access is no longer active.'};
  }
  const credential = dashboardCredentialForEmail_(email);
  if (!credential || String(credential.updatedAt||'') !== String(row.credentialUpdatedAt||'')) {
    const now = new Date().toISOString();
    data.sheet.getRange(row.row,6,1,2).setValues([[now,'CREDENTIAL_CHANGED']]);
    dashboardSecurityAudit_('SESSION_REVOKED','DENIED',email,email,row.sessionId,access.role,clientMeta,{reason:'CREDENTIAL_CHANGED'});
    return {success:false,code:'SESSION_REVOKED',message:'Your dashboard credential changed. Sign in again.'};
  }
  if (dashboardCredentialExpired_(credential)) {
    const now = new Date().toISOString();
    data.sheet.getRange(row.row,6,1,2).setValues([[now,'PASSWORD_EXPIRED']]);
    dashboardSecurityAudit_('SESSION_REVOKED','DENIED',email,email,row.sessionId,access.role,clientMeta,{reason:'PASSWORD_EXPIRED'});
    return {success:false,code:'PASSWORD_EXPIRED',message:'Your dashboard password has expired. Sign in and change it to continue.'};
  }
  const meta = dashboardClientMeta_(clientMeta);
  const seen = new Date().toISOString();
  data.sheet.getRange(row.row,5).setValue(seen);
  if (clientMeta) {
    data.sheet.getRange(row.row,9,1,6).setValues([[
      String(access.name||row.displayName||''),String(access.role||row.role||''),meta.browser,meta.platform,meta.device,meta.userAgent
    ]]);
  }
  return {success:true,code:'SESSION_RESUMED',message:'Dashboard session restored.',access:access,sessionId:row.sessionId,lastSeenAt:seen};
}

function revokeDashboardSession(token, reason, clientMeta) {
  const parsed = dashboardParseSessionToken_(token);
  if (!parsed) return {success:true};
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const data = dashboardSessionRows_();
  const row = data.rows.find(function(item){return item.sessionId===parsed.sessionId && item.email===email;});
  if (!row) return {success:true};
  const expected = dashboardSessionTokenHash_(row.sessionId, parsed.secret, row.email);
  if (!dashboardPasswordEqual_(expected,row.tokenHash)) return {success:true};
  if (!row.revokedAt) {
    const now = new Date().toISOString();
    const revokeReason = String(reason || 'USER_SIGN_OUT').slice(0,80);
    data.sheet.getRange(row.row,6,1,2).setValues([[now,revokeReason]]);
    dashboardSecurityAudit_(revokeReason === 'USER_LOCK' ? 'DASHBOARD_LOCK' : 'SIGN_OUT','SUCCESS',email,email,row.sessionId,row.role,clientMeta,{reason:revokeReason});
  }
  return {success:true};
}

function loginDashboardWithCredentials(input) {
  input = input || {};
  const meta = dashboardClientMeta_(input.clientMeta);
  let access;
  try { access = dashboardValidateLoginIdentity_(input.email); }
  catch (error) {
    dashboardSecurityAudit_('ACCESS_DENIED','DENIED',String(Session.getActiveUser().getEmail()||''),String(input.email||''),'','',meta,{message:error.message||String(error)});
    return {success:false,code:'LOGIN_IDENTITY_ERROR',message:error.message || String(error)};
  }

  const credential = dashboardCredentialForEmail_(access.email);
  if (!credential) {
    dashboardSecurityAudit_('FIRST_TIME_LOGIN_REQUIRED','INFO',access.email,access.email,'',access.role,meta,{});
    return {success:false,code:'PASSWORD_SETUP_REQUIRED',message:'No dashboard password is configured yet. Use First-time login to create your dashboard password.'};
  }
  const verified = verifyDashboardPassword({password:String(input.password||'')});
  if (!verified.success) {
    dashboardSecurityAudit_('LOGIN_FAILED','FAILED',access.email,access.email,'',access.role,meta,{code:verified.code||'LOGIN_FAILED'});
    return verified;
  }
  const session = dashboardCreateSession_(access.email,meta);
  dashboardSecurityAudit_('LOGIN_SUCCESS','SUCCESS',access.email,access.email,session.sessionId,access.role,meta,{});
  return {success:true,code:'AUTHORIZED',message:'Dashboard sign-in successful.',access:verified.access,sessionToken:session.token};
}

/* TC-19 / V60: serialize OTP issuing, verification, and password reset.
 * Flush any pending Sheets writes while holding the lock. This ensures a
 * subsequent request cannot read an earlier OTP state after lock release.
 */
function dashboardResetWithLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

/* ---------- Password reset OTP ---------- */
function dashboardResetSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(DASHBOARD_RESET_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_RESET_SHEET);
    sheet.getRange(1,1,1,DASHBOARD_RESET_HEADERS.length).setValues([DASHBOARD_RESET_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = Math.max(sheet.getLastColumn(), DASHBOARD_RESET_HEADERS.length);
  const row = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  for (let i=0;i<DASHBOARD_RESET_HEADERS.length;i++) {
    if (row[i] && row[i] !== DASHBOARD_RESET_HEADERS[i]) throw new Error('Dashboard_Password_Reset headers are invalid. Contact the Compliance Administrator.');
  }
  sheet.getRange(1,1,1,DASHBOARD_RESET_HEADERS.length).setValues([DASHBOARD_RESET_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function dashboardResetRows_() {
  const sheet = dashboardResetSheet_();
  if (sheet.getLastRow() < 2) return {sheet:sheet,rows:[]};
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,DASHBOARD_RESET_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(row,index){
    return {
      email:String(row[0]||'').trim().toLowerCase(), otpSalt:String(row[1]||''), otpHash:String(row[2]||''),
      otpExpiresAt:row[3] instanceof Date ? row[3].toISOString() : String(row[3]||''), otpAttempts:Number(row[4]||0)||0,
      windowStartedAt:row[5] instanceof Date ? row[5].toISOString() : String(row[5]||''), windowSendCount:Number(row[6]||0)||0,
      lastSentAt:row[7] instanceof Date ? row[7].toISOString() : String(row[7]||''), resetTokenHash:String(row[8]||''),
      resetTokenExpiresAt:row[9] instanceof Date ? row[9].toISOString() : String(row[9]||''),
      verifiedAt:row[10] instanceof Date ? row[10].toISOString() : String(row[10]||''),
      completedAt:row[11] instanceof Date ? row[11].toISOString() : String(row[11]||''), row:index+2
    };
  }).filter(function(row){return !!row.email;})};
}

function dashboardResetForEmail_(email) {
  email = dashboardNormalizeLoginEmail_(email);
  const matches = dashboardResetRows_().rows.filter(function(row){return row.email===email;});
  if (matches.length > 1) throw new Error('Duplicate password-reset entries exist for this account. Contact the Compliance Administrator.');
  return matches[0] || null;
}

function dashboardGenerateOtp_() {
  const seed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid()+Utilities.getUuid()+String(Date.now()));
  let n = 0;
  for (let i=0;i<4;i++) n = ((n << 8) | (seed[i] & 255)) >>> 0;
  return String(n % 1000000).padStart(6,'0');
}

function dashboardOtpHash_(email, otp, salt) {
  return dashboardHmac_(String(email||'').toLowerCase()+'\n'+String(salt||'')+'\n'+String(otp||''), DASHBOARD_RESET_PEPPER_KEY);
}

function dashboardResetTokenHash_(email, token) {
  return dashboardHmac_(String(email||'').toLowerCase()+'\n'+String(token||''), DASHBOARD_RESET_PEPPER_KEY);
}

function dashboardMaskedEmail_(email) {
  const parts = String(email||'').split('@');
  if (parts.length !== 2) return email;
  const local = parts[0];
  const shown = local.length <= 2 ? local.charAt(0)+'*' : local.slice(0,2)+'*'.repeat(Math.min(8,Math.max(2,local.length-2)));
  return shown+'@'+parts[1];
}

function requestDashboardPasswordResetOtp(input) {
  input = input || {};
  return dashboardResetWithLock_(function() {
  const access = dashboardValidateLoginIdentity_(input.email);
  const email = access.email;
  if (!dashboardCredentialForEmail_(email)) {
    throw new Error('No dashboard password exists yet for this account. Use First-time login to create your dashboard password.');
  }
  const data = dashboardResetRows_();
  const existing = data.rows.find(function(row){return row.email===email;}) || null;
  const now = Date.now();
  if (existing && existing.lastSentAt) {
    const last = new Date(existing.lastSentAt).getTime();
    if (!isNaN(last) && now-last < DASHBOARD_RESET_RESEND_SECONDS*1000) {
      throw new Error('Please wait ' + Math.ceil((DASHBOARD_RESET_RESEND_SECONDS*1000-(now-last))/1000) + ' seconds before requesting another OTP.');
    }
  }
  let windowStarted = existing && existing.windowStartedAt ? new Date(existing.windowStartedAt).getTime() : 0;
  let windowCount = existing ? Number(existing.windowSendCount||0) : 0;
  if (!windowStarted || isNaN(windowStarted) || now-windowStarted >= 3600000) { windowStarted=now; windowCount=0; }
  if (windowCount >= DASHBOARD_RESET_MAX_SENDS_PER_HOUR) throw new Error('Too many OTP requests. Please wait before requesting another code.');

  const otp = dashboardGenerateOtp_();
  const salt = Utilities.getUuid().replace(/-/g,'');
  const otpHash = dashboardOtpHash_(email,otp,salt);
  const sentAt = new Date(now).toISOString();
  const expiresAt = new Date(now + DASHBOARD_RESET_OTP_MINUTES*60000).toISOString();
  const subject = 'DigiClarity Compliance Dashboard password reset code';
  const displayName = access.name || email;
  const plain = 'Hello '+displayName+',\n\nYour Compliance Dashboard password reset code is: '+otp+'\n\nThis code expires in '+DASHBOARD_RESET_OTP_MINUTES+' minutes. If you did not request this reset, you can ignore this email.\n\nDigiClarity Compliance';
  const html = '<div style="font-family:Arial,sans-serif;color:#17213a;line-height:1.6"><h2 style="margin:0 0 12px">Compliance Dashboard password reset</h2><p>Hello '+String(displayName).replace(/[<>&]/g,'')+',</p><p>Use the following one-time code to reset your Compliance Dashboard password:</p><div style="font-size:30px;font-weight:700;letter-spacing:8px;padding:16px 18px;border-radius:12px;background:#f2f5ff;display:inline-block">'+otp+'</div><p>This code expires in <b>'+DASHBOARD_RESET_OTP_MINUTES+' minutes</b>.</p><p style="color:#667085">If you did not request this reset, ignore this email. Your Google Workspace password is not affected.</p><p>DigiClarity Compliance</p></div>';
  MailApp.sendEmail({to:email,subject:subject,body:plain,htmlBody:html,name:'DigiClarity Compliance'});
  dashboardSecurityAudit_('PASSWORD_RESET_OTP_SENT','SUCCESS',email,email,'',access.role,input.clientMeta,{expiresMinutes:DASHBOARD_RESET_OTP_MINUTES});

  const row = [email,salt,otpHash,expiresAt,0,new Date(windowStarted).toISOString(),windowCount+1,sentAt,'','','',''];
  if (existing) data.sheet.getRange(existing.row,1,1,DASHBOARD_RESET_HEADERS.length).setValues([row]);
  else data.sheet.appendRow(row);
  try { rbacAudit_(email,email,null,{action:'DASHBOARD PASSWORD RESET OTP SENT',sentAt:sentAt}); } catch (error) {}
  return {success:true,message:'A 6-digit OTP was sent to '+dashboardMaskedEmail_(email)+'.',email:email,name:access.name||'',maskedEmail:dashboardMaskedEmail_(email),expiresMinutes:DASHBOARD_RESET_OTP_MINUTES};
  });
}

function verifyDashboardPasswordResetOtp(input) {
  input = input || {};
  return dashboardResetWithLock_(function() {
  const access = dashboardValidateLoginIdentity_(input.email);
  const email = access.email;
  const data = dashboardResetRows_();
  const row = data.rows.find(function(item){return item.email===email;});
  if (row && row.verifiedAt && row.resetTokenHash && !row.completedAt && !row.otpHash) {
    throw new Error('OTP_ALREADY_USED: This OTP was already verified. Continue in the original tab or request a new code.');
  }
  if (!row || !row.otpHash || row.completedAt) throw new Error('Request a new OTP before continuing.');
  const expires = new Date(row.otpExpiresAt).getTime();
  if (!expires || isNaN(expires) || expires <= Date.now()) throw new Error('This OTP has expired. Request a new code.');
  if (row.otpAttempts >= DASHBOARD_RESET_MAX_OTP_ATTEMPTS) throw new Error('Too many incorrect OTP attempts. Request a new code.');
  const otp = String(input.otp||'').replace(/\D/g,'');
  if (!/^\d{6}$/.test(otp)) throw new Error('Enter the 6-digit OTP sent to your email.');
  const candidate = dashboardOtpHash_(email,otp,row.otpSalt);
  if (!dashboardPasswordEqual_(candidate,row.otpHash)) {
    data.sheet.getRange(row.row,5).setValue(row.otpAttempts+1);
    dashboardSecurityAudit_('PASSWORD_RESET_OTP_FAILED','FAILED',email,email,'',access.role,input.clientMeta,{attempt:row.otpAttempts+1});
    throw new Error('Incorrect OTP. '+Math.max(0,DASHBOARD_RESET_MAX_OTP_ATTEMPTS-row.otpAttempts-1)+' attempt(s) remaining.');
  }
  const resetToken = Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  const tokenHash = dashboardResetTokenHash_(email,resetToken);
  const verifiedAt = new Date().toISOString();
  const tokenExpires = new Date(Date.now()+DASHBOARD_RESET_TOKEN_MINUTES*60000).toISOString();
  // Consume the OTP BEFORE returning authorization. A second tab cannot use it,
  // even if it submits the same correct digits within milliseconds.
  data.sheet.getRange(row.row,2,1,4).setValues([['','','',0]]);
  data.sheet.getRange(row.row,9,1,4).setValues([[tokenHash,tokenExpires,verifiedAt,'']]);
  try { rbacAudit_(email,email,null,{action:'DASHBOARD PASSWORD RESET OTP VERIFIED',verifiedAt:verifiedAt}); } catch (error) {}
  dashboardSecurityAudit_('PASSWORD_RESET_OTP_VERIFIED','SUCCESS',email,email,'',access.role,input.clientMeta,{});
  return {success:true,message:'OTP verified. Create your new dashboard password.',resetToken:resetToken,email:email,name:access.name||'',role:access.role||''};
  });
}

function completeDashboardPasswordReset(input) {
  input = input || {};
  return dashboardResetWithLock_(function() {
  const access = dashboardValidateLoginIdentity_(input.email);
  const email = access.email;
  const data = dashboardResetRows_();
  const reset = data.rows.find(function(item){return item.email===email;});
  if (!reset || !reset.resetTokenHash || reset.completedAt) throw new Error('Password reset authorization is unavailable. Request a new OTP.');
  const expiry = new Date(reset.resetTokenExpiresAt).getTime();
  if (!expiry || isNaN(expiry) || expiry <= Date.now()) throw new Error('Password reset authorization expired. Request a new OTP.');
  const providedHash = dashboardResetTokenHash_(email,String(input.resetToken||''));
  if (!dashboardPasswordEqual_(providedHash,reset.resetTokenHash)) throw new Error('Password reset authorization is invalid.');

  const password = dashboardPasswordPolicy_(input.newPassword,access);
  if (password !== String(input.confirmPassword||'')) throw new Error('The new passwords do not match.');
  const existing = dashboardCredentialForEmail_(email);
  if (existing) dashboardAssertPasswordNotReused_(password,existing);
  const result = dashboardWriteCredential_(email,email,password,existing);
  const completedAt = new Date().toISOString();
  // Redeem the reset grant once. Lock prevents simultaneous password commits.
  data.sheet.getRange(reset.row,9,1,4).setValues([['','',reset.verifiedAt,completedAt]]);
  dashboardRevokeAllSessionsForEmail_(email,'PASSWORD_RESET');
  const session = dashboardCreateSession_(email, input.clientMeta);
  dashboardSecurityAudit_('PASSWORD_RESET','SUCCESS',email,email,session.sessionId,access.role,input.clientMeta,{});
  try { rbacAudit_(email,email,{action:existing?'DASHBOARD PASSWORD ACTIVE':'NO DASHBOARD PASSWORD'},{action:'DASHBOARD PASSWORD RESET VIA OTP',updatedAt:result.updatedAt,expiresAt:result.expiresAt}); } catch (error) {}
  return {success:true,message:'Dashboard password updated successfully.',access:access,sessionToken:session.token,expiresAt:result.expiresAt};
  });
}

/***************************************************************
 * PHASE 8.3 — SESSION SECURITY & ACCESS AUDIT
 ***************************************************************/
function dashboardClientMeta_(input) {
  input = input || {};
  function clean(value, max) { return String(value || '').replace(/[\r\n\t]+/g,' ').trim().slice(0,max); }
  return {
    browser: clean(input.browser,80) || 'Unknown browser',
    platform: clean(input.platform,80) || 'Unknown platform',
    device: clean(input.device,40) || 'Unknown device',
    userAgent: clean(input.userAgent,300)
  };
}

function dashboardSecurityAuditSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(DASHBOARD_SECURITY_AUDIT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_SECURITY_AUDIT_SHEET);
    sheet.getRange(1,1,1,DASHBOARD_SECURITY_AUDIT_HEADERS.length).setValues([DASHBOARD_SECURITY_AUDIT_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = Math.max(sheet.getLastColumn(), DASHBOARD_SECURITY_AUDIT_HEADERS.length);
  const headers = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  for (let i=0;i<DASHBOARD_SECURITY_AUDIT_HEADERS.length;i++) {
    if (headers[i] && headers[i] !== DASHBOARD_SECURITY_AUDIT_HEADERS[i]) throw new Error('Dashboard_Security_Audit headers are invalid. Contact the Compliance Administrator.');
  }
  sheet.getRange(1,1,1,DASHBOARD_SECURITY_AUDIT_HEADERS.length).setValues([DASHBOARD_SECURITY_AUDIT_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function dashboardSecurityDetailsJson_(details) {
  try {
    const json = JSON.stringify(details || {});
    return json.length > 3000 ? json.slice(0,3000) : json;
  } catch (error) { return '{}'; }
}

function dashboardSecurityAudit_(eventType, result, actorEmail, targetEmail, sessionId, role, clientMeta, details) {
  try {
    const meta = dashboardClientMeta_(clientMeta);
    dashboardSecurityAuditSheet_().appendRow([
      Utilities.getUuid(), new Date().toISOString(), String(actorEmail||'').toLowerCase(), String(targetEmail||'').toLowerCase(),
      String(eventType||'SECURITY_EVENT').slice(0,80), String(result||'INFO').slice(0,30), String(sessionId||'').slice(0,100),
      String(role||'').slice(0,120), meta.browser, meta.platform, meta.device, dashboardSecurityDetailsJson_(details)
    ]);
  } catch (error) {
    console.log('Security audit write failed: ' + (error && error.message ? error.message : error));
  }
}

function dashboardSecurityAuditRows_(limit) {
  const sheet = dashboardSecurityAuditSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const count = Math.min(Math.max(1,Number(limit||DASHBOARD_SECURITY_AUDIT_LIMIT)), last-1);
  const start = Math.max(2,last-count+1);
  const values = sheet.getRange(start,1,last-start+1,DASHBOARD_SECURITY_AUDIT_HEADERS.length).getValues();
  return values.map(function(row){
    return {
      eventId:String(row[0]||''), timestamp:row[1] instanceof Date ? row[1].toISOString() : String(row[1]||''),
      actorEmail:String(row[2]||''), targetEmail:String(row[3]||''), eventType:String(row[4]||''), result:String(row[5]||''),
      sessionId:String(row[6]||''), role:String(row[7]||''), browser:String(row[8]||''), platform:String(row[9]||''),
      device:String(row[10]||''), detailsJson:String(row[11]||'')
    };
  }).reverse();
}

function dashboardSessionUserFallback_(email) {
  try {
    const row = rbacRows_().rows.find(function(item){return item.email===email;});
    return row ? {name:row.name||'',role:row.role||''} : {name:'',role:''};
  } catch (error) { return {name:'',role:''}; }
}

function dashboardSessionPublic_(row, currentSessionId) {
  const fallback = (!row.displayName || !row.role) ? dashboardSessionUserFallback_(row.email) : {name:'',role:''};
  return {
    sessionId:row.sessionId,email:row.email,displayName:row.displayName||fallback.name||'',role:row.role||fallback.role||'',
    createdAt:row.createdAt,lastSeenAt:row.lastSeenAt,revokedAt:row.revokedAt,revokeReason:row.revokeReason,
    browser:row.browser||'Unknown browser',platform:row.platform||'Unknown platform',device:row.device||'Unknown device',
    active:!row.revokedAt,current:!!currentSessionId && row.sessionId===currentSessionId
  };
}

function getDashboardSecurityCenter_() {
  const actor = rbacRequirePermission_('manageUsers');
  const sessions = dashboardSessionRows_().rows.map(function(row){return dashboardSessionPublic_(row,'');}).sort(function(a,b){
    return String(b.lastSeenAt||b.createdAt).localeCompare(String(a.lastSeenAt||a.createdAt));
  }).slice(0,250);
  const audits = dashboardSecurityAuditRows_(DASHBOARD_SECURITY_AUDIT_LIMIT);
  const active = sessions.filter(function(row){return row.active;});
  const activeUsers = {};
  active.forEach(function(row){activeUsers[row.email]=true;});
  const since = Date.now()-86400000;
  const recent = audits.filter(function(item){const t=new Date(item.timestamp).getTime();return !isNaN(t)&&t>=since;});
  const failed24h = recent.filter(function(item){return item.result==='FAILED'||item.result==='DENIED';}).length;
  const reset24h = recent.filter(function(item){return item.eventType.indexOf('PASSWORD_RESET')===0;}).length;
  return {success:true,actor:actor.email,summary:{activeSessions:active.length,activeUsers:Object.keys(activeUsers).length,failed24h:failed24h,reset24h:reset24h},sessions:sessions,audit:audits};
}

function adminRevokeDashboardSession_(sessionId, reason) {
  const actor = rbacRequirePermission_('manageUsers');
  const id = String(sessionId||'').trim();
  const data = dashboardSessionRows_();
  const row = data.rows.find(function(item){return item.sessionId===id;});
  if (!row) throw new Error('Dashboard session was not found. Refresh Security & Sessions.');
  if (row.revokedAt) return {success:true,message:'This dashboard session is already revoked.'};
  const now = new Date().toISOString();
  const why = String(reason||'ADMIN_REVOKE').trim().slice(0,200) || 'ADMIN_REVOKE';
  data.sheet.getRange(row.row,6,1,2).setValues([[now,'ADMIN_REVOKE']]);
  dashboardSecurityAudit_('ADMIN_SESSION_REVOKE','SUCCESS',actor.email,row.email,row.sessionId,row.role,null,{reason:why});
  return {success:true,message:'Session revoked for '+row.email+'.'};
}

function getMyDashboardSessions_(currentToken) {
  const access = rbacRequire_('read');
  const parsed = dashboardParseSessionToken_(currentToken);
  const currentId = parsed ? parsed.sessionId : '';
  const sessions = dashboardSessionRows_().rows.filter(function(row){return row.email===access.email && !row.revokedAt;}).map(function(row){
    return dashboardSessionPublic_(row,currentId);
  }).sort(function(a,b){return String(b.lastSeenAt||b.createdAt).localeCompare(String(a.lastSeenAt||a.createdAt));});
  return {success:true,sessions:sessions};
}

function revokeMyDashboardSession_(sessionId) {
  const access = rbacRequire_('read');
  const data = dashboardSessionRows_();
  const row = data.rows.find(function(item){return item.sessionId===String(sessionId||'') && item.email===access.email;});
  if (!row) throw new Error('Session not found.');
  if (!row.revokedAt) {
    data.sheet.getRange(row.row,6,1,2).setValues([[new Date().toISOString(),'USER_SESSION_REVOKE']]);
    dashboardSecurityAudit_('USER_SESSION_REVOKE','SUCCESS',access.email,access.email,row.sessionId,access.role,null,{});
  }
  return {success:true,message:'Session signed out.'};
}

function revokeOtherDashboardSessions_(currentToken) {
  const access = rbacRequire_('read');
  const parsed = dashboardParseSessionToken_(currentToken);
  if (!parsed) throw new Error('Current session could not be verified.');
  const data = dashboardSessionRows_();
  const current = data.rows.find(function(row){return row.sessionId===parsed.sessionId && row.email===access.email && !row.revokedAt;});
  if (!current) throw new Error('Current session is no longer active.');
  const expected = dashboardSessionTokenHash_(current.sessionId,parsed.secret,current.email);
  if (!dashboardPasswordEqual_(expected,current.tokenHash)) throw new Error('Current session could not be verified.');
  let count=0; const now=new Date().toISOString();
  data.rows.forEach(function(row){
    if(row.email===access.email && row.sessionId!==current.sessionId && !row.revokedAt){
      data.sheet.getRange(row.row,6,1,2).setValues([[now,'USER_REVOKE_OTHER_SESSIONS']]);count++;
    }
  });
  dashboardSecurityAudit_('REVOKE_OTHER_SESSIONS','SUCCESS',access.email,access.email,current.sessionId,access.role,null,{revokedCount:count});
  return {success:true,count:count,message:count ? (count+' other session'+(count===1?'':'s')+' signed out.') : 'No other active sessions were found.'};
}

function heartbeatDashboardSession(token, clientMeta) {
  return resumeDashboardSession(token,clientMeta);
}

// Optional one-time editor helper. Run only if Apps Script asks for the new
// MailApp permission after Phase 8.3 is deployed. It sends no email.
function authorizePhase83MailOnce_() {
  rbacRequire_('admin');
  return 'Mail authorization ready. Remaining daily email quota: ' + MailApp.getRemainingDailyQuota();
}

/***************************************************************
 * PHASE 8.4 — ACCESS GOVERNANCE & RECERTIFICATION
 * Periodic user-access review cycles with immutable access
 * snapshots, risk flags, reviewer decisions and audit evidence.
 ***************************************************************/
const ACCESS_REVIEW_CYCLE_SHEET = 'Access_Review_Cycles';
const ACCESS_REVIEW_CYCLE_HEADERS = [
  'Cycle ID','Cycle Name','Frequency','Start At','Due At','Status',
  'Created By','Created At','Completed At','Total Users','Reviewed Users','Changes Required'
];
const ACCESS_REVIEW_ITEM_SHEET = 'Access_Review_Items';
const ACCESS_REVIEW_ITEM_HEADERS = [
  'Item ID','Cycle ID','User Email','Display Name','Role','Departments JSON','Active',
  'Last Login At','Risk Flags JSON','Decision','Proposed Role','Proposed Departments JSON',
  'Reviewer Comment','Reviewed By','Reviewed At','Snapshot Updated At',
  'Previous Review At','Previous Role','Previous Departments JSON','Previous Active'
];

function accessReviewIso_(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString();
  const d = new Date(value);
  return isNaN(d.getTime()) ? String(value || '') : d.toISOString();
}

function accessReviewJsonArray_(value) {
  if (Array.isArray(value)) return value.map(function(v){return String(v||'').trim();}).filter(Boolean);
  const raw = String(value || '').trim();
  if (!raw) return [];
  try { const p = JSON.parse(raw); return Array.isArray(p) ? p.map(function(v){return String(v||'').trim();}).filter(Boolean) : []; }
  catch (error) { return raw.split(',').map(function(v){return v.trim();}).filter(Boolean); }
}

function accessReviewEnsureSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  for (let i=0;i<headers.length;i++) {
    if (current[i] && current[i] !== headers[i]) throw new Error(name + ' headers are invalid. Contact the Compliance Administrator.');
  }
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function accessReviewSheets_() {
  return {
    cycles:accessReviewEnsureSheet_(ACCESS_REVIEW_CYCLE_SHEET, ACCESS_REVIEW_CYCLE_HEADERS),
    items:accessReviewEnsureSheet_(ACCESS_REVIEW_ITEM_SHEET, ACCESS_REVIEW_ITEM_HEADERS)
  };
}

function accessReviewCycleRows_() {
  const sheet = accessReviewSheets_().cycles;
  if (sheet.getLastRow() < 2) return {sheet:sheet,rows:[]};
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,ACCESS_REVIEW_CYCLE_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(row,index){return {
    cycleId:String(row[0]||''), cycleName:String(row[1]||''), frequency:String(row[2]||''),
    startAt:accessReviewIso_(row[3]), dueAt:accessReviewIso_(row[4]), status:String(row[5]||''),
    createdBy:String(row[6]||''), createdAt:accessReviewIso_(row[7]), completedAt:accessReviewIso_(row[8]),
    totalUsers:Number(row[9]||0), reviewedUsers:Number(row[10]||0), changesRequired:Number(row[11]||0), row:index+2
  };}).filter(function(row){return !!row.cycleId;})};
}

function accessReviewItemRows_() {
  const sheet = accessReviewSheets_().items;
  if (sheet.getLastRow() < 2) return {sheet:sheet,rows:[]};
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,ACCESS_REVIEW_ITEM_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(row,index){return {
    itemId:String(row[0]||''), cycleId:String(row[1]||''), email:String(row[2]||'').trim().toLowerCase(),
    name:String(row[3]||''), role:String(row[4]||''), departments:accessReviewJsonArray_(row[5]),
    active:row[6]===true || String(row[6]).toLowerCase()==='true', lastLoginAt:accessReviewIso_(row[7]),
    riskFlags:accessReviewJsonArray_(row[8]), decision:String(row[9]||''), proposedRole:String(row[10]||''),
    proposedDepartments:accessReviewJsonArray_(row[11]), reviewerComment:String(row[12]||''),
    reviewedBy:String(row[13]||''), reviewedAt:accessReviewIso_(row[14]), snapshotUpdatedAt:accessReviewIso_(row[15]),
    previousReviewAt:accessReviewIso_(row[16]), previousRole:String(row[17]||''),
    previousDepartments:accessReviewJsonArray_(row[18]), previousActive:row[19]===true || String(row[19]).toLowerCase()==='true',
    row:index+2
  };}).filter(function(row){return !!row.itemId;})};
}

function accessReviewCycleName_(frequency, date) {
  const d = date instanceof Date ? date : new Date();
  const year = d.getFullYear();
  const month = d.getMonth();
  if (frequency === 'Half-Yearly') return (month < 6 ? 'H1 ' : 'H2 ') + year + ' Access Review';
  if (frequency === 'Annual') return year + ' Annual Access Review';
  return 'Q' + (Math.floor(month/3)+1) + ' ' + year + ' Access Review';
}

function accessReviewLatestPreviousMap_(items) {
  const map = {};
  (items || []).forEach(function(item){
    if (!item.email || !item.reviewedAt) return;
    const existing = map[item.email];
    if (!existing || String(item.reviewedAt) > String(existing.reviewedAt)) map[item.email] = item;
  });
  return map;
}

function accessReviewCredentialMap_() {
  const map = {};
  try { dashboardCredentialRows_().rows.forEach(function(row){map[row.email]=row;}); } catch (error) {}
  return map;
}

function accessReviewRoleMap_() {
  const map = {};
  rbacReadRoleDefinitions_().forEach(function(role){map[String(role.name||'').toLowerCase()]=role;});
  return map;
}

function accessReviewSameArray_(a,b) {
  const aa=(a||[]).map(String).sort(), bb=(b||[]).map(String).sort();
  return JSON.stringify(aa)===JSON.stringify(bb);
}

function accessReviewRiskFlags_(user, roleDef, lastLoginAt, previous) {
  const flags = [];
  if (!user.active) flags.push('INACTIVE_ACCOUNT');
  const permissions = roleDef ? rbacNormalizePermissions_(roleDef.permissions) : {};
  if (user.role === 'Admin' || permissions.manageUsers || permissions.assignRoles) flags.push('PRIVILEGED_ACCESS');
  if (roleDef && roleDef.scopeType === 'DEPARTMENTS' && !(user.departments||[]).length) flags.push('MISSING_DEPARTMENT_SCOPE');
  if (!lastLoginAt) flags.push('NEVER_LOGGED_IN');
  else {
    const age = Math.floor((Date.now()-new Date(lastLoginAt).getTime())/86400000);
    if (!isNaN(age)) {
      if (age >= 90) flags.push('DORMANT_90_DAYS');
      else if (age >= 60) flags.push('DORMANT_60_DAYS');
      else if (age >= 30) flags.push('DORMANT_30_DAYS');
    }
  }
  if (previous && (previous.role !== user.role || previous.active !== user.active || !accessReviewSameArray_(previous.departments,user.departments))) {
    flags.push('ACCESS_CHANGED_SINCE_REVIEW');
  }
  if (!roleDef) flags.push('ROLE_CONFIGURATION_MISSING');
  return flags;
}

function accessReviewPublicCycle_(cycle) {
  const overdue = cycle.status === 'Open' && cycle.dueAt && new Date(cycle.dueAt).getTime() < Date.now();
  return {
    cycleId:cycle.cycleId, cycleName:cycle.cycleName, frequency:cycle.frequency, startAt:cycle.startAt,
    dueAt:cycle.dueAt, status:overdue?'Overdue':cycle.status, storedStatus:cycle.status,
    createdBy:cycle.createdBy, createdAt:cycle.createdAt, completedAt:cycle.completedAt,
    totalUsers:cycle.totalUsers, reviewedUsers:cycle.reviewedUsers, changesRequired:cycle.changesRequired,
    pendingUsers:Math.max(0,cycle.totalUsers-cycle.reviewedUsers)
  };
}

function accessReviewPublicItem_(item) {
  return {
    itemId:item.itemId, cycleId:item.cycleId, email:item.email, name:item.name, role:item.role,
    departments:item.departments, active:item.active, lastLoginAt:item.lastLoginAt, riskFlags:item.riskFlags,
    decision:item.decision, proposedRole:item.proposedRole, proposedDepartments:item.proposedDepartments,
    reviewerComment:item.reviewerComment, reviewedBy:item.reviewedBy, reviewedAt:item.reviewedAt,
    snapshotUpdatedAt:item.snapshotUpdatedAt, previousReviewAt:item.previousReviewAt,
    previousRole:item.previousRole, previousDepartments:item.previousDepartments, previousActive:item.previousActive
  };
}

function startAccessReviewCycle_(input) {
  const actor = rbacRequirePermission_('manageUsers');
  input = input || {};
  const frequency = ['Quarterly','Half-Yearly','Annual'].indexOf(String(input.frequency||'')) !== -1 ? String(input.frequency) : 'Quarterly';
  const dueDays = [7,14,30,45].indexOf(Number(input.dueDays)) !== -1 ? Number(input.dueDays) : 14;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cyclesData = accessReviewCycleRows_();
    const open = cyclesData.rows.find(function(row){return row.status==='Open';});
    if (open) throw new Error('An access review cycle is already open: '+open.cycleName+'. Complete it before starting another cycle.');
    const users = rbacRows_().rows.slice().sort(function(a,b){return String(a.name||a.email).localeCompare(String(b.name||b.email));});
    if (!users.length) throw new Error('No dashboard users are available for access review.');
    const itemsData = accessReviewItemRows_();
    const previous = accessReviewLatestPreviousMap_(itemsData.rows);
    const credentialMap = accessReviewCredentialMap_();
    const roleMap = accessReviewRoleMap_();
    const now = new Date();
    const due = new Date(now.getTime()+dueDays*86400000);
    const cycleId = 'AR-' + Utilities.formatDate(now, Session.getScriptTimeZone()||'Asia/Kolkata','yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0,6).toUpperCase();
    const cycleName = accessReviewCycleName_(frequency,now);
    cyclesData.sheet.appendRow([cycleId,cycleName,frequency,now.toISOString(),due.toISOString(),'Open',actor.email,now.toISOString(),'',users.length,0,0]);
    const rows = users.map(function(user){
      const roleDef = roleMap[String(user.role||'').toLowerCase()] || null;
      const cred = credentialMap[user.email] || null;
      const lastLoginAt = cred ? cred.lastLoginAt : '';
      const prev = previous[user.email] || null;
      const flags = accessReviewRiskFlags_(user,roleDef,lastLoginAt,prev);
      return [Utilities.getUuid(),cycleId,user.email,user.name,user.role,JSON.stringify(user.departments||[]),user.active,lastLoginAt,
        JSON.stringify(flags),'','','[]','','','',user.updatedAt||'',prev?prev.reviewedAt:'',prev?prev.role:'',JSON.stringify(prev?prev.departments:[]),prev?prev.active:''];
    });
    if (rows.length) itemsData.sheet.getRange(itemsData.sheet.getLastRow()+1,1,rows.length,ACCESS_REVIEW_ITEM_HEADERS.length).setValues(rows);
    dashboardSecurityAudit_('ACCESS_REVIEW_CYCLE_STARTED','SUCCESS',actor.email,actor.email,'',actor.role,null,{cycleId:cycleId,cycleName:cycleName,totalUsers:users.length,dueAt:due.toISOString()});
    return getAccessReviewCenter_(cycleId);
  } finally { lock.releaseLock(); }
}

function getAccessReviewCenter_(cycleId) {
  const actor = rbacRequirePermission_('manageUsers');
  const cycles = accessReviewCycleRows_().rows.sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));});
  let selected = null;
  if (cycleId) selected = cycles.find(function(row){return row.cycleId===String(cycleId);}) || null;
  if (!selected) selected = cycles.find(function(row){return row.status==='Open';}) || cycles[0] || null;
  const items = selected ? accessReviewItemRows_().rows.filter(function(row){return row.cycleId===selected.cycleId;}) : [];
  return {
    success:true, actor:actor.email,
    cycles:cycles.slice(0,24).map(accessReviewPublicCycle_),
    cycle:selected?accessReviewPublicCycle_(selected):null,
    items:items.map(accessReviewPublicItem_),
    roles:rbacReadRoleDefinitions_().filter(function(role){return role.active;}).map(function(role){return {name:role.name,scopeType:role.scopeType};})
  };
}

function accessReviewUpdateCycleSummary_(cycleId) {
  const cyclesData = accessReviewCycleRows_();
  const cycle = cyclesData.rows.find(function(row){return row.cycleId===cycleId;});
  if (!cycle) throw new Error('Access review cycle not found.');
  const items = accessReviewItemRows_().rows.filter(function(row){return row.cycleId===cycleId;});
  const reviewed = items.filter(function(row){return !!row.decision;}).length;
  const changes = items.filter(function(row){return row.decision==='Modify'||row.decision==='Remove';}).length;
  const complete = items.length>0 && reviewed===items.length;
  cyclesData.sheet.getRange(cycle.row,6,1,7).setValues([[
    complete?'Completed':'Open',cycle.createdBy,cycle.createdAt,complete?(cycle.completedAt||new Date().toISOString()):'',items.length,reviewed,changes
  ]]);
  return {reviewed:reviewed,changes:changes,complete:complete,total:items.length};
}

function submitAccessReviewDecision_(input) {
  const actor = rbacRequirePermission_('manageUsers');
  input = input || {};
  const cycleId = String(input.cycleId||'').trim();
  const itemId = String(input.itemId||'').trim();
  const decision = String(input.decision||'').trim();
  const comment = String(input.comment||'').trim();
  if (['Retain','Modify','Remove'].indexOf(decision)===-1) throw new Error('Choose Retain Access, Modify Access, or Remove Access.');
  if ((decision==='Modify'||decision==='Remove') && comment.length<5) throw new Error('Add a reviewer comment of at least 5 characters for access changes.');
  if (comment.length>1500) throw new Error('Reviewer comment must be 1500 characters or fewer.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cyclesData = accessReviewCycleRows_();
    const cycle = cyclesData.rows.find(function(row){return row.cycleId===cycleId;});
    if (!cycle) throw new Error('Access review cycle not found. Refresh Access Review.');
    if (cycle.status!=='Open') throw new Error('This access review cycle is already completed and is read-only.');
    const itemsData = accessReviewItemRows_();
    const item = itemsData.rows.find(function(row){return row.cycleId===cycleId&&row.itemId===itemId;});
    if (!item) throw new Error('Access review item not found. Refresh Access Review.');
    let proposedRole = '', proposedDepartments = [];
    if (decision==='Modify') {
      proposedRole = String(input.proposedRole||'').trim();
      const role = rbacGetRoleDefinition_(proposedRole);
      if (!role || !role.active) throw new Error('Select an active proposed role.');
      proposedDepartments = accessReviewJsonArray_(input.proposedDepartments);
      if (role.scopeType==='DEPARTMENTS') proposedDepartments = iamValidateDepartments_(proposedDepartments);
      else proposedDepartments = [];
    }
    const now = new Date().toISOString();
    itemsData.sheet.getRange(item.row,10,1,6).setValues([[
      decision,proposedRole,JSON.stringify(proposedDepartments),comment,actor.email,now
    ]]);
    const summary = accessReviewUpdateCycleSummary_(cycleId);
    let linkedRequest = null;
    if ((decision==='Modify'||decision==='Remove') && typeof accessRequestCreateFromReview_ === 'function') {
      try { linkedRequest = accessRequestCreateFromReview_(actor,cycle,item,decision,proposedRole,proposedDepartments,comment); }
      catch (requestError) { console.log('Access Review decision saved, but Access Request creation failed: ' + (requestError && requestError.message ? requestError.message : requestError)); }
    }
    dashboardSecurityAudit_('ACCESS_REVIEW_DECISION','SUCCESS',actor.email,item.email,'',actor.role,null,{cycleId:cycleId,decision:decision,proposedRole:proposedRole,proposedDepartments:proposedDepartments,cycleCompleted:summary.complete,accessRequestId:linkedRequest&&linkedRequest.requestId?linkedRequest.requestId:''});
    let message = decision==='Retain' ? 'Access retained for '+item.email+'.' : 'Review decision recorded for '+item.email+'.';
    if (linkedRequest && linkedRequest.requestId) message += ' Access request '+linkedRequest.requestId+' created for controlled approval.';
    else if (decision!=='Retain') message += ' Open Access Requests to implement the required change.';
    return {success:true,message:message,linkedRequestId:linkedRequest&&linkedRequest.requestId?linkedRequest.requestId:'',center:getAccessReviewCenter_(cycleId)};
  } finally { lock.releaseLock(); }
}

/***************************************************************
 * PHASE 8.5 — ACCESS REQUEST & APPROVAL WORKFLOW
 * Controlled access requests with independent approval, automatic
 * implementation, review-cycle integration and temporary access.
 ***************************************************************/
const ACCESS_REQUEST_SHEET = 'Access_Requests';
const ACCESS_REQUEST_HEADERS = [
  'Request ID','Created At','Requested By','Requested By Name','Target Email','Target Name','Request Type',
  'Current Role','Current Departments JSON','Current Active','Requested Role','Requested Departments JSON',
  'Temporary Start At','Temporary End At','Business Justification','Approver Email','Status','Reviewer Comment',
  'Reviewed By','Reviewed At','Implemented At','Implementation Summary','Source Review Cycle ID','Source Review Item ID',
  'Privileged','Updated At'
];
const ACCESS_TEMP_GRANT_SHEET = 'Temporary_Access_Grants';
const ACCESS_TEMP_GRANT_HEADERS = [
  'Grant ID','Request ID','User Email','Original Role','Original Departments JSON','Original Active',
  'Granted Role','Granted Departments JSON','Start At','End At','Status','Created At','Activated At','Reverted At','Last Error'
];
const ACCESS_REQUEST_TYPES = ['NEW_ACCESS','ROLE_CHANGE','DEPARTMENT_CHANGE','TEMPORARY_ACCESS','ACCESS_REMOVAL','PRIVILEGED_ACCESS'];
const ACCESS_REQUEST_STATUSES = ['AWAITING_APPROVER','PENDING_APPROVAL','RETURNED','APPROVED','REJECTED','CANCELLED'];

function accessRequestEnsureSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const width = Math.max(sheet.getLastColumn(),headers.length);
  const current = sheet.getRange(1,1,1,width).getDisplayValues()[0];
  for (let i=0;i<headers.length;i++) if (current[i] && current[i] !== headers[i]) throw new Error(name+' headers are invalid. Contact the Compliance Administrator.');
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function accessRequestSheet_(){ return accessRequestEnsureSheet_(ACCESS_REQUEST_SHEET,ACCESS_REQUEST_HEADERS); }
function accessTempGrantSheet_(){ return accessRequestEnsureSheet_(ACCESS_TEMP_GRANT_SHEET,ACCESS_TEMP_GRANT_HEADERS); }
function accessRequestIso_(v){ if(!v)return ''; if(v instanceof Date)return v.toISOString(); const d=new Date(v); return isNaN(d.getTime())?String(v||''):d.toISOString(); }
function accessRequestJsonArray_(v){ return accessReviewJsonArray_(v); }

function accessRequestRows_(){
  const sheet=accessRequestSheet_();
  if(sheet.getLastRow()<2)return {sheet:sheet,rows:[]};
  const values=sheet.getRange(2,1,sheet.getLastRow()-1,ACCESS_REQUEST_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(r,i){return {
    requestId:String(r[0]||''),createdAt:accessRequestIso_(r[1]),requestedBy:String(r[2]||'').trim().toLowerCase(),requestedByName:String(r[3]||''),
    targetEmail:String(r[4]||'').trim().toLowerCase(),targetName:String(r[5]||''),requestType:String(r[6]||''),currentRole:String(r[7]||''),
    currentDepartments:accessRequestJsonArray_(r[8]),currentActive:r[9]===true||String(r[9]).toLowerCase()==='true',requestedRole:String(r[10]||''),
    requestedDepartments:accessRequestJsonArray_(r[11]),temporaryStartAt:accessRequestIso_(r[12]),temporaryEndAt:accessRequestIso_(r[13]),
    justification:String(r[14]||''),approverEmail:String(r[15]||'').trim().toLowerCase(),status:String(r[16]||''),reviewerComment:String(r[17]||''),
    reviewedBy:String(r[18]||'').trim().toLowerCase(),reviewedAt:accessRequestIso_(r[19]),implementedAt:accessRequestIso_(r[20]),
    implementationSummary:String(r[21]||''),sourceReviewCycleId:String(r[22]||''),sourceReviewItemId:String(r[23]||''),
    privileged:r[24]===true||String(r[24]).toLowerCase()==='true',updatedAt:accessRequestIso_(r[25]),row:i+2
  };}).filter(function(r){return !!r.requestId;})};
}

function accessTempGrantRows_(){
  const sheet=accessTempGrantSheet_();
  if(sheet.getLastRow()<2)return {sheet:sheet,rows:[]};
  const values=sheet.getRange(2,1,sheet.getLastRow()-1,ACCESS_TEMP_GRANT_HEADERS.length).getValues();
  return {sheet:sheet,rows:values.map(function(r,i){return {
    grantId:String(r[0]||''),requestId:String(r[1]||''),email:String(r[2]||'').trim().toLowerCase(),originalRole:String(r[3]||''),
    originalDepartments:accessRequestJsonArray_(r[4]),originalActive:r[5]===true||String(r[5]).toLowerCase()==='true',grantedRole:String(r[6]||''),
    grantedDepartments:accessRequestJsonArray_(r[7]),startAt:accessRequestIso_(r[8]),endAt:accessRequestIso_(r[9]),status:String(r[10]||''),
    createdAt:accessRequestIso_(r[11]),activatedAt:accessRequestIso_(r[12]),revertedAt:accessRequestIso_(r[13]),lastError:String(r[14]||''),row:i+2
  };}).filter(function(r){return !!r.grantId;})};
}

function accessRequestTypeLabel_(type){
  return ({NEW_ACCESS:'New Dashboard Access',ROLE_CHANGE:'Role Change',DEPARTMENT_CHANGE:'Department Access Change',TEMPORARY_ACCESS:'Temporary Access',ACCESS_REMOVAL:'Access Removal',PRIVILEGED_ACCESS:'Privileged Access'})[type]||type;
}
function accessRequestStatusLabel_(status){
  return ({AWAITING_APPROVER:'Awaiting Approver',PENDING_APPROVAL:'Pending Approval',RETURNED:'Returned',APPROVED:'Approved',REJECTED:'Rejected',CANCELLED:'Cancelled'})[status]||status;
}
function accessRequestRoleIsPrivileged_(role){
  if(!role)return false;
  const p=rbacNormalizePermissions_(role.permissions);
  return role.name==='Admin'||!!p.manageUsers||!!p.assignRoles||!!p.manageAccounts;
}

// Phase 8.5/8.7 approval policy:
// - Standard requests: one independent Compliance Manager or Admin.
// - Any TEMPORARY privileged elevation, including temporary Admin: one independent Compliance Manager or Admin.
// - Permanent privileged access and permanent changes involving Admin: two independent permanent Admins.
function accessRequestApprovalPolicy_(request){
  request=request||{};
  const type=String(request.requestType||'').trim().toUpperCase();
  const requestedRole=String(request.requestedRole||'').trim();
  const currentRole=String(request.currentRole||'').trim();
  const requestedPrivileged=!!requestedRole&&phase87RolePrivileged_(requestedRole);
  const currentPrivileged=!!currentRole&&phase87RolePrivileged_(currentRole);
  const privileged=type==='PRIVILEGED_ACCESS'||requestedPrivileged||currentPrivileged||!!request.privileged;
  if(type==='TEMPORARY_ACCESS'&&requestedPrivileged){
    return {key:'TEMP_PRIVILEGED',privileged:true,dualAdmin:false,requiredApprovals:1,reviewerLabel:'Compliance Manager or Admin'};
  }
  if(privileged){
    return {key:'PRIVILEGED',privileged:true,dualAdmin:true,requiredApprovals:2,reviewerLabel:'permanent Admin'};
  }
  return {key:'STANDARD',privileged:false,dualAdmin:false,requiredApprovals:1,reviewerLabel:'Compliance Manager or Admin'};
}
function accessRequestCanApprove_(access,requiresDualAdmin){
  if(!access)return false;
  if(requiresDualAdmin)return access.role==='Admin'&&!!access.isSystemAdmin;
  return access.role==='Admin'||access.role==='Compliance Manager';
}
function accessRequestEligibleApprovers_(requesterEmail,targetEmail,requiresDualAdmin){
  const users=rbacRows_().rows.filter(function(u){return u.active;});
  const out=[];
  users.forEach(function(u){
    if(u.email===requesterEmail||u.email===targetEmail)return;
    try {
      const access=rbacAccessForEmail_(u.email);
      if(accessRequestCanApprove_(access,requiresDualAdmin)&&(!requiresDualAdmin||phase87PermanentAdminEmail_(u.email))){
        out.push({email:u.email,name:u.name,role:u.role,canApprovePrivileged:phase87PermanentAdminEmail_(u.email)});
      }
    } catch(error){}
  });
  out.sort(function(a,b){return String(a.name||a.email).localeCompare(String(b.name||b.email));});
  return out;
}
function accessRequestPickApprover_(requested,requesterEmail,targetEmail,requiresDualAdmin){
  const eligible=accessRequestEligibleApprovers_(requesterEmail,targetEmail,requiresDualAdmin);
  const wanted=String(requested||'').trim().toLowerCase();
  if(wanted){const found=eligible.find(function(a){return a.email===wanted;}); if(!found)throw new Error('Selected approver is not eligible for this request.'); return found.email;}
  return eligible.length?eligible[0].email:'';
}
function accessRequestValidateDates_(start,end){
  const s=new Date(start),e=new Date(end);
  if(!start||isNaN(s.getTime())||!end||isNaN(e.getTime()))throw new Error('Choose valid temporary-access start and end dates.');
  if(e.getTime()<=s.getTime())throw new Error('Temporary access end must be after the start time.');
  if(e.getTime()-s.getTime()>90*86400000)throw new Error('Temporary access cannot exceed 90 days.');
  if(e.getTime()<Date.now()+60000)throw new Error('Temporary access end must be in the future.');
  return {startAt:s.toISOString(),endAt:e.toISOString()};
}

function accessRequestPublic_(r,actor){
  r=Object.assign({},r,{privileged:phase87IsPrivileged_(r)});
  const policy=accessRequestApprovalPolicy_(r);
  const approval=policy.dualAdmin?phase87ApprovalState_(r):{required:1,count:r.status==='APPROVED'?1:0,approvals:[]};
  const canManage=!!(actor&&actor.permissions&&actor.permissions.manageUsers);
  const canReview=!!actor && r.status==='PENDING_APPROVAL' && r.approverEmail===actor.email && accessRequestCanApprove_(actor,policy.dualAdmin) && actor.email!==r.requestedBy && actor.email!==r.targetEmail && (!policy.dualAdmin||(phase87PermanentAdminEmail_(actor.email)&&!approval.approvals.some(function(a){return a.email===actor.email;})));
  const canAssign=canManage && (r.status==='AWAITING_APPROVER'||r.status==='PENDING_APPROVAL');
  const canCancel=!!actor && (r.status==='AWAITING_APPROVER'||r.status==='PENDING_APPROVAL'||r.status==='RETURNED') && (r.requestedBy===actor.email||canManage);
  const canResubmit=!!actor && r.status==='RETURNED' && (r.requestedBy===actor.email||canManage);
  return {
    requestId:r.requestId,createdAt:r.createdAt,requestedBy:r.requestedBy,requestedByName:r.requestedByName,targetEmail:r.targetEmail,targetName:r.targetName,
    requestType:r.requestType,requestTypeLabel:accessRequestTypeLabel_(r.requestType),currentRole:r.currentRole,currentDepartments:r.currentDepartments,currentActive:r.currentActive,
    requestedRole:r.requestedRole,requestedDepartments:r.requestedDepartments,temporaryStartAt:r.temporaryStartAt,temporaryEndAt:r.temporaryEndAt,justification:r.justification,
    approverEmail:r.approverEmail,status:r.status,statusLabel:accessRequestStatusLabel_(r.status),reviewerComment:r.reviewerComment,reviewedBy:r.reviewedBy,reviewedAt:r.reviewedAt,
    implementedAt:r.implementedAt,implementationSummary:r.implementationSummary,sourceReviewCycleId:r.sourceReviewCycleId,sourceReviewItemId:r.sourceReviewItemId,
    requiredApprovals:policy.requiredApprovals,approvalCount:approval.count,approvals:approval.approvals,approvalPolicy:policy.key,approvalReviewerLabel:policy.reviewerLabel,
    privileged:r.privileged,updatedAt:r.updatedAt,canReview:canReview,canAssignApprover:canAssign,canCancel:canCancel,canResubmit:canResubmit
  };
}

function accessRequestActor_(){ return rbacRequire_('read'); }
function accessRequestUserMap_(){ const m={}; rbacRows_().rows.forEach(function(u){m[u.email]=u;}); return m; }
function accessRequestRoleList_(){ return rbacReadRoleDefinitions_().filter(function(r){return r.active;}); }

function accessRequestBuild_(actor,input,existingRequest){
  input=input||{};
  const type=String(input.requestType||'').trim().toUpperCase();
  if(ACCESS_REQUEST_TYPES.indexOf(type)===-1)throw new Error('Choose a valid access request type.');
  const canManage=!!actor.permissions.manageUsers;
  let targetEmail=String(input.targetEmail||actor.email).trim().toLowerCase();
  let targetName=String(input.targetName||'').trim();
  if(!rbacAllowedEmail_(targetEmail))throw new Error('Target user must use an approved company email domain.');
  if(!canManage && targetEmail!==actor.email)throw new Error('You can submit access requests only for your own dashboard account.');
  if(!canManage && type==='NEW_ACCESS')throw new Error('New dashboard access must be requested by an authorized administrator or manager.');
  const userMap=accessRequestUserMap_();
  const current=userMap[targetEmail]||null;
  if(type==='NEW_ACCESS' && current)throw new Error('This user already has dashboard access. Choose a change request type instead.');
  if(type!=='NEW_ACCESS' && !current)throw new Error('Target dashboard user was not found. Use New Dashboard Access for a new user.');
  if(current && !targetName)targetName=current.name||targetEmail.split('@')[0];
  if(!targetName && type==='NEW_ACCESS')throw new Error('Enter the target user display name.');
  if(targetName.length>120||/^[=+@-]/.test(targetName))throw new Error('Target display name is invalid.');
  const justification=String(input.justification||'').trim();
  if(justification.length<10)throw new Error('Business justification must be at least 10 characters.');
  if(justification.length>1500)throw new Error('Business justification must be 1500 characters or fewer.');
  const roles=accessRequestRoleList_();
  let requestedRole='',requestedDepartments=[],tempStart='',tempEnd='',privileged=false;
  if(type==='ACCESS_REMOVAL'){
    requestedRole=current.role; requestedDepartments=current.departments||[];
  } else if(type==='DEPARTMENT_CHANGE'){
    requestedRole=current.role;
    const role=rbacGetRoleDefinition_(requestedRole); if(!role||!role.active)throw new Error('Current role is missing or inactive.');
    if(role.scopeType!=='DEPARTMENTS')throw new Error('Department Access Change applies only to department-scoped roles. Use Role Change if the scope model must change.');
    requestedDepartments=rbacValidateRoleScope_(role,targetEmail,input.requestedDepartments||[]);
  } else {
    requestedRole=String(input.requestedRole||'').trim();
    const role=rbacGetRoleDefinition_(requestedRole); if(!role||!role.active)throw new Error('Select an active requested role.');
    requestedDepartments=rbacValidateRoleScope_(role,targetEmail,input.requestedDepartments||[]);
    privileged=accessRequestRoleIsPrivileged_(role) || type==='PRIVILEGED_ACCESS';
    if(type==='PRIVILEGED_ACCESS' && !accessRequestRoleIsPrivileged_(role))throw new Error('Privileged Access must request a role with administrative IAM permissions.');
    if(type==='ROLE_CHANGE' && current && requestedRole===current.role && accessReviewSameArray_(requestedDepartments,current.departments||[]))throw new Error('Requested access is the same as the current access.');
    if(type==='NEW_ACCESS' && !requestedRole)throw new Error('Select a role for the new dashboard user.');
    if(type==='TEMPORARY_ACCESS'){
      const dates=accessRequestValidateDates_(input.temporaryStartAt,input.temporaryEndAt); tempStart=dates.startAt;tempEnd=dates.endAt;
    }
  }
  if(type==='DEPARTMENT_CHANGE' && accessReviewSameArray_(requestedDepartments,current.departments||[]))throw new Error('Requested departments are the same as the current department scope.');
  privileged=privileged||!!(current&&phase87RolePrivileged_(current.role))||phase87RolePrivileged_(requestedRole);
  phase87AssertNoGrant_(targetEmail);
  if(type==='TEMPORARY_ACCESS'&&privileged&&Date.parse(tempEnd)-Date.parse(tempStart)>24*3600000)throw new Error('Temporary privileged access is limited to 24 hours.');
  const approvalPolicy=accessRequestApprovalPolicy_({requestType:type,currentRole:current?current.role:'',requestedRole:requestedRole,privileged:privileged});
  const approver=accessRequestPickApprover_(input.approverEmail,existingRequest?existingRequest.requestedBy:actor.email,targetEmail,approvalPolicy.dualAdmin);
  return {
    requestType:type,targetEmail:targetEmail,targetName:targetName,currentRole:current?current.role:'',currentDepartments:current?current.departments||[]:[],currentActive:current?current.active:false,
    requestedRole:requestedRole,requestedDepartments:requestedDepartments,temporaryStartAt:tempStart,temporaryEndAt:tempEnd,justification:justification,
    approverEmail:approver,status:approver?'PENDING_APPROVAL':'AWAITING_APPROVER',privileged:privileged
  };
}

function accessRequestAppend_(actor,built,sourceCycleId,sourceItemId){
  const now=new Date().toISOString();
  const id='ARQ-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,5).toUpperCase();
  accessRequestSheet_().appendRow([
    id,now,actor.email,actor.name||'',built.targetEmail,built.targetName,built.requestType,built.currentRole,JSON.stringify(built.currentDepartments||[]),built.currentActive,
    built.requestedRole,JSON.stringify(built.requestedDepartments||[]),built.temporaryStartAt,built.temporaryEndAt,built.justification,built.approverEmail,built.status,'','','','','',
    String(sourceCycleId||''),String(sourceItemId||''),built.privileged,now
  ]);
  dashboardSecurityAudit_('ACCESS_REQUEST_SUBMITTED','SUCCESS',actor.email,built.targetEmail,'',actor.role,null,{requestId:id,type:built.requestType,approver:built.approverEmail,status:built.status,privileged:built.privileged});
  return {requestId:id,status:built.status,approverEmail:built.approverEmail};
}

function submitAccessRequest_(input){
  const actor=accessRequestActor_();
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try { const built=accessRequestBuild_(actor,input||{},null); const created=accessRequestAppend_(actor,built,input&&input.sourceReviewCycleId,input&&input.sourceReviewItemId); return {success:true,message:created.status==='PENDING_APPROVAL'?'Access request '+created.requestId+' submitted for approval.':'Access request '+created.requestId+' created. Assign an independent approver to continue.',requestId:created.requestId,status:created.status}; }
  finally {lock.releaseLock();}
}

function accessRequestCreateFromReview_(actor,cycle,item,decision,proposedRole,proposedDepartments,comment){
  const existing=accessRequestRows_().rows.find(function(r){return r.sourceReviewCycleId===cycle.cycleId&&r.sourceReviewItemId===item.itemId&&['AWAITING_APPROVER','PENDING_APPROVAL','RETURNED','APPROVED'].indexOf(r.status)!==-1;});
  if(existing)return {requestId:existing.requestId,status:existing.status};
  const input={
    requestType:decision==='Remove'?'ACCESS_REMOVAL':(proposedRole!==item.role?'ROLE_CHANGE':'DEPARTMENT_CHANGE'),
    targetEmail:item.email,targetName:item.name,requestedRole:proposedRole,requestedDepartments:proposedDepartments,
    justification:'Access Review '+cycle.cycleName+': '+comment,sourceReviewCycleId:cycle.cycleId,sourceReviewItemId:item.itemId
  };
  const built=accessRequestBuild_(actor,input,null);
  return accessRequestAppend_(actor,built,cycle.cycleId,item.itemId);
}

function getAccessRequestCenter_(){
  const actor=accessRequestActor_();
  if(actor.role!=='Admin'&&actor.role!=='Compliance Manager')throw new Error('Only Compliance Manager or Admin can review the access request center.');
  accessRequestApplyDueTemporaryGrants_();
  const rows=accessRequestRows_().rows.sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));});
  const publicRows=rows.map(function(r){return accessRequestPublic_(r,actor);});
  const pending=rows.filter(function(r){return r.status==='PENDING_APPROVAL'||r.status==='AWAITING_APPROVER';});
  const approvers=accessRequestEligibleApprovers_('','',false);
  return {success:true,actor:{email:actor.email,name:actor.name,role:actor.role},summary:{total:rows.length,pending:pending.length,approved:rows.filter(function(r){return r.status==='APPROVED';}).length,returned:rows.filter(function(r){return r.status==='RETURNED';}).length,rejected:rows.filter(function(r){return r.status==='REJECTED'||r.status==='CANCELLED';}).length,privilegedPending:pending.filter(function(r){return phase87IsPrivileged_(r);}).length},requests:publicRows,users:rbacRows_().rows.map(function(u){return {email:u.email,name:u.name,role:u.role,departments:u.departments,active:u.active};}),roles:accessRequestRoleList_().map(function(r){return {name:r.name,scopeType:r.scopeType,privileged:accessRequestRoleIsPrivileged_(r)};}),approvers:approvers,departments:iamDepartmentCatalog_()};
}

function getMyAccessRequests_(){
  const actor=accessRequestActor_();
  accessRequestApplyDueTemporaryGrants_();
  const rows=accessRequestRows_().rows.filter(function(r){return r.requestedBy===actor.email||r.targetEmail===actor.email||r.approverEmail===actor.email;}).sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));}).slice(0,30);
  const eligible=accessRequestEligibleApprovers_(actor.email,actor.email,false);
  return {success:true,approvalAvailability:{standard:eligible.length,privileged:eligible.filter(function(a){return a.canApprovePrivileged;}).length},actor:{email:actor.email,name:actor.name,role:actor.role,departments:actor.departments},requests:rows.map(function(r){return accessRequestPublic_(r,actor);}),roles:accessRequestRoleList_().map(function(r){return {name:r.name,scopeType:r.scopeType,privileged:accessRequestRoleIsPrivileged_(r)};}),departments:iamDepartmentCatalog_()};
}

function assignAccessRequestApprover_(requestId,approverEmail){
  const actor=rbacRequirePermission_('manageUsers');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const data=accessRequestRows_();const row=data.rows.find(function(r){return r.requestId===String(requestId||'');});
    if(!row)throw new Error('Access request not found. Refresh Access Requests.');
    if(['AWAITING_APPROVER','PENDING_APPROVAL'].indexOf(row.status)===-1)throw new Error('Approver can be changed only while the request is awaiting approval.');
    row.privileged=phase87IsPrivileged_(row);
    const policy=accessRequestApprovalPolicy_(row);
    const selected=accessRequestPickApprover_(approverEmail,row.requestedBy,row.targetEmail,policy.dualAdmin);
    if(policy.dualAdmin&&!phase87Candidates_(row).some(function(a){return a.email===selected;}))throw new Error('Choose a different independent permanent Admin who has not already approved.');
    if(!selected)throw new Error('No independent eligible approver is currently available.');
    data.sheet.getRange(row.row,16,1,11).setValues([[selected,'PENDING_APPROVAL',row.reviewerComment,row.reviewedBy,row.reviewedAt,row.implementedAt,row.implementationSummary,row.sourceReviewCycleId,row.sourceReviewItemId,row.privileged,new Date().toISOString()]]);
    dashboardSecurityAudit_('ACCESS_REQUEST_APPROVER_ASSIGNED','SUCCESS',actor.email,row.targetEmail,'',actor.role,null,{requestId:row.requestId,approver:selected});
    return {success:true,message:'Approver assigned to '+row.requestId+'.'};
  }finally{lock.releaseLock();}
}

function accessRequestApplyUserState_(actorEmail,targetEmail,targetName,roleName,departments,active,actionLabel){
  const role=rbacGetRoleDefinition_(roleName); if(!role||!role.active)throw new Error('Requested role is missing or inactive.');
  const depts=rbacValidateRoleScope_(role,targetEmail,departments||[]);
  const data=rbacRows_(); const matches=data.rows.filter(function(r){return r.email===targetEmail;}); if(matches.length>1)throw new Error('Duplicate User_Access entries exist for the target user.');
  const before=matches[0]||null;
  phase87ProtectAnchor_(targetEmail,roleName,active);
  if(before&&before.role==='Admin'&&before.active&&(roleName!=='Admin'||!active)&&!data.rows.some(function(r){return r.email!==targetEmail&&r.role==='Admin'&&r.active;}))throw new Error('The last active Admin cannot be removed, disabled, or demoted.');
  if(!getSpreadsheet_().getSheetByName('Access_Audit_Log'))throw new Error('Access_Audit_Log is missing. Access change was not applied.');
  const after={email:targetEmail,name:targetName||before&&before.name||'',role:roleName,active:!!active,departments:depts};
  const values=[targetEmail,after.name,roleName,!!active,actorEmail,new Date().toISOString(),JSON.stringify(depts)];
  const rowNo=before?before.row:data.sheet.getLastRow()+1; const old=before?data.sheet.getRange(rowNo,1,1,RBAC_HEADERS.length).getValues():null;
  data.sheet.getRange(rowNo,1,1,RBAC_HEADERS.length).setValues([values]);
  try{rbacAudit_(actorEmail,targetEmail,before,after);dashboardRevokeAllSessionsForEmail_(targetEmail,actionLabel||'ACCESS_CHANGED');}catch(error){if(old)data.sheet.getRange(rowNo,1,1,RBAC_HEADERS.length).setValues(old);else data.sheet.getRange(rowNo,1,1,RBAC_HEADERS.length).clearContent();throw new Error('Access change rolled back because mandatory auditing or session revocation failed.');}
  return {before:before,after:after};
}

function accessRequestApplyRemoval_(actorEmail,row){
  const data=rbacRows_(); const target=data.rows.find(function(r){return r.email===row.targetEmail;}); if(!target)throw new Error('Target user no longer exists.');
  if(target.role==='Admin'&&target.active&&!data.rows.some(function(r){return r.email!==target.email&&r.role==='Admin'&&r.active;}))throw new Error('The last active Admin cannot be removed.');
  phase87ProtectAnchor_(target.email,target.role,false);
  const before=Object.assign({},target); const old=data.sheet.getRange(target.row,1,1,RBAC_HEADERS.length).getValues();
  const after={email:target.email,name:target.name,role:target.role,active:false,departments:target.departments};
  data.sheet.getRange(target.row,1,1,RBAC_HEADERS.length).setValues([[target.email,target.name,target.role,false,actorEmail,new Date().toISOString(),JSON.stringify(target.departments||[])]]);
  try{rbacAudit_(actorEmail,target.email,before,after);dashboardRevokeAllSessionsForEmail_(target.email,'ACCESS_REMOVED');}catch(error){data.sheet.getRange(target.row,1,1,RBAC_HEADERS.length).setValues(old);throw new Error('Access removal rolled back because mandatory auditing or session revocation failed.');}
  return {before:before,after:after};
}



function accessRequestImplement_(actor,row){
  phase87AssertRequestCurrent_(row);
  if(row.requestType==='TEMPORARY_ACCESS'&&!phase87TriggerReady_())throw new Error('Initialize Phase 8.7 to enable background expiry before approving temporary access.');
  if(row.requestType==='ACCESS_REMOVAL'){accessRequestApplyRemoval_(actor.email,row);return 'Dashboard account deactivated and active sessions revoked.';}
  if(row.requestType==='TEMPORARY_ACCESS'){const grant=accessRequestCreateTempGrant_(row,actor.email);return grant.status==='ACTIVE'?'Temporary access activated until '+row.temporaryEndAt+'.':'Temporary access scheduled from '+row.temporaryStartAt+' until '+row.temporaryEndAt+'.';}
  const role=row.requestedRole; const active=row.requestType==='NEW_ACCESS'?true:row.currentActive;
  accessRequestApplyUserState_(actor.email,row.targetEmail,row.targetName,role,row.requestedDepartments,active,'ACCESS_REQUEST_IMPLEMENTED');
  return row.requestType==='NEW_ACCESS'?'Dashboard access created.':'Requested access applied and prior sessions revoked.';
}

function reviewAccessRequest_(requestId,action,comment){
  const actor=accessRequestActor_(); const id=String(requestId||'').trim(); const act=String(action||'').trim().toUpperCase(); const note=String(comment||'').trim();
  if(['APPROVE','RETURN','REJECT'].indexOf(act)===-1)throw new Error('Choose Approve, Return, or Reject.');
  if((act==='RETURN'||act==='REJECT')&&note.length<5)throw new Error('Add a reviewer comment of at least 5 characters.');
  if(note.length>1500)throw new Error('Reviewer comment must be 1500 characters or fewer.');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const data=accessRequestRows_();const row=data.rows.find(function(r){return r.requestId===id;});if(!row)throw new Error('Access request not found. Refresh Access Requests.');
    if(row.status!=='PENDING_APPROVAL')throw new Error('Only pending requests can be reviewed.');
    if(row.approverEmail!==actor.email)throw new Error('This request is assigned to another approver.');
    if(actor.email===row.requestedBy||actor.email===row.targetEmail)throw new Error('You cannot approve or review your own access request.');
    row.privileged=phase87IsPrivileged_(row);
    const policy=accessRequestApprovalPolicy_(row);
    if(!accessRequestCanApprove_(actor,policy.dualAdmin))throw new Error(policy.dualAdmin?'This privileged/Admin request requires an independent permanent Admin approver.':'Only a Compliance Manager or Admin can review this access request.');
    if(policy.dualAdmin){const progress=phase87ReviewPrivileged_(actor,row,act,note,data);if(progress)return progress;}
    const now=new Date().toISOString();let status='',summary='';
    if(act==='RETURN')status='RETURNED';
    else if(act==='REJECT')status='REJECTED';
    else { try{summary=accessRequestImplement_(actor,row);status='APPROVED';}catch(error){if(policy.dualAdmin)phase87ResetApprovals_(row,actor.email,'Implementation failed; fresh approvals required');throw error;} }
    data.sheet.getRange(row.row,17,1,10).setValues([[status,note,actor.email,now,act==='APPROVE'?now:'',summary,row.sourceReviewCycleId,row.sourceReviewItemId,row.privileged,now]]);
    dashboardSecurityAudit_('ACCESS_REQUEST_'+(act==='APPROVE'?'APPROVED':act),'SUCCESS',actor.email,row.targetEmail,'',actor.role,null,{requestId:row.requestId,type:row.requestType,privileged:row.privileged,implementation:summary});
    return {success:true,message:act==='APPROVE'?'Access request approved and implemented.':act==='RETURN'?'Access request returned to the requester.':'Access request rejected.'};
  }finally{lock.releaseLock();}
}

function cancelAccessRequest_(requestId,reason){
  const actor=accessRequestActor_();const note=String(reason||'').trim();if(note.length>500)throw new Error('Cancellation reason is too long.');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{const data=accessRequestRows_();const row=data.rows.find(function(r){return r.requestId===String(requestId||'');});if(!row)throw new Error('Access request not found.');
    const canManage=!!actor.permissions.manageUsers;if(row.requestedBy!==actor.email&&!canManage)throw new Error('Only the requester or an access administrator can cancel this request.');
    if(['AWAITING_APPROVER','PENDING_APPROVAL','RETURNED'].indexOf(row.status)===-1)throw new Error('This request can no longer be cancelled.');
    const now=new Date().toISOString();data.sheet.getRange(row.row,17,1,10).setValues([['CANCELLED',note,actor.email,now,'','',row.sourceReviewCycleId,row.sourceReviewItemId,row.privileged,now]]);
    dashboardSecurityAudit_('ACCESS_REQUEST_CANCELLED','SUCCESS',actor.email,row.targetEmail,'',actor.role,null,{requestId:row.requestId,reason:note});return {success:true,message:'Access request cancelled.'};
  }finally{lock.releaseLock();}
}

function resubmitAccessRequest_(input){
  const actor=accessRequestActor_();input=input||{};const id=String(input.requestId||'').trim();
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{const data=accessRequestRows_();const row=data.rows.find(function(r){return r.requestId===id;});if(!row)throw new Error('Access request not found.');
    const canManage=!!actor.permissions.manageUsers;if(row.status!=='RETURNED')throw new Error('Only returned requests can be edited and resubmitted.');if(row.requestedBy!==actor.email&&!canManage)throw new Error('Only the requester or an access administrator can resubmit this request.');
    if(accessRequestApprovalPolicy_(row).dualAdmin)phase87ResetApprovals_(row,actor.email,'Resubmitted');
    const built=accessRequestBuild_(actor,{requestType:row.requestType,targetEmail:row.targetEmail,targetName:row.targetName,requestedRole:input.requestedRole||row.requestedRole,requestedDepartments:input.requestedDepartments===undefined?row.requestedDepartments:input.requestedDepartments,temporaryStartAt:input.temporaryStartAt||row.temporaryStartAt,temporaryEndAt:input.temporaryEndAt||row.temporaryEndAt,justification:input.justification||row.justification,approverEmail:input.approverEmail===undefined?row.approverEmail:input.approverEmail},row);
    const now=new Date().toISOString();data.sheet.getRange(row.row,8,1,19).setValues([[built.currentRole,JSON.stringify(built.currentDepartments),built.currentActive,built.requestedRole,JSON.stringify(built.requestedDepartments),built.temporaryStartAt,built.temporaryEndAt,built.justification,built.approverEmail,built.status,'','','','','',row.sourceReviewCycleId,row.sourceReviewItemId,built.privileged,now]]);
    dashboardSecurityAudit_('ACCESS_REQUEST_RESUBMITTED','SUCCESS',actor.email,row.targetEmail,'',actor.role,null,{requestId:row.requestId,approver:built.approverEmail});return {success:true,status:built.status,message:built.status==='PENDING_APPROVAL'?'Access request resubmitted for approval.':'Access request resubmitted. Awaiting an independent eligible approver; access has not changed.'};
  }finally{lock.releaseLock();}
}



/***************************************************************
 * PHASE 8.7 — PRIVILEGED ACCESS CONTROLS
 * Dual independent permanent-Admin approval; bounded emergency
 * grants; fail-closed expiry; protected permanent Admin anchor.
 ***************************************************************/
const P87_APPROVAL_HEADERS=['Request ID','Fingerprint','Action','Actor Email','At','Comment'];
const P87_EMERGENCY_HEADERS=['Emergency ID','Target Email','Authorized By','Incident Reference','Justification','Start At','End At','Grant ID','Review Status','Reviewed By','Reviewed At','Review Comment'];
const P87_AUDIT_HEADERS=['At','Action','Actor Email','Target Email','Reference','Details'];
let phase87LifecycleDone_=false,phase87LifecycleBusy_=false;
function phase87Read_(name,headers){const sheet=getSpreadsheet_().getSheetByName(name);if(!sheet)return [];if(sheet.getLastRow()<2)return [];return sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();}
function phase87Audit_(action,actor,target,reference,details){accessRequestEnsureSheet_('Privileged_Audit_Log',P87_AUDIT_HEADERS).appendRow([new Date().toISOString(),action,actor,target,reference,JSON.stringify(details||{})]);}
function phase87OpenGrants_(){return accessTempGrantRows_().rows.filter(function(g){return ['REVERTED','CANCELLED'].indexOf(g.status)===-1;});}
function phase87RolePrivileged_(name){return accessRequestRoleIsPrivileged_(rbacGetRoleDefinition_(name));}
function phase87PermanentAdminEmail_(email){
 const matches=rbacRows_().rows.filter(function(u){return u.email===email;});
 const user=matches.length===1&&matches[0].active&&matches[0].role==='Admin'?matches[0]:null;
 const role=rbacGetRoleDefinition_('Admin');
 return !!user&&!!role&&role.active&&role.builtIn&&!phase87OpenGrants_().some(function(g){return g.email===email;});
}
function phase87RequireAdmin_(){const actor=rbacRequire_('admin');if(!phase87PermanentAdminEmail_(actor.email))throw new Error('This action requires a permanent Admin. Temporary and emergency Admins cannot authorize privileged access.');return actor;}
function phase87IsPrivileged_(r){return !!r.privileged||phase87RolePrivileged_(r.currentRole)||phase87RolePrivileged_(r.requestedRole);}
function phase87Fingerprint_(r){return JSON.stringify([r.requestId,r.requestType,r.requestedBy,r.targetEmail,r.currentRole,r.currentDepartments,r.currentActive,r.requestedRole,r.requestedDepartments,r.temporaryStartAt,r.temporaryEndAt,r.justification,rbacGetRoleDefinition_(r.requestedRole)]);}
function phase87ApprovalState_(r){
 const fingerprint=phase87Fingerprint_(r);let approvals=[];
 phase87Read_('Privileged_Approvals',P87_APPROVAL_HEADERS).forEach(function(v){if(String(v[0])!==r.requestId)return;if(v[2]==='RESET'){approvals=[];return;}if(v[2]==='APPROVE'&&v[1]===fingerprint&&v[3]!==r.requestedBy&&v[3]!==r.targetEmail&&(r.status==='APPROVED'||phase87PermanentAdminEmail_(String(v[3])))&&!approvals.some(function(a){return a.email===v[3];}))approvals.push({email:String(v[3]),at:accessRequestIso_(v[4])});});
 return {required:2,count:approvals.length,approvals:approvals};
}
function phase87ResetApprovals_(r,actor,reason){accessRequestEnsureSheet_('Privileged_Approvals',P87_APPROVAL_HEADERS).appendRow([r.requestId,'','RESET',actor,new Date().toISOString(),reason]);}
function phase87Candidates_(r){const approved=phase87ApprovalState_(r).approvals.map(function(a){return a.email;});return accessRequestEligibleApprovers_(r.requestedBy,r.targetEmail,true).filter(function(a){return approved.indexOf(a.email)===-1;});}
function phase87AssertNoGrant_(email){if(phase87OpenGrants_().some(function(g){return g.email===email;}))throw new Error('This user has an active, scheduled, or unresolved temporary grant. End or restore that grant in Privileged Access before changing access.');}
function phase87GuardDirectUser_(before,roleName,active,departments){
 phase87AssertNoGrant_(before?before.email:'');
 if((before&&phase87RolePrivileged_(before.role))||phase87RolePrivileged_(roleName)){
  const same=before&&before.role===roleName&&before.active===active&&accessReviewSameArray_(before.departments||[],departments||[]);
  if(!same)throw new Error('Privileged access changes require Access Requests and two independent permanent Admin approvals. Use Access Requests instead of direct editing.');
 }
}
function phase87ProtectAnchor_(email,role,active){
 const current=rbacRows_().rows.find(function(u){return u.email===email;});
 if(current&&current.active&&current.role==='Admin'&&(role!=='Admin'||!active)){
  const others=rbacRows_().rows.filter(function(u){return u.email!==email&&phase87PermanentAdminEmail_(u.email);});
  if(!others.length)throw new Error('Keep at least one permanent active Admin. A temporary or emergency Admin does not satisfy this requirement.');
 }
}
function phase87GuardRole_(actor,existing,role){
 if(role.name==='Admin'&&(!role.active||role.scopeType!=='ANY'))throw new Error('The built-in Admin role must remain active and unrestricted.');
 const privileged=accessRequestRoleIsPrivileged_(role)||(existing&&accessRequestRoleIsPrivileged_(existing));
 if(privileged&&!phase87PermanentAdminEmail_(actor.email))throw new Error('Only a permanent Admin may maintain privileged role definitions.');
 if(existing&&JSON.stringify([existing.permissions,existing.scopeType,existing.allowedUsers,existing.allowedDepartments,existing.active])!==JSON.stringify([role.permissions,role.scopeType,role.allowedUsers,role.allowedDepartments,role.active])){
  if(rbacRows_().rows.some(function(u){return u.role===existing.name;})&&(privileged||accessRequestRoleIsPrivileged_(existing)))throw new Error('An assigned privileged role cannot be changed in place. Create a separate role and migrate users through dual-approved Access Requests.');
  if(accessTempGrantRows_().rows.some(function(g){return ['REVERTED','CANCELLED'].indexOf(g.status)===-1&&(g.originalRole===existing.name||g.grantedRole===existing.name);}))throw new Error('This role is referenced by an open temporary grant. Resolve the grant before editing this role.');
 }
}
function phase87AssertRequestCurrent_(r){
 const user=rbacRows_().rows.find(function(u){return u.email===r.targetEmail;});
 if(r.requestType==='NEW_ACCESS'){if(user)throw new Error('This user was created after the request. Return the request and refresh its details.');}
 else if(!user||user.role!==r.currentRole||user.active!==r.currentActive||!accessReviewSameArray_(user.departments||[],r.currentDepartments||[]))throw new Error('Current access changed since this request was submitted. Return it for an updated request.');
 phase87AssertNoGrant_(r.targetEmail);
}
function phase87ReviewPrivileged_(actor,row,act,note,data){
 if(!phase87PermanentAdminEmail_(actor.email))throw new Error('Privileged requests require a permanent Admin reviewer.');
 if(act==='RETURN'||act==='REJECT'){phase87ResetApprovals_(row,actor.email,act);return null;}
 phase87AssertRequestCurrent_(row);
 const state=phase87ApprovalState_(row);
 if(state.approvals.some(function(a){return a.email===actor.email;}))throw new Error('You already approved this request. A different permanent Admin must provide the second approval.');
 phase87Audit_('PRIVILEGED_APPROVAL',actor.email,row.targetEmail,row.requestId,{approvalNumber:state.count+1,comment:note});
 accessRequestEnsureSheet_('Privileged_Approvals',P87_APPROVAL_HEADERS).appendRow([row.requestId,phase87Fingerprint_(row),'APPROVE',actor.email,new Date().toISOString(),note]);
 if(state.count===0){
  const candidates=phase87Candidates_(row),next=candidates.length?candidates[0].email:'';
  data.sheet.getRange(row.row,16,1,2).setValues([[next,next?'PENDING_APPROVAL':'AWAITING_APPROVER']]);
  data.sheet.getRange(row.row,26).setValue(new Date().toISOString());
  return {success:true,message:'First approval recorded. Access is unchanged. '+(next?'A second independent Admin must approve.':'Awaiting another independent permanent Admin.')};
 }
 return null;
}
function phase87AssertExpirySafe_(email){
 if(phase87OpenGrants_().some(function(g){return g.email===email&&(g.status==='ERROR'||g.status==='APPLYING'||!Number.isFinite(Date.parse(g.endAt))||Date.parse(g.endAt)<=Date.now());}))throw new Error('Access is paused because a temporary grant expired or needs recovery. A permanent Admin must restore it in Privileged Access.');
}
function accessRequestApplyDueTemporaryGrants_(){
 if(phase87LifecycleBusy_||phase87LifecycleDone_)return;
 const lock=LockService.getScriptLock(),owned=!lock.hasLock();if(owned&&!lock.tryLock(5000))return;
 phase87LifecycleBusy_=true;
 try{
  const data=accessTempGrantRows_(),now=Date.now();
  data.rows.forEach(function(g){
   if(['ACTIVE','SCHEDULED'].indexOf(g.status)===-1)return;
   try{
    const start=Date.parse(g.startAt),end=Date.parse(g.endAt);if(!Number.isFinite(start)||!Number.isFinite(end))throw new Error('Invalid grant dates. Restore this grant.');
    const current=rbacRows_().rows.find(function(u){return u.email===g.email;});if(!current)throw new Error('Target user no longer exists.');
    if(end<=now){
     if(g.status==='ACTIVE'){
      if(current.role!==g.grantedRole||!accessReviewSameArray_(current.departments,g.grantedDepartments))throw new Error('Access drift detected. Review and restore the original access explicitly.');
      phase87Audit_('GRANT_EXPIRY_STARTED','SYSTEM',g.email,g.grantId,{});
      accessRequestApplyUserState_('SYSTEM_TEMP_ACCESS',g.email,current.name,g.originalRole,g.originalDepartments,g.originalActive,'TEMP_ACCESS_REVERTED');
     }
     data.sheet.getRange(g.row,11,1,5).setValues([['REVERTED',g.createdAt,g.activatedAt,new Date().toISOString(),'']]);
    }else if(g.status==='SCHEDULED'&&start<=now){
     if(current.role!==g.originalRole||current.active!==g.originalActive||!accessReviewSameArray_(current.departments,g.originalDepartments))throw new Error('Access changed after approval. Scheduled grant was not activated.');
     phase87Audit_('GRANT_ACTIVATION_STARTED','SYSTEM',g.email,g.grantId,{});
     accessRequestApplyUserState_('SYSTEM_TEMP_ACCESS',g.email,current.name,g.grantedRole,g.grantedDepartments,true,'TEMP_ACCESS_ACTIVATED');
     data.sheet.getRange(g.row,11,1,5).setValues([['ACTIVE',g.createdAt,new Date().toISOString(),'','']]);
    }
   }catch(error){data.sheet.getRange(g.row,11).setValue('ERROR');data.sheet.getRange(g.row,15).setValue(String(error.message||error).slice(0,500));}
  });phase87LifecycleDone_=true;
 }finally{phase87LifecycleBusy_=false;if(owned)lock.releaseLock();}
}
function accessRequestCreateTempGrant_(row,actorEmail){
 phase87AssertNoGrant_(row.targetEmail);
 const current=rbacRows_().rows.find(function(u){return u.email===row.targetEmail;});if(!current)throw new Error('Target user no longer exists.');
 if(phase87PermanentAdminEmail_(current.email))throw new Error('Permanent Admin accounts cannot receive temporary replacement roles. Use a separate active non-Admin account.');
 const now=new Date(),start=new Date(row.temporaryStartAt),end=new Date(row.temporaryEndAt);if(!Number.isFinite(start.getTime())||!Number.isFinite(end.getTime())||end<=now||end<=start)throw new Error('Temporary access dates are no longer valid.');
 if(phase87RolePrivileged_(row.requestedRole)&&end-start>24*3600000)throw new Error('Temporary privileged access is limited to 24 hours.');
 const grantId='TAG-'+Utilities.getUuid().slice(0,12).toUpperCase(),sheet=accessTempGrantSheet_(),rowNo=sheet.getLastRow()+1,status=start<=now?'APPLYING':'SCHEDULED';
 phase87Audit_('GRANT_CREATED',actorEmail,row.targetEmail,grantId,{requestId:row.requestId,endAt:row.temporaryEndAt});
 sheet.appendRow([grantId,row.requestId,row.targetEmail,current.role,JSON.stringify(current.departments||[]),current.active,row.requestedRole,JSON.stringify(row.requestedDepartments||[]),row.temporaryStartAt,row.temporaryEndAt,status,now.toISOString(),'','','']);
 if(status==='APPLYING'){
  try{accessRequestApplyUserState_(actorEmail,row.targetEmail,row.targetName||current.name,row.requestedRole,row.requestedDepartments,true,'TEMP_ACCESS_ACTIVATED');sheet.getRange(rowNo,11).setValue('ACTIVE');sheet.getRange(rowNo,13).setValue(new Date().toISOString());}
  catch(error){sheet.getRange(rowNo,11).setValue('ERROR');sheet.getRange(rowNo,15).setValue(String(error.message||error).slice(0,500));throw new Error('Grant needs recovery in Privileged Access: '+(error.message||error));}
 }
 return {grantId:grantId,status:status==='APPLYING'?'ACTIVE':status};
}
function phase87ExpiryWorker_(){phase87LifecycleDone_=false;accessRequestApplyDueTemporaryGrants_();}
function initializePhase87_(){
 const actor=phase87RequireAdmin_(),lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  accessRequestEnsureSheet_('Privileged_Approvals',P87_APPROVAL_HEADERS);accessRequestEnsureSheet_('Emergency_Access_Log',P87_EMERGENCY_HEADERS);accessRequestEnsureSheet_('Privileged_Audit_Log',P87_AUDIT_HEADERS);accessTempGrantSheet_();
  if(!ScriptApp.getProjectTriggers().some(function(t){return t.getHandlerFunction()==='phase87ExpiryWorker_';}))ScriptApp.newTrigger('phase87ExpiryWorker_').timeBased().everyMinutes(5).create();
  phase87Audit_('CONTROLS_INITIALIZED',actor.email,'','',{expirySweepMinutes:5});return {success:true,message:'Privileged controls initialized. Background expiry checks are enabled.'};
 }finally{lock.releaseLock();}
}
function phase87TriggerReady_(){return ScriptApp.getProjectTriggers().some(function(t){return t.getHandlerFunction()==='phase87ExpiryWorker_';});}
function getPrivilegedAccessCenter_(){
 const actor=rbacRequirePermission_('manageUsers'),grants=accessTempGrantRows_().rows,users=rbacRows_().rows;
 const openGrantEmails={};grants.forEach(function(g){if(['REVERTED','CANCELLED'].indexOf(g.status)===-1)openGrantEmails[g.email]=true;});
 const permanentAdmins=users.filter(function(u){const role=rbacGetRoleDefinition_(u.role);return u.active&&role&&role.active&&role.builtIn&&role.name==='Admin'&&!openGrantEmails[u.email];});
 const permanentAdminEmails={};permanentAdmins.forEach(function(u){permanentAdminEmails[u.email]=true;});
 const canAdmin=!!permanentAdminEmails[actor.email];
 const emergencies=phase87Read_('Emergency_Access_Log',P87_EMERGENCY_HEADERS).map(function(v){return {id:String(v[0]),target:String(v[1]),actor:String(v[2]),incident:String(v[3]),reason:String(v[4]),start:accessRequestIso_(v[5]),end:accessRequestIso_(v[6]),grantId:String(v[7]),reviewStatus:String(v[8]),reviewer:String(v[9]),reviewedAt:accessRequestIso_(v[10]),reviewComment:String(v[11]),canReview:canAdmin&&v[8]==='PENDING'&&v[1]!==actor.email&&v[2]!==actor.email};}).reverse();
 const pending=accessRequestRows_().rows.filter(function(r){return phase87IsPrivileged_(r)&&['PENDING_APPROVAL','AWAITING_APPROVER','RETURNED'].indexOf(r.status)!==-1;}).map(function(r){const policy=accessRequestApprovalPolicy_(r),state=policy.dualAdmin?phase87ApprovalState_(r):{count:0};return {id:r.requestId,target:r.targetEmail,type:r.requestTypeLabel||accessRequestTypeLabel_(r.requestType),role:r.requestedRole,status:r.status,approver:r.approverEmail,approved:state.count,required:policy.requiredApprovals,policy:policy.key};});
 return {success:true,canAdmin:canAdmin,triggerReady:canAdmin?phase87TriggerReady_():null,adminCount:permanentAdmins.length,privilegedCount:users.filter(function(u){return u.active&&phase87RolePrivileged_(u.role);}).length,pending:pending,grants:grants.slice(-60).reverse(),emergencies:emergencies.slice(0,40),users:canAdmin?users.filter(function(u){return u.active&&!phase87PermanentAdminEmail_(u.email);}).map(function(u){return {email:u.email,name:u.name};}):[],audit:phase87Read_('Privileged_Audit_Log',P87_AUDIT_HEADERS).slice(-40).reverse().map(function(v){return {at:accessRequestIso_(v[0]),action:String(v[1]),actor:String(v[2]),target:String(v[3]),reference:String(v[4]),details:String(v[5])};})};
}
function grantEmergencyAccess_(input){
 const actor=phase87RequireAdmin_(),lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  input=input||{};const target=String(input.targetEmail||'').trim().toLowerCase(),incident=String(input.incident||'').trim(),reason=String(input.reason||'').trim(),minutes=Number(input.minutes);
  if(target===actor.email)throw new Error('You cannot grant emergency access to yourself.');
  if(!Number.isInteger(minutes)||minutes<15||minutes>240)throw new Error('Emergency access must be between 15 and 240 minutes.');
  if(incident.length<3||incident.length>120||/^[=+@-]/.test(incident))throw new Error('Enter a valid incident or ticket reference (3–120 characters).');
  if(reason.length<20||reason.length>1500||/^[=+@-]/.test(reason))throw new Error('Explain the emergency in 20–1500 characters.');
  if(!phase87TriggerReady_())throw new Error('Initialize Phase 8.7 first to enable background expiry.');
  const user=rbacRows_().rows.find(function(u){return u.email===target&&u.active;});if(!user)throw new Error('Choose an existing active dashboard user.');
  if(phase87PermanentAdminEmail_(target))throw new Error('This user already has permanent Admin access.');phase87AssertNoGrant_(target);
  if(phase87Read_('Emergency_Access_Log',P87_EMERGENCY_HEADERS).some(function(v){return v[1]===target&&v[8]==='PENDING';}))throw new Error('Complete the independent review of this user’s earlier emergency before granting another.');
  const id='BG-'+Utilities.getUuid().slice(0,12).toUpperCase(),now=new Date(),end=new Date(now.getTime()+minutes*60000),sheet=accessRequestEnsureSheet_('Emergency_Access_Log',P87_EMERGENCY_HEADERS),rowNo=sheet.getLastRow()+1;
  sheet.appendRow([id,target,actor.email,incident,reason,now.toISOString(),end.toISOString(),'','PENDING','','','']);
  phase87Audit_('EMERGENCY_AUTHORIZED',actor.email,target,id,{incident:incident,minutes:minutes,reason:reason});
  const grant=accessRequestCreateTempGrant_({requestId:id,targetEmail:target,targetName:user.name,requestedRole:'Admin',requestedDepartments:[],temporaryStartAt:now.toISOString(),temporaryEndAt:end.toISOString()},actor.email);
  sheet.getRange(rowNo,8).setValue(grant.grantId);
  return {success:true,message:'Emergency Admin access activated until '+end.toISOString()+'. The user must sign in again. Independent review is required.'};
 }finally{lock.releaseLock();}
}
function restorePrivilegedGrant_(grantId,reason){
 const actor=phase87RequireAdmin_(),lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  reason=String(reason||'').trim();if(reason.length<10||reason.length>1000)throw new Error('Enter a restoration reason of 10–1000 characters.');
  const data=accessTempGrantRows_(),g=data.rows.find(function(v){return v.grantId===String(grantId);});if(!g||['REVERTED','CANCELLED'].indexOf(g.status)!==-1)throw new Error('This grant is already closed or was not found.');
  if(g.email===actor.email)throw new Error('Another permanent Admin must restore your access.');
  const current=rbacRows_().rows.find(function(u){return u.email===g.email;});if(!current)throw new Error('Target user is missing; reconcile User_Access before restoring.');
  phase87Audit_('GRANT_RESTORE_STARTED',actor.email,g.email,g.grantId,{reason:reason,previousStatus:g.status});
  if(g.status!=='SCHEDULED')accessRequestApplyUserState_(actor.email,g.email,current.name,g.originalRole,g.originalDepartments,g.originalActive,'PRIVILEGED_GRANT_ENDED');
  data.sheet.getRange(g.row,11,1,5).setValues([[g.status==='SCHEDULED'?'CANCELLED':'REVERTED',g.createdAt,g.activatedAt,new Date().toISOString(),'']]);
  return {success:true,message:g.status==='SCHEDULED'?'Scheduled grant cancelled.':'Original access restored and sessions revoked.'};
 }finally{lock.releaseLock();}
}
function reviewEmergencyAccess_(id,comment){
 const actor=phase87RequireAdmin_(),lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  comment=String(comment||'').trim();if(comment.length<20||comment.length>1500||/^[=+@-]/.test(comment))throw new Error('Enter review findings in 20–1500 characters.');
  const rows=phase87Read_('Emergency_Access_Log',P87_EMERGENCY_HEADERS),index=rows.findIndex(function(v){return v[0]===String(id);}),r=rows[index];
  if(!r||r[8]!=='PENDING')throw new Error('This emergency was not found or is already reviewed.');
  if(r[1]===actor.email||r[2]===actor.email)throw new Error('An Admin independent of both the authorizer and target must perform this review.');
  phase87Audit_('EMERGENCY_REVIEWED',actor.email,r[1],id,{comment:comment});
  accessRequestEnsureSheet_('Emergency_Access_Log',P87_EMERGENCY_HEADERS).getRange(index+2,9,1,4).setValues([['REVIEWED',actor.email,new Date().toISOString(),comment]]);
  return {success:true,message:'Independent emergency review recorded.'};
 }finally{lock.releaseLock();}
}

/***************************************************************
 * PHASE 8.8 — IAM REPORTING
 * Read-only consolidated IAM governance reporting.
 * No access state is changed by this module.
 ***************************************************************/
function phase88Iso_(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString();
  const d = new Date(value);
  return isNaN(d.getTime()) ? String(value || '') : d.toISOString();
}
function phase88DaysSince_(value, nowMs) {
  if (!value) return null;
  const t = Date.parse(value);
  if (!Number.isFinite(t)) return null;
  return Math.max(0, Math.floor((nowMs - t) / 86400000));
}
function phase88Percent_(numerator, denominator) {
  if (!denominator) return null;
  return Math.max(0, Math.min(100, Math.round((Number(numerator || 0) / Number(denominator)) * 100)));
}
function phase88RoleScope_(role) {
  if (!role) return 'Role configuration missing';
  if (role.scopeType === 'DEPARTMENTS') return role.allowedDepartments && role.allowedDepartments.length ? 'Role departments: ' + role.allowedDepartments.join(', ') : 'Assigned departments';
  if (role.scopeType === 'USERS') return 'Restricted users';
  return 'All departments';
}
function phase88RequestPrivileged_(row, roleMap) {
  const type = String(row.requestType || '').toUpperCase();
  const requested = roleMap[String(row.requestedRole || '').toLowerCase()] || null;
  const current = roleMap[String(row.currentRole || '').toLowerCase()] || null;
  return !!row.privileged || type === 'PRIVILEGED_ACCESS' || accessRequestRoleIsPrivileged_(requested) || accessRequestRoleIsPrivileged_(current);
}
function phase88FastDepartmentNames_(users, roles) {
  // Phase 8.8 must not scan Compliance_Master just to build a department report.
  // That older path can be very expensive because readMasterRecords_ also applies IAM scope.
  const map = {};
  (users || []).forEach(function(user){ (user.departments || []).forEach(function(name){ name=String(name||'').trim(); if(name && !/^(all|\*)$/i.test(name)) map[name]=true; }); });
  (roles || []).forEach(function(role){ (role.allowedDepartments || []).forEach(function(name){ name=String(name||'').trim(); if(name && !/^(all|\*)$/i.test(name)) map[name]=true; }); });
  return Object.keys(map).sort(function(a,b){return a.localeCompare(b);});
}
function phase88CacheKey_(days){ return 'P88_V35_'+String(days); }
function phase88CacheGet_(days){
  try { const raw=CacheService.getScriptCache().get(phase88CacheKey_(days)); return raw?JSON.parse(raw):null; } catch(error){ return null; }
}
function phase88CachePut_(days,value){
  try { const raw=JSON.stringify(value); if(raw.length<95000) CacheService.getScriptCache().put(phase88CacheKey_(days),raw,90); } catch(error){}
}
function getIamReportingCenter_(input) {
  const startedMs = Date.now();
  const actor = rbacRequirePermission_('viewIamReports');
  input = input || {};
  const allowed = [0,7,30,90,365];
  let days = Number(input.days);
  if (allowed.indexOf(days) === -1) days = 30;
  if (!input.force) {
    const cached = phase88CacheGet_(days);
    if (cached && cached.success) {
      cached.actor = {email:actor.email,name:actor.name || '',role:actor.role || ''};
      cached.cached = true;
      cached.loadMs = Date.now() - startedMs;
      return cached;
    }
  }
  const now = new Date();
  const nowMs = now.getTime();
  const fromMs = days ? nowMs - days * 86400000 : 0;
  const periodLabel = days === 0 ? 'All available data' : (days === 365 ? 'Last 12 months' : 'Last ' + days + ' days');
  function inPeriod(value) {
    if (!days) return true;
    const t = Date.parse(value || '');
    return Number.isFinite(t) && t >= fromMs && t <= nowMs + 60000;
  }

  const users = rbacRows_().rows.slice();
  const roles = rbacReadRoleDefinitions_();
  const roleMap = {};
  roles.forEach(function(role){ roleMap[String(role.name || '').toLowerCase()] = role; });

  let credentials = [];
  try { credentials = dashboardCredentialRows_().rows; } catch (error) {}
  const credentialMap = {};
  credentials.forEach(function(row){ credentialMap[row.email] = row; });

  let sessions = [];
  try { sessions = dashboardSessionRows_().rows; } catch (error) {}
  const activeSessions = sessions.filter(function(row){ return !row.revokedAt; });
  const activeSessionCount = {};
  activeSessions.forEach(function(row){ activeSessionCount[row.email] = (activeSessionCount[row.email] || 0) + 1; });

  let audits = [];
  try { audits = dashboardSecurityAuditRows_(days && days <= 30 ? 120 : (days && days <= 90 ? 200 : DASHBOARD_SECURITY_AUDIT_LIMIT)); } catch (error) {}
  const securityPeriod = audits.filter(function(row){ return inPeriod(row.timestamp); });
  const failedSecurity = securityPeriod.filter(function(row){ return row.result === 'FAILED' || row.result === 'DENIED'; });

  let cycles = [], reviewItems = [];
  try { cycles = accessReviewCycleRows_().rows; reviewItems = accessReviewItemRows_().rows; } catch (error) {}
  cycles.sort(function(a,b){ return String(b.startAt || b.createdAt).localeCompare(String(a.startAt || a.createdAt)); });
  const latestCycle = cycles[0] || null;
  const latestReviewByEmail = {};
  reviewItems.forEach(function(item){
    const stamp = item.reviewedAt || '';
    if (!item.email || !stamp) return;
    const existing = latestReviewByEmail[item.email];
    const existingStamp = existing ? (existing.reviewedAt || '') : '';
    if (!existing || String(stamp) > String(existingStamp)) latestReviewByEmail[item.email] = item;
  });

  let requests = [];
  try { requests = accessRequestRows_().rows; } catch (error) {}
  const requestPeriod = requests.filter(function(row){ return inPeriod(row.createdAt); });
  const pendingRequests = requests.filter(function(row){ return row.status === 'PENDING_APPROVAL' || row.status === 'AWAITING_APPROVER'; });

  let grants = [];
  try { grants = accessTempGrantRows_().rows; } catch (error) {}
  const openGrants = grants.filter(function(row){ return ['REVERTED','CANCELLED'].indexOf(row.status) === -1; });

  let emergencies = [];
  try {
    emergencies = phase87Read_('Emergency_Access_Log',P87_EMERGENCY_HEADERS).map(function(v){ return {
      id:String(v[0] || ''), target:String(v[1] || '').trim().toLowerCase(), actor:String(v[2] || '').trim().toLowerCase(),
      incident:String(v[3] || ''), reason:String(v[4] || ''), startAt:phase88Iso_(v[5]), endAt:phase88Iso_(v[6]),
      grantId:String(v[7] || ''), reviewStatus:String(v[8] || ''), reviewer:String(v[9] || ''), reviewedAt:phase88Iso_(v[10])
    }; });
  } catch (error) {}

  function rolePrivileged(name) {
    return accessRequestRoleIsPrivileged_(roleMap[String(name || '').toLowerCase()] || null);
  }

  const userRows = users.map(function(user){
    const role = roleMap[String(user.role || '').toLowerCase()] || null;
    const cred = credentialMap[user.email] || null;
    const lastReview = latestReviewByEmail[user.email] || null;
    const lastLoginAt = cred ? phase88Iso_(cred.lastLoginAt) : '';
    const daysSinceLogin = phase88DaysSince_(lastLoginAt, nowMs);
    let loginState = 'Never logged in';
    if (lastLoginAt) loginState = daysSinceLogin === 0 ? 'Today' : (daysSinceLogin + ' days ago');
    const locked = !!(cred && cred.lockedUntil && Date.parse(cred.lockedUntil) > nowMs);
    const expired = !!(cred && cred.expiresAt && Date.parse(cred.expiresAt) <= nowMs);
    let credentialState = !cred ? 'Not configured' : (locked ? 'Locked' : (expired ? 'Expired' : 'Current'));
    const risks = [];
    if (!role) risks.push('Role missing');
    if (role && role.scopeType === 'DEPARTMENTS' && !(user.departments || []).length) risks.push('No department scope');
    if (user.active && !lastLoginAt) risks.push('Never logged in');
    if (user.active && daysSinceLogin !== null && daysSinceLogin >= 90) risks.push('Dormant 90+ days');
    if (locked) risks.push('Locked');
    if (expired) risks.push('Password expired');
    if (!user.active && (activeSessionCount[user.email] || 0) > 0) risks.push('Inactive with active session');
    return {
      email:user.email,name:user.name,role:user.role,departments:user.departments || [],scope:role && role.scopeType === 'DEPARTMENTS' ? ((user.departments || []).join(', ') || 'No department assigned') : phase88RoleScope_(role),
      active:!!user.active,privileged:rolePrivileged(user.role),lastLoginAt:lastLoginAt,daysSinceLogin:daysSinceLogin,loginState:loginState,
      credentialState:credentialState,passwordExpiresAt:cred ? phase88Iso_(cred.expiresAt) : '',lockedUntil:cred ? phase88Iso_(cred.lockedUntil) : '',
      activeSessions:activeSessionCount[user.email] || 0,lastReviewedAt:lastReview ? phase88Iso_(lastReview.reviewedAt) : '',lastReviewDecision:lastReview ? String(lastReview.decision || '') : '',
      riskSummary:risks.length ? risks.join(' · ') : 'None',riskCount:risks.length
    };
  }).sort(function(a,b){ return String(a.name || a.email).localeCompare(String(b.name || b.email)); });

  const roleRows = roles.map(function(role){
    const assigned = users.filter(function(user){ return user.role === role.name; });
    const permissionCount = Object.keys(rbacNormalizePermissions_(role.permissions)).filter(function(key){ return rbacNormalizePermissions_(role.permissions)[key]; }).length;
    return {name:role.name,description:role.description,builtIn:!!role.builtIn,active:!!role.active,scope:phase88RoleScope_(role),assignedUsers:assigned.length,activeUsers:assigned.filter(function(user){return user.active;}).length,privileged:accessRequestRoleIsPrivileged_(role),permissionCount:permissionCount};
  }).sort(function(a,b){ return String(a.name).localeCompare(String(b.name)); });

  const departmentSet = {};
  phase88FastDepartmentNames_(users,roles).forEach(function(name){ if(name) departmentSet[String(name)] = true; });
  users.forEach(function(user){ (user.departments || []).forEach(function(name){ if (name) departmentSet[String(name)] = true; }); });
  const departmentRows = Object.keys(departmentSet).sort().map(function(department){
    const assigned = users.filter(function(user){ return (user.departments || []).indexOf(department) !== -1; });
    return {department:department,scopeType:'Explicit department assignment',activeUsers:assigned.filter(function(user){return user.active;}).length,totalUsers:assigned.length,roles:Array.from(new Set(assigned.map(function(user){return user.role;}))).sort(),privilegedUsers:assigned.filter(function(user){return user.active && rolePrivileged(user.role);}).length};
  });
  const unrestricted = users.filter(function(user){ const role=roleMap[String(user.role||'').toLowerCase()]; return !role || role.scopeType !== 'DEPARTMENTS'; });
  if (unrestricted.length) departmentRows.unshift({department:'All Departments / Unrestricted',scopeType:'Role scope',activeUsers:unrestricted.filter(function(user){return user.active;}).length,totalUsers:unrestricted.length,roles:Array.from(new Set(unrestricted.map(function(user){return user.role;}))).sort(),privilegedUsers:unrestricted.filter(function(user){return user.active && rolePrivileged(user.role);}).length});

  const reviewRows = cycles.map(function(cycle){
    const overdue = cycle.status === 'Open' && cycle.dueAt && Date.parse(cycle.dueAt) < nowMs;
    return {cycleId:cycle.cycleId,cycleName:cycle.cycleName,frequency:cycle.frequency,status:overdue?'overdue':String(cycle.status || '').toLowerCase(),totalUsers:cycle.totalUsers,reviewedUsers:cycle.reviewedUsers,changesRequired:cycle.changesRequired,pendingUsers:Math.max(0,cycle.totalUsers-cycle.reviewedUsers),completionPercent:phase88Percent_(cycle.reviewedUsers,cycle.totalUsers) || 0,startAt:cycle.startAt,dueAt:cycle.dueAt,completedAt:cycle.completedAt,createdBy:cycle.createdBy};
  });

  const requestReportSource = requests.filter(function(row){ return !days || inPeriod(row.createdAt) || row.status === 'PENDING_APPROVAL' || row.status === 'AWAITING_APPROVER' || row.status === 'RETURNED'; });
  const requestRows = requestReportSource.slice().sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));}).slice(0,300).map(function(row){
    return {requestId:row.requestId,requestedBy:row.requestedBy,targetEmail:row.targetEmail,targetName:row.targetName,requestType:row.requestType,requestTypeLabel:accessRequestTypeLabel_(row.requestType),requestedRole:row.requestedRole,approverEmail:row.approverEmail,status:row.status,statusLabel:accessRequestStatusLabel_(row.status),privileged:phase88RequestPrivileged_(row,roleMap),createdAt:row.createdAt,reviewedAt:row.reviewedAt,implementedAt:row.implementedAt,temporaryStartAt:row.temporaryStartAt,temporaryEndAt:row.temporaryEndAt,ageDays:phase88DaysSince_(row.createdAt,nowMs) || 0};
  });

  const privilegedRows = [];
  userRows.filter(function(row){return row.active && row.privileged;}).forEach(function(row){privilegedRows.push({kind:'user',kindLabel:'Current privileged identity',title:row.name || row.email,target:row.email,access:row.role,status:row.active?'ACTIVE':'INACTIVE',startAt:'',endAt:'',reference:row.lastReviewDecision ? 'Review: '+row.lastReviewDecision : 'Not recertified'});});
  grants.filter(function(g){return !days || inPeriod(g.createdAt) || ['ACTIVE','SCHEDULED','APPLYING','ERROR'].indexOf(g.status)!==-1;}).sort(function(a,b){return String(b.createdAt).localeCompare(String(a.createdAt));}).slice(0,120).forEach(function(g){privilegedRows.push({kind:'grant',kindLabel:'Temporary access grant',title:g.grantId,target:g.email,access:g.grantedRole,status:g.status,startAt:g.startAt,endAt:g.endAt,reference:g.requestId});});
  emergencies.filter(function(e){return !days || inPeriod(e.startAt) || e.reviewStatus==='PENDING';}).sort(function(a,b){return String(b.startAt).localeCompare(String(a.startAt));}).slice(0,80).forEach(function(e){privilegedRows.push({kind:'emergency',kindLabel:'Emergency Admin access',title:e.id,target:e.target,access:'Admin · emergency',status:e.reviewStatus === 'PENDING' ? 'PENDING' : 'REVIEWED',startAt:e.startAt,endAt:e.endAt,reference:e.incident || e.grantId});});

  const securityRows = securityPeriod.slice(0,250).map(function(row){return {eventId:row.eventId,timestamp:row.timestamp,eventType:row.eventType,result:row.result,actorEmail:row.actorEmail,targetEmail:row.targetEmail,role:row.role,browser:row.browser,platform:row.platform,device:row.device};});

  const exceptions = [];
  const exceptionKeys = {};
  function addException(severity, area, title, target, detail, reference) {
    const key = [severity,area,title,target,reference].join('|');
    if (exceptionKeys[key]) return;
    exceptionKeys[key] = true;
    exceptions.push({severity:severity,area:area,title:title,target:target || '',detail:detail || '',reference:reference || ''});
  }

  userRows.forEach(function(row){
    const role = roleMap[String(row.role || '').toLowerCase()] || null;
    if (row.active && !role) addException('Critical','User access','Role configuration missing',row.email,'The active account references a role that is not present in Role_Definitions.',row.role);
    if (row.active && role && role.scopeType === 'DEPARTMENTS' && !row.departments.length) addException('High','User access','Department scope missing',row.email,'A department-scoped role is active without an assigned department.',row.role);
    if (!row.active && row.activeSessions > 0) addException('Critical','Session security','Inactive account has an active session',row.email,'Revoke active dashboard sessions for this inactive identity.',row.activeSessions+' session(s)');
    if (row.active && row.credentialState === 'Locked') addException('High','Authentication','Account is currently locked',row.email,'The dashboard credential is locked following authentication failures.',row.lockedUntil);
    if (row.active && row.credentialState === 'Expired') addException('High','Authentication','Dashboard password expired',row.email,'The user must update the dashboard password before normal access can continue.',row.passwordExpiresAt);
    if (row.active && row.loginState === 'Never logged in') addException('Medium','Account hygiene','Active account has never logged in',row.email,'Confirm whether the dashboard access is still required.',row.role);
    if (row.active && row.daysSinceLogin !== null && row.daysSinceLogin >= 90) addException('Medium','Account hygiene','Dormant active account',row.email,'No dashboard login has been recorded for '+row.daysSinceLogin+' days.',row.role);
    if (row.active && row.privileged) {
      const age = phase88DaysSince_(row.lastReviewedAt,nowMs);
      if (!row.lastReviewedAt || age === null || age >= 90) addException('High','Privileged access','Privileged access needs recertification',row.email,row.lastReviewedAt?'Last access review was '+age+' days ago.':'No completed access review is recorded for this privileged identity.',row.role);
    }
    if (row.active && row.passwordExpiresAt) {
      const expiry = Date.parse(row.passwordExpiresAt);
      const daysLeft = Number.isFinite(expiry) ? Math.ceil((expiry-nowMs)/86400000) : null;
      if (daysLeft !== null && daysLeft >= 0 && daysLeft <= 14) addException('Low','Authentication','Dashboard password expires soon',row.email,'Password expiry is due in '+daysLeft+' day'+(daysLeft===1?'':'s')+'.',row.passwordExpiresAt);
    }
  });

  cycles.forEach(function(cycle){ if (cycle.status === 'Open' && cycle.dueAt && Date.parse(cycle.dueAt) < nowMs) addException('High','Access review','Access review cycle overdue',cycle.cycleName,'The recertification due date has passed with '+Math.max(0,cycle.totalUsers-cycle.reviewedUsers)+' user(s) still pending.',cycle.cycleId); });

  pendingRequests.forEach(function(row){
    const age = phase88DaysSince_(row.createdAt,nowMs) || 0;
    const privileged = phase88RequestPrivileged_(row,roleMap);
    if (privileged && age >= 2) addException('High','Access requests','Privileged request awaiting action',row.targetEmail,'A privileged access request has remained open for '+age+' days.',row.requestId);
    else if (!privileged && age >= 3) addException('Medium','Access requests','Access request awaiting action',row.targetEmail,'The request has remained open for '+age+' days.',row.requestId);
  });

  openGrants.forEach(function(grant){
    const end = Date.parse(grant.endAt || '');
    if (grant.status === 'ERROR' || grant.status === 'APPLYING') addException('Critical','Privileged access','Temporary grant requires recovery',grant.email,grant.lastError || 'The grant is in a recovery-required state.',grant.grantId);
    else if (Number.isFinite(end) && end <= nowMs) addException('Critical','Privileged access','Temporary grant passed expiry',grant.email,'The grant is still open even though its configured end time has passed.',grant.grantId);
    else if (Number.isFinite(end) && end-nowMs <= 86400000) addException('Medium','Privileged access','Temporary grant expires within 24 hours',grant.email,'Confirm the activity is complete and original access will be restored.',grant.grantId);
  });

  emergencies.filter(function(row){return row.reviewStatus === 'PENDING';}).forEach(function(row){addException('High','Emergency access','Independent emergency review pending',row.target,'Emergency Admin access must be independently reviewed after use.',row.incident || row.id);});

  const failedByActor = {};
  failedSecurity.forEach(function(row){const key=row.actorEmail||row.targetEmail||'Unknown';failedByActor[key]=(failedByActor[key]||0)+1;});
  Object.keys(failedByActor).forEach(function(email){if(failedByActor[email]>=3)addException('High','Authentication','Repeated failed or denied authentication',email,failedByActor[email]+' failed/denied events were recorded during '+periodLabel+'.',periodLabel);});

  const severityRank = {Critical:4,High:3,Medium:2,Low:1};
  exceptions.sort(function(a,b){return (severityRank[b.severity]||0)-(severityRank[a.severity]||0) || String(a.area).localeCompare(String(b.area)) || String(a.target).localeCompare(String(b.target));});

  const latestCycleItems = latestCycle ? reviewItems.filter(function(item){return item.cycleId === latestCycle.cycleId;}) : [];
  const latestPrivilegedItems = latestCycleItems.filter(function(item){return rolePrivileged(item.role);});
  const reviewedPrivileged = latestPrivilegedItems.filter(function(item){return !!item.decision;}).length;
  const dormant90 = userRows.filter(function(row){return row.active && row.daysSinceLogin !== null && row.daysSinceLogin >= 90;}).length;
  const neverLogged = userRows.filter(function(row){return row.active && row.loginState === 'Never logged in';}).length;
  const activeUsers = userRows.filter(function(row){return row.active;}).length;
  const closedRequestPeriod = requestPeriod.filter(function(row){return ['APPROVED','REJECTED','CANCELLED'].indexOf(row.status)!==-1;}).length;
  const latestReviewPending = latestCycle ? Math.max(0,Number(latestCycle.totalUsers||0)-Number(latestCycle.reviewedUsers||0)) : 0;

  const result = {
    success:true,
    generatedAt:now.toISOString(),
    periodDays:days,
    periodLabel:periodLabel,
    actor:{email:actor.email,name:actor.name || '',role:actor.role || ''},
    summary:{
      totalUsers:userRows.length,
      activeUsers:activeUsers,
      inactiveUsers:userRows.length-activeUsers,
      privilegedUsers:userRows.filter(function(row){return row.active && row.privileged;}).length,
      pendingRequests:pendingRequests.length,
      reviewPendingUsers:latestReviewPending,
      dormant90:dormant90,
      neverLoggedIn:neverLogged,
      openTemporaryGrants:openGrants.filter(function(row){return ['ACTIVE','SCHEDULED','APPLYING','ERROR'].indexOf(row.status)!==-1;}).length,
      failedAuth:failedSecurity.length,
      lockedAccounts:userRows.filter(function(row){return row.active && row.credentialState==='Locked';}).length,
      activeSessions:activeSessions.length
    },
    health:{
      reviewCompletion:latestCycle ? phase88Percent_(latestCycle.reviewedUsers,latestCycle.totalUsers) : null,
      privilegedReviewCoverage:latestPrivilegedItems.length ? phase88Percent_(reviewedPrivileged,latestPrivilegedItems.length) : null,
      accountHygiene:activeUsers ? phase88Percent_(Math.max(0,activeUsers-dormant90-neverLogged),activeUsers) : null,
      requestClosure:requestPeriod.length ? phase88Percent_(closedRequestPeriod,requestPeriod.length) : null
    },
    exceptions:exceptions.slice(0,100),
    reports:{
      users:userRows,
      roles:roleRows,
      departments:departmentRows,
      reviews:reviewRows,
      requests:requestRows,
      privileged:privilegedRows,
      security:securityRows
    }
  };
  result.cached = false;
  result.loadMs = Date.now() - startedMs;
  phase88CachePut_(days,result);
  return result;
}


/***************************************************************
 * REPORTS HUB — IAM EVIDENCE EXPORT (V36)
 * Read-only XLSX workbook generated from Phase 8.8 reporting.
 ***************************************************************/
function exportIamReportingExcel_(input) {
  const actor = rbacRequirePermission_('manageUsers');
  const days = Number(input && input.days);
  const report = getIamReportingCenter_({days:isFinite(days)?days:30, force:false});
  const allowedScopes = ['users','roles','departments','reviews','requests','privileged','security'];
  const requestedScope = String(input && input.scope || 'all').toLowerCase();
  const scope = requestedScope === 'all' || allowedScopes.indexOf(requestedScope) !== -1 ? requestedScope : 'all';
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  const suffixMap = {users:'User_Access',roles:'Role_Matrix',departments:'Department_Access',reviews:'Access_Reviews',requests:'Access_Requests',privileged:'Privileged_Access',security:'Security_Activity'};
  const name = 'DigiClarity_IAM_' + (scope === 'all' ? 'Detailed_Reports' : (suffixMap[scope] || 'Report')) + '_' + stamp;
  const ss = SpreadsheetApp.create(name);
  try {
    const first = ss.getSheets()[0];
    first.setName('Executive Summary');
    const s = report.summary || {}, h = report.health || {};
    const summary = [
      ['DigiClarity Identity & Access Management Report',''],
      ['Reporting Period',report.periodLabel || ''],
      ['Generated At',report.generatedAt || ''],
      ['Generated By',actor.email || ''],['',''],
      ['Metric','Value'],
      ['Total Users',s.totalUsers || 0],['Active Users',s.activeUsers || 0],['Privileged Users',s.privilegedUsers || 0],
      ['Pending Requests',s.pendingRequests || 0],['Review Pending',s.reviewPendingUsers || 0],['Dormant 90+ Days',s.dormant90 || 0],
      ['Open Temporary Grants',s.openTemporaryGrants || 0],['Failed Auth Events',s.failedAuth || 0],['',''],
      ['Governance Health','Percent'],['Access Review Completion',h.reviewCompletion == null ? '' : h.reviewCompletion],
      ['Privileged Recertification',h.privilegedReviewCoverage == null ? '' : h.privilegedReviewCoverage],
      ['Active Account Hygiene',h.accountHygiene == null ? '' : h.accountHygiene],['Request Closure Rate',h.requestClosure == null ? '' : h.requestClosure]
    ];
    first.getRange(1,1,summary.length,2).setValues(summary);
    first.getRange('A1:B1').setFontWeight('bold').setFontSize(14);
    first.getRange(6,1,1,2).setFontWeight('bold'); first.getRange(16,1,1,2).setFontWeight('bold');
    first.setFrozenRows(1); first.autoResizeColumns(1,2);
    function addSheet_(title, rows) {
      const sh = ss.insertSheet(title); rows = Array.isArray(rows) ? rows : [];
      if (!rows.length) { sh.getRange(1,1).setValue('No records available for the selected reporting period.'); return; }
      const keys = [];
      rows.forEach(function(r){ Object.keys(r || {}).forEach(function(k){ if(keys.indexOf(k)<0)keys.push(k); }); });
      const values = [keys.map(function(k){return k.replace(/([A-Z])/g,' $1').replace(/^./,function(m){return m.toUpperCase();});})];
      rows.forEach(function(r){ values.push(keys.map(function(k){ const v=r[k]; return Array.isArray(v)?v.join(', '):(v===null||v===undefined?'':(typeof v==='object'?JSON.stringify(v):v)); })); });
      sh.getRange(1,1,values.length,keys.length).setValues(values);
      sh.getRange(1,1,1,keys.length).setFontWeight('bold'); sh.setFrozenRows(1); sh.autoResizeColumns(1,Math.min(keys.length,20));
    }
    const r = report.reports || {};
    const sheetMap = {
      users:['User Access',r.users], roles:['Role Matrix',r.roles], departments:['Department Access',r.departments],
      reviews:['Access Reviews',r.reviews], requests:['Access Requests',r.requests], privileged:['Privileged Access',r.privileged], security:['Security Activity',r.security]
    };
    const selected = scope === 'all' ? allowedScopes : [scope];
    selected.forEach(function(key){ const def = sheetMap[key]; if(def) addSheet_(def[0], def[1]); });
    if (scope === 'all') addSheet_('IAM Exceptions',report.exceptions || []);
    SpreadsheetApp.flush();

    // Google Sheets files cannot be converted to XLSX with Drive File#getAs().
    // Export the temporary spreadsheet through the authenticated Sheets export
    // endpoint instead. This preserves a real .xlsx workbook with all tabs.
    Utilities.sleep(350);
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(ss.getId()) + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(exportUrl, {
      method:'get',
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
      followRedirects:true
    });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      const body = String(response.getContentText() || '').slice(0,500);
      throw new Error('XLSX export failed (HTTP ' + status + '). ' + body);
    }
    const blob = response.getBlob().setName(name + '.xlsx');
    const bytes = blob.getBytes();
    if (!bytes || bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) {
      throw new Error('XLSX export returned an invalid workbook payload.');
    }
    return {
      fileName:name + '.xlsx',
      mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64:Utilities.base64Encode(bytes)
    };
  } finally {
    try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch(e) {}
  }
}


/***************************************************************
 * REPORTS HUB — SELECTED CSV / ALL CSV ZIP (V38)
 ***************************************************************/
function exportIamReportingCsvBundle_(input) {
  const actor = rbacRequirePermission_('manageUsers');
  const days = Number(input && input.days);
  const report = getIamReportingCenter_({days:isFinite(days)?days:30, force:false});
  const allowedScopes = ['users','roles','departments','reviews','requests','privileged','security'];
  const requestedScope = String(input && input.scope || 'all').toLowerCase();
  const scope = requestedScope === 'all' || allowedScopes.indexOf(requestedScope) !== -1 ? requestedScope : 'all';
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  const r = report.reports || {};
  const defs = {
    users:['User_Access',r.users], roles:['Role_Matrix',r.roles], departments:['Department_Access',r.departments],
    reviews:['Access_Reviews',r.reviews], requests:['Access_Requests',r.requests], privileged:['Privileged_Access',r.privileged], security:['Security_Activity',r.security]
  };
  function safeCell_(value) {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value)) value = value.join(', ');
    else if (typeof value === 'object') value = JSON.stringify(value);
    let s = String(value);
    if (/^[=+@-]/.test(s)) s = "'" + s;
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g,'""') + '"';
    return s;
  }
  function csvBlob_(key) {
    const def = defs[key];
    const rows = def && Array.isArray(def[1]) ? def[1] : [];
    const keys = [];
    rows.forEach(function(row){ Object.keys(row || {}).forEach(function(k){ if(keys.indexOf(k)<0)keys.push(k); }); });
    let csv;
    if (!rows.length) {
      csv = 'Message\r\n"No records available for the selected reporting period."';
    } else {
      const header = keys.map(function(k){ return safeCell_(k.replace(/([A-Z])/g,' $1').replace(/^./,function(m){return m.toUpperCase();})); }).join(',');
      const body = rows.map(function(row){ return keys.map(function(k){return safeCell_(row[k]);}).join(','); }).join('\r\n');
      csv = header + '\r\n' + body;
    }
    return Utilities.newBlob('\uFEFF' + csv, 'text/csv', 'DigiClarity_IAM_' + def[0] + '_' + stamp + '.csv');
  }
  if (scope !== 'all') {
    const blob = csvBlob_(scope);
    return {fileName:blob.getName(),mimeType:'text/csv',base64:Utilities.base64Encode(blob.getBytes())};
  }
  const blobs = allowedScopes.map(csvBlob_);
  const zip = Utilities.zip(blobs, 'DigiClarity_IAM_Detailed_Reports_' + stamp + '.zip');
  return {fileName:zip.getName(),mimeType:'application/zip',base64:Utilities.base64Encode(zip.getBytes()),generatedBy:actor.email || ''};
}


/***************************************************************
 * PHASE 9.1 — CENTRAL EVIDENCE REGISTER
 * References existing private IAM uploads. Metadata proposals use
 * IAM_Responses + IAM_Response_Reviews; no parallel approval engine.
 * Effective metadata is derived only from accepted proposals.
 ***************************************************************/
const E91_REGISTER_HEADERS=['Evidence ID','Record ID','Source Submission ID','File ID','File Name','MIME Type','Size Bytes','Uploaded By','Uploaded At','Indexed At'];
const E91_CHANGE_HEADERS=['Change ID','Evidence ID','Review Submission ID','Base Version','Metadata JSON','Proposed By','Proposed At','Client Token'];
function e91Read_(name,headers){
 const sheet=getSpreadsheet_().getSheetByName(name);if(!sheet)return [];
 const actual=sheet.getRange(1,1,1,headers.length).getDisplayValues()[0];
 if(headers.some(function(h,i){return actual[i]!==h;}))throw new Error(name+' headers are invalid. No evidence changes were made.');
 return sheet.getLastRow()<2?[]:sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();
}
function e91Actor_(){return rbacRequirePermission_('viewEvidence');}
function e91CanPropose_(a){return !!a.permissions.manageEvidence;}
function e91Text_(value){return String(value==null?'':value).trim();}
function e91Responses_(){return iamResponseRows_();}
function e91Index_(responses,records){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const sheet=accessRequestEnsureSheet_('Evidence_Register',E91_REGISTER_HEADERS),rows=e91Read_('Evidence_Register',E91_REGISTER_HEADERS),seen={};
  rows.forEach(function(r){seen[JSON.stringify([String(r[2]),String(r[3])])]=true;});
  const visible={};records.forEach(function(r){visible[r['Record ID']]=true;});
  responses.forEach(function(r){
   if(!visible[r['Record ID']]||String(r['Submission ID']).indexOf('EMD-')===0)return;
   iamPublicResponse_(r).files.forEach(function(f){
    if(!f.id)return;const key=JSON.stringify([r['Submission ID'],f.id]);if(seen[key])return;
    sheet.appendRow(['EV-'+Utilities.getUuid(),r['Record ID'],r['Submission ID'],f.id,iamSafeCell_(f.name),f.mime,Number(f.size)||0,r['Submitted By'],r['Submitted At'],new Date().toISOString()]);seen[key]=true;
   });
  });return e91Read_('Evidence_Register',E91_REGISTER_HEADERS);
 }finally{lock.releaseLock();}
}
function e91Context_(access,index){
 const records=readMasterRecords_().records.filter(function(r){return iamCanReadRecord_(access,r);}),responses=e91Responses_(),byRecord={},byResponse={};
 records.forEach(function(r){byRecord[r['Record ID']]=r;});responses.forEach(function(r){byResponse[r['Submission ID']]=r;});
 return {access:access,records:records,byRecord:byRecord,responses:responses,byResponse:byResponse,registry:index?e91Index_(responses,records):e91Read_('Evidence_Register',E91_REGISTER_HEADERS),changes:e91Read_('Evidence_Metadata_Changes',E91_CHANGE_HEADERS),users:rbacRows_().rows,roles:rbacReadRoleDefinitions_(),today:Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd')};
}
function e91Find_(id,c){
 const hits=c.registry.filter(function(r){return r[0]===id;});if(hits.length!==1)throw new Error('Evidence is unavailable. Refresh the register.');
 const row=hits[0],record=c.byRecord[row[1]],source=c.byResponse[row[2]];
 if(!record||!source||source['Record ID']!==row[1]||!iamCanReadRecord_(c.access,record)||!iamVisibleResponse_(source,record,c.access)||!iamPublicResponse_(source).files.some(function(f){return f.id===row[3];}))throw new Error('Evidence is unavailable or outside your department scope.');
 return {row:row,record:record,source:source};
}
function e91EligibleOwners_(record,c){
 return c.users.filter(function(u){
  if(!u.active||c.users.filter(function(v){return v.email===u.email;}).length!==1)return false;
  const role=c.roles.find(function(r){return r.name===u.role&&r.active;});if(!role)return false;
  const p=rbacNormalizeRole_(role).permissions;if(!p.viewEvidence)return false;
  return iamCanReadRecord_({departmentScoped:role.scopeType==='DEPARTMENTS',departments:u.departments||[]},record);
 }).map(function(u){return {email:u.email,name:u.name};});
}
function e91Date_(v,label){
 v=e91Text_(v);if(!/^\d{4}-\d{2}-\d{2}$/.test(v))throw new Error('Choose a valid '+label+'.');
 const d=new Date(v+'T00:00:00Z');if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==v)throw new Error('Choose a valid '+label+'.');return v;
}
function e91ValidateMetadata_(input,record,c){
 const title=e91Text_(input.title),owner=e91Text_(input.owner).toLowerCase(),notes=e91Text_(input.notes),reason=e91Text_(input.reason),noExpiry=input.noExpiry===true,exception=e91Text_(input.noExpiryReason);
 if(title.length<3||title.length>160)throw new Error('Evidence title must be 3–160 characters.');
 if(!e91EligibleOwners_(record,c).some(function(u){return u.email===owner;}))throw new Error('Choose an active owner with access to the linked record’s department.');
 if(reason.length<10||reason.length>1000)throw new Error('Explain this metadata change in 10–1000 characters.');
 if(notes.length>2000)throw new Error('Notes must be 2,000 characters or fewer.');
 const validFrom=e91Date_(input.validFrom,'valid-from date');let validUntil='';
 if(noExpiry){if(exception.length<10||exception.length>500)throw new Error('Explain why expiry is not applicable in 10–500 characters.');}
 else{validUntil=e91Date_(input.validUntil,'expiry date');if(validUntil<validFrom)throw new Error('Expiry cannot be before the valid-from date.');}
 return {title:title,owner:owner,validFrom:validFrom,validUntil:validUntil,noExpiry:noExpiry,noExpiryReason:noExpiry?exception:'',notes:notes,reason:reason};
}
function e91State_(entry,c){
 const history=[];let effective=null,version=0,pending=null;
 c.changes.filter(function(r){return r[1]===entry.row[0];}).forEach(function(r){
  const response=c.byResponse[r[2]];if(!response||response['Record ID']!==entry.row[1])return;
  let metadata;try{metadata=JSON.parse(String(r[4]));}catch(error){throw new Error('Evidence metadata is invalid. Ask Compliance to inspect the change history.');}
  const item={id:String(r[0]),submissionId:String(r[2]),baseVersion:Number(r[3]),metadata:metadata,proposedBy:String(r[5]),proposedAt:accessRequestIso_(r[6]),status:response.Status,reviewedBy:response['Reviewed By']||'',reviewedAt:response['Reviewed At']||'',reviewComment:response['Review Comment']||''};
  history.push(item);
  if(response.Status==='Accepted'&&Number(r[3])===version){effective=metadata;version++;}
  if(response.Status==='Submitted')pending=item;
 });
 let validity='Needs metadata';
 if(effective){
  if(!e91EligibleOwners_(entry.record,c).some(function(u){return u.email===effective.owner;}))validity='Owner unavailable';
  else if(effective.validFrom>c.today)validity='Not yet valid';
  else if(!effective.noExpiry&&effective.validUntil<c.today)validity='Expired';
  else if(!effective.noExpiry&&Date.parse(effective.validUntil+'T00:00:00Z')-Date.parse(c.today+'T00:00:00Z')<=30*86400000)validity='Expiring soon';
  else validity='Current';
 }
 const review=entry.source.Status;
 const status=review==='More information required'?'Returned':review!=='Accepted'?'Pending review':validity;
 return {version:version,metadata:effective,pending:pending,history:history,validity:validity,status:status,reviewStatus:review};
}
function e91Public_(entry,c,detail){
 const state=e91State_(entry,c),r=entry.row,m=state.metadata||{},record=entry.record;
 const item={id:r[0],recordId:r[1],submissionId:r[2],fileId:r[3],fileName:r[4],mime:r[5],size:Number(r[6]),uploadedBy:r[7],uploadedAt:accessRequestIso_(r[8]),title:m.title||r[4],owner:m.owner||'',department:record.Department||'',recordTitle:record['Record Title']||'',framework:record['Primary Compliance Framework']||'',controlId:record['Requirement / Control ID']||'',status:state.status,validity:state.validity,reviewStatus:state.reviewStatus,validFrom:m.validFrom||'',validUntil:m.validUntil||'',noExpiry:!!m.noExpiry,version:state.version,pending:!!state.pending,canPropose:e91CanPropose_(c.access)&&!state.pending,canReview:!!state.pending&&c.access.permissions.manageEvidence&&c.access.permissions.reviewResponses&&state.pending.proposedBy!==c.access.email};
 if(detail){item.metadata=state.metadata;item.proposal=state.pending;item.history=state.history.slice().reverse();item.owners=e91CanPropose_(c.access)?e91EligibleOwners_(record,c):[];}
 return item;
}
function getEvidenceManagement_(){
 const access=e91Actor_(),c=e91Context_(access,true),items=[];
 c.registry.forEach(function(r){if(!c.byRecord[r[1]])return;const source=c.byResponse[r[2]];if(!source||!iamVisibleResponse_(source,c.byRecord[r[1]],access))return;const entry=e91Find_(r[0],c);items.push(e91Public_(entry,c,false));});
 items.sort(function(a,b){return String(b.uploadedAt).localeCompare(String(a.uploadedAt));});
 return {success:true,items:items,today:c.today,canUpload:!!access.permissions.uploadEvidence&&!!access.permissions.submitEvidence,canPropose:e91CanPropose_(access),records:c.records.map(function(r){return {id:r['Record ID'],title:r['Record Title'],department:r.Department,canUpload:!!access.permissions.uploadEvidence&&!!access.permissions.submitEvidence&&!phase3IsTerminal_(r)};}),summary:{total:items.length,current:items.filter(function(i){return ['Current','Expiring soon'].indexOf(i.status)!==-1;}).length,expiring:items.filter(function(i){return i.status==='Expiring soon';}).length,expired:items.filter(function(i){return i.status==='Expired';}).length,needsMetadata:items.filter(function(i){return ['Needs metadata','Owner unavailable'].indexOf(i.validity)!==-1;}).length,pending:items.filter(function(i){return i.pending||i.reviewStatus==='Submitted';}).length}};
}
function getEvidenceDetails_(id){
 const c=e91Context_(e91Actor_(),false),item=e91Public_(e91Find_(e91Text_(id),c),c,true);
 if(item.framework&&item.controlId){const control=readFrameworkControlMaster_().controls.find(function(v){return phase3FrameworkKey_(v.framework)===phase3FrameworkKey_(item.framework)&&String(v.controlId).toLowerCase()===String(item.controlId).toLowerCase();});item.controlName=control?control.controlName:'';item.commonControlArea=control?control.commonControlArea:'';}
 return item;
}
function proposeEvidenceMetadata_(input){
 const access=e91Actor_();if(!e91CanPropose_(access))throw new Error('Your role can view evidence but cannot propose metadata changes.');
 input=input||{};const token=e91Text_(input.clientToken);if(!/^[a-zA-Z0-9_-]{16,100}$/.test(token))throw new Error('Reload the evidence form before submitting.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const c=e91Context_(access,false),entry=e91Find_(e91Text_(input.id),c),state=e91State_(entry,c),retry=c.changes.find(function(r){return r[7]===token&&r[5]===access.email;});
  if(retry){if(!c.byResponse[retry[2]])throw new Error('An incomplete metadata submission needs administrator recovery before retrying.');if(retry[1]!==entry.row[0])throw new Error('This submission token belongs to another evidence item.');return {success:true,message:'This metadata proposal was already submitted.',submissionId:retry[2]};}
  if(state.pending)throw new Error('This evidence already has a pending metadata proposal. Review it before submitting another.');
  if(Number(input.expectedVersion)!==state.version)throw new Error('Evidence was updated by another user. Refresh before editing.');
  const metadata=e91ValidateMetadata_(input,entry.record,c),changeId='EMC-'+Utilities.getUuid(),submissionId='EMD-'+Utilities.getUuid(),now=new Date().toISOString();
  const changes=accessRequestEnsureSheet_('Evidence_Metadata_Changes',E91_CHANGE_HEADERS),responses=iamSheet_(IAM_RESPONSE_SHEET,IAM_RESPONSE_HEADERS),changeRow=changes.getLastRow()+1,responseRow=responses.getLastRow()+1;
  const description='EVIDENCE METADATA REVIEW — '+entry.row[0]+'\nFile: '+entry.row[4]+'\nApproved version: '+state.version+'\nProposed title: '+metadata.title+'\nOwner: '+metadata.owner+'\nValid from: '+metadata.validFrom+'\nExpiry: '+(metadata.noExpiry?'Not applicable — '+metadata.noExpiryReason:metadata.validUntil)+'\nReason: '+metadata.reason+'\nNotes: '+metadata.notes;
  const file=iamPublicResponse_(entry.source).files.find(function(f){return f.id===entry.row[3];});
  try{
   changes.getRange(changeRow,1,1,E91_CHANGE_HEADERS.length).setValues([[changeId,entry.row[0],submissionId,state.version,JSON.stringify(metadata),access.email,now,token]]);
   responses.getRange(responseRow,1,1,IAM_RESPONSE_HEADERS.length).setValues([[submissionId,entry.row[1],iamSafeCell_(entry.record.Department),access.email,iamSafeCell_(access.name),iamSafeCell_(description),JSON.stringify([file]),now,'Submitted','','','','',token]]);
  }catch(error){changes.getRange(changeRow,1,1,E91_CHANGE_HEADERS.length).clearContent();responses.getRange(responseRow,1,1,IAM_RESPONSE_HEADERS.length).clearContent();throw new Error('Metadata proposal was not saved. '+(error.message||error));}
  return {success:true,submissionId:submissionId,message:'Metadata proposal sent for independent review. Approved evidence details remain unchanged until acceptance.'};
 }finally{lock.releaseLock();}
}
function e91ValidateReview_(item,decision,access){
 if(String(item['Submission ID']).indexOf('EMD-')!==0)return;
 const c=e91Context_(access,false),matches=c.changes.filter(function(r){return r[2]===item['Submission ID'];});if(matches.length!==1)throw new Error('Evidence metadata proposal is missing or ambiguous.');
 const change=matches[0],entry=e91Find_(String(change[1]),c),state=e91State_(entry,c);
 if(item['Submitted By']!==change[5])throw new Error('Metadata proposer does not match the review submission.');
 if(decision==='Accepted'){
  if(Number(change[3])!==state.version)throw new Error('Approved metadata changed after this proposal. Return it for a fresh submission.');
  e91ValidateMetadata_(JSON.parse(String(change[4])),entry.record,c);
 }
}

/***************************************************************
 * PHASE 9.2 — EVIDENCE REQUEST WORKFLOW
 * Immutable requests + append-only response/cancellation events.
 * Uploads and independent approvals retain the existing IAM flow.
 * Completion is derived from the accepted linked response, never
 * from a second status write. Creation emails are delivered to the
 * assigned responder(s) and audited; reminders belong to Phase 9.4.
 ***************************************************************/
const E92_REQUEST_HEADERS=['Request ID','Record ID','Department','Title','Instructions','Assigned To','Due Date','Created By','Creator Name','Created At','Client Token'];
const E92_EVENT_HEADERS=['Event ID','Request ID','Event Type','Submission ID','Reason','Actor','Created At'];

/***************************************************************
 * PHASE 9.2.1 — COMPACT EVIDENCE REQUEST MAIL
 * No finding narrative, title, instructions, PHI or attachments in email.
 * The dashboard performs record-level authorization after login.
 ***************************************************************/
const E92_MAIL_HEADERS=['Notification ID','Request ID','Recipient','Result','Sent At','Error'];
function e92MailEsc_(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function e92MailUrl_(){
 try{const configured=String(PropertiesService.getScriptProperties().getProperty('DASHBOARD_PUBLIC_URL')||'').trim();if(/^https:\/\//.test(configured))return configured;const url=String(ScriptApp.getService().getUrl()||'');return /^https:\/\/script\.google\.com\/macros\//.test(url)?url:'';}catch(e){return '';}
}
function e92MailShortType_(type){
 return normalizeRecordTypeKey_(String(type||''))||'Compliance Record';
}
function e92MailFormatDue_(due){
 const d=String(due||'');return /^\d{4}-\d{2}-\d{2}$/.test(d)?Utilities.formatDate(new Date(d+'T12:00:00+05:30'),'Asia/Kolkata','dd MMM yyyy'):d;
}

/** Shared DigiClarity email brand asset. Also usable by future mail templates. */
function e92MailInlineImages_(){
 const encoded='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';
 return {dgsLogo:Utilities.newBlob(Utilities.base64Decode(encoded),'image/png','DigiClarity.png')};
}

function e92MailTemplate_(data){
 const type=e92MailShortType_(data.recordType),due=e92MailFormatDue_(data.dueDate),dept=String(data.department||'Assigned department');
 const shortName=String(data.name||'').trim().split(/\s+/)[0].slice(0,60)||'Team member';
 const test=!!data.test,base=e92MailUrl_();
 if (!base) throw new Error('Email action URL is not configured. Run configureDashboardEmailLinkOnce_() with your HTTPS dashboard URL.');
 const link=base+(base.indexOf('?')>=0?'&':'?')+'route=evidence-requests';
 const badge=test?'DESIGN TEST · NO ACTION REQUIRED':'ACTION REQUIRED';
 const heading=test?'Your evidence request email preview':'Evidence requested for '+type;
 const description=test?'This is a sample of the automated evidence request notification. No actual compliance request was created.':'The Compliance team has requested supporting evidence for a '+type.toLowerCase()+' in your department. Please review the requirements securely in the dashboard and submit the requested evidence.';
 const cta=link?'<tr><td style="padding:24px 0 3px;text-align:center"><a href="'+e92MailEsc_(link)+'" style="display:inline-block;background:#336cf0;color:#ffffff;text-decoration:none;font-size:14px;font-weight:700;padding:15px 25px;border-radius:11px">SIGN IN &amp; VIEW REQUEST&nbsp; &#8594;</a></td></tr>':'';
 const html='<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head><body style="margin:0;background:#eef3fb;padding:30px 12px;font-family:Arial,Helvetica,sans-serif;color:#1a2740">'+
 '<table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width:600px;margin:0 auto;border-collapse:separate;border-spacing:0;background:#ffffff;border:1px solid #e3eaf4;border-radius:20px;overflow:hidden">'+
 '<tr><td style="background:#142442;padding:27px 32px 28px;border-top:5px solid #4483ff"><table role="presentation" width="100%"><tr><td><img src="cid:dgsLogo" alt="DigiClarity" width="265" style="display:block;width:265px;max-width:100%;height:auto;border:0"></td></tr><tr><td style="color:#b9c9e8;font-size:11px;letter-spacing:1.8px;padding-top:9px">COMPLIANCE CENTER&nbsp;&nbsp; / &nbsp;&nbsp;EVIDENCE MANAGEMENT</td></tr></table></td></tr>'+
 '<tr><td style="padding:32px 32px 14px"><span style="display:inline-block;padding:7px 11px;background:'+(test?'#e8f0ff':'#eaf1ff')+';color:#245ad1;border-radius:6px;font-size:10px;font-weight:800;letter-spacing:1.1px">'+e92MailEsc_(badge)+'</span><h1 style="font-size:24px;line-height:1.24;letter-spacing:-.55px;margin:18px 0 13px;color:#15223b">'+e92MailEsc_(heading)+'</h1><p style="font-size:14px;line-height:1.75;color:#596981;margin:0">Hi '+e92MailEsc_(shortName)+',<br>'+e92MailEsc_(description)+'</p></td></tr>'+
 '<tr><td style="padding:9px 32px 0"><table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border:1px solid #e6ecf6;border-radius:13px;background:#f7f9fe"><tr><td colspan="2" style="padding:19px 19px 11px;font-size:11px;font-weight:800;color:#6780a3;letter-spacing:1.2px">REQUEST SNAPSHOT</td></tr>'+
 '<tr><td style="padding:7px 19px;width:39%;font-size:12px;color:#71829c">Record type</td><td style="padding:7px 19px 7px 0;font-size:13px;font-weight:700;color:#182944">'+e92MailEsc_(type)+'</td></tr>'+
 '<tr><td style="padding:7px 19px;font-size:12px;color:#71829c">Department</td><td style="padding:7px 19px 7px 0;font-size:13px;font-weight:700;color:#182944">'+e92MailEsc_(dept)+'</td></tr>'+
 '<tr><td style="padding:7px 19px 21px;font-size:12px;color:#71829c">Submit by</td><td style="padding:7px 19px 21px 0;font-size:13px;font-weight:700;color:#182944">'+e92MailEsc_(due)+'</td></tr></table></td></tr>'+
 '<tr><td style="padding:0 32px"><table role="presentation" width="100%">'+cta+'</table><p style="font-size:12px;color:#6b7890;text-align:center;line-height:1.6;margin:14px 0 26px">Use the button above to securely open your assigned request after signing in.</p></td></tr>'+
 '<tr><td style="padding:18px 32px;background:#f7faff;border-top:1px solid #e4eaf5;font-size:11px;line-height:1.7;color:#74839d">This notification intentionally excludes finding descriptions, detailed evidence requirements and any sensitive records. View them only in the access-controlled dashboard.<br><b style="color:#4c5f7e">DigiClarity Compliance Team</b> &middot; Automated notification</td></tr></table><p style="max-width:600px;margin:15px auto 0;text-align:center;color:#8795ab;font-size:11px">Please do not send confidential records or patient information by email.</p></body></html>';
 const subject=test?'[TEST] Evidence request notification preview':'Action required | Evidence requested for '+type;
 const text='Hi '+shortName+',\n\n'+description+'\n\nRecord type: '+type+'\nDepartment: '+dept+'\nSubmit by: '+due+'\n\nOpen your Compliance Dashboard'+(link?': '+link:'')+'\nNavigate to Evidence Management > Requests.\n\nPlease do not reply with evidence or patient information.\nDigiClarity Compliance Team';
 return {subject:subject,html:html,text:text};
}
function e92MailLog_(id,email,status,error){
 try{e92Sheet_('Evidence_Request_Notifications',E92_MAIL_HEADERS).appendRow(['ERN-'+Utilities.getUuid(),id,email,status,new Date().toISOString(),iamSafeCell_(String(error||'').slice(0,250))]);}catch(e){Logger.log('Evidence mail logging failed: '+e);}
}
function e92MailHistory_(id){return e91Read_('Evidence_Request_Notifications',E92_MAIL_HEADERS).filter(function(row){return row[1]===id;});}
function e92MailSend_(data,eligible,onlyUnsent){
 const history=e92MailHistory_(data.id),sent={};history.forEach(function(row){if(row[3]==='SENT')sent[String(row[2]).toLowerCase()]=true;});
 const list=eligible.filter(function(u){return !onlyUnsent||!sent[u.email.toLowerCase()];});
 const result={sent:0,failed:0,skipped:eligible.length-list.length};
 if(list.length>40)throw new Error('Request is assigned to more than 40 responders. Assign a specific user before sending notifications.');
 list.forEach(function(u){
  try{
   const msg=e92MailTemplate_(Object.assign({},data,{name:u.name}));
   MailApp.sendEmail({to:u.email,subject:msg.subject,body:msg.text,htmlBody:msg.html,inlineImages:e92MailInlineImages_(),name:'DigiClarity Compliance'});
   e92MailLog_(data.id,u.email,'SENT','');result.sent++;
  }catch(err){e92MailLog_(data.id,u.email,'FAILED',err&&err.message?err.message:String(err));result.failed++;}
 });
 return result;
}
/** Safe self-mail only. Never accepts an arbitrary test recipient. */
function sendEvidenceRequestTestEmail_(){
 const access=rbacRequirePermission_('manageEvidenceRequests');if(!access.permissions.reviewResponses)throw new Error('Response review permission is required.');
 if(!/^[^\s@]+@(digiclarity\.in|clarityrcm\.com)$/i.test(access.email))throw new Error('Test mail requires your signed-in company account.');
 const sample=e92MailTemplate_({test:true,name:access.name||'Compliance',recordType:'Audit Finding',department:'Compliance',dueDate:Utilities.formatDate(new Date(Date.now()+2*86400000),'Asia/Kolkata','yyyy-MM-dd')});
 MailApp.sendEmail({to:access.email,subject:sample.subject,body:sample.text,htmlBody:sample.html,inlineImages:e92MailInlineImages_(),name:'DigiClarity Compliance'});
 return {success:true,recipient:access.email,message:'Test email sent to your company mailbox: '+access.email};
}
/** Re-send only to eligible users not already confirmed SENT; scope is revalidated. */
function retryEvidenceRequestNotification_(id){
 const access=rbacRequirePermission_('manageEvidenceRequests'),c=e92Context_(access),row=e92Find_(e91Text_(id),c),state=e92State_(row,c);if(!access.permissions.reviewResponses)throw new Error('Response review permission is required.');
 if(['Cancelled','Completed'].indexOf(state.status)!==-1)throw new Error('Cannot notify for a completed or cancelled request.');
 const record=c.byRecord[row[1]],eligible=e92Responders_(record,c).filter(function(u){return !row[5]||u.email===row[5];});
 if(!eligible.length)throw new Error('No eligible responder is currently assigned.');
 const summary=e92MailSend_({id:row[0],recordType:record['Record Type'],department:row[2],dueDate:row[6]},eligible,true);
 return {success:true,email:summary,message:summary.sent+' email(s) delivered; '+summary.failed+' failed; '+summary.skipped+' previously sent.'};
}

function e92Sheet_(name,headers){
 const existing=getSpreadsheet_().getSheetByName(name);
 if(existing){e91Read_(name,headers);return existing;}
 return accessRequestEnsureSheet_(name,headers);
}
function e92Due_(value){
 return value&&typeof value.getTime==='function'?Utilities.formatDate(value,Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd'):e91Date_(value,'request due date');
}
function e92Context_(access){
 const records=readMasterRecords_().records.filter(function(r){return iamCanReadRecord_(access,r);}),byRecord={},byResponse={};
 records.forEach(function(r){byRecord[r['Record ID']]=r;});
 iamResponseRows_().forEach(function(r){byResponse[r['Submission ID']]=r;});
 const requests=e91Read_('Evidence_Requests',E92_REQUEST_HEADERS).map(function(r){r[6]=e92Due_(r[6]);return r;});
 return {access:access,records:records,byRecord:byRecord,byResponse:byResponse,requests:requests,events:e91Read_('Evidence_Request_Events',E92_EVENT_HEADERS),users:rbacRows_().rows,roles:rbacReadRoleDefinitions_(),today:Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd')};
}
function e92Visible_(row,c){
 const record=c.byRecord[row[1]];
 return !!record&&iamDepartmentKey_(record.Department)===iamDepartmentKey_(row[2])&&iamCanReadRecord_(c.access,record);
}
function e92Find_(id,c){
 const rows=c.requests.filter(function(r){return r[0]===id;});
 if(rows.length!==1||!e92Visible_(rows[0],c))throw new Error('Request is unavailable, outside your departments, or its record has changed department.');
 return rows[0];
}
function e92Responders_(record,c){
 return e91EligibleOwners_(record,c).filter(function(u){
  const user=c.users.find(function(v){return v.email===u.email;}),role=c.roles.find(function(r){return r.name===user.role&&r.active;});
  return !!(rbacNormalizeRole_(role).permissions.submitEvidence&&rbacNormalizeRole_(role).permissions.uploadEvidence);
 });
}
function e92State_(row,c){
 const events=c.events.filter(function(e){return e[1]===row[0];}),history=[];let cancelled=null;
 events.forEach(function(e){
  if(e[2]==='Cancelled'){cancelled={reason:String(e[4]),by:String(e[5]),at:accessRequestIso_(e[6])};return;}
  if(e[2]!=='Submitted')throw new Error('Request history contains an unknown event. Ask Compliance to check it.');
  const response=c.byResponse[e[3]];
  if(!response||response['Record ID']!==row[1]||response['Submitted By']!==e[5]||iamDepartmentKey_(response.Department)!==iamDepartmentKey_(row[2])||!iamVisibleResponse_(response,c.byRecord[row[1]],c.access))throw new Error('A request submission link is incomplete. Ask Compliance to inspect its history.');
  if(!iamPublicResponse_(response).files.length)throw new Error('The request submission has no evidence file.');
  history.push(response);
 });
 const latest=history.length?history[history.length-1]:null;
 if(latest&&['Accepted','Submitted','More information required'].indexOf(latest.Status)===-1)throw new Error('Request review status is invalid.');
 const status=cancelled?'Cancelled':!latest?'Open':latest.Status==='Accepted'?'Completed':latest.Status==='Submitted'?'In review':'Returned';
 return {status:status,latest:latest,history:history,cancelled:cancelled,overdue:['Completed','Cancelled'].indexOf(status)===-1&&String(row[6])<c.today};
}
function e92Public_(row,c,detail){
 const state=e92State_(row,c),record=c.byRecord[row[1]],responders=e92Responders_(record,c),assignee=responders.find(function(u){return u.email===row[5];}),open=['Open','Returned'].indexOf(state.status)!==-1,terminal=phase3IsTerminal_(record);
 const canRespond=open&&!terminal&&!!c.access.permissions.submitEvidence&&!!c.access.permissions.uploadEvidence&&(!row[5]||row[5]===c.access.email);
 const item={id:row[0],recordId:row[1],department:row[2],title:row[3],assignee:row[5],assigneeName:assignee?assignee.name:row[5]||'Department team',assigneeUnavailable:!!row[5]&&!assignee,dueDate:String(row[6]),createdBy:row[7],creatorName:row[8],createdAt:accessRequestIso_(row[9]),recordTitle:record['Record Title']||'',framework:record['Primary Compliance Framework']||'',controlId:record['Requirement / Control ID']||'',status:state.status,overdue:state.overdue,canRespond:canRespond,canReview:state.status==='In review'&&!!c.access.permissions.manageEvidenceRequests&&!!c.access.permissions.reviewResponses&&state.latest['Submitted By']!==c.access.email,canCancel:open&&!!c.access.permissions.manageEvidenceRequests&&!!c.access.permissions.reviewResponses,submissionCount:state.history.length,latestSubmissionId:state.latest?state.latest['Submission ID']:'',recordClosed:terminal};
 if(detail){item.instructions=row[4];item.history=state.history.map(iamPublicResponse_).reverse();item.cancellation=state.cancelled;}
 return item;
}
function getEvidenceRequests_(){
 const access=e91Actor_(),c=e92Context_(access),items=c.requests.filter(function(r){return e92Visible_(r,c);}).map(function(r){return e92Public_(r,c,false);});
 items.sort(function(a,b){const closed=function(x){return ['Completed','Cancelled'].indexOf(x.status)!==-1?1:0;};return closed(a)-closed(b)||a.dueDate.localeCompare(b.dueDate)||b.createdAt.localeCompare(a.createdAt);});
 return {items:items,today:c.today,canCreate:!!access.permissions.manageEvidenceRequests&&!!access.permissions.reviewResponses,summary:{open:items.filter(function(i){return i.status==='Open';}).length,inReview:items.filter(function(i){return i.status==='In review';}).length,returned:items.filter(function(i){return i.status==='Returned';}).length,overdue:items.filter(function(i){return i.overdue;}).length,completed:items.filter(function(i){return i.status==='Completed';}).length},records:access.permissions.manageEvidenceRequests&&access.permissions.reviewResponses?c.records.filter(function(r){return !phase3IsTerminal_(r)&&e91Text_(r.Department);}).map(function(r){return {id:r['Record ID'],title:r['Record Title'],department:r.Department,responders:e92Responders_(r,c)};}):[]};
}
function getEvidenceRequestDetails_(id){const c=e92Context_(e91Actor_()),item=e92Public_(e92Find_(e91Text_(id),c),c,true);if(c.access.permissions.manageEvidenceRequests&&c.access.permissions.reviewResponses){const mail=e92MailHistory_(item.id),latest={};mail.forEach(function(x){latest[String(x[2]).toLowerCase()]=x[3];});const statuses=Object.keys(latest).map(function(email){return latest[email];});item.mail={sent:statuses.filter(function(x){return x==='SENT';}).length,failed:statuses.filter(function(x){return x==='FAILED';}).length};item.canRetryMail=['Open','Returned'].indexOf(item.status)!==-1;}return item;}
function createEvidenceRequest_(input){
 const access=rbacRequirePermission_('manageEvidenceRequests');if(!access.permissions.reviewResponses)throw new Error('Independent evidence request review permission is required.');input=input||{};
 const token=e91Text_(input.clientToken);if(!/^[a-zA-Z0-9_-]{16,100}$/.test(token))throw new Error('Reload the request form before submitting.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 let notice=null,existingId='';
 try{
  const c=e92Context_(access),record=iamRequireRecord_(input.recordId,access),retry=c.requests.find(function(r){return r[10]===token&&r[7]===access.email;});
  if(retry){if(retry[1]!==record['Record ID'])throw new Error('This token belongs to another record.');existingId=retry[0];}
  else{
   if(phase3IsTerminal_(record)||!e91Text_(record.Department))throw new Error('Choose an open compliance record with an assigned department.');
   const title=e91Text_(input.title),instructions=e91Text_(input.instructions),assignee=e91Text_(input.assignee).toLowerCase(),due=e91Date_(input.dueDate,'due date'),responders=e92Responders_(record,c);
   if(title.length<3||title.length>160)throw new Error('Request title must be 3–160 characters.');
   if(instructions.length<10||instructions.length>2000)throw new Error('Explain the evidence required in 10–2,000 characters.');
   if(due<c.today)throw new Error('The due date cannot be in the past.');
   if(!responders.length||assignee&&!responders.some(function(u){return u.email===assignee;}))throw new Error('Choose an active responder with upload permission and access to this department.');
   const eligible=responders.filter(function(u){return !assignee||u.email===assignee;});
   if(eligible.length>40)throw new Error('More than 40 department responders are eligible. Please assign a specific responder.');
   const sheet=e92Sheet_('Evidence_Requests',E92_REQUEST_HEADERS),id='ERQ-'+Utilities.getUuid();
   sheet.appendRow([id,record['Record ID'],iamSafeCell_(record.Department),iamSafeCell_(title),iamSafeCell_(instructions),assignee,due,access.email,iamSafeCell_(access.name),new Date().toISOString(),token]);
   notice={data:{id:id,recordType:record['Record Type'],department:record.Department,dueDate:due},eligible:eligible};
  }
 }finally{lock.releaseLock();}
 if(existingId)return {success:true,id:existingId,message:'This evidence request already exists. No duplicate notification was sent.'};
 // The request is committed first. Mail failure never rolls back its evidence audit trail.
 let mail;
 try{mail=e92MailSend_(notice.data,notice.eligible,true);}catch(error){mail={sent:0,failed:notice.eligible.length,skipped:0};e92MailLog_(notice.data.id,'','FAILED',error.message||String(error));}
 return {success:true,id:notice.data.id,email:mail,message:'Evidence request created. '+mail.sent+' notification email(s) sent'+(mail.failed?'; '+mail.failed+' failed — use Retry email in request details.':'.')};
}
function cancelEvidenceRequest_(input){
 const access=rbacRequirePermission_('manageEvidenceRequests');if(!access.permissions.reviewResponses)throw new Error('Evidence request management requires response review permission.');input=input||{};
 const reason=e91Text_(input.reason);if(reason.length<10||reason.length>1000)throw new Error('Give a cancellation reason of 10–1,000 characters.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const c=e92Context_(access),row=e92Find_(e91Text_(input.id),c),state=e92State_(row,c);
  if(state.cancelled&&state.cancelled.by===access.email&&state.cancelled.reason===reason)return {success:true,message:'This request was already cancelled.'};
  if(['Open','Returned'].indexOf(state.status)===-1)throw new Error('Only open or returned requests can be cancelled. Refresh the request.');
  e92Sheet_('Evidence_Request_Events',E92_EVENT_HEADERS).appendRow(['ERE-'+Utilities.getUuid(),row[0],'Cancelled','',iamSafeCell_(reason),access.email,new Date().toISOString()]);
  return {success:true,message:'Request cancelled. Its history is retained.'};
 }finally{lock.releaseLock();}
}
// Called inside submitIAMResponse's existing lock, before any file creation.
function e92PrepareSubmission_(input,record,access,duplicate){
 const id=e91Text_(input.requestId);
 if(!id){
  if(duplicate&&String(duplicate['Submission ID']).indexOf('ERS-')===0)throw new Error('This submission token belongs to an evidence request. Reopen that request.');
  return null;
 }
 if(!access.permissions.uploadEvidence||!access.permissions.submitEvidence)throw new Error('Your role cannot upload evidence to this request.');
 const c=e92Context_(access),row=e92Find_(id,c);
 if(row[1]!==record['Record ID'])throw new Error('The request belongs to another compliance record.');
 const state=e92State_(row,c);
 if(duplicate){
  const links=c.events.filter(function(e){return e[2]==='Submitted'&&e[3]===duplicate['Submission ID'];});
  if(links.length!==1||links[0][1]!==id)throw new Error('Submission token belongs to another request or its link needs administrator recovery.');
  return row;
 }
 if(['Open','Returned'].indexOf(state.status)===-1)throw new Error('This request is already submitted, completed or cancelled. Refresh it.');
 if(row[5]&&row[5]!==access.email)throw new Error('Only the assigned responder can submit for this request.');
 if(!input.files.length)throw new Error('Attach at least one evidence file to fulfil this request.');
 return row;
}
function e92SaveSubmission_(sheet,values,request){
 if(!request){sheet.appendRow(values);return;}
 const events=e92Sheet_('Evidence_Request_Events',E92_EVENT_HEADERS),eventRow=events.getLastRow()+1,responseRow=sheet.getLastRow()+1;
 try{
  events.getRange(eventRow,1,1,E92_EVENT_HEADERS.length).setValues([['ERE-'+Utilities.getUuid(),request[0],'Submitted',values[0],'',values[3],values[7]]]);
  sheet.getRange(responseRow,1,1,IAM_RESPONSE_HEADERS.length).setValues([values]);
 }catch(error){
  events.getRange(eventRow,1,1,E92_EVENT_HEADERS.length).clearContent();
  sheet.getRange(responseRow,1,1,IAM_RESPONSE_HEADERS.length).clearContent();
  throw new Error('Request response was not saved. '+(error.message||error));
 }
}
function e92ValidateReview_(item,access){
 if(String(item['Submission ID']).indexOf('ERS-')!==0)return;
 const c=e92Context_(access),links=c.events.filter(function(e){return e[2]==='Submitted'&&e[3]===item['Submission ID'];});
 if(links.length!==1)throw new Error('Request submission link is missing or ambiguous.');
 const row=e92Find_(String(links[0][1]),c),state=e92State_(row,c);
 if(state.status!=='In review'||state.latest['Submission ID']!==item['Submission ID'])throw new Error('This is no longer the active request submission. Refresh its history.');
}

/** One-time administrator configuration for email links. Paste your actual deployed dashboard URL. */
function configureDashboardEmailLinkOnce_(){
  const dashboardUrl='PASTE_YOUR_DASHBOARD_URL_HERE';
  if(!/^https:\/\//.test(dashboardUrl))throw new Error('Replace PASTE_YOUR_DASHBOARD_URL_HERE with your deployed dashboard URL before running.');
  PropertiesService.getScriptProperties().setProperty('DASHBOARD_PUBLIC_URL',dashboardUrl);
  return 'Dashboard email URL configured';
}


/** Phase 9: remove an unassigned custom role; built-in roles remain immutable identities. */
function deleteCustomDashboardRole_(name, reason) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const actor = rbacRequirePermission_('assignRoles');
    const roleName = String(name||'').trim();
    const why = String(reason||'').trim();
    if (why.length < 8 || why.length > 500) throw new Error('Enter a deletion reason (8–500 characters).');
    const role = rbacGetRoleDefinition_(roleName);
    if (!role) throw new Error('Role not found. Refresh and retry.');
    if (role.builtIn || ['Admin','Compliance Manager','Auditor','Department User'].indexOf(role.name) !== -1) throw new Error('Built-in roles cannot be deleted.');
    if (accessRequestRoleIsPrivileged_(role) && !phase87PermanentAdminEmail_(actor.email)) throw new Error('Only a permanent Admin can delete a privileged role.');
    if (rbacRows_().rows.some(function(u){return u.role === role.name;})) throw new Error('Role is assigned to users. Reassign them before deletion.');
    if (accessTempGrantRows_().rows.some(function(grant){return ['REVERTED','CANCELLED'].indexOf(grant.status)===-1 && (grant.originalRole===role.name||grant.grantedRole===role.name);})) throw new Error('Role is referenced by an unresolved temporary grant.');
    const sheet = rbacEnsureRoleDefinitionsSheet_(), values=sheet.getDataRange().getValues();
    const index=values.findIndex(function(row,i){return i>0&&String(row[0]).toLowerCase()===role.name.toLowerCase();});
    if(index<1) throw new Error('Role row not found.');
    rbacAudit_(actor.email,'ROLE::'+role.name,role,{action:'DELETE_ROLE',reason:why});
    sheet.deleteRow(index+1);
    return {success:true,message:'Custom role deleted: '+role.name};
  } finally {lock.releaseLock();}
}


/***************************************************************
 * PHASE 9.3 — CONTROL EFFECTIVENESS TESTING
 * Append-only assessments, independent review, scoped evidence.
 ***************************************************************/
const E93_TEST_HEADERS=['Test ID','Record ID','Framework','Control ID','Test Date','Conclusion','Procedure','Observation','Remediation','Evidence IDs JSON','Tester Email','Tester Name','Created At','Client Token'];
const E93_REVIEW_HEADERS=['Review ID','Test ID','Decision','Reason','Reviewer Email','Reviewed At'];
function e93Actor_(){return rbacRequirePermission_('viewControlTests');}
function e93VisibleRecords_(access){return readMasterRecords_().records.filter(function(r){return iamCanReadRecord_(access,r);});}
function e93AccessRecord_(id,access){const rows=e93VisibleRecords_(access).filter(function(r){return String(r['Record ID'])===id;});if(rows.length!==1)throw new Error('Record unavailable or outside your department scope.');return rows[0];}
function e93Control_(record){const framework=e91Text_(record['Primary Compliance Framework']),controlId=e91Text_(record['Requirement / Control ID']);if(!framework||!controlId)throw new Error('The selected record must have a primary framework and control ID.');const matches=readFrameworkControlMaster_().controls.filter(function(c){return phase3FrameworkKey_(c.framework)===phase3FrameworkKey_(framework)&&e91Text_(c.controlId).toLowerCase()===controlId.toLowerCase();});if(matches.length!==1)throw new Error('Linked framework control is missing or ambiguous in Framework_Control_Master.');return {framework:framework,controlId:controlId};}
function e93Rows_(){return e91Read_('Control_Tests',E93_TEST_HEADERS);}
function e93Reviews_(){return e91Read_('Control_Test_Reviews',E93_REVIEW_HEADERS);}
function e93Find_(id,access){const matches=e93Rows_().filter(function(r){return String(r[0])===id;});if(matches.length!==1)throw new Error('Control test not found. Refresh and retry.');const record=e93AccessRecord_(String(matches[0][1]),access);const ctrl=e93Control_(record);if(phase3FrameworkKey_(matches[0][2])!==phase3FrameworkKey_(ctrl.framework)||String(matches[0][3]).toLowerCase()!==ctrl.controlId.toLowerCase())throw new Error('Control link has changed since testing. Ask Compliance to reconcile this record.');return {row:matches[0],record:record};}
function e93ReviewFor_(id,reviews){const hits=reviews.filter(function(r){return r[1]===id;});if(hits.length>1)throw new Error('Duplicate control review detected. Ask Admin to inspect the audit sheet.');return hits[0]||null;}
function e93Public_(row,record,review,access){return {id:row[0],recordId:row[1],recordTitle:record['Record Title']||'',department:record.Department||'',framework:row[2],controlId:row[3],testDate:e92Due_(row[4]),result:row[5],tester:row[10],testerName:row[11],createdAt:accessRequestIso_(row[12]),evidenceCount:JSON.parse(String(row[9]||'[]')).length,status:review?review[2]:'Pending review',canReview:!review&&!!access.permissions.manageControlTests&&!!access.permissions.reviewResponses&&String(row[10]).toLowerCase()!==String(access.email).toLowerCase()};}
function e93Evidence_(recordId,access){const ctx=e91Context_(access,false),available=[];ctx.registry.forEach(function(row){if(row[1]!==recordId)return;const entry=e91Find_(String(row[0]),ctx),state=e91State_(entry,ctx);if(state.status==='Current'||state.status==='Expiring soon')available.push(e91Public_(entry,ctx,false));});return available;}
function getControlTests_(){const access=e93Actor_(),records=e93VisibleRecords_(access),byId={};records.forEach(function(r){byId[String(r['Record ID'])]=r;});const reviews=e93Reviews_(),items=e93Rows_().filter(function(row){return !!byId[String(row[1])];}).map(function(row){return e93Public_(row,byId[String(row[1])],e93ReviewFor_(row[0],reviews),access);});items.sort(function(a,b){return b.createdAt.localeCompare(a.createdAt);});let eligible=[];if(access.permissions.manageControlTests&&access.permissions.reviewResponses){eligible=records.filter(function(r){return !!r['Primary Compliance Framework']&&!!r['Requirement / Control ID'];}).map(function(r){return {id:String(r['Record ID']),title:r['Record Title']||'',department:r.Department||'',framework:r['Primary Compliance Framework'],controlId:r['Requirement / Control ID']};});}let evidence=[];if(eligible.length){const ids={};eligible.forEach(function(r){ids[r.id]=true;});const ctx=e91Context_(access,false);ctx.registry.forEach(function(row){if(!ids[String(row[1])])return;const entry=e91Find_(String(row[0]),ctx),state=e91State_(entry,ctx);if(['Current','Expiring soon'].indexOf(state.status)!==-1)evidence.push({id:row[0],recordId:row[1],title:state.metadata?state.metadata.title:row[4]});});}return {success:true,items:items,records:eligible,evidence:evidence,canTest:!!access.permissions.manageControlTests&&!!access.permissions.reviewResponses,today:Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd'),summary:{total:items.length,pending:items.filter(function(x){return x.status==='Pending review';}).length,approved:items.filter(function(x){return x.status==='Approved';}).length,returned:items.filter(function(x){return x.status==='Returned';}).length}};}
function createControlTest_(input){const access=rbacRequirePermission_('manageControlTests');if(!access.permissions.reviewResponses)throw new Error('Control testing requires review permission.');input=input||{};const token=e91Text_(input.clientToken);if(!/^[a-zA-Z0-9_-]{16,100}$/.test(token))throw new Error('Refresh the control test form.');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const rows=e93Rows_(),retry=rows.filter(function(r){return r[13]===token&&r[10]===access.email;});if(retry.length){if(retry.length!==1||retry[0][1]!==e91Text_(input.recordId))throw new Error('Duplicate token belongs to a different test.');return {success:true,id:retry[0][0],message:'Control test was already submitted.'};}const record=e93AccessRecord_(e91Text_(input.recordId),access),ctrl=e93Control_(record),result=e91Text_(input.result),procedure=e91Text_(input.procedure),observation=e91Text_(input.observation),remediation=e91Text_(input.remediation),testDate=e91Date_(input.testDate,'test date');const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd');if(testDate>today)throw new Error('Test date cannot be in the future.');if(['Effective','Partially effective','Ineffective','Not tested'].indexOf(result)===-1)throw new Error('Choose a valid effectiveness conclusion.');if(procedure.length<15||procedure.length>3000||observation.length<15||observation.length>3000||remediation.length>2000)throw new Error('Procedure and observations must be 15–3,000 characters; remediation at most 2,000.');const evidenceIds=input.evidenceIds;if(!Array.isArray(evidenceIds)||evidenceIds.length>30||new Set(evidenceIds).size!==evidenceIds.length||evidenceIds.some(function(x){return !/^EV-[a-f0-9-]{36}$/i.test(x);}))throw new Error('Invalid evidence selection.');const allowed={};e93Evidence_(record['Record ID'],access).forEach(function(e){allowed[e.id]=true;});if(evidenceIds.some(function(id){return !allowed[id];}))throw new Error('Evidence is not accepted, currently valid, or linked to this record. Refresh before submitting.');if(result!=='Not tested'&&!evidenceIds.length)throw new Error('A control conclusion requires at least one accepted, currently valid evidence item. Select Not tested otherwise.');const id='CT-'+Utilities.getUuid(),now=new Date().toISOString();e92Sheet_('Control_Tests',E93_TEST_HEADERS).appendRow([id,record['Record ID'],iamSafeCell_(ctrl.framework),iamSafeCell_(ctrl.controlId),testDate,result,iamSafeCell_(procedure),iamSafeCell_(observation),iamSafeCell_(remediation),JSON.stringify(evidenceIds),access.email,iamSafeCell_(access.name),now,token]);return {success:true,id:id,message:'Control test submitted. A different authorized reviewer must approve its conclusion.'};}finally{lock.releaseLock();}}
function getControlTestDetails_(id){const access=e93Actor_(),item=e93Find_(e91Text_(id),access),row=item.row,review=e93ReviewFor_(row[0],e93Reviews_()),r=e93Public_(row,item.record,review,access),ctx=e91Context_(access,false);r.procedure=String(row[6]);r.observation=String(row[7]);r.remediation=String(row[8]);r.evidence=JSON.parse(String(row[9]||'[]')).map(function(eid){try{const entry=e91Find_(eid,ctx),state=e91State_(entry,ctx);if(entry.row[1]!==r.recordId)throw new Error('Evidence link mismatch.');return {id:eid,title:state.metadata?state.metadata.title:entry.row[4],status:state.status};}catch(e){return {id:eid,title:'Evidence unavailable',status:'Unavailable or out of scope'};}});r.review=review?{by:review[4],at:accessRequestIso_(review[5]),comment:review[3]}:null;return r;}
function reviewControlTest_(input){const access=rbacRequirePermission_('manageControlTests');if(!access.permissions.reviewResponses)throw new Error('Control testing requires review permission.');input=input||{};const decision=e91Text_(input.decision),reason=e91Text_(input.reason);if(['Approved','Returned'].indexOf(decision)===-1)throw new Error('Invalid test review decision.');if(reason.length<10||reason.length>2000)throw new Error('Review reason must be 10–2,000 characters.');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const item=e93Find_(e91Text_(input.id),access),row=item.row;if(String(row[10]).toLowerCase()===String(access.email).toLowerCase())throw new Error('You cannot review your own control test.');const existing=e93ReviewFor_(row[0],e93Reviews_());if(existing){if(existing[2]===decision&&existing[3]===reason&&existing[4]===access.email)return {success:true,message:'Review already recorded.'};throw new Error('This test was already reviewed. Review decisions are immutable.');}const ids=JSON.parse(String(row[9]||'[]'));if(decision==='Approved'&&row[5]!=='Not tested'){if(!ids.length)throw new Error('The assessment has no linked evidence.');const valid={};e93Evidence_(String(row[1]),access).forEach(function(x){valid[x.id]=true;});if(ids.some(function(id){return !valid[id];}))throw new Error('Linked evidence is no longer accepted or current. Return the test for correction.');}e92Sheet_('Control_Test_Reviews',E93_REVIEW_HEADERS).appendRow(['CTR-'+Utilities.getUuid(),row[0],decision,iamSafeCell_(reason),access.email,new Date().toISOString()]);return {success:true,message:'Control test '+decision.toLowerCase()+'; independent decision recorded.'};}finally{lock.releaseLock();}}

/***************************************************************
 * PHASE 9.4 — AUTOMATED COMPLIANCE REMINDERS
 * Installable daily trigger, safe recipients, immutable mail log.
 ***************************************************************/
const E94_LOG_HEADERS=['Notification Key','Category','Item ID','Recipient','Milestone','Status','Created At','Updated At','Error'];
const E94_RUN_HEADERS=['Run ID','Mode','Started At','Finished At','Candidates','Sent','Skipped','Failed','Notes'];
const E94_HANDLER='phase94DailyReminderJob';
function e94Today_(){return Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Kolkata','yyyy-MM-dd');}
function e94Days_(day,today){return Math.round((Date.parse(day+'T00:00:00Z')-Date.parse(today+'T00:00:00Z'))/86400000);}
function e94Admin_(){return phase87RequireAdmin_();}
function e94Rows_(name,headers){return e91Read_(name,headers);}
function e94Roles_(){const roles={};rbacReadRoleDefinitions_().filter(r=>r.active).forEach(r=>roles[r.name]=r);return roles;}
function e94Users_(roles){return rbacRows_().rows.filter(u=>u.active&&roles[u.role]&&u.email&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(u.email));}
function e94Can_(u,roles,permission){return !!rbacNormalizePermissions_(roles[u.role].permissions)[permission];}
function e94CanSee_(u,roles,record){const role=roles[u.role];if(!role)return false;return iamCanReadRecord_({departmentScoped:role.scopeType==='DEPARTMENTS',departments:u.departments||[]},record);}
function e94Notice_(list,category,id,due,offset,recipient,record,owner){if(!recipient||!recipient.email)return;list.push({category:category,id:String(id),due:due,offset:offset,email:recipient.email.toLowerCase(),name:recipient.name||'Team member',department:String(record&&record.Department||''),owner:owner||'',key:[category,id,recipient.email.toLowerCase(),offset,due].join('|')});}
function e94Candidates_(){
 const today=e94Today_(),roles=e94Roles_(),users=e94Users_(roles),records=readMasterRecords_().records,byId={};records.forEach(r=>byId[String(r['Record ID'])]=r);
 const output=[];
 // Due / overdue evidence requests: reminders only for requests still waiting for the responder.
 const requests=e94Rows_('Evidence_Requests',E92_REQUEST_HEADERS),events=e94Rows_('Evidence_Request_Events',E92_EVENT_HEADERS),eventMap={};events.forEach(e=>(eventMap[e[1]]||(eventMap[e[1]]=[])).push(e));
 const responses={};iamResponseRows_().forEach(r=>responses[r['Submission ID']]=r);
 requests.forEach(row=>{const record=byId[String(row[1])];if(!record||iamDepartmentKey_(record.Department)!==iamDepartmentKey_(row[2]))return;
   const history=eventMap[row[0]]||[];if(history.some(e=>e[2]==='Cancelled'))return;
   const submissions=history.filter(e=>e[2]==='Submitted');const last=submissions.length?responses[submissions[submissions.length-1][3]]:null;
   if(last&&last.Status!=='More information required')return;
   const due=e92Due_(row[6]),offset=e94Days_(due,today);if([7,3,1,0,-1,-3,-7,-14].indexOf(offset)<0)return;
   const eligible=users.filter(u=>e94Can_(u,roles,'submitEvidence')&&e94CanSee_(u,roles,record));
   const assignee=String(row[5]||'').toLowerCase();const recipients=assignee?eligible.filter(u=>u.email===assignee):eligible;
   recipients.slice(0,20).forEach(u=>e94Notice_(output,'EVIDENCE_REQUEST',row[0],due,offset,u,record));
 });
 // Pending independent control-test reviews: only reviewers other than the tester.
 const reviews={},testRows=e94Rows_('Control_Tests',E93_TEST_HEADERS);e94Rows_('Control_Test_Reviews',E93_REVIEW_HEADERS).forEach(r=>reviews[r[1]]=true);
 testRows.forEach(row=>{if(reviews[row[0]])return;const record=byId[String(row[1])];if(!record)return;const submitted=String(row[12]||'').slice(0,10);if(!/^\d{4}-\d{2}-\d{2}$/.test(submitted))return;
   const age=-e94Days_(submitted,today);if([2,5,10,15].indexOf(age)<0)return;
   users.filter(u=>u.email!==String(row[10]).toLowerCase()&&e94Can_(u,roles,'reviewResponses')&&e94CanSee_(u,roles,record)).slice(0,12).forEach(u=>e94Notice_(output,'CONTROL_REVIEW',row[0],submitted,-age,u,record));
 });
 // Evidence expiry: only accepted evidence with approved metadata and a current owner.
 const metadata=e94Rows_('Evidence_Metadata_Changes',E91_CHANGE_HEADERS),byEvidence={};metadata.forEach(r=>(byEvidence[r[1]]||(byEvidence[r[1]]=[])).push(r));
 const accepted={};iamResponseRows_().forEach(r=>accepted[r['Submission ID']]=r);
 e94Rows_('Evidence_Register',E91_REGISTER_HEADERS).forEach(row=>{const record=byId[String(row[1])],source=accepted[row[2]];if(!record||!source||source.Status!=='Accepted'||source['Record ID']!==row[1])return;
   let version=0,current=null;(byEvidence[row[0]]||[]).forEach(change=>{const response=accepted[change[2]];if(response&&response.Status==='Accepted'&&Number(change[3])===version){try{current=JSON.parse(String(change[4]));version++;}catch(e){current=null;}}});
   if(!current||current.noExpiry||!current.validUntil||!current.owner)return;
   const due=String(current.validUntil).slice(0,10);if(!/^\d{4}-\d{2}-\d{2}$/.test(due))return;const offset=e94Days_(due,today);if([30,14,7,1,0].indexOf(offset)<0)return;
   const recipients=users.filter(u=>u.email===String(current.owner).toLowerCase()&&e94CanSee_(u,roles,record));
   recipients.forEach(u=>e94Notice_(output,'EVIDENCE_EXPIRY',row[0],due,offset,u,record));
 });
 const seen={};return output.filter(c=>{if(seen[c.key])return false;seen[c.key]=true;return true;});
}
function e94Template_(c){
 const url=e92MailUrl_();if(!url)throw new Error('Dashboard email URL is not configured.');
 const dest=c.category==='CONTROL_REVIEW'?'Control testing':c.category==='EVIDENCE_EXPIRY'?'Evidence register':'Requests';
 const link=url+(url.includes('?')?'&':'?')+'route='+(c.category==='EVIDENCE_REQUEST'?'evidence-requests':'evidence');
 const subject=c.category==='CONTROL_REVIEW'?'Control test awaiting independent review':c.category==='EVIDENCE_EXPIRY'?'Evidence validity expires soon':'Evidence request deadline reminder';
 const text=c.category==='CONTROL_REVIEW'?'An independent control assessment review is pending.':c.category==='EVIDENCE_EXPIRY'?'An approved evidence item is approaching its validity end date.':'Your assigned evidence request needs attention.';
 const when=c.category==='CONTROL_REVIEW'?'Submitted':c.category==='EVIDENCE_EXPIRY'?'Valid until':'Due date';
 const html='<!doctype html><html><body style="margin:0;background:#edf3fb;padding:26px 10px;font-family:Arial,Helvetica,sans-serif;color:#162849"><table role="presentation" width="100%" style="max-width:560px;background:white;border-radius:18px;margin:auto;border-collapse:collapse;overflow:hidden"><tr><td style="background:#142442;padding:25px;color:white;border-top:5px solid #3b75ef"><img src="cid:dgsLogo" alt="DigiClarity" width="230" style="max-width:100%"></td></tr><tr><td style="padding:30px"><p style="font-size:11px;color:#3769d8;font-weight:700;letter-spacing:1px">COMPLIANCE REMINDER</p><h2 style="font-size:22px;margin:10px 0">'+e92MailEsc_(subject)+'</h2><p style="font-size:14px;line-height:1.7;color:#52637e">Hi '+e92MailEsc_(c.name.split(/\s+/)[0])+', '+e92MailEsc_(text)+'</p><p style="background:#f3f7fe;padding:16px;border-radius:10px;font-size:13px;line-height:1.9">Department: <b>'+e92MailEsc_(c.department)+'</b><br>'+when+': <b>'+e92MailEsc_(e92MailFormatDue_(c.due))+'</b></p><p style="text-align:center;padding:18px"><a href="'+e92MailEsc_(link)+'" style="background:#336cf0;padding:14px 23px;color:white;text-decoration:none;border-radius:10px;font-weight:700">Sign In &amp; View Dashboard</a></p><p style="color:#74829a;font-size:12px">Open Evidence Management → '+e92MailEsc_(dest)+'. This email does not contain patient information or detailed finding content.</p></td></tr></table></body></html>';
 return {subject:subject,html:html,body:'Hi '+c.name.split(/\s+/)[0]+',\n\n'+text+'\nDepartment: '+c.department+'\n'+when+': '+c.due+'\n\nSign in: '+link+'\nOpen Evidence Management > '+dest+'.'};
}
const E94_ARCHIVE_HEADERS=['Notification Key','Archived At','Archived By','Reason','Previous Status'];
function e94Status_(){
 const handler=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===E94_HANDLER);
 const archived={};e94Rows_('Compliance_Reminder_Archives',E94_ARCHIVE_HEADERS).forEach(r=>{archived[String(r[0])]=true;});
 const log=e94Rows_('Compliance_Reminder_Log',E94_LOG_HEADERS).filter(r=>!archived[String(r[0])]).slice(-12).reverse().map(r=>({key:String(r[0]),category:String(r[1]),recipient:String(r[3]),milestone:String(r[4]),status:String(r[5]),at:String(r[7]||'').slice(0,19)}));
 return {enabled:handler.length>0,triggerCount:handler.length,log:log};
}
/** Removes a reminder from the visible history only; never erases a delivery record or dedupe key. */
function archiveComplianceReminderLog_(input){
 const actor=e94Admin_(),payload=input||{},key=String(payload.key||'').trim(),reason=String(payload.reason||'').trim();
 if(!key||key.length>500)throw new Error('Invalid reminder identifier.');
 if(reason.length<10||reason.length>1000)throw new Error('Enter a removal reason of 10–1,000 characters.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const row=e94Rows_('Compliance_Reminder_Log',E94_LOG_HEADERS).find(r=>String(r[0])===key);
  if(!row)throw new Error('The reminder delivery record was not found.');
  const archived=e94Rows_('Compliance_Reminder_Archives',E94_ARCHIVE_HEADERS).some(r=>String(r[0])===key);
  if(archived)return {success:true,message:'This reminder is already removed from the visible history.'};
  e92Sheet_('Compliance_Reminder_Archives',E94_ARCHIVE_HEADERS).appendRow([iamSafeCell_(key),new Date().toISOString(),iamSafeCell_(actor.email),iamSafeCell_(reason),iamSafeCell_(String(row[5]))]);
  return {success:true,message:'Reminder removed from visible history. Delivery and audit records are retained.'};
 }finally{lock.releaseLock();}
}

function getComplianceReminderSettings_(){e94Admin_();const state=e94Status_();state.success=true;state.today=e94Today_();return state;}
function previewComplianceReminders_(){e94Admin_();const items=e94Candidates_(),sent={};e94Rows_('Compliance_Reminder_Log',E94_LOG_HEADERS).forEach(r=>{if(['SENT','CLAIMED'].indexOf(r[5])>=0)sent[r[0]]=true;});return {success:true,total:items.length,ready:items.filter(c=>!sent[c.key]).length,byCategory:['EVIDENCE_REQUEST','CONTROL_REVIEW','EVIDENCE_EXPIRY'].map(category=>({category:category,count:items.filter(c=>c.category===category).length})),sample:items.slice(0,12).map(c=>({category:c.category,recipient:c.email,due:c.due,alreadySent:!!sent[c.key]}))};}
function setComplianceReminderSchedule_(enabled){const actor=e94Admin_();const current=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===E94_HANDLER);current.forEach(t=>ScriptApp.deleteTrigger(t));if(enabled){if(!e92MailUrl_())throw new Error('Configure the dashboard email URL before enabling reminders.');const trigger=ScriptApp.newTrigger(E94_HANDLER).timeBased().everyDays(1).atHour(9).create();PropertiesService.getScriptProperties().setProperty('E94_TRIGGER_UID',trigger.getUniqueId());}else{PropertiesService.getScriptProperties().deleteProperty('E94_TRIGGER_UID');}e92Sheet_('Compliance_Reminder_Runs',E94_RUN_HEADERS).appendRow(['SET-'+Utilities.getUuid(),enabled?'ENABLED':'DISABLED',new Date().toISOString(),new Date().toISOString(),0,0,0,0,iamSafeCell_(actor.email)]);return {success:true,enabled:!!enabled,message:enabled?'Daily reminder trigger enabled (approximately 9 AM project timezone).':'Daily reminder trigger disabled.'};}
function phase94DailyReminderJob(event){const expected=PropertiesService.getScriptProperties().getProperty('E94_TRIGGER_UID');if(!expected||!event||String(event.triggerUid||'')!==expected)throw new Error('Unauthorized reminder trigger invocation.');return e94Run_(false);}
function runComplianceRemindersNow_(){e94Admin_();return e94Run_(true);}
function e94Run_(manual){const lock=LockService.getScriptLock();if(!lock.tryLock(2000))throw new Error('Reminder job is already running.');const runId='RUN-'+Utilities.getUuid(),start=new Date().toISOString();let sent=0,failed=0,skipped=0,candidates=[];try{
  if(!e92MailUrl_())throw new Error('Dashboard email URL is not configured.');candidates=e94Candidates_();if(candidates.length>150)throw new Error('More than 150 reminder candidates. Review the scope before mailing.');
  const log=e92Sheet_('Compliance_Reminder_Log',E94_LOG_HEADERS),rows=e94Rows_('Compliance_Reminder_Log',E94_LOG_HEADERS),existing={};rows.forEach((r,i)=>{existing[r[0]]={status:r[5],row:i+2};});
  candidates.forEach(c=>{const old=existing[c.key];if(old&&old.status!=='FAILED'){skipped++;return;}
   let sheetRow;if(old){sheetRow=old.row;log.getRange(sheetRow,6,1,3).setValues([['CLAIMED',new Date().toISOString(),'']]);}else{log.appendRow([iamSafeCell_(c.key),c.category,iamSafeCell_(c.id),c.email,String(c.offset)+'d / '+c.due,'CLAIMED',new Date().toISOString(),new Date().toISOString(),'']);sheetRow=log.getLastRow();existing[c.key]={status:'CLAIMED',row:sheetRow};}
   SpreadsheetApp.flush();try{const msg=e94Template_(c);MailApp.sendEmail({to:c.email,subject:msg.subject,body:msg.body,htmlBody:msg.html,inlineImages:e92MailInlineImages_(),name:'DigiClarity Compliance'});log.getRange(sheetRow,6,1,3).setValues([['SENT',new Date().toISOString(),'']]);sent++;}catch(e){log.getRange(sheetRow,6,1,3).setValues([['FAILED',new Date().toISOString(),iamSafeCell_(String(e.message||e).slice(0,200))]]);failed++;}
  });
  e92Sheet_('Compliance_Reminder_Runs',E94_RUN_HEADERS).appendRow([runId,manual?'MANUAL':'SCHEDULED',start,new Date().toISOString(),candidates.length,sent,skipped,failed,'']);return {success:true,candidates:candidates.length,sent:sent,skipped:skipped,failed:failed};
 }catch(err){try{e92Sheet_('Compliance_Reminder_Runs',E94_RUN_HEADERS).appendRow([runId,manual?'MANUAL':'SCHEDULED',start,new Date().toISOString(),candidates.length,sent,skipped,failed,iamSafeCell_(String(err.message||err).slice(0,300))]);}catch(ignore){}throw err;}finally{lock.releaseLock();}}


/***************************************************************
 * PHASE 9.5 — CONTINUOUS COMPLIANCE MONITORING
 * Deterministic exception keys; append-only operator activity.
 * Rule snapshots are recalculated from source; no records mutated.
 ***************************************************************/
const E95_EVENT_HEADERS=['Event ID','Exception Key','Action','Actor Email','Actor Name','Reason','Owner Email','Target Date','At'];
const E95_RUN_HEADERS=['Run ID','Run At','Actor','Detected','Open','Acknowledged','Resolved','Notes'];
const E95_RULES={OVERDUE_REQUEST:'Overdue evidence request',EVIDENCE_EXPIRED:'Expired evidence',EVIDENCE_GAP:'Evidence metadata gap',INEFFECTIVE_CONTROL:'Ineffective control',PARTIAL_CONTROL:'Partially effective control',PENDING_TEST:'Control test awaiting review'};
function e95Actor_(){return rbacRequirePermission_('viewMonitoring');}
function e95Manage_(actor){if(!actor.permissions.manageMonitoring||!actor.permissions.reviewResponses)throw new Error('Control monitoring updates require the reviewResponses permission.');}
function e95Key_(rule,source){return rule+'|'+String(source);}
function e95Source_(){
 const records=readMasterRecords_().records,byId={};records.forEach(r=>byId[String(r['Record ID'])]=r);
 const today=e94Today_(),found=[],put=(rule,id,record,detail,severity)=>{
   if(!record)return;found.push({key:e95Key_(rule,id),rule:rule,label:E95_RULES[rule],sourceId:String(id),recordId:String(record['Record ID']),recordTitle:String(record['Record Title']||''),department:String(record.Department||''),framework:String(record['Primary Compliance Framework']||''),controlId:String(record['Requirement / Control ID']||''),severity:severity,detail:detail});
 };
 const req=e94Rows_('Evidence_Requests',E92_REQUEST_HEADERS),events=e94Rows_('Evidence_Request_Events',E92_EVENT_HEADERS),evBy={};events.forEach(e=>(evBy[e[1]]||(evBy[e[1]]=[])).push(e));
 const responseBy={};iamResponseRows_().forEach(r=>responseBy[String(r['Submission ID'])]=r);
 req.forEach(row=>{
  const record=byId[String(row[1])];if(!record||String(row[6])>=today)return;
  const ev=evBy[String(row[0])]||[];if(ev.some(e=>e[2]==='Cancelled'))return;
  const subs=ev.filter(e=>e[2]==='Submitted'),last=subs.length?responseBy[String(subs[subs.length-1][3])]:null;
  if(last&&['Accepted','Submitted'].indexOf(last.Status)>=0)return;
  put('OVERDUE_REQUEST',row[0],record,'Evidence request is past its due date '+String(row[6])+'.','High');
 });
 const allowed={};e94Rows_('Control_Test_Reviews',E93_REVIEW_HEADERS).forEach(r=>allowed[String(r[1])]=r);
 e94Rows_('Control_Tests',E93_TEST_HEADERS).forEach(row=>{
  const record=byId[String(row[1])],review=allowed[String(row[0])];if(!record)return;
  if(!review){const submitted=String(row[12]||'').slice(0,10);if(/^\d{4}-\d{2}-\d{2}$/.test(submitted)&&e94Days_(submitted,today)<=-5)put('PENDING_TEST',row[0],record,'Independent control review has been pending at least 5 days.','Medium');return;}
  if(review[2]!=='Approved')return;
  if(row[5]==='Ineffective')put('INEFFECTIVE_CONTROL',row[0],record,'Independently approved control test concluded Ineffective.','Critical');
  if(row[5]==='Partially effective')put('PARTIAL_CONTROL',row[0],record,'Independently approved control test concluded Partially effective.','High');
 });
 const metadata=e94Rows_('Evidence_Metadata_Changes',E91_CHANGE_HEADERS),byEvidence={};metadata.forEach(r=>(byEvidence[String(r[1])]||(byEvidence[String(r[1])]=[])).push(r));
 e94Rows_('Evidence_Register',E91_REGISTER_HEADERS).forEach(row=>{
  const record=byId[String(row[1])],source=responseBy[String(row[2])];if(!record||!source||source.Status!=='Accepted'||source['Record ID']!==row[1])return;
  let version=0,current=null;(byEvidence[String(row[0])]||[]).forEach(change=>{const review=responseBy[String(change[2])];if(review&&review.Status==='Accepted'&&Number(change[3])===version){try{current=JSON.parse(String(change[4]));version++;}catch(e){current=null;}}});
  if(!current||!current.owner||(!current.noExpiry&&!current.validUntil))put('EVIDENCE_GAP',row[0],record,'Accepted evidence is missing approved owner or validity metadata.','Medium');
  else if(!current.noExpiry&&String(current.validUntil).slice(0,10)<today)put('EVIDENCE_EXPIRED',row[0],record,'Approved evidence validity ended on '+String(current.validUntil).slice(0,10)+'.','High');
 });
 return found;
}
function e95Events_(){return e91Read_('Compliance_Monitor_Events',E95_EVENT_HEADERS);}
function e95Materialize_(access){
 const events=e95Events_(),latest={};events.forEach(row=>latest[String(row[1])]={action:String(row[2]),actor:String(row[3]),actorName:String(row[4]),reason:String(row[5]),owner:String(row[6]),target:String(row[7]),at:String(row[8])});
 const visible=e95Source_().filter(f=>iamCanReadRecord_(access,{Department:f.department}));
 const active={};visible.forEach(x=>active[x.key]=true);
 const results=visible.map(f=>Object.assign({},f,{state:latest[f.key]&&latest[f.key].action==='ACKNOWLEDGE'?'Acknowledged':'Open',owner:latest[f.key]?latest[f.key].owner:'',targetDate:latest[f.key]?latest[f.key].target:'',lastAction:latest[f.key]||null}));
 // A previously resolved exception that recurs is reopened by the detection engine, not silently hidden.
 const summary={total:results.length,critical:results.filter(x=>x.severity==='Critical').length,high:results.filter(x=>x.severity==='High').length,acknowledged:results.filter(x=>x.state==='Acknowledged').length};
 return {success:true,items:results,summary:summary,canManage:!!access.permissions.manageMonitoring&&!!access.permissions.reviewResponses,today:e94Today_(),lastScan:e95LatestRun_()};
}
function e95LatestRun_(){const rows=e91Read_('Compliance_Monitor_Runs',E95_RUN_HEADERS);return rows.length?String(rows[rows.length-1][1]):'';}
function getComplianceMonitoring_(){return e95Materialize_(e95Actor_());}
function e95WriteRun_(actor){const items=e95Source_(),now=new Date().toISOString();e92Sheet_('Compliance_Monitor_Runs',E95_RUN_HEADERS).appendRow(['CMR-'+Utilities.getUuid(),now,actor,items.length,items.length,0,0,'Read-only detection scan. Exceptions remain sourced from authoritative records.']);return {success:true,detected:items.length,scannedAt:now};}
function scanComplianceMonitoring_(){const actor=e95Actor_();e95Manage_(actor);return e95WriteRun_(actor.email);}
function updateComplianceMonitorException_(input){
 const actor=e95Actor_();e95Manage_(actor);input=input||{};
 const key=String(input.key||'').trim(),action=String(input.action||''),reason=String(input.reason||'').trim(),owner=String(input.owner||'').trim().toLowerCase(),target=String(input.targetDate||'').trim();
 if(!['ACKNOWLEDGE','REOPEN'].includes(action))throw new Error('Invalid monitoring action.');
 if(reason.length<12||reason.length>2000)throw new Error('Reason must be 12–2,000 characters.');
 if(owner&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(owner))throw new Error('Owner must be a valid email address.');
 if(target&&(!/^\d{4}-\d{2}-\d{2}$/.test(target)||target<e94Today_()))throw new Error('Choose a future target date.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const item=e95Source_().find(f=>f.key===key);if(!item)throw new Error('This exception is no longer active. Refresh monitoring.');
  if(!iamCanReadRecord_(actor,{Department:item.department}))throw new Error('Exception is outside your department access.');
  const seen=e95Events_().filter(r=>String(r[1])===key),last=seen[seen.length-1];
  const state=last&&String(last[2])==='ACKNOWLEDGE'?'Acknowledged':'Open';
  if((action==='ACKNOWLEDGE'&&state==='Acknowledged')||(action==='REOPEN'&&state==='Open'))throw new Error('Exception state has changed; refresh before acting.');
  e92Sheet_('Compliance_Monitor_Events',E95_EVENT_HEADERS).appendRow(['CME-'+Utilities.getUuid(),iamSafeCell_(key),action,actor.email,iamSafeCell_(actor.name),iamSafeCell_(reason),iamSafeCell_(owner),target,new Date().toISOString()]);
  return {success:true,message:action==='ACKNOWLEDGE'?'Exception acknowledged. The underlying issue remains open until its source is corrected.':'Acknowledgement cleared; exception reopened.'};
 }finally{lock.releaseLock();}
}

// Optional daily snapshot trigger. Deliberately does not send emails or update source records.
const E95_TRIGGER_HANDLER='phase95DailyMonitoringJob';
function e95ScheduleState_(){const ts=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===E95_TRIGGER_HANDLER);return {enabled:ts.length>0,triggerCount:ts.length};}
function getComplianceMonitoringSchedule_(){phase87RequireAdmin_();return e95ScheduleState_();}
function setComplianceMonitoringSchedule_(enabled){const admin=phase87RequireAdmin_(),lock=LockService.getScriptLock();lock.waitLock(20000);try{
 const props=PropertiesService.getScriptProperties(),existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===E95_TRIGGER_HANDLER);
 if(enabled===true){if(existing.length>1)throw new Error('Multiple monitoring triggers detected; inspect Apps Script triggers.');if(!existing.length){const trigger=ScriptApp.newTrigger(E95_TRIGGER_HANDLER).timeBased().atHour(9).everyDays(1).create();props.setProperty('E95_TRIGGER_UID',trigger.getUniqueId());}}
 else if(enabled===false){existing.forEach(t=>ScriptApp.deleteTrigger(t));props.deleteProperty('E95_TRIGGER_UID');}
 else throw new Error('Specify whether monitoring should be enabled.');
 return Object.assign({success:true,message:enabled?'Daily monitoring scan enabled.':'Daily monitoring scan disabled.'},e95ScheduleState_());
 }finally{lock.releaseLock();}}
function phase95DailyMonitoringJob(event){const uid=PropertiesService.getScriptProperties().getProperty('E95_TRIGGER_UID');if(!uid||!event||String(event.triggerUid||'')!==uid)throw new Error('Unauthorized monitoring trigger.');return e95WriteRun_('SCHEDULED_TRIGGER');}


/***************************************************************
 * PHASE 9.6 — READ-ONLY AUDIT READINESS & SCOPED CSV EXPORT
 * An operational readiness indicator, NOT certification.
 ***************************************************************/
const E96_EXPORT_HEADERS=['Export ID','At','Actor Email','Framework Filter','Readiness Filter','Record Count','Format'];
function e96Actor_(){return rbacRequirePermission_('viewAuditReadiness');}
function e96Snapshot_(access){
 // Phase 9.6.1: join metadata once rather than repeatedly scanning all
 // evidence changes, register entries and user rows for every file.
 const ctx=e91Context_(access,false),visible=ctx.records,byId={};
 visible.forEach(r=>{byId[String(r['Record ID'])]=r;});
 const valid={},pending={},effective={},byEvidence={},ownersByRecord={};
 ctx.changes.forEach(row=>{const id=String(row[1]);(byEvidence[id]||(byEvidence[id]=[])).push(row);});
 ctx.registry.forEach(row=>{
   const rid=String(row[1]),record=byId[rid],response=ctx.byResponse[String(row[2])];
   if(!record||!response||response.Status!=='Accepted'||String(response['Record ID'])!==rid)return;
   if(!iamVisibleResponse_(response,record,access))return;
   const published=iamPublicResponse_(response);
   if(!published||!Array.isArray(published.files)||!published.files.some(f=>String(f.id)===String(row[3])))return;
   let version=0,metadata=null;
   (byEvidence[String(row[0])]||[]).forEach(change=>{
     const review=ctx.byResponse[String(change[2])];
     if(!review||String(review['Record ID'])!==rid||review.Status!=='Accepted'||Number(change[3])!==version)return;
     try{metadata=JSON.parse(String(change[4]));version++;}catch(ignore){metadata=null;}
   });
   if(!metadata)return;
   if(!ownersByRecord[rid])ownersByRecord[rid]=e91EligibleOwners_(record,ctx).map(u=>u.email);
   const dateFrom=String(metadata.validFrom||''),dateTo=String(metadata.validUntil||'');
   const owner=String(metadata.owner||'').toLowerCase();
   if(!ownersByRecord[rid].includes(owner)||!dateFrom||dateFrom>ctx.today||(!metadata.noExpiry&&(!dateTo||dateTo<ctx.today)))return;
   valid[rid]=(valid[rid]||0)+1;
 });
 const reviews={};e94Rows_('Control_Test_Reviews',E93_REVIEW_HEADERS).forEach(r=>{reviews[String(r[1])]=r;});
 e94Rows_('Control_Tests',E93_TEST_HEADERS).forEach(row=>{
   const rid=String(row[1]);if(!byId[rid])return;
   const review=reviews[String(row[0])];
   if(!review)pending[rid]=(pending[rid]||0)+1;
   else if(review[2]==='Approved'&&row[5]==='Effective')effective[rid]=(effective[rid]||0)+1;
 });
 const exceptions={};e95Source_().forEach(x=>{if(byId[x.recordId])exceptions[x.recordId]=(exceptions[x.recordId]||0)+1;});
 const items=visible.map(r=>{
   const id=String(r['Record ID']),ve=valid[id]||0,et=effective[id]||0,pt=pending[id]||0,ex=exceptions[id]||0;
   const framework=String(r['Primary Compliance Framework']||''),controlId=String(r['Requirement / Control ID']||'');
   const gaps=[];if(!framework||!controlId)gaps.push('Framework/control mapping needed');
   if(!ve)gaps.push('No currently valid accepted evidence');
   if(!et)gaps.push('No approved effective control test');
   if(pt)gaps.push(pt+' test(s) awaiting independent review');
   if(ex)gaps.push(ex+' active monitoring exception(s)');
   return {id:id,title:String(r['Record Title']||''),department:String(r.Department||''),framework:framework,controlId:controlId,validEvidence:ve,effectiveTests:et,pendingTests:pt,exceptions:ex,readiness:gaps.length?'Action required':'Ready for review',gaps:gaps};
 });
 items.sort((a,b)=>a.framework.localeCompare(b.framework)||a.controlId.localeCompare(b.controlId)||a.id.localeCompare(b.id));
 return {success:true,generatedAt:new Date().toISOString(),items:items};
}
function getAuditReadinessSnapshot_(){return e96Snapshot_(e96Actor_());}
function e96CsvCell_(value){let s=String(value==null?'':value);if(/^[\s]*[=+@\-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
function exportAuditReadinessCsv_(input){const actor=e96Actor_();if(!actor.permissions.exportAuditReadiness)throw new Error('Your role cannot export audit readiness data.');input=input||{};const framework=String(input.framework||'').trim(),status=String(input.status||'').trim(),search=String(input.search||'').trim().toLowerCase();if(search.length>200)throw new Error('Audit search filter is too long.');if(framework.length>150||!['','Ready for review','Action required'].includes(status))throw new Error('Invalid audit export filter.');const snapshot=e96Snapshot_(actor);if(framework&&!snapshot.items.some(x=>x.framework===framework))throw new Error('Framework outside authorized audit scope.');const items=snapshot.items.filter(x=>(!framework||x.framework===framework)&&(!status||x.readiness===status)&&(!search||[x.id,x.title,x.department,x.framework,x.controlId].join(' ').toLowerCase().includes(search)));const head=['Record ID','Record title','Department','Framework','Control ID','Readiness','Valid evidence','Approved effective tests','Pending tests','Active exceptions','Gaps'];const csv=[head,...items.map(x=>[x.id,x.title,x.department,x.framework,x.controlId,x.readiness,x.validEvidence,x.effectiveTests,x.pendingTests,x.exceptions,x.gaps.join('; ')])].map(row=>row.map(e96CsvCell_).join(',')).join('\r\n');const bytes=Utilities.newBlob('\ufeff'+csv,'text/csv').getBytes();const fileName='DigiClarity_Audit_Readiness_'+e94Today_()+'.csv';e92Sheet_('Audit_Readiness_Exports',E96_EXPORT_HEADERS).appendRow(['ARE-'+Utilities.getUuid(),new Date().toISOString(),actor.email,iamSafeCell_(framework||'ALL'),status||'ALL',items.length,'CSV']);return {success:true,fileName:fileName,mimeType:'text/csv',base64:Utilities.base64Encode(bytes)};}


/** V57: Staged login. Verify only the currently signed-in Google Workspace identity. */
function verifyDashboardLoginEmail(input) {
  const typed = dashboardNormalizeLoginEmail_((input || {}).email);
  if (!typed) return {ok:false,code:'EMAIL_REQUIRED',message:'Enter your work email.'};
  try {
    // Requires an approved company domain, matching Google identity, and active RBAC membership.
    dashboardValidateLoginIdentity_(typed);
    const context = getLoginGatewayContext();
    if (!context || !context.ok || dashboardNormalizeLoginEmail_(context.email) !== typed) {
      return {ok:false,code:'ACCESS_DENIED',message:'Unable to verify dashboard access for this account.'};
    }
    return {
      ok:true, email:typed, name:context.name || '', role:context.role || '',
      requiresPasswordSetup:!!context.requiresPasswordSetup,
      requiresDashboardPassword:!!context.requiresDashboardPassword
    };
  } catch (error) {
    return {ok:false,code:'LOGIN_IDENTITY_ERROR',message:error && error.message ? error.message : 'Unable to verify your account.'};
  }
}

/** V57: Corporate branding only. No account information is returned. */
function getLoginBrandLogo() {
  try {
    const file=DriveApp.getFileById(LOGO_FILE_ID);
    const blob=file.getBlob();
    const mime=String(blob.getContentType()||'image/png');
    if (!/^image\/(?:png|jpeg|webp|svg\+xml)$/i.test(mime)) return '';
    const bytes=blob.getBytes();
    if (bytes.length > 5000000) return '';
    return 'data:'+mime+';base64,'+Utilities.base64Encode(bytes);
  } catch (error) {
    console.error('Login branding could not be loaded: '+String(error));
    return '';
  }
}



/***************************************************************
 * TC-06 SECURITY PATCH — SERVER-ENFORCED SESSION BOUNDARY
 * Code.gs V58 candidate / Index.html V103 candidate
 *
 * All user-facing protected backend entry points are private
 * implementations, accessible only via dashboardSecureInvoke.
 * Each RPC request must present its session token, checked BEFORE
 * any protected read or write. Authentication and trigger handlers
 * retain their existing independent authorization checks.
 *
 * IMPORTANT: Session validation does NOT refresh Last Seen At.
 * Only the existing human-activity heartbeat can do that.
 ***************************************************************/
function dashboardTc06AssertSession_(token) {
  if (typeof token !== 'string' || token.length > 2048 || !token.trim()) {
    throw new Error('SESSION_REQUIRED: Sign in to access the Compliance Dashboard.');
  }
  const parsed = dashboardParseSessionToken_(token);
  if (!parsed || !parsed.sessionId || !parsed.secret) {
    throw new Error('SESSION_INVALID: Invalid dashboard session.');
  }
  // Bind token to the actual Workspace identity, not a claimed email from the browser.
  const email = rbacIdentity_();
  const data = dashboardSessionRows_();
  const rows = data.rows.filter(function(row){ return row.sessionId === parsed.sessionId; });
  if (rows.length !== 1) throw new Error('SESSION_INVALID: Unknown dashboard session.');
  const row = rows[0];
  if (String(row.email||'').toLowerCase() !== String(email||'').toLowerCase()) {
    throw new Error('SESSION_INVALID: Session identity does not match.');
  }
  // Validate the unforgeable token secret before any use of the session record.
  const expected = dashboardSessionTokenHash_(row.sessionId, parsed.secret, row.email);
  if (!dashboardPasswordEqual_(expected, row.tokenHash)) {
    throw new Error('SESSION_INVALID: Token signature failed verification.');
  }
  if (row.revokedAt) {
    throw new Error('SESSION_REVOKED: This dashboard session has been revoked. Sign in again.');
  }
  const lastSeen = new Date(row.lastSeenAt || row.createdAt).getTime();
  if (!Number.isFinite(lastSeen) || lastSeen > Date.now() + 60000 || Date.now() - lastSeen >= 300000) {
    // Refuse execution even if the browser bypasses its five-minute UI lock.
    throw new Error('SESSION_EXPIRED: Five-minute dashboard session expired. Sign in again.');
  }
  const credential = dashboardCredentialForEmail_(email);
  if (!credential || String(credential.updatedAt||'') !== String(row.credentialUpdatedAt||'')) {
    throw new Error('SESSION_REVOKED: Dashboard credentials changed. Sign in again.');
  }
  if (dashboardCredentialExpired_(credential)) {
    throw new Error('SESSION_EXPIRED: Dashboard password expired. Sign in again.');
  }
  // RBAC and department scope are still checked by the original method.
  // This extra check immediately rejects inactive/deleted dashboard accounts.
  return rbacAccessForEmail_(email);
}

function dashboardSecureInvoke(token, method, params) {
  if (typeof method !== 'string' || !Array.isArray(params) || params.length > 30) {
    throw new Error('Invalid dashboard API request.');
  }
  // Authenticate FIRST, even when the caller guesses or tampers with a method name.
  dashboardTc06AssertSession_(token);
  switch (method) {
    case "migrateMasterForSubDepartment": return migrateMasterForSubDepartment_(...params);
    // TC-24: all requests require a live dashboard session and server-side Admin/Manager RBAC.
    case "tc24GetCenter": return tc24GetCenter_(...params);
    case "tc24Initialize": return tc24Initialize_(...params);
    case "tc24SavePolicy": return tc24SavePolicy_(...params);
    case "tc24SetLegalHold": return tc24SetLegalHold_(...params);
    case "tc24RequestDisposalReview": return tc24RequestDisposalReview_(...params);
    case "tc24ReviewDisposalRequest": return tc24ReviewDisposalRequest_(...params);
    case "getRecords": return getRecords_(...params);
    case "getLookupValues": return getLookupValues_(...params);
    case "getDependentLookupValues": return getDependentLookupValues_(...params);
    case "getFrameworkControls": return getFrameworkControls_(...params);
    case "getCommonControlAreas": return getCommonControlAreas_(...params);
    case "getControlsByCommonControlArea": return getControlsByCommonControlArea_(...params);
    case "getFrameworkControlById": return getFrameworkControlById_(...params);
    case "getFilterOptions": return getFilterOptions_(...params);
    case "getDashboardData": return getDashboardData_(...params);
    case "getRecordById": return getRecordById_(...params);
    case "deleteComplianceRecord": return deleteComplianceRecord_(...params);
    case "getDeletedRecordsForRecovery": return getDeletedRecordsForRecovery_(...params);
    case "restoreComplianceRecord": return restoreComplianceRecord_(...params);
    case "getRecordHistory": return getRecordHistory_(...params);
    case "saveComplianceDraft": return saveComplianceDraft_(...params);
    case "getComplianceDrafts": return getComplianceDrafts_(...params);
    case "getComplianceDraftById": return getComplianceDraftById_(...params);
    case "deleteComplianceDraft": return deleteComplianceDraft_(...params);
    case "markComplianceDraftSubmitted": return markComplianceDraftSubmitted_(...params);
    case "getLogoDataUrl": return getLogoDataUrl_(...params);
    case "getConfig": return getConfig_(...params);
    case "getRecordTypeConfig": return getRecordTypeConfig_(...params);
    case "getNextRecordIdPreview": return getNextRecordIdPreview_(...params);
    case "createComplianceRecord": return createComplianceRecord_(...params);
    case "testCreateRecordAPI": return testCreateRecordAPI_(...params);
    case "testCreateComplianceRecord": return testCreateComplianceRecord_(...params);
    case "testMasterSheet": return testMasterSheet_(...params);
    case "testLookupSheet": return testLookupSheet_(...params);
    case "testDashboardData": return testDashboardData_(...params);
    case "testFilterOptions": return testFilterOptions_(...params);
    case "testLogo": return testLogo_(...params);
    case "testRiskCalculation": return testRiskCalculation_(...params);
    case "testFrameworkControlMaster": return testFrameworkControlMaster_(...params);
    case "testDynamicLookupIntegration": return testDynamicLookupIntegration_(...params);
    case "runFinalSystemTest": return runFinalSystemTest_(...params);
    case "testCrossFrameworkControlMapping": return testCrossFrameworkControlMapping_(...params);
    case "setupMakerCheckerSheets": return setupMakerCheckerSheets_(...params);
    case "submitRecordChangeRequest": return submitRecordChangeRequest_(...params);
    case "getChangeRequestsForCurrentUser": return getChangeRequestsForCurrentUser_(...params);
    case "getChangeRequestDetails": return getChangeRequestDetails_(...params);
    case "approveRecordChangeRequest": return approveRecordChangeRequest_(...params);
    case "rejectRecordChangeRequest": return rejectRecordChangeRequest_(...params);
    case "returnRecordChangeRequest": return returnRecordChangeRequest_(...params);
    case "runMakerCheckerSystemTest": return runMakerCheckerSystemTest_(...params);
    case "getAccessContext": return getAccessContext_(...params);
    case "changeDashboardPassword": return changeDashboardPassword_(...params);
    case "adminResetDashboardPassword": return adminResetDashboardPassword_(...params);
    case "listDashboardUsers": return listDashboardUsers_(...params);
    case "listRoleDefinitions": return listRoleDefinitions_(...params);
    case "saveDashboardUser": return saveDashboardUser_(...params);
    case "saveRoleDefinition": return saveRoleDefinition_(...params);
    case "initializeCustomRoleRbac": return initializeCustomRoleRbac_(...params);
    case "initializeRbacOnce": return initializeRbacOnce_(...params);
    case "reassignPendingChangeRequest": return reassignPendingChangeRequest_(...params);
    case "getReturnedRequestForRevision": return getReturnedRequestForRevision_(...params);
    case "deleteDashboardUser": return deleteDashboardUser_(...params);
    case "deleteDashboardUserV9": return deleteDashboardUserV9_(...params);
    case "getCrossFrameworkExplorerBootstrap": return getCrossFrameworkExplorerBootstrap_(...params);
    case "getCrossFrameworkMappings": return getCrossFrameworkMappings_(...params);
    case "testCrossFrameworkExplorerV16": return testCrossFrameworkExplorerV16_(...params);
    case "getFrameworkControlsAliasSafe": return getFrameworkControlsAliasSafe_(...params);
    case "getCrossFrameworkExplorerBootstrapV20": return getCrossFrameworkExplorerBootstrapV20_(...params);
    case "getCrossFrameworkMappingsV20": return getCrossFrameworkMappingsV20_(...params);
    case "testCrossFrameworkExplorerV20": return testCrossFrameworkExplorerV20_(...params);
    case "getPreviousRecordOptions": return getPreviousRecordOptions_(...params);
    case "getAnalyticsRecords": return getAnalyticsRecords_(...params);
    case "getIAMDepartmentOptions": return getIAMDepartmentOptions_(...params);
    case "initializeIamPhase": return initializeIamPhase_(...params);
    case "getIAMWorkspace": return getIAMWorkspace_(...params);
    case "submitIAMResponse": return submitIAMResponse_(...params);
    case "reviewIAMResponse": return reviewIAMResponse_(...params);
    case "getIAMEvidence": return getIAMEvidence_(...params);
    case "getDashboardSecurityCenter": return getDashboardSecurityCenter_(...params);
    case "adminRevokeDashboardSession": return adminRevokeDashboardSession_(...params);
    case "getMyDashboardSessions": return getMyDashboardSessions_(...params);
    case "revokeMyDashboardSession": return revokeMyDashboardSession_(...params);
    case "revokeOtherDashboardSessions": return revokeOtherDashboardSessions_(...params);
    case "authorizePhase83MailOnce": return authorizePhase83MailOnce_(...params);
    case "startAccessReviewCycle": return startAccessReviewCycle_(...params);
    case "getAccessReviewCenter": return getAccessReviewCenter_(...params);
    case "submitAccessReviewDecision": return submitAccessReviewDecision_(...params);
    case "submitAccessRequest": return submitAccessRequest_(...params);
    case "getAccessRequestCenter": return getAccessRequestCenter_(...params);
    case "getMyAccessRequests": return getMyAccessRequests_(...params);
    case "assignAccessRequestApprover": return assignAccessRequestApprover_(...params);
    case "reviewAccessRequest": return reviewAccessRequest_(...params);
    case "cancelAccessRequest": return cancelAccessRequest_(...params);
    case "resubmitAccessRequest": return resubmitAccessRequest_(...params);
    case "initializePhase87": return initializePhase87_(...params);
    case "getPrivilegedAccessCenter": return getPrivilegedAccessCenter_(...params);
    case "grantEmergencyAccess": return grantEmergencyAccess_(...params);
    case "restorePrivilegedGrant": return restorePrivilegedGrant_(...params);
    case "reviewEmergencyAccess": return reviewEmergencyAccess_(...params);
    case "getIamReportingCenter": return getIamReportingCenter_(...params);
    case "exportIamReportingExcel": return exportIamReportingExcel_(...params);
    case "exportIamReportingCsvBundle": return exportIamReportingCsvBundle_(...params);
    case "getEvidenceManagement": return getEvidenceManagement_(...params);
    case "getEvidenceDetails": return getEvidenceDetails_(...params);
    case "proposeEvidenceMetadata": return proposeEvidenceMetadata_(...params);
    case "sendEvidenceRequestTestEmail": return sendEvidenceRequestTestEmail_(...params);
    case "retryEvidenceRequestNotification": return retryEvidenceRequestNotification_(...params);
    case "getEvidenceRequests": return getEvidenceRequests_(...params);
    case "getEvidenceRequestDetails": return getEvidenceRequestDetails_(...params);
    case "createEvidenceRequest": return createEvidenceRequest_(...params);
    case "cancelEvidenceRequest": return cancelEvidenceRequest_(...params);
    case "configureDashboardEmailLinkOnce": return configureDashboardEmailLinkOnce_(...params);
    case "deleteCustomDashboardRole": return deleteCustomDashboardRole_(...params);
    case "getControlTests": return getControlTests_(...params);
    case "createControlTest": return createControlTest_(...params);
    case "getControlTestDetails": return getControlTestDetails_(...params);
    case "reviewControlTest": return reviewControlTest_(...params);
    case "archiveComplianceReminderLog": return archiveComplianceReminderLog_(...params);
    case "getComplianceReminderSettings": return getComplianceReminderSettings_(...params);
    case "previewComplianceReminders": return previewComplianceReminders_(...params);
    case "setComplianceReminderSchedule": return setComplianceReminderSchedule_(...params);
    case "runComplianceRemindersNow": return runComplianceRemindersNow_(...params);
    case "getComplianceMonitoring": return getComplianceMonitoring_(...params);
    case "scanComplianceMonitoring": return scanComplianceMonitoring_(...params);
    case "updateComplianceMonitorException": return updateComplianceMonitorException_(...params);
    case "getComplianceMonitoringSchedule": return getComplianceMonitoringSchedule_(...params);
    case "setComplianceMonitoringSchedule": return setComplianceMonitoringSchedule_(...params);
    case "getAuditReadinessSnapshot": return getAuditReadinessSnapshot_(...params);
    case "exportAuditReadinessCsv": return exportAuditReadinessCsv_(...params);
    default: throw new Error('Unknown or inaccessible dashboard operation.');
  }
}
