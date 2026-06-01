// ============================================================
//  ATHLETIC STIPEND MANAGER — Code.gs
//  Stand-alone Apps Script Web App
// ============================================================

// ── CONFIGURATION ────────────────────────────────────────────
const CONFIG = {
  MASTER_LOOKUP_SS_ID: 'YOUR_MASTER_LOOKUP_SPREADSHEET_ID_HERE',

  // Map each authorized user email → their school's data spreadsheet ID
  SCHOOL_MAP: {
    'bmayeski@guhsd.net':       '1tbBGYuq9dYqOaW00Ysv_g4pILV3rlh1y-3wJmNAOta4',
    'secretary@schoolb.edu':    'SCHOOL_B_SPREADSHEET_ID',
    'secretary@schoolc.edu':    'SCHOOL_C_SPREADSHEET_ID',
    // Add more schools here
  },

  // Sheet tab names — Master Lookup Spreadsheet
  SHEETS: {
    STIPEND_LOOKUP:  'StipendLookup',   // Position × Experience tier → dollar amounts
    SPORTS_LIST:     'SportsList',       // Sports per season, boys/girls
    COACH_ROSTER:    'CoachRoster',      // All coach employee data
    SCHOOL_CONFIG:   'SchoolConfig',     // Per-school overrides (future use)
  },

  // Sheet tab names — each School Spreadsheet
  SCHOOL_SHEETS: {
    FALL:   'Fall',
    WINTER: 'Winter',
    SPRING: 'Spring',
    SPORTS_OVERRIDE: 'SportsOverride',   // School-level sport add/remove
  },

  SEASONS: ['Fall', 'Winter', 'Spring'],

  POSITIONS: ['Head Coach', 'Assistant Coach', 'JV Coach', 'Freshman Coach'],

  EXP_TIERS: ['1-3', '4-6', '7-9', '10+'],

  PAYMENT_MONTHS: {
    Fall:   ['October', 'November', 'December'],
    Winter: ['January', 'February', 'March'],
    Spring: ['April',   'May',      'June'],
  },
};

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Athletic Stipend Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Include helper for HTML templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── AUTH & SCHOOL RESOLUTION ─────────────────────────────────
function getCurrentUserInfo() {
  const email = Session.getActiveUser().getEmail();
  const ssId  = CONFIG.SCHOOL_MAP[email];

  if (!ssId) {
    return { authorized: false, email, school: null };
  }

  // Pull school display name from the School Data spreadsheet metadata
  try {
    const ss   = SpreadsheetApp.openById(ssId);
    const name = ss.getName();
    return { authorized: true, email, school: name, ssId };
  } catch(err) {
    return { authorized: false, email, school: null, error: err.message };
  }
}

// ── SPREADSHEET HELPERS ──────────────────────────────────────
function getMasterSS() {
  return SpreadsheetApp.openById(CONFIG.MASTER_LOOKUP_SS_ID);
}

function getSchoolSS() {
  const email = Session.getActiveUser().getEmail();
  const ssId  = CONFIG.SCHOOL_MAP[email];
  if (!ssId) throw new Error('Unauthorized: no spreadsheet mapped for ' + email);
  return SpreadsheetApp.openById(ssId);
}

function sheetData(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] ?? '');
    return obj;
  });
}

// ── LOOKUP DATA ──────────────────────────────────────────────

/**
 * Returns stipend lookup table:
 * [{ Position, '1-3', '4-6', '7-9', '10+' }, ...]
 * Columns in sheet: Position | 1-3 | 4-6 | 7-9 | 10+
 */
function getStipendLookup() {
  const sheet = getMasterSS().getSheetByName(CONFIG.SHEETS.STIPEND_LOOKUP);
  return sheetData(sheet);
}

/**
 * Returns sports list grouped by season and gender.
 * Sheet columns: Sport | Season | Gender | Active
 * Returns: { Fall: { Boys: [...], Girls: [...] }, Winter: {...}, Spring: {...} }
 */
function getSportsList() {
  const master  = getMasterSS().getSheetByName(CONFIG.SHEETS.SPORTS_LIST);
  const rows    = sheetData(master);
  const result  = { Fall: {Boys:[], Girls:[]}, Winter: {Boys:[], Girls:[]}, Spring: {Boys:[], Girls:[]} };

  // Apply school-level overrides (add/remove per site)
  let overrides = [];
  try {
    const ss = getSchoolSS();
    const ov = ss.getSheetByName(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
    if (ov) overrides = sheetData(ov);
  } catch(_) {}

  const removedKeys = new Set(
    overrides.filter(r => r['Action'] === 'REMOVE').map(r => `${r['Season']}|${r['Gender']}|${r['Sport']}`)
  );
  const added = overrides.filter(r => r['Action'] === 'ADD');

  rows.forEach(r => {
    const key = `${r['Season']}|${r['Gender']}|${r['Sport']}`;
    if (!removedKeys.has(key) && result[r['Season']] && r['Gender']) {
      const genderKey = r['Gender'] === 'Boys' ? 'Boys' : 'Girls';
      result[r['Season']][genderKey].push(String(r['Sport']));
    }
  });

  added.forEach(r => {
    if (result[r['Season']] && r['Gender']) {
      const genderKey = r['Gender'] === 'Boys' ? 'Boys' : 'Girls';
      if (!result[r['Season']][genderKey].includes(r['Sport'])) {
        result[r['Season']][genderKey].push(String(r['Sport']));
      }
    }
  });

  return result;
}

/**
 * Returns full coach roster for autocomplete/lookup.
 * Sheet columns: Name | EmployeeID | EmployeeRecNum | PositionNum |
 *                ComboCode | Classification | PositionCode | Email
 */
function getCoachRoster() {
  const sheet = getMasterSS().getSheetByName(CONFIG.SHEETS.COACH_ROSTER);
  return sheetData(sheet);
}

// ── SEASON DATA (per school) ──────────────────────────────────

/**
 * Returns all coach assignments for a given season.
 * Each row: { Sport, Gender, CoachName, EmployeeID, EmployeeRecNum,
 *             PositionNum, ComboCode, Classification, PositionCode,
 *             YearsExp, ExpTier, StipendPct, StipendAmt,
 *             PAFNumber, Month1Amt, Month2Amt, Month3Amt,
 *             Month1Name, Month2Name, Month3Name, RowIndex }
 */
function getSeasonData(season) {
  if (!CONFIG.SEASONS.includes(season)) throw new Error('Invalid season: ' + season);
  const ss    = getSchoolSS();
  const sheet = ss.getSheetByName(season);
  if (!sheet) return [];

  const rows    = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  return rows.slice(1).map((row, i) => {
    const obj = { RowIndex: i + 2 }; // 1-based, +1 for header
    headers.forEach((h, j) => obj[h] = row[j] ?? '');
    return obj;
  });
}

/**
 * Saves (upsert) a single coach assignment row.
 * data: { RowIndex (if update), Sport, Gender, CoachName, EmployeeID,
 *         EmployeeRecNum, PositionNum, ComboCode, Classification,
 *         PositionCode, YearsExp, ExpTier, StipendPct, StipendAmt,
 *         PAFNumber, Month1Amt, Month2Amt, Month3Amt }
 */
function saveCoachAssignment(season, data) {
  if (!CONFIG.SEASONS.includes(season)) throw new Error('Invalid season');
  const ss    = getSchoolSS();
  const sheet = ss.getSheetByName(season);
  if (!sheet) throw new Error('Season sheet not found: ' + season);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(h => data[h] !== undefined ? data[h] : '');

  if (data.RowIndex && data.RowIndex >= 2) {
    // Update existing row
    sheet.getRange(data.RowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Append new row
    sheet.appendRow(rowData);
  }
  return { success: true };
}

/**
 * Deletes a coach assignment row by RowIndex.
 */
function deleteCoachAssignment(season, rowIndex) {
  if (!CONFIG.SEASONS.includes(season)) throw new Error('Invalid season');
  const ss    = getSchoolSS();
  const sheet = ss.getSheetByName(season);
  if (!sheet) throw new Error('Season sheet not found');
  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ── SPORT MANAGEMENT (school-level overrides) ─────────────────

function addSportOverride(season, gender, sport) {
  const ss    = getSchoolSS();
  let   sheet = ss.getSheetByName(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
    sheet.appendRow(['Season', 'Gender', 'Sport', 'Action']);
  }
  sheet.appendRow([season, gender, sport, 'ADD']);
  return { success: true };
}

function removeSportOverride(season, gender, sport) {
  const ss    = getSchoolSS();
  let   sheet = ss.getSheetByName(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
    sheet.appendRow(['Season', 'Gender', 'Sport', 'Action']);
  }
  sheet.appendRow([season, gender, sport, 'REMOVE']);
  return { success: true };
}

// ── STIPEND CALCULATION ───────────────────────────────────────

/**
 * Calculates stipend amount from lookup table.
 * positionCode: e.g. 'Head Coach'
 * expTier: e.g. '4-6'
 * pct: percentage as decimal 0–1
 */
function calculateStipend(positionCode, expTier, pct) {
  const table = getStipendLookup();
  const row   = table.find(r => r['Position'] === positionCode);
  if (!row) return 0;
  const base  = parseFloat(row[expTier]) || 0;
  return Math.round(base * pct * 100) / 100;
}

/**
 * Splits a stipend amount evenly into 3 monthly payments (33/33/34 default).
 * Allows custom split percentages per month.
 */
function splitStipend(totalAmt, customSplit) {
  // customSplit: [pct1, pct2, pct3] summing to 1, or null for even
  const split = customSplit || [1/3, 1/3, 1/3];
  const m1 = Math.round(totalAmt * split[0] * 100) / 100;
  const m2 = Math.round(totalAmt * split[1] * 100) / 100;
  const m3 = Math.round((totalAmt - m1 - m2) * 100) / 100; // ensure exact total
  return [m1, m2, m3];
}

// ── DASHBOARD SUMMARY ─────────────────────────────────────────

/**
 * Returns aggregate stats for the dashboard header.
 */
function getDashboardSummary() {
  const summary = {
    school: '',
    totalCoaches:   { Fall: 0, Winter: 0, Spring: 0 },
    totalStipends:  { Fall: 0, Winter: 0, Spring: 0 },
    pendingPAF:     { Fall: 0, Winter: 0, Spring: 0 },
    sportsCounts:   { Fall: 0, Winter: 0, Spring: 0 },
  };

  try {
    const info = getCurrentUserInfo();
    summary.school = info.school || info.email;
  } catch(_) {}

  CONFIG.SEASONS.forEach(season => {
    try {
      const rows = getSeasonData(season);
      summary.totalCoaches[season]  = rows.length;
      summary.totalStipends[season] = rows.reduce((sum, r) => sum + (parseFloat(r['StipendAmt']) || 0), 0);
      summary.pendingPAF[season]    = rows.filter(r => !r['PAFNumber']).length;
    } catch(_) {}
  });

  try {
    const sports = getSportsList();
    CONFIG.SEASONS.forEach(s => {
      summary.sportsCounts[s] = (sports[s]?.Boys?.length || 0) + (sports[s]?.Girls?.length || 0);
    });
  } catch(_) {}

  return summary;
}

// ── SHEET INITIALIZATION ──────────────────────────────────────
/**
 * Call once to set up a new school's spreadsheet with correct headers.
 * Run from the Apps Script editor, not from the web app.
 */
function initializeSchoolSpreadsheet() {
  const ss = getSchoolSS();

  const seasonHeaders = [
    'Sport', 'Gender', 'CoachName', 'EmployeeID', 'EmployeeRecNum',
    'PositionNum', 'ComboCode', 'Classification', 'PositionCode',
    'YearsExp', 'ExpTier', 'StipendPct', 'StipendAmt',
    'PAFNumber', 'Month1Name', 'Month1Amt', 'Month2Name', 'Month2Amt',
    'Month3Name', 'Month3Amt'
  ];

  CONFIG.SEASONS.forEach(season => {
    let sheet = ss.getSheetByName(season);
    if (!sheet) sheet = ss.insertSheet(season);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(seasonHeaders);
      sheet.getRange(1, 1, 1, seasonHeaders.length).setFontWeight('bold');
    }
  });

  // Sports override sheet
  let ovSheet = ss.getSheetByName(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
  if (!ovSheet) {
    ovSheet = ss.insertSheet(CONFIG.SCHOOL_SHEETS.SPORTS_OVERRIDE);
    ovSheet.appendRow(['Season', 'Gender', 'Sport', 'Action']);
    ovSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  return 'Initialization complete.';
}

/**
 * Call once to set up the Master Lookup spreadsheet with sample data.
 * Run from the Apps Script editor.
 */
function initializeMasterLookup() {
  const ss = getMasterSS();

  // StipendLookup
  let sl = ss.getSheetByName(CONFIG.SHEETS.STIPEND_LOOKUP);
  if (!sl) sl = ss.insertSheet(CONFIG.SHEETS.STIPEND_LOOKUP);
  sl.clearContents();
  sl.appendRow(['Position', '1-3', '4-6', '7-9', '10+']);
  sl.appendRow(['Head Coach',       3000, 3500, 4000, 4500]);
  sl.appendRow(['Assistant Coach',  2000, 2300, 2600, 3000]);
  sl.appendRow(['JV Coach',         1800, 2100, 2400, 2700]);
  sl.appendRow(['Freshman Coach',   1500, 1800, 2100, 2400]);
  sl.getRange(1,1,1,5).setFontWeight('bold');

  // SportsList
  let sp = ss.getSheetByName(CONFIG.SHEETS.SPORTS_LIST);
  if (!sp) sp = ss.insertSheet(CONFIG.SHEETS.SPORTS_LIST);
  sp.clearContents();
  const sportsData = [
    ['Sport','Season','Gender'],
    ['Football','Fall','Boys'],['Soccer','Fall','Boys'],['Cross Country','Fall','Boys'],['Water Polo','Fall','Boys'],['Golf','Fall','Boys'],
    ['Volleyball','Fall','Girls'],['Soccer','Fall','Girls'],['Cross Country','Fall','Girls'],['Tennis','Fall','Girls'],['Golf','Fall','Girls'],
    ['Basketball','Winter','Boys'],['Wrestling','Winter','Boys'],['Swimming','Winter','Boys'],['Soccer (Indoor)','Winter','Boys'],
    ['Basketball','Winter','Girls'],['Soccer (Indoor)','Winter','Girls'],['Swimming','Winter','Girls'],['Volleyball (Indoor)','Winter','Girls'],
    ['Baseball','Spring','Boys'],['Track & Field','Spring','Boys'],['Tennis','Spring','Boys'],['Golf','Spring','Boys'],['Lacrosse','Spring','Boys'],
    ['Softball','Spring','Girls'],['Track & Field','Spring','Girls'],['Tennis','Spring','Girls'],['Golf','Spring','Girls'],['Lacrosse','Spring','Girls'],
  ];
  sp.getRange(1, 1, sportsData.length, 3).setValues(sportsData);
  sp.getRange(1,1,1,3).setFontWeight('bold');

  // CoachRoster (sample)
  let cr = ss.getSheetByName(CONFIG.SHEETS.COACH_ROSTER);
  if (!cr) cr = ss.insertSheet(CONFIG.SHEETS.COACH_ROSTER);
  cr.clearContents();
  cr.appendRow(['Name','EmployeeID','EmployeeRecNum','PositionNum','ComboCode','Classification','PositionCode','Email']);
  cr.appendRow(['Jane Smith','E10001','R001','P4421','CC-001','Certificated','Head Coach','jsmith@school.edu']);
  cr.appendRow(['John Doe','E10002','R002','P4422','CC-002','Classified','Assistant Coach','jdoe@school.edu']);
  cr.getRange(1,1,1,8).setFontWeight('bold');

  return 'Master lookup initialized.';
}
