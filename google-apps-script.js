// ═══════════════════════════════════════════════════════════════
//  CHURCH ATTENDANCE SYSTEM — Google Apps Script Backend
//  Paste this entire file into Extensions → Apps Script
//  Then deploy as a Web App (Anyone can access)
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Sheet names — created automatically on first run
const SHEETS = {
  CONFIG:   'Config',
  MEMBERS:  'Members',
  SESSIONS: 'Sessions',
  ATTENDANCE: 'Attendance'
};

// ── ENTRY POINT ──────────────────────────────────────────────
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  let result;

  try {
    switch(data.action) {
      case 'load':        result = loadAll();                    break;
      case 'saveSession': result = saveSession(data);            break;
      case 'saveMember':  result = saveMember(data.member);      break;
      case 'saveConfig':  result = saveConfig(data);             break;
      default:            result = {error: 'Unknown action'};
    }
  } catch(err) {
    result = {error: err.message};
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Health check
  return ContentService
    .createTextOutput(JSON.stringify({status:'ok', message:'Church Attendance API running'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LOAD ALL DATA ────────────────────────────────────────────
function loadAll() {
  ensureSheetsExist();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Load config (classes + departments)
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const configData  = configSheet.getDataRange().getValues();
  const classes = [], departments = [];

  for (let i = 1; i < configData.length; i++) {
    const [type, id, name, color] = configData[i];
    if (type === 'class')  classes.push({id, name, color});
    if (type === 'dept')   departments.push({id, name});
  }

  // Load members
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const memberData  = memberSheet.getDataRange().getValues();
  const members = [];
  for (let i = 1; i < memberData.length; i++) {
    const [id, name, classId, deptsRaw] = memberData[i];
    if (!id) continue;
    const depts = deptsRaw ? deptsRaw.split(',').filter(Boolean) : [];
    members.push({id: String(id), name, classId, depts});
  }

  // Load attendance sessions
  const attSheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const attData  = attSheet.getDataRange().getValues();
  const sessionsMap = {};

  for (let i = 1; i < attData.length; i++) {
    const [date, memberId, status] = attData[i];
    if (!date || !memberId) continue;
    if (!sessionsMap[date]) sessionsMap[date] = {};
    sessionsMap[date][String(memberId)] = status;
  }

  const sessions = Object.entries(sessionsMap)
    .map(([date, attendance]) => ({date, attendance}))
    .sort((a,b) => a.date.localeCompare(b.date));

  return {classes, departments, members, sessions};
}

// ── SAVE SESSION ─────────────────────────────────────────────
function saveSession(data) {
  ensureSheetsExist();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const attSheet = ss.getSheetByName(SHEETS.ATTENDANCE);

  // Remove existing rows for this date
  const existing = attSheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = existing.length - 1; i >= 1; i--) {
    if (existing[i][0] === data.date) toDelete.push(i + 1);
  }
  toDelete.forEach(row => attSheet.deleteRow(row));

  // Write new attendance rows
  const rows = [];
  for (const [memberId, status] of Object.entries(data.attendance)) {
    if (status) rows.push([data.date, memberId, status]);
  }
  if (rows.length) {
    attSheet.getRange(attSheet.getLastRow()+1, 1, rows.length, 3).setValues(rows);
  }

  // Also write a summary row to Sessions sheet
  const sessionSheet = ss.getSheetByName(SHEETS.SESSIONS);
  const present = Object.values(data.attendance).filter(s => s==='present').length;
  const absent  = Object.values(data.attendance).filter(s => s==='absent').length;
  // Check if date already exists in sessions
  const sessData = sessionSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < sessData.length; i++) {
    if (sessData[i][0] === data.date) {
      sessionSheet.getRange(i+1, 2, 1, 2).setValues([[present, absent]]);
      found = true; break;
    }
  }
  if (!found) sessionSheet.appendRow([data.date, present, absent]);

  return {success: true, date: data.date, rows: rows.length};
}

// ── SAVE MEMBER ──────────────────────────────────────────────
function saveMember(member) {
  ensureSheetsExist();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);

  // Check if member already exists (update) or is new (append)
  const existing = memberSheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(member.id)) {
      memberSheet.getRange(i+1, 1, 1, 4).setValues([
        [member.id, member.name, member.classId, (member.depts||[]).join(',')]
      ]);
      return {success: true, action: 'updated'};
    }
  }
  memberSheet.appendRow([member.id, member.name, member.classId, (member.depts||[]).join(',')]);
  return {success: true, action: 'created'};
}

// ── SAVE CONFIG (classes + departments) ──────────────────────
function saveConfig(data) {
  ensureSheetsExist();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);

  // Clear and rewrite
  configSheet.clearContents();
  configSheet.appendRow(['type', 'id', 'name', 'color']);
  const rows = [];
  (data.classes||[]).forEach(c => rows.push(['class', c.id, c.name, c.color||'']));
  (data.departments||[]).forEach(d => rows.push(['dept', d.id, d.name, '']));
  if (rows.length) configSheet.getRange(2, 1, rows.length, 4).setValues(rows);

  return {success: true};
}

// ── ENSURE SHEETS EXIST ──────────────────────────────────────
function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function getOrCreate(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#f0f4ff');
    }
    return sheet;
  }

  getOrCreate(SHEETS.CONFIG,    ['type','id','name','color']);
  getOrCreate(SHEETS.MEMBERS,   ['id','name','classId','departments']);
  getOrCreate(SHEETS.SESSIONS,  ['date','present','absent']);
  getOrCreate(SHEETS.ATTENDANCE,['date','memberId','status']);

  // Seed default classes/depts if Config is empty
  const config = ss.getSheetByName(SHEETS.CONFIG);
  if (config.getLastRow() <= 1) {
    const defaults = [
      ['class','c1','GA Class','#2563eb'],
      ['class','c2','Twi Class','#16a34a'],
      ['class','c3','Ewe Class','#7c3aed'],
      ['class','c4','Youth Class','#dc2626'],
      ['class','c5','Children','#0891b2'],
      ['class','c6','Elders','#d97706'],
      ['dept','d1','Ushers',''],
      ['dept','d2','Choristers',''],
      ['dept','d3','Youth Choir',''],
      ['dept','d4','Audio Visual',''],
      ['dept','d5','Finance Team',''],
      ['dept','d6','Podium Directors',''],
    ];
    config.getRange(2, 1, defaults.length, 4).setValues(defaults);
  }
}
