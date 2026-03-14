// ── HOA Manager – Apps Script Backend ────────────────────────────────────────

const S_SETTINGS  = 'Settings';
const S_PROPS     = 'Properties';
const S_PMTS      = 'Payments';
const S_VIOLS     = 'Violations';
const S_TASKS     = 'Tasks';
const S_TEMPLATES = 'TaskTemplates';
const S_MEMBERS   = 'BoardMembers';
const S_VENDORS   = 'Vendors';
const S_MINUTES   = 'Minutes';
const S_DOCS      = 'PropertyDocs';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HOA Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Schema & Migration ────────────────────────────────────────────────────────
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schema = {
    [S_SETTINGS]:  ['key','value'],
    [S_PROPS]:     ['id','address','lotNumber','ownerName','ownerEmail','ownerPhone',
                    'moveInDate','duesStatus','ownerOccupied',
                    'contactName','contactPhone','contactEmail','contactAddress','notes','status'],
    [S_PMTS]:      ['id','propertyId','date','period','amount','status','note'],
    [S_VIOLS]:     ['id','propertyId','date','description','status'],
    [S_TASKS]:     ['id','year','month','description','assignee','dueDate','status','notes'],
    [S_TEMPLATES]: ['id','month','description','assignee','notes'],
    [S_MEMBERS]:   ['id','name','role','email','phone','termEnd','notes'],
    [S_VENDORS]:   ['id','name','category','phone','email','notes'],
    [S_MINUTES]:   ['id','date','title','attendees','notes'],
    [S_DOCS]:      ['id','propertyId','date','description']
  };

  Object.entries(schema).forEach(([name, headers]) => {
    if (!ss.getSheetByName(name)) {
      const sh = ss.insertSheet(name);
      sh.appendRow(headers);
      sh.getRange(1,1,1,headers.length)
        .setFontWeight('bold').setBackground('#4F46E5').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    } else {
      // Migrate: add any missing columns to existing sheets
      const sh = ss.getSheetByName(name);
      const lastCol = Math.max(1, sh.getLastColumn());
      const existing = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h));
      headers.forEach(h => {
        if (!existing.includes(h)) {
          const col = sh.getLastColumn() + 1;
          sh.getRange(1, col).setValue(h)
            .setFontWeight('bold').setBackground('#4F46E5').setFontColor('#ffffff');
        }
      });
    }
  });

  // Default settings
  const existingKeys = sheetToObjects(ss.getSheetByName(S_SETTINGS)).map(r => r.key);
  const defaults = {
    hoaName:    'Homeowners Association',
    duesAmount: '0',
    duesDueDate:'March 1',
    hoaPhone:   '',
    hoaAddress: '',
    appUrl:     ''
  };
  Object.entries(defaults).forEach(([k,v]) => {
    if (!existingKeys.includes(k)) ss.getSheetByName(S_SETTINGS).appendRow([k, v]);
  });

  // Seed default task templates if empty
  const tmpl = ss.getSheetByName(S_TEMPLATES);
  if (sheetToObjects(tmpl).length === 0) {
    [['1','Send annual dues notices to all homeowners','','Send by mail and email'],
     ['2','Follow up on unpaid dues','',''],
     ['3','Begin dues collection — contact delinquent owners','','Check unpaid after 60 days'],
     ['3','Schedule spring common area inspection','',''],
     ['4','Spring common area inspection','','Document any violations found'],
     ['5','Send violation notices if applicable','','Allow 30 days to remediate'],
     ['6','Mid-year financial review','','Review budget vs actuals'],
     ['7','Plan annual meeting agenda','',''],
     ['8','Send annual meeting notice to homeowners','','Check bylaws for required notice period'],
     ['9','Annual homeowners meeting','',''],
     ['9','Fall common area inspection','','Before winter sets in'],
     ['10','Budget planning for next year','',''],
     ['11','Send budget and dues notice for next year','',''],
     ['12','Year-end financial summary','','Prepare for records'],
     ['12','Confirm board officer roles for next year','','']
    ].forEach(r => tmpl.appendRow([Utilities.getUuid(), ...r]));
  }
}

function sheetToObjects(sheet) {
  const vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return [];
  const headers = vals[0].map(h => String(h));
  const tz = Session.getScriptTimeZone();
  return vals.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      obj[h] = (v === null || v === undefined) ? '' : String(v);
    });
    return obj;
  });
}

// ── Main Data Fetch ───────────────────────────────────────────────────────────
function getAllData() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings  = {};
  sheetToObjects(ss.getSheetByName(S_SETTINGS)).forEach(r => { settings[r.key] = r.value; });
  const properties = sheetToObjects(ss.getSheetByName(S_PROPS));
  const payments   = sheetToObjects(ss.getSheetByName(S_PMTS));
  const violations = sheetToObjects(ss.getSheetByName(S_VIOLS));
  const docs       = sheetToObjects(ss.getSheetByName(S_DOCS));
  return {
    settings,
    properties: properties.map(p => ({
      ...p,
      status:     p.status || 'active',
      payments:   payments.filter(pm => pm.propertyId === p.id),
      violations: violations.filter(v  => v.propertyId  === p.id),
      docs:       docs.filter(d => d.propertyId === p.id)
    })),
    tasks:     sheetToObjects(ss.getSheetByName(S_TASKS)),
    templates: sheetToObjects(ss.getSheetByName(S_TEMPLATES)),
    members:   sheetToObjects(ss.getSheetByName(S_MEMBERS)),
    vendors:   sheetToObjects(ss.getSheetByName(S_VENDORS)),
    minutes:   sheetToObjects(ss.getSheetByName(S_MINUTES))
  };
}

// ── Properties ────────────────────────────────────────────────────────────────
function addProperty(p) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PROPS).appendRow([
    id, p.address||'', p.lotNumber||'', p.ownerName||'', p.ownerEmail||'', p.ownerPhone||'',
    p.moveInDate||'', p.duesStatus||'unpaid', p.ownerOccupied||'yes',
    p.contactName||'', p.contactPhone||'', p.contactEmail||'', p.contactAddress||'',
    p.notes||'', p.status||'active'
  ]);
  return id;
}

function updateProperty(p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PROPS);
  const vals  = sheet.getDataRange().getValues();
  const hdr   = vals[0].map(h => String(h));
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(p.id)) {
      const fields = {
        address:p.address||'', lotNumber:p.lotNumber||'', ownerName:p.ownerName||'',
        ownerEmail:p.ownerEmail||'', ownerPhone:p.ownerPhone||'', moveInDate:p.moveInDate||'',
        duesStatus:p.duesStatus||'unpaid', ownerOccupied:p.ownerOccupied||'yes',
        contactName:p.contactName||'', contactPhone:p.contactPhone||'',
        contactEmail:p.contactEmail||'', contactAddress:p.contactAddress||'',
        notes:p.notes||'', status:p.status||'active'
      };
      // Build full row, preserving any columns we don't manage (e.g. future migrations)
      const newRow = hdr.map((col, idx) => fields.hasOwnProperty(col) ? fields[col] : vals[i][idx]);
      sheet.getRange(i+1, 1, 1, hdr.length).setValues([newRow]);
      return true;
    }
  }
  return false;
}

function deleteProperty(id) {
  _deleteByCol(S_PROPS, id, 0);
  _deleteByCol(S_PMTS,  id, 1);
  _deleteByCol(S_VIOLS, id, 1);
  _deleteByCol(S_DOCS,  id, 1);
}

function bulkResetDues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PROPS);
  const vals  = sheet.getDataRange().getValues();
  const hdr   = vals[0].map(h => String(h));
  const duesCol   = hdr.indexOf('duesStatus');
  const statusCol = hdr.indexOf('status');
  if (duesCol === -1) return 0;
  let count = 0;
  // Collect all rows that need updating, then write in one batch per contiguous block
  // Simplest correct approach: build full updated data array and write once
  const updates = []; // { rowIndex, colIndex, value }
  for (let i = 1; i < vals.length; i++) {
    const st = statusCol > -1 ? String(vals[i][statusCol]) : 'active';
    if (st === 'active') {
      vals[i][duesCol] = 'unpaid';
      count++;
    }
  }
  if (count > 0) {
    // Write all data rows back in a single call
    sheet.getRange(2, 1, vals.length - 1, hdr.length).setValues(vals.slice(1));
  }
  return count;
}

// ── Payments ──────────────────────────────────────────────────────────────────
function addPayment(propertyId, pmt) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PMTS).appendRow([
    id, propertyId, pmt.date||'', pmt.period||'', pmt.amount||0, pmt.status||'paid', pmt.note||''
  ]);
  return id;
}
function deletePayment(id) { _deleteByCol(S_PMTS, id, 0); }

// ── Violations ────────────────────────────────────────────────────────────────
function addViolation(propertyId, v) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_VIOLS).appendRow([
    id, propertyId, v.date||'', v.description||'', v.status||'open'
  ]);
  return id;
}
function deleteViolation(id) { _deleteByCol(S_VIOLS, id, 0); }

// ── Tasks ─────────────────────────────────────────────────────────────────────
function addTask(t) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_TASKS).appendRow([
    id, t.year||'', t.month||'', t.description||'',
    t.assignee||'', t.dueDate||'', t.status||'pending', t.notes||''
  ]);
  return id;
}
function updateTask(t) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_TASKS);
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(t.id)) {
      sheet.getRange(i+1,1,1,8).setValues([[
        t.id, t.year||'', t.month||'', t.description||'',
        t.assignee||'', t.dueDate||'', t.status||'pending', t.notes||''
      ]]);
      return true;
    }
  }
  return false;
}
function deleteTask(id) { _deleteByCol(S_TASKS, id, 0); }

function loadTasksForYear(year) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const tmpl      = sheetToObjects(ss.getSheetByName(S_TEMPLATES));
  const taskSheet = ss.getSheetByName(S_TASKS);
  const existingDescs = sheetToObjects(taskSheet)
    .filter(t => String(t.year) === String(year))
    .map(t => t.description.trim().toLowerCase());
  let created = 0, skipped = 0;
  tmpl.forEach(t => {
    if (existingDescs.includes(t.description.trim().toLowerCase())) { skipped++; return; }
    taskSheet.appendRow([
      Utilities.getUuid(), year, t.month, t.description,
      t.assignee||'', '', 'pending', t.notes||''
    ]);
    created++;
  });
  return { created, skipped };
}

// ── Task Templates ────────────────────────────────────────────────────────────
function addTemplate(t) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_TEMPLATES).appendRow([
    id, t.month||'0', t.description||'', t.assignee||'', t.notes||''
  ]);
  return id;
}
function updateTemplate(t) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_TEMPLATES);
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(t.id)) {
      sheet.getRange(i+1,1,1,5).setValues([[t.id,t.month||'0',t.description||'',t.assignee||'',t.notes||'']]);
      return true;
    }
  }
  return false;
}
function deleteTemplate(id) { _deleteByCol(S_TEMPLATES, id, 0); }

// ── Board Members ─────────────────────────────────────────────────────────────
function addMember(m) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MEMBERS).appendRow([
    id, m.name||'', m.role||'', m.email||'', m.phone||'', m.termEnd||'', m.notes||''
  ]);
  return id;
}
function updateMember(m) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MEMBERS);
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(m.id)) {
      sheet.getRange(i+1,1,1,7).setValues([[m.id,m.name||'',m.role||'',m.email||'',m.phone||'',m.termEnd||'',m.notes||'']]);
      return true;
    }
  }
  return false;
}
function deleteMember(id) { _deleteByCol(S_MEMBERS, id, 0); }

// ── Vendors ───────────────────────────────────────────────────────────────────
function addVendor(v) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_VENDORS).appendRow([
    id, v.name||'', v.category||'', v.phone||'', v.email||'', v.notes||''
  ]);
  return id;
}
function updateVendor(v) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_VENDORS);
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(v.id)) {
      sheet.getRange(i+1,1,1,6).setValues([[v.id,v.name||'',v.category||'',v.phone||'',v.email||'',v.notes||'']]);
      return true;
    }
  }
  return false;
}
function deleteVendor(id) { _deleteByCol(S_VENDORS, id, 0); }

// ── Meeting Minutes ───────────────────────────────────────────────────────────
function addMinutes(m) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MINUTES).appendRow([
    id, m.date||'', m.title||'', m.attendees||'', m.notes||''
  ]);
  return id;
}
function updateMinutes(m) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MINUTES);
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(m.id)) {
      sheet.getRange(i+1,1,1,5).setValues([[m.id,m.date||'',m.title||'',m.attendees||'',m.notes||'']]);
      return true;
    }
  }
  return false;
}
function deleteMinutes(id) { _deleteByCol(S_MINUTES, id, 0); }

// ── Property Documents ────────────────────────────────────────────────────────
function addDoc(propertyId, doc) {
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_DOCS).appendRow([
    id, propertyId, doc.date||'', doc.description||''
  ]);
  return id;
}
function deleteDoc(id) { _deleteByCol(S_DOCS, id, 0); }

// ── Settings ──────────────────────────────────────────────────────────────────
function updateSettings(settings) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_SETTINGS);
  const vals  = sheet.getDataRange().getValues();
  Object.entries(settings).forEach(([key, val]) => {
    let found = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === key) { sheet.getRange(i+1, 2).setValue(val); found = true; break; }
    }
    if (!found) sheet.appendRow([key, val]);
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _deleteByCol(sheetName, value, colIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const vals = sheet.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][colIndex]) === String(value)) sheet.deleteRow(i+1);
  }
}
