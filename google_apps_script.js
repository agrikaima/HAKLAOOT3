// ────────────────────────────────────────────────────────────────
// יומן שדה — Google Apps Script
// הוראות:
//   1. פתח Google Sheets חדש
//   2. תוספות → Apps Script → מחק קוד קיים → הדבק קוד זה
//   3. שמור (Ctrl+S)
//   4. פרסם → פרוס כאפליקציית אינטרנט
//      ● הרץ כ: אני (Me)
//      ● מי יכול לגשת: כל אחד (Anyone)  ← חשוב!
//   5. אשר הרשאות → העתק את ה-URL → הדבק בהגדרות האפליקציה
// ────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    switch (data._type) {
      case 'task':        writeTask(ss, data);        break;
      case 'procurement': writeProcurement(ss, data); break;
      case 'orders':      writeOrders(ss, data);      break;
      case 'rotation':    writeRotation(ss, data);    break;
      case 'catalog':     writeCatalog(ss, data);     break;
      case 'fieldmap':       writeFieldMap(ss, data);       break;
      case 'weeklySchedule': writeWeeklySchedule(ss, data); break;
      default:            writePlanting(ss, data);    break;
    }

    // יומן קטיף — שורות מהלוג
    if (!data._type && data.log && data.log.length) {
      const logSheet = getOrCreate(ss, 'יומן קטיף', LOG_HEADERS);
      data.log.forEach(entry => {
        if (entry.harvest) {
          logSheet.appendRow([
            data.id, data.field, data.bed,
            data.variety || data.clearVariety || '',
            data.season, entry.date, entry.time,
            entry.harvest, entry.text || ''
          ]);
        }
      });
    }

    return jsonOk();
  } catch (err) {
    return jsonError(err.message);
  }
}

function doGet(e) {
  return jsonOk('יומן שדה פעיל');
}

// ─── שתילות ──────────────────────────────────────────────────────
const PLANTING_HEADERS = [
  'מזהה','סוג תיעוד','עונה','תאריך','שם חלקה','מספר ערוגה',
  'זן / סוג צמח','כמות שתילים','מספר שורות','מספר שלוחות',
  'מרחק בין צמחים (ס"מ)','זן שחוסל','הערות','תאריך תיעוד'
];

function writePlanting(ss, d) {
  const sheet = getOrCreate(ss, 'שתילות', PLANTING_HEADERS);
  const typeLabel = d.recType==='rest'?'מנוחה':d.recType==='clear'?'חיסול':'שתילה';
  sheet.appendRow([
    d.id||'', typeLabel, d.season||'', d.date||'', d.field||'', d.bed||'',
    d.variety||'', d.qty||'', d.rows||'', d.shoots||'', d.distPlant||'',
    d.clearVariety||'', d.notes||'', new Date().toLocaleString('he-IL')
  ]);
}

// ─── יומן קטיף ───────────────────────────────────────────────────
const LOG_HEADERS = [
  'מזהה רשומה','שם חלקה','ערוגה','זן','עונה',
  'תאריך קטיף','שעה','משקל קטיף (ק"ג)','הערה'
];

// ─── משימות ──────────────────────────────────────────────────────
const TASK_HEADERS = [
  'מזהה','תיאור משימה','חלקה','עדיפות','סטטוס',
  'תאריך רישום','תאריך יעד','תאריך ביצוע','הערות','תאריך תיעוד'
];

const PRIORITY_HE = { urgent:'דחוף', normal:'רגיל', low:'נמוך' };
const STATUS_HE   = { open:'פתוח', wip:'בביצוע', done:'הושלם' };

function writeTask(ss, d) {
  const sheet = getOrCreate(ss, 'משימות', TASK_HEADERS);
  sheet.appendRow([
    d.id||'',
    d.desc||'',
    d.field||'',
    PRIORITY_HE[d.priority]||d.priority||'',
    STATUS_HE[d.status]||d.status||'',
    d.regDate||'',
    d.dueDate||'',
    d.doneDate||'',
    d.notes||'',
    new Date().toLocaleString('he-IL')
  ]);
}

// ─── רכש ─────────────────────────────────────────────────────────
const PROC_HEADERS = [
  'מזהה','קטגוריה','תאריך','שם פריט','כמות','יחידה',
  'מחיר (₪)','ספק','הערות','תאריך תיעוד'
];

function writeProcurement(ss, d) {
  const sheet = getOrCreate(ss, 'רכש', PROC_HEADERS);
  sheet.appendRow([
    d.id||'', d.cat||'', d.date||'', d.item||'',
    d.qty||'', d.unit||'', d.price||'', d.supplier||'',
    d.notes||'', new Date().toLocaleString('he-IL')
  ]);
}

// ─── מיפוי רוטציה ────────────────────────────────────────────────
const ROTATION_HEADERS = ['חלקה','ערוגה','אזור','ירק לעונה','אורך ערוגה (מ׳)'];

function writeRotation(ss, d) {
  const sheetName = 'מיפוי ' + (d.season || '');
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sheet = getOrCreate(ss, sheetName, ROTATION_HEADERS);
  d.rows.forEach(r => {
    sheet.appendRow([r.field, r.bed, r.zone, r.crop || '—', r.len]);
  });
  // Color coding by zone
  const zoneColors = { 'ירוקים':'#d1fae5', 'עונתיים':'#fee2e2', 'חול':'#fef3c7', 'מיוחד':'#f3e8ff' };
  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  for (let row = 2; row <= numRows; row++) {
    const zone = sheet.getRange(row, 3).getValue();
    if (zoneColors[zone]) {
      sheet.getRange(row, 1, 1, ROTATION_HEADERS.length).setBackground(zoneColors[zone]);
    }
  }
}

// ─── הזמנות ──────────────────────────────────────────────────────
const ORDER_HEADERS = [
  'צמח','עונה','אזור','משתל','שתילים','מ׳ ערוגות','שלוחות','כל (שבועות)'
];

function writeOrders(ss, d) {
  const sheetName = 'הזמנות ' + (d.season || '');
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sheet = getOrCreate(ss, sheetName, ORDER_HEADERS);
  d.rows.forEach(r => {
    sheet.appendRow([
      r.name, r.season, r.field||'', 'משתל ' + r.wave,
      r.plants, r.meters, r.shoots, r.intervalWeeks||''
    ]);
  });
  sheet.appendRow([]);
  sheet.appendRow(['סה"כ','','','',
    d.rows.reduce((s,r)=>s+(r.plants||0),0),
    d.rows.reduce((s,r)=>s+(r.meters||0),0).toFixed(1),
    d.rows.reduce((s,r)=>s+(r.shoots||0),0),''
  ]);
}

// ─── קטלוג ירקות ─────────────────────────────────────────────────
const CATALOG_BASE_HEADERS = [
  'שם','עונה','שלוחות','שורות','מרחק (סמ)','גלים','כל (שבועות)','תאריך ראשון',
  'שתילים/גל','מטר/גל','סה"כ שתילים','סה"כ מטר'
];

function writeCatalog(ss, d) {
  const sheetName = 'קטלוג ירקות';
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  // find max waves to build dynamic date columns
  const maxWaves = d.rows.reduce((m, r) => Math.max(m, (r.waveDates || []).length), 0);
  const dateHeaders = [];
  for (let i = 1; i <= maxWaves; i++) dateHeaders.push('תאריך משתל ' + i);
  const headers = CATALOG_BASE_HEADERS.concat(dateHeaders);

  const sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);

  // header row
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setBackground('#1b4332');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setHorizontalAlignment('center');

  // group by season — each crop is one row
  const seasons = [...new Set(d.rows.map(r => r.season))];
  let rowIdx = 2;

  seasons.forEach(season => {
    const seasonRows = d.rows.filter(r => r.season === season);

    seasonRows.forEach(r => {
      const waveDates = r.waveDates || [];
      const rowData = [
        r.name, r.season,
        r.drips !== undefined && r.drips !== '' ? r.drips : '',
        r.rows  !== undefined && r.rows  !== '' ? r.rows  : '',
        r.dist  !== undefined && r.dist  !== '' ? r.dist  : '',
        r.waves !== undefined && r.waves !== '' ? r.waves : '',
        r.interval !== undefined && r.interval !== '' ? r.interval : '',
        r.startDate || '',
        r.plantsWave || '',
        r.metersPerWave || '',
        r.plantsTotal || '',
        r.metersTotal || ''
      ].concat(waveDates);

      sheet.getRange(rowIdx, 1, 1, headers.length)
        .setValues([rowData.concat(Array(headers.length - rowData.length).fill(''))]);
      rowIdx++;
    });

    // totals row per season
    const totalPlants = seasonRows.reduce((s, r) => s + (r.plantsTotal || 0), 0);
    const totalMeters = seasonRows.reduce((s, r) => s + (r.metersTotal || 0), 0);
    const totalsRow = Array(headers.length).fill('');
    totalsRow[0] = 'סה"כ עונת ' + season;
    totalsRow[8] = '';
    totalsRow[10] = totalPlants;
    totalsRow[11] = Math.round(totalMeters * 10) / 10;

    const totRange = sheet.getRange(rowIdx, 1, 1, headers.length);
    totRange.setValues([totalsRow]);
    totRange.setBackground('#d1fae5');
    totRange.setFontWeight('bold');
    totRange.setBorder(true, true, true, true, false, false);
    rowIdx++;

    // blank separator
    rowIdx++;
  });

  // column widths
  sheet.setColumnWidth(1, 140); // שם
  sheet.setColumnWidth(2, 100); // עונה
  for (let i = 3; i <= 12; i++) sheet.setColumnWidth(i, 90);
  for (let i = 13; i <= headers.length; i++) sheet.setColumnWidth(i, 120);

  // borders around all data
  if (rowIdx > 2) {
    sheet.getRange(1, 1, rowIdx - 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  }
}

// ─── מפת שדה ─────────────────────────────────────────────────────
function writeFieldMap(ss, d) {
  const sheetName = 'מפת שדה ' + (d.season || '');
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);

  const ZONE_BG   = { 'עונתיים':'#2d6a4f', 'ירוקים':'#2f9e44', 'חול':'#b45309', 'מיוחד':'#7c3aed' };
  const FIELD_BG  = { 'עונתיים':'#10b981', 'ירוקים':'#22c55e', 'חול':'#d97706', 'מיוחד':'#8b5cf6' };
  const BED_COL_W = 54; // px per bed column
  const GAP       = 1;  // blank columns between paired fields

  // Group fields into ימין/שמאל pairs, preserving FIELD_DEFS order
  const pairs = [];
  const used  = new Set();
  d.fields.forEach(f => {
    if (used.has(f.name)) return;
    const base    = f.name.replace(/ ימין$/, '').replace(/ שמאל$/, '');
    const partner = d.fields.find(f2 => f2.name !== f.name &&
      f2.name.replace(/ ימין$/, '').replace(/ שמאל$/, '') === base &&
      !used.has(f2.name));
    if (partner) {
      const right = f.name.endsWith('ימין')    ? f : partner;
      const left  = f.name.endsWith('שמאל')   ? f : partner;
      pairs.push({ type:'pair', right, left, zone: f.zone, base });
      used.add(f.name); used.add(partner.name);
    } else {
      pairs.push({ type:'single', field: f, zone: f.zone });
      used.add(f.name);
    }
  });

  let row = 1;
  let prevZone = null;

  pairs.forEach(g => {
    // Zone header row
    if (g.zone !== prevZone) {
      if (prevZone !== null) { row++; } // blank separator
      prevZone = g.zone;
      const maxW = g.type === 'pair'
        ? g.right.beds + GAP + g.left.beds
        : g.field.beds;
      const zr = sheet.getRange(row, 1, 1, maxW + 2);
      zr.merge();
      zr.setValue('▸ ' + (g.zone || ''));
      zr.setBackground(ZONE_BG[g.zone] || '#1f2937');
      zr.setFontColor('#ffffff');
      zr.setFontWeight('bold');
      zr.setFontSize(10);
      zr.setHorizontalAlignment('right');
      zr.setVerticalAlignment('middle');
      sheet.setRowHeight(row, 24);
      row++;
    }

    if (g.type === 'pair') {
      _renderBlock(sheet, g.right, row, 1, BED_COL_W, FIELD_BG);
      row += 3;
      _renderBlock(sheet, g.left,  row, 1, BED_COL_W, FIELD_BG);
    } else {
      _renderBlock(sheet, g.field, row, 1, BED_COL_W, FIELD_BG);
    }
    row += 3; // header + bed# row + content row
  });

  // gap column between paired fields (leave narrow)
  sheet.setColumnWidth(1, 4);
}

function _renderBlock(sheet, field, startRow, startCol, colW, FIELD_BG) {
  const n = field.beds;

  // ── Header ──
  const hdr = n > 1 ? sheet.getRange(startRow, startCol, 1, n) : sheet.getRange(startRow, startCol);
  if (n > 1) hdr.merge();
  hdr.setValue(field.name + '  (' + n + ' ע׳ × ' + field.len + 'מ׳)');
  hdr.setBackground(FIELD_BG[field.zone] || '#374151');
  hdr.setFontColor('#ffffff');
  hdr.setFontWeight('bold');
  hdr.setFontSize(9);
  hdr.setHorizontalAlignment('center');
  hdr.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 22);

  // ── Bed numbers ──
  for (let b = 0; b < n; b++) {
    const c = sheet.getRange(startRow + 1, startCol + b);
    c.setValue(b + 1);
    c.setBackground('#f1f5f9');
    c.setFontSize(8);
    c.setHorizontalAlignment('center');
    c.setFontColor('#374151');
    sheet.setColumnWidth(startCol + b, colW);
  }
  sheet.setRowHeight(startRow + 1, 18);

  // ── Bed content ──
  for (let b = 0; b < n; b++) {
    const bedInfo = (field.bedData || []).find(bd => bd.bedNum === b + 1);
    const cell    = sheet.getRange(startRow + 2, startCol + b);
    const segs    = bedInfo ? (bedInfo.segments || []) : [];

    if (segs.length > 0) {
      const bg = segs[0].color || '#86efac';
      cell.setBackground(bg);

      const lines = segs.map(s => {
        let t = s.crop;
        if (s.wave && s.wave > 1) t += ' מ' + s.wave;
        if (s.arrivalDate)        t += '\n' + s.arrivalDate;
        if (s.seq)                t += '\n(חליפה)';
        return t;
      });
      cell.setValue(lines.join('\n─\n'));
      cell.setWrap(true);
      cell.setFontSize(7);
      cell.setHorizontalAlignment('center');
      cell.setVerticalAlignment('middle');

      // auto text color based on background luminance
      const hex = bg.replace('#', '');
      const lum = (0.299 * parseInt(hex.substr(0,2),16) +
                   0.587 * parseInt(hex.substr(2,2),16) +
                   0.114 * parseInt(hex.substr(4,2),16)) / 255;
      cell.setFontColor(lum > 0.55 ? '#111827' : '#ffffff');
    } else {
      cell.setBackground('#f9fafb');
    }
  }
  sheet.setRowHeight(startRow + 2, 46);
}

// ─── עזרים ───────────────────────────────────────────────────────
function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setBackground('#1b4332');
    r.setFontColor('#ffffff');
    r.setFontWeight('bold');
    sheet.setRightToLeft(true);
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.setColumnWidth(i, 130);
  }
  return sheet;
}

function jsonOk(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:'ok', message: msg||'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── לוח חגים וחופשות ישראל תשפ"ז (2026-2027) ──────────────────
// מקור: לוח חגים רשמי תשפ"ז
// type: 'holiday' = חג (שורה כתומה)
//       'vacation' = חופשת לימודים (שורה סגולה) — כולל תאריך פתיחה מחדש
// endDate: תאריך אחרון של החופשה (כולל)
// resumeDate: יום ראשון ללימודים אחרי החופשה
var EVENTS_IL_2026_2027 = [

  // ══ ראש השנה ══
  { date:'2026-09-11', label:'ערב ראש השנה', type:'holiday' },
  { date:'2026-09-12', label:'ראש השנה א׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-13', label:'ראש השנה ב׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-11', endDate:'2026-09-13', label:'חופשת ראש השנה', resumeDate:'14/9/2026', type:'vacation' },

  // ══ יום הכיפורים ══
  { date:'2026-09-20', label:'ערב יום הכיפורים', type:'holiday' },
  { date:'2026-09-21', label:'יום הכיפורים', type:'holiday' },
  { date:'2026-09-20', endDate:'2026-09-21', label:'יום הכיפורים — אין לימודים', resumeDate:'22/9/2026', type:'vacation' },

  // ══ ימי חופשה בין יוה"כ לסוכות ══
  { date:'2026-09-22', endDate:'2026-09-24', label:'ימי חופשה בין יוה"כ לסוכות (שלישי–חמישי)', resumeDate:'25/9/2026', type:'vacation' },

  // ══ חג הסוכות ══
  { date:'2026-09-25', label:'ערב חג הסוכות', type:'holiday' },
  { date:'2026-09-26', label:'סוכות א׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-27', label:'סוכות ב׳ (חו"מ)', type:'holiday' },
  { date:'2026-10-01', label:'הושענא רבה', type:'holiday' },
  { date:'2026-10-02', label:'שמיני עצרת', type:'holiday' },
  { date:'2026-10-03', label:'שמחת תורה', type:'holiday' },
  { date:'2026-09-25', endDate:'2026-10-03', label:'חופשת סוכות ושמחת תורה', resumeDate:'4/10/2026', type:'vacation' },

  // ══ חג הסיגד ══
  { date:'2026-11-09', label:'חג הסיגד (יום לימודים — אירועים)', type:'holiday' },

  // ══ חנוכה ══
  { date:'2026-12-06', label:'חנוכה א׳ (כ"ו כסלו)', type:'holiday' },
  { date:'2026-12-12', label:'חנוכה ז׳ (ב׳ טבת)', type:'holiday' },
  { date:'2026-12-06', endDate:'2026-12-12', label:'חופשת חנוכה', resumeDate:'13/12/2026', type:'vacation' },

  // ══ ט"ו בשבט ══
  { date:'2027-01-23', label:'ט"ו בשבט (שבת)', type:'holiday' },

  // ══ פורים ══
  { date:'2027-03-23', label:'פורים (י"ד אדר ב׳)', type:'holiday' },
  { date:'2027-03-24', label:'שושן פורים (ט"ו אדר ב׳)', type:'holiday' },
  { date:'2027-03-22', endDate:'2027-03-24', label:'חופשת פורים', resumeDate:'25/3/2027', type:'vacation' },

  // ══ פסח ══
  { date:'2027-04-13', label:'ערב פסח (י"ג ניסן)', type:'holiday' },
  { date:'2027-04-14', label:'פסח א׳', type:'holiday' },
  { date:'2027-04-15', label:'פסח ב׳', type:'holiday' },
  { date:'2027-04-20', label:'פסח ז׳ (חול המועד אחרון)', type:'holiday' },
  { date:'2027-04-21', label:'פסח ז׳ / שביעי של פסח', type:'holiday' },
  { date:'2027-04-22', label:'אסרו חג פסח (יום לימודים)', type:'holiday' },
  { date:'2027-04-13', endDate:'2027-04-28', label:'חופשת פסח', resumeDate:'29/4/2027', type:'vacation' },

  // ══ יום הזיכרון / יום העצמאות ══
  { date:'2027-05-11', label:'יום הזיכרון', type:'holiday' },
  { date:'2027-05-12', label:'יום העצמאות (ה׳ אייר) — אין לימודים', type:'holiday' },

  // ══ ל"ג בעומר ══
  { date:'2027-05-25', label:'ל"ג בעומר (יום לימודים)', type:'holiday' },

  // ══ יום ירושלים ══
  { date:'2027-06-04', label:'יום ירושלים (כ"ח אייר, שישי)', type:'holiday' },

  // ══ שבועות ══
  { date:'2027-06-10', label:'ערב שבועות / שבועות א׳', type:'holiday' },
  { date:'2027-06-11', label:'שבועות ב׳ (שישי)', type:'holiday' },
  { date:'2027-06-10', endDate:'2027-06-12', label:'חופשת שבועות', resumeDate:'13/6/2027', type:'vacation' },
];

// ─── לוח שתילה שבועי ─────────────────────────────────────────────
function writeWeeklySchedule(ss, d) {
  var tabName = 'לוח שתילה ' + (d.season || '');
  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);
  var sh = ss.insertSheet(tabName);
  sh.setRightToLeft(true);

  var headers = ['#', 'תאריך', 'יום', 'חג / חופשה', 'ירק / אירוע', 'שדה', 'משתל', 'כמות שתילים', 'הערות'];
  var DAY_HE  = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];

  var timeline = [];

  var weekNum = 0;
  (d.weeks || []).forEach(function(week) {
    weekNum++;
    var items = (week.items || []).slice();
    items.sort(function(a,b){ return a.crop.localeCompare(b.crop,'he'); });
    timeline.push({ sortKey: week.date, kind:'planting', week: week, items: items, weekNum: weekNum });
  });

  var firstDate = d.firstDate || '2000-01-01';
  var lastDate  = d.lastDate  || '2099-12-31';
  EVENTS_IL_2026_2027.forEach(function(ev) {
    if (ev.date > lastDate || (ev.endDate || ev.date) < firstDate) return;
    timeline.push({ sortKey: ev.date, kind: ev.type, ev: ev });
  });

  timeline.sort(function(a,b){ return a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0; });

  var rows      = [];
  var rowMeta   = [];
  var rowIndex  = 0;

  var shownEvents = {};

  timeline.forEach(function(entry) {
    if (entry.kind === 'planting') {
      var week    = entry.week;
      var dateObj = new Date(week.date + 'T00:00:00');
      var dateStr = _fmtDate(dateObj);
      var dayStr  = DAY_HE[dateObj.getDay()];
      var isHol   = week.holidays && week.holidays.length > 0;
      var holNote = isHol ? (week.holidays || []).join(' | ') : '';

      entry.items.forEach(function(item, i) {
        rows.push([
          i === 0 ? entry.weekNum : '',
          i === 0 ? dateStr       : '',
          i === 0 ? dayStr        : '',
          i === 0 ? holNote       : '',
          item.crop,
          item.field || '',
          'משתל ' + item.wave + '/' + item.totalWaves,
          item.plants,
          ''
        ]);
        rowMeta.push({
          bg: isHol ? '#fff3cd' : (entry.weekNum % 2 === 0 ? '#f0fdf4' : '#ffffff'),
          bold: false, italic: false,
          borderTop: i === 0
        });
      });

      if (!entry.items.length) {
        rows.push([entry.weekNum, dateStr, dayStr, holNote, '(אין שתילות)', '', '', '', '']);
        rowMeta.push({ bg: '#f8fafc', bold: false, italic: true, borderTop: true });
      }

      rows.push(['','','','','','','','','']);
      rowMeta.push({ bg: '#e2e8f0', bold: false, italic: false, borderTop: false });

    } else if (entry.kind === 'holiday') {
      var evKey = entry.ev.date + '_' + entry.ev.label;
      if (shownEvents[evKey]) return;
      shownEvents[evKey] = true;
      var edo  = new Date(entry.ev.date + 'T00:00:00');
      rows.push(['', _fmtDate(edo), DAY_HE[edo.getDay()], entry.ev.label, '🕍 חג', '', '', '', '']);
      rowMeta.push({ bg: '#ffedd5', bold: true, italic: false, borderTop: true });

    } else if (entry.kind === 'vacation') {
      var evKey2 = entry.ev.date + '_' + entry.ev.label;
      if (shownEvents[evKey2]) return;
      shownEvents[evKey2] = true;
      var vs  = new Date(entry.ev.date    + 'T00:00:00');
      var ve  = entry.ev.endDate ? new Date(entry.ev.endDate + 'T00:00:00') : vs;
      var rangeStr = _fmtDate(vs) + (entry.ev.endDate ? ' — ' + _fmtDate(ve) : '');
      rows.push(['', rangeStr, '', entry.ev.label, '📚 חופשת לימודים', '', '', '', 'חופשת מוסדות לימוד (בקירוב)']);
      rowMeta.push({ bg: '#ede9fe', bold: true, italic: false, borderTop: true });
    }
  });

  if (!rows.length) return;

  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1e3a2f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
  sh.setFrozenRows(1);

  var dataRange = sh.getRange(2, 1, rows.length, headers.length);
  dataRange.setValues(rows);

  for (var r = 0; r < rows.length; r++) {
    var m  = rowMeta[r];
    var rg = sh.getRange(r + 2, 1, 1, headers.length);
    rg.setBackground(m.bg);
    if (m.bold)   rg.setFontWeight('bold');
    if (m.italic) rg.setFontStyle('italic');
    if (m.borderTop) {
      rg.setBorder(true, false, false, false, false, false, '#9ca3af', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  sh.setColumnWidth(1, 35);
  sh.setColumnWidth(2, 105);
  sh.setColumnWidth(3, 55);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 155);
  sh.setColumnWidth(6, 80);
  sh.setColumnWidth(7, 85);
  sh.setColumnWidth(8, 100);
  sh.setColumnWidth(9, 200);

  var noteRow = rows.length + 3;
  sh.getRange(noteRow, 1, 1, headers.length).merge()
    .setValue('צהוב = שתילה בשבוע חג  |  כתום = חג  |  סגול = חופשת לימודים  |  * תאריכים בקירוב — יש לאמת')
    .setFontColor('#6b7280').setFontStyle('italic').setFontSize(8).setHorizontalAlignment('right');

  sh.insertRowBefore(1);
  sh.getRange(1, 1, 1, headers.length).merge()
    .setValue('לוח שתילה — ' + (d.season || '') + ' | שתילות: ימי שני')
    .setBackground('#14532d').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  sh.setFrozenRows(2);

  ss.setActiveSheet(sh);
}

function _fmtDate(d) {
  return d.getDate() + '/' + (d.getMonth()+1) + '/' + d.getFullYear();
}

function jsonError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
