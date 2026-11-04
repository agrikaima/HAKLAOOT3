// ────────────────────────────────────────────────────────────────
// יומן שדה — Google Apps Script  (גרסה מתוקנת)
//
// מה תוקן בגרסה זו:
//   ✅ היו שתי פונקציות doPost — השנייה דרסה את הראשונה, ולכן
//      שתילות/משימות/רכש/הזמנות/רוטציה/קטלוג/מפת שדה/לוח שתילה
//      לא נכתבו לגיליון כלל. עכשיו יש doPost אחת שמטפלת בהכל.
//   ✅ saveFinanceData חיפשה שדות בשמות vat / total — האפליקציה
//      שולחת vatEnabled / vatAmount / totalWithVat. תוקן.
//   ✅ נוספה תמיכה בלשונית מלאי   (_type: 'inventory')
//   ✅ נוספה תמיכה בלשונית ספקים  (_type: 'suppliers')
//   ✅ הסנכרון האוטומטי מעדכן גם מלאי וגם ספקים
//
// הוראות פריסה:
//   1. הדבק קוד זה במקום הקוד הקיים ב-Apps Script
//   2. שמור (Ctrl+S)
//   3. פרוס → ניהול פריסות → עריכה → גרסה חדשה → פרוס
//      ● הרץ כ: אני (Me)
//      ● מי יכול לגשת: כל אחד (Anyone)  ← חשוב!
//   4. ה-URL נשאר אותו URL — אין צורך לעדכן באפליקציה
// ────────────────────────────────────────────────────────────────

// ═══════════════════ נקודת כניסה יחידה ═══════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    // ── סנכרון אוטומטי מלא מהאפליקציה ──
      if (data.action === 'saveIrrigation') {
      saveIrrigation(ss, data);
      return jsonOk('נשמר');
    }
        if (data.action === 'saveSeasonPlan') {
      const n = saveSeasonPlan(ss, data);
      return jsonOk(n + ' שורות');
    }
    if (data.action === 'syncData') {
      const dd = data.data || {};
      saveRawSnapshot(ss, dd, data.by);
      if (dd.financeData) saveFinanceData(dd.financeData);
      if (dd.inventory)   snapshotInventory(ss, dd.inventory);
      if (dd.suppliers)   snapshotSuppliers(ss, dd.suppliers, dd.procurements || []);
      if (dd.fieldOps)    snapshotFieldOps(ss, dd.fieldOps);
      if (dd.catalog)     snapshotProductCatalog(ss, dd.catalog);
      return jsonOk('synced');
    }

    // ── שליחות ממוקדות (כפתורי "שלח ל-Sheets") ──
    switch (data._type) {
      case 'task':           writeTask(ss, data);           break;
      case 'procurement':    writeProcurement(ss, data);    break;
      case 'orders':         writeOrders(ss, data);         break;
      case 'rotation':       writeRotation(ss, data);       break;
      case 'catalog':        writeCatalog(ss, data);        break;
      case 'fieldmap':       writeFieldMap(ss, data);       break;
      case 'weeklySchedule': writeWeeklySchedule(ss, data); break;
      case 'inventory':      snapshotInventory(ss, data.rows, true);      break;
      case 'suppliers':      snapshotSuppliers(ss, data.rows, [], true);  break;
      case 'fieldops':       snapshotFieldOps(ss, data.rows, true);       break;
      case 'productcatalog': snapshotProductCatalog(ss, data.rows, true);  break;
      default:               writePlanting(ss, data);       break;
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
  const p = (e && e.parameter) || {};
  if (p.action === 'listVersions') {
    return ContentService
      .createTextOutput(JSON.stringify({ status:'ok', versions: listSnapshots() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (p.action === 'getVersion') {
    return ContentService
      .createTextOutput(JSON.stringify({ status:'ok', data: readSnapshot(p.row) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
    if (p.action === 'getHarvest') {
    try {
      return ContentService
        .createTextOutput(JSON.stringify({ status:'ok', harvest: readHarvest() }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return jsonError(err.message);
    }
  }
    if (p.action === 'getIrrigation') {
    return ContentService
      .createTextOutput(JSON.stringify({ status:'ok', log: readIrrigation() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (p.action === 'getData') {
    try {
      return ContentService
        .createTextOutput(JSON.stringify({ status:'ok', data: readRawSnapshot() }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return jsonError(err.message);
    }
  }
  return jsonOk('יומן שדה פעיל');
}

// ═══════════════════ מלאי ═══════════════════
const INVENTORY_HEADERS = [
  'שם פריט','קטגוריה','כמות','יחידה','סף התראה','מלאי נמוך',
  'מיקום','מחיר ליחידה (₪)','שווי (₪)','הערות','עודכן'
];

function snapshotInventory(ss, items, force) {
  items = items || [];
  if (!items.length && !force) return;
  const sheet = getOrCreate(ss, 'מלאי', INVENTORY_HEADERS);
  if (!items.length) return;

  // ניקוי הנתונים הישנים (משאיר את שורת הכותרות)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, INVENTORY_HEADERS.length).clear();
  }

  const rows = items.map(function(i) {
    const qty   = Number(i.qty)   || 0;
    const min   = Number(i.min)   || 0;
    const price = Number(i.price) || 0;
    const isLow = (i.low === 'כן') || (min > 0 && qty <= min);
    return [
      i.name || '', i.cat || '', qty, i.unit || '', min,
      isLow ? 'כן' : '',
      i.loc || '', price,
      i.value !== undefined ? i.value : Math.round(qty * price),
      i.notes || '',
      _fmtStamp(i.updatedAt)
    ];
  });

  // מיון: מלאי נמוך קודם, אחר כך לפי קטגוריה ושם
  rows.sort(function(a, b) {
    if (a[5] !== b[5]) return a[5] ? -1 : 1;
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]), 'he');
    return String(a[0]).localeCompare(String(b[0]), 'he');
  });

  sheet.getRange(2, 1, rows.length, INVENTORY_HEADERS.length).setValues(rows);

  // צביעה אדומה לשורות מלאי נמוך
  for (let r = 0; r < rows.length; r++) {
    const range = sheet.getRange(r + 2, 1, 1, INVENTORY_HEADERS.length);
    if (rows[r][5] === 'כן') {
      range.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
    } else {
      range.setBackground(r % 2 === 0 ? '#ffffff' : '#f6fdf8')
           .setFontColor('#111827').setFontWeight('normal');
    }
  }

  // שורת סיכום
  const totalValue = rows.reduce(function(s, r) { return s + (Number(r[8]) || 0); }, 0);
  const lowCount   = rows.filter(function(r) { return r[5] === 'כן'; }).length;
  const sumRow = sheet.getRange(rows.length + 2, 1, 1, INVENTORY_HEADERS.length);
  sumRow.setValues([[
    'סה"כ', rows.length + ' פריטים', '', '', '',
    lowCount ? lowCount + ' נמוכים' : '', '', '', totalValue, '', ''
  ]]);
  sumRow.setBackground('#d1fae5').setFontWeight('bold').setFontColor('#065f46');
}

// ═══════════════════ ספקים ═══════════════════
const SUPPLIER_HEADERS = [
  'שם ספק','איש קשר','טלפון','טלפון נוסף','אימייל','אתר',
  'כתובת','מה מספק','תנאי תשלום','מס׳ לקוח','רכישות','הערות'
];

function snapshotSuppliers(ss, sups, procurements, force) {
  sups = sups || [];
  if (!sups.length && !force) return;
  const sheet = getOrCreate(ss, 'ספקים', SUPPLIER_HEADERS);
  if (!sups.length) return;

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, SUPPLIER_HEADERS.length).clear();
  }

  const procs = procurements || [];
  const rows = sups.map(function(s) {
    let purchases = s.purchases;
    if (purchases === undefined) {
      purchases = procs.filter(function(p) {
        return String(p.supplier || '').trim() === String(s.name || '').trim();
      }).length;
    }
    return [
      s.name || '', s.contact || '', s.phone || '', s.phone2 || '',
      s.email || '', s.web || '', s.address || '', s.supplies || '',
      s.terms || '', s.account || '', purchases || 0, s.notes || ''
    ];
  });

  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0]), 'he'); });
  sheet.getRange(2, 1, rows.length, SUPPLIER_HEADERS.length).setValues(rows);

  for (let r = 0; r < rows.length; r++) {
    sheet.getRange(r + 2, 1, 1, SUPPLIER_HEADERS.length)
      .setBackground(r % 2 === 0 ? '#ffffff' : '#eef2ff');
  }
}


// ═══════════════════ שיגרת שדה ═══════════════════
const FIELDOPS_HEADERS = [
  'תאריך','סוג עבודה','חלקה','ערוגות','מבצע','ציוד / כלי',
  'שעות','סטטוס','בוצע בפועל','חוזר כל (ימים)','הערות'
];

function snapshotFieldOps(ss, ops, force) {
  ops = ops || [];
  if (!ops.length && !force) return;
  const sheet = getOrCreate(ss, 'שיגרת שדה', FIELDOPS_HEADERS);
  if (!ops.length) return;   // הגיליון נוצר עם כותרות, אין עדיין שורות

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, FIELDOPS_HEADERS.length).clear();
  }

  const todayISO = _todayISO();
  const rows = ops.map(function(o) {
    const status = (o.status === 'done' || o.status === 'בוצע') ? 'בוצע' : 'מתוכנן';
    return [
      o.date || '', o.type || '', o.field || 'כל השדה', o.beds || '',
      o.by || '', o.equipment || '', o.hours || '', status,
      o.doneDate || '', o.repeat || '', o.notes || ''
    ];
  });

  // מתוכננות קודם לפי תאריך עולה, אחר כך שבוצעו לפי תאריך יורד
  const planned = rows.filter(function(r){ return r[7] === 'מתוכנן'; })
                      .sort(function(a,b){ return String(a[0]).localeCompare(String(b[0])); });
  const done    = rows.filter(function(r){ return r[7] === 'בוצע'; })
                      .sort(function(a,b){ return String(b[0]).localeCompare(String(a[0])); });
  const all = planned.concat(done);

  sheet.getRange(2, 1, all.length, FIELDOPS_HEADERS.length).setValues(all);

  for (let r = 0; r < all.length; r++) {
    const range = sheet.getRange(r + 2, 1, 1, FIELDOPS_HEADERS.length);
    const isPlanned = all[r][7] === 'מתוכנן';
    const overdue   = isPlanned && String(all[r][0]) < todayISO;
    if (overdue) {
      range.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
    } else if (isPlanned) {
      range.setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('normal');
    } else {
      range.setBackground(r % 2 === 0 ? '#ffffff' : '#f6fdf8')
           .setFontColor('#111827').setFontWeight('normal');
    }
  }
}

function _todayISO() {
  const d = new Date();
  return d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
}


// ═══════════════════ קטלוג מוצרים ═══════════════════
const PRODCAT_HEADERS = [
  'שם מוצר','מק״ט','יצרן','ספק','קטגוריה','יחידה',
  'מחיר מחירון (₪)','מחיר אחרון (₪)','תאריך רכישה אחרון','רכישות','קישור','מפרט / הערות','תמונה'
];

function snapshotProductCatalog(ss, items, force) {
  items = items || [];
  if (!items.length && !force) return;
  const sheet = getOrCreate(ss, 'קטלוג מוצרים', PRODCAT_HEADERS);
  if (!items.length) return;

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, PRODCAT_HEADERS.length).clear();
  }

  const rows = items.map(function(c) {
    return [
      c.name || '', c.sku || '', c.maker || '', c.supplier || '',
      c.cat || '', c.unit || '',
      c.price || '', c.lastPrice || '', c.lastDate || '',
      c.purchases || 0, c.link || '', c.notes || '', c.hasImg || ''
    ];
  });

  rows.sort(function(a, b) {
    if (a[4] !== b[4]) return String(a[4]).localeCompare(String(b[4]), 'he');
    return String(a[0]).localeCompare(String(b[0]), 'he');
  });

  sheet.getRange(2, 1, rows.length, PRODCAT_HEADERS.length).setValues(rows);

  for (let r = 0; r < rows.length; r++) {
    const range = sheet.getRange(r + 2, 1, 1, PRODCAT_HEADERS.length);
    range.setBackground(r % 2 === 0 ? '#ffffff' : '#ecfdf5');
    // התייקרות מול המחירון — מסומן
    const listP = Number(rows[r][6]) || 0;
    const lastP = Number(rows[r][7]) || 0;
    if (listP && lastP && lastP > listP * 1.1) {
      sheet.getRange(r + 2, 8).setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
    }
  }
}

// ═══════════════════ כספים ═══════════════════
function saveFinanceData(finData) {
  if (!finData) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('כספים');
  if (!sheet) {
    sheet = ss.insertSheet('כספים');
    sheet.setRightToLeft(true);
  }
  sheet.clear();

  const headers = ['עונה', 'סה"כ עלות (₪)', 'מע"מ פעיל', 'סכום מע"מ (₪)', 'סה"כ עם מע"מ (₪)', 'עודכן'];
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  for (let i = 1; i <= headers.length; i++) sheet.setColumnWidth(i, 140);

  const rows = [];
  for (const season in finData) {
    const fd = finData[season] || {};
    // תמיכה גם בשמות השדות הישנים, למקרה של נתונים ישנים
    const vatAmount    = fd.vatAmount    !== undefined ? fd.vatAmount    : (fd.vat   || 0);
    const totalWithVat = fd.totalWithVat !== undefined ? fd.totalWithVat : (fd.total || 0);
    const vatOn        = fd.vatEnabled !== undefined ? fd.vatEnabled : (vatAmount > 0);
    rows.push([
      season,
      fd.totalCost || 0,
      vatOn ? 'כן (18%)' : 'לא',
      vatAmount,
      totalWithVat,
      _fmtStamp(fd.updatedAt)
    ]);
  }

  if (!rows.length) return;
  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0]), 'he'); });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ═══════════════════ שתילות ═══════════════════
const PLANTING_HEADERS = [
  'מזהה','סוג תיעוד','עונה','תאריך','שם חלקה','מספר ערוגה',
  'זן / סוג צמח','כמות שתילים','מספר שורות','מספר שלוחות',
  'מרחק בין צמחים (ס"מ)','זן שחוסל',
  'זן זרע','כמות זרעים','יחידת זרעים','עומק זריעה (ס"מ)','שורות זריעה','מרחק שורות (ס"מ)','שיטת זריעה',
  'חומר ריסוס','מזיק / מחלה','מינון','נפח תרסיס (ל׳)','ימי המתנה','מותר לקטיף מ-','מבצע הריסוס',
  'הערות','תאריך תיעוד'
];

const REC_TYPE_HE = {
  plant:'שתילה', seed:'זריעה', spray:'ריסוס', rest:'מנוחה', clear:'חיסול'
};

// שם הגיליון תואם לשם הלשונית באפליקציה.
// אם קיים גיליון ישן בשם 'שתילות' — הוא ישונה אוטומטית (הנתונים נשמרים).
const PLANTING_SHEET        = 'מלאכת ערוגות';
const PLANTING_SHEET_LEGACY = 'שתילות';

function ensurePlantingSheet(ss) {
  let sheet = ss.getSheetByName(PLANTING_SHEET);

  // שינוי שם לגיליון הישן, בלי לאבד נתונים
  if (!sheet) {
    const legacy = ss.getSheetByName(PLANTING_SHEET_LEGACY);
    if (legacy) { legacy.setName(PLANTING_SHEET); sheet = legacy; }
  }

  // אין גיליון בכלל — נוצר חדש עם כל הכותרות
  if (!sheet) return getOrCreate(ss, PLANTING_SHEET, PLANTING_HEADERS);

  // עדכון שורת הכותרות אם המבנה השתנה (למשל אחרי הוספת זריעה/ריסוס)
  const need    = PLANTING_HEADERS.length;
  const lastCol = sheet.getLastColumn();
  let current = [];
  if (lastCol > 0) current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const same = current.length === need &&
    current.every(function(v, i) { return String(v) === PLANTING_HEADERS[i]; });

  if (!same) {
    const r = sheet.getRange(1, 1, 1, need);
    r.setValues([PLANTING_HEADERS]);
    r.setBackground('#1b4332');
    r.setFontColor('#ffffff');
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function writePlanting(ss, d) {
  const sheet = ensurePlantingSheet(ss);
  const typeLabel = REC_TYPE_HE[d.recType] || 'שתילה';
  sheet.appendRow([
    d.id||'', typeLabel, d.season||'', d.date||'', d.field||'', d.bed||'',
    d.variety||'', d.qty||'', d.rows||'', d.shoots||'', d.distPlant||'',
    d.clearVariety||'',
    d.seedVariety||'', d.seedQty||'', d.seedUnit||'', d.seedDepth||'',
    d.seedRows||'', d.seedDist||'', d.seedMethod||'',
    d.sprayMaterial||'', d.sprayTarget||'', d.sprayDose||'', d.sprayVolume||'',
    d.sprayPHI||'', d.sprayUntil||'', d.sprayBy||'',
    d.notes||'', new Date().toLocaleString('he-IL')
  ]);

  // הדגשת שורות ריסוס — כדי שיהיה קל לאתר תקופות המתנה לקטיף
  if (d.recType === 'spray') {
    sheet.getRange(sheet.getLastRow(), 1, 1, PLANTING_HEADERS.length)
      .setBackground('#eef2ff');
  }
}

// ═══════════════════ יומן קטיף ═══════════════════
const LOG_HEADERS = [
  'מזהה רשומה','שם חלקה','ערוגה','זן','עונה',
  'תאריך קטיף','שעה','משקל קטיף (ק"ג)','הערה'
];

// ═══════════════════ משימות ═══════════════════
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

// ═══════════════════ רכש ═══════════════════
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

// ═══════════════════ מיפוי רוטציה ═══════════════════
const ROTATION_HEADERS = ['חלקה','ערוגה','אזור','ירק לעונה','אורך ערוגה (מ׳)'];

function writeRotation(ss, d) {
  const sheetName = 'מיפוי ' + (d.season || '');
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sheet = getOrCreate(ss, sheetName, ROTATION_HEADERS);
  d.rows.forEach(r => {
    sheet.appendRow([r.field, r.bed, r.zone, r.crop || '—', r.len]);
  });
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

// ═══════════════════ הזמנות ═══════════════════
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

// ═══════════════════ קטלוג ירקות ═══════════════════
const CATALOG_BASE_HEADERS = [
  'שם','עונה','שלוחות','שורות','מרחק (סמ)','גלים','כל (שבועות)','תאריך ראשון',
  'שתילים/גל','מטר/גל','סה"כ שתילים','סה"כ מטר'
];

function writeCatalog(ss, d) {
  const sheetName = 'קטלוג ירקות';
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const maxWaves = d.rows.reduce((m, r) => Math.max(m, (r.waveDates || []).length), 0);
  const dateHeaders = [];
  for (let i = 1; i <= maxWaves; i++) dateHeaders.push('תאריך משתל ' + i);
  const headers = CATALOG_BASE_HEADERS.concat(dateHeaders);

  const sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setBackground('#1b4332');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setHorizontalAlignment('center');

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

    rowIdx++;
  });

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 100);
  for (let i = 3; i <= 12; i++) sheet.setColumnWidth(i, 90);
  for (let i = 13; i <= headers.length; i++) sheet.setColumnWidth(i, 120);

  if (rowIdx > 2) {
    sheet.getRange(1, 1, rowIdx - 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  }
}

// ═══════════════════ מפת שדה ═══════════════════
function writeFieldMap(ss, d) {
  const sheetName = 'מפת שדה ' + (d.season || '');
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(sheetName);
  sheet.setRightToLeft(true);

  const ZONE_BG   = { 'עונתיים':'#2d6a4f', 'ירוקים':'#2f9e44', 'חול':'#b45309', 'מיוחד':'#7c3aed' };
  const FIELD_BG  = { 'עונתיים':'#10b981', 'ירוקים':'#22c55e', 'חול':'#d97706', 'מיוחד':'#8b5cf6' };
  const BED_COL_W = 54;
  const GAP       = 1;

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
    if (g.zone !== prevZone) {
      if (prevZone !== null) { row++; }
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
    row += 3;
  });

  sheet.setColumnWidth(1, 4);
}

function _renderBlock(sheet, field, startRow, startCol, colW, FIELD_BG) {
  const n = field.beds;

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

// ═══════════════════ לוח חגים וחופשות ═══════════════════
var EVENTS_IL_2026_2027 = [

  { date:'2026-09-11', label:'ערב ראש השנה', type:'holiday' },
  { date:'2026-09-12', label:'ראש השנה א׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-13', label:'ראש השנה ב׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-11', endDate:'2026-09-13', label:'חופשת ראש השנה', resumeDate:'14/9/2026', type:'vacation' },

  { date:'2026-09-20', label:'ערב יום הכיפורים', type:'holiday' },
  { date:'2026-09-21', label:'יום הכיפורים', type:'holiday' },
  { date:'2026-09-20', endDate:'2026-09-21', label:'יום הכיפורים — אין לימודים', resumeDate:'22/9/2026', type:'vacation' },

  { date:'2026-09-22', endDate:'2026-09-24', label:'ימי חופשה בין יוה"כ לסוכות (שלישי–חמישי)', resumeDate:'25/9/2026', type:'vacation' },

  { date:'2026-09-25', label:'ערב חג הסוכות', type:'holiday' },
  { date:'2026-09-26', label:'סוכות א׳ תשפ"ז', type:'holiday' },
  { date:'2026-09-27', label:'סוכות ב׳ (חו"מ)', type:'holiday' },
  { date:'2026-10-01', label:'הושענא רבה', type:'holiday' },
  { date:'2026-10-02', label:'שמיני עצרת', type:'holiday' },
  { date:'2026-10-03', label:'שמחת תורה', type:'holiday' },
  { date:'2026-09-25', endDate:'2026-10-03', label:'חופשת סוכות ושמחת תורה', resumeDate:'4/10/2026', type:'vacation' },

  { date:'2026-11-09', label:'חג הסיגד (יום לימודים — אירועים)', type:'holiday' },

  { date:'2026-12-06', label:'חנוכה א׳ (כ"ו כסלו)', type:'holiday' },
  { date:'2026-12-12', label:'חנוכה ז׳ (ב׳ טבת)', type:'holiday' },
  { date:'2026-12-06', endDate:'2026-12-12', label:'חופשת חנוכה', resumeDate:'13/12/2026', type:'vacation' },

  { date:'2027-01-23', label:'ט"ו בשבט (שבת)', type:'holiday' },

  { date:'2027-03-23', label:'פורים (י"ד אדר ב׳)', type:'holiday' },
  { date:'2027-03-24', label:'שושן פורים (ט"ו אדר ב׳)', type:'holiday' },
  { date:'2027-03-22', endDate:'2027-03-24', label:'חופשת פורים', resumeDate:'25/3/2027', type:'vacation' },

  { date:'2027-04-13', label:'ערב פסח (י"ג ניסן)', type:'holiday' },
  { date:'2027-04-14', label:'פסח א׳', type:'holiday' },
  { date:'2027-04-15', label:'פסח ב׳', type:'holiday' },
  { date:'2027-04-20', label:'פסח ז׳ (חול המועד אחרון)', type:'holiday' },
  { date:'2027-04-21', label:'פסח ז׳ / שביעי של פסח', type:'holiday' },
  { date:'2027-04-22', label:'אסרו חג פסח (יום לימודים)', type:'holiday' },
  { date:'2027-04-13', endDate:'2027-04-28', label:'חופשת פסח', resumeDate:'29/4/2027', type:'vacation' },

  { date:'2027-05-11', label:'יום הזיכרון', type:'holiday' },
  { date:'2027-05-12', label:'יום העצמאות (ה׳ אייר) — אין לימודים', type:'holiday' },

  { date:'2027-05-25', label:'ל"ג בעומר (יום לימודים)', type:'holiday' },

  { date:'2027-06-04', label:'יום ירושלים (כ"ח אייר, שישי)', type:'holiday' },

  { date:'2027-06-10', label:'ערב שבועות / שבועות א׳', type:'holiday' },
  { date:'2027-06-11', label:'שבועות ב׳ (שישי)', type:'holiday' },
  { date:'2027-06-10', endDate:'2027-06-12', label:'חופשת שבועות', resumeDate:'13/6/2027', type:'vacation' },
];

// ═══════════════════ לוח שתילה שבועי ═══════════════════
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

// ═══════════════════ עזרים ═══════════════════
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

function _fmtDate(d) {
  return d.getDate() + '/' + (d.getMonth()+1) + '/' + d.getFullYear();
}

function _fmtStamp(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    return _fmtDate(d) + ' ' +
      String(d.getHours()).padStart(2,'0') + ':' +
      String(d.getMinutes()).padStart(2,'0');
  } catch (e) {
    return String(iso);
  }
}

function jsonOk(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:'ok', message: msg||'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
// ═══════════════════ עותק גולמי לסנכרון בין מכשירים ═══════════════════
const RAW_SHEET = '_נתונים_גולמיים';
const CHUNK_SIZE = 45000;
const HIST_SHEET = '_היסטוריה';
const HIST_KEEP  = 20;
function saveRawSnapshot(ss, dd, who) {
  const stamp = new Date().toISOString();
  const json  = JSON.stringify({ savedAt: stamp, by: who || '', data: dd });
  let changes = [];
  try {
    const prev = readRawSnapshot();
    changes = diffData(prev && prev.data, dd);
  } catch (e) { changes = ['—']; }

  // העותק העדכני (כמו קודם)
  let sheet = ss.getSheetByName(RAW_SHEET);
  if (!sheet) { sheet = ss.insertSheet(RAW_SHEET); sheet.hideSheet(); }
  sheet.clear();
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push([json.substr(i, CHUNK_SIZE)]);
  }
  if (chunks.length) sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);

  // היסטוריה — שורה לכל סנכרון
  let hist = ss.getSheetByName(HIST_SHEET);
  if (!hist) {
    hist = ss.insertSheet(HIST_SHEET);
    hist.getRange(1, 1, 1, 5).setValues([['תאריך ושעה','מי','מה השתנה','גודל','נתונים']]);
    hist.setFrozenRows(1);
    hist.hideSheet();
  }
  hist.insertRowAfter(1);
  hist.getRange(2, 1, 1, 5).setValues([[
    _fmtStamp(stamp), who || 'לא ידוע', changes.join(' | ').substr(0, 4000),
    json.length, json.substr(0, 45000)
  ]]);

  // שמירת 20 האחרונים בלבד
  const last = hist.getLastRow();
  if (last > HIST_KEEP + 1) {
    hist.deleteRows(HIST_KEEP + 2, last - (HIST_KEEP + 1));
  }
}

function listSnapshots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName(HIST_SHEET);
  if (!hist || hist.getLastRow() < 2) return [];
  const n = hist.getLastRow() - 1;
  const vals = hist.getRange(2, 1, n, 4).getValues();
  return vals.map(function(r, i) {
    return { row: i + 2, when: r[0], by: r[1], what: r[2], size: r[3] };
  });
}

function readSnapshot(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName(HIST_SHEET);
  if (!hist) return null;
  const txt = hist.getRange(Number(row), 5).getValue();
  if (!txt) return null;
  return JSON.parse(txt);
}

function readRawSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RAW_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return null;
  const vals = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const json = vals.map(function(r) { return r[0]; }).join('');
  if (!json) return null;
  return JSON.parse(json);
}
// ═══════════════════ השוואת גרסאות ═══════════════════
const DIFF_LABELS = {
  plantings:'מלאכת ערוגות', tasks:'משימות', procurements:'רכש',
  inventory:'מלאי', suppliers:'ספקים', fieldOps:'שיגרת שדה',
  catalog:'קטלוג מוצרים', plantSpecies:'קטלוג ירקות',
  plantDefs:'תכנון עונה', financeData:'כספים', settings:'הגדרות'
};

function _label(item, key) {
  if (!item) return '';
  return item.name || item.item || item.desc || item.variety ||
         item.clearVariety || item.type || ('#' + (item.id || ''));
}

function diffData(oldD, newD) {
  if (!oldD) return ['סנכרון ראשון'];
  const notes = [];
  Object.keys(DIFF_LABELS).forEach(function(key) {
    const a = oldD[key], b = newD[key];
    if (!Array.isArray(a) || !Array.isArray(b)) {
      if (JSON.stringify(a) !== JSON.stringify(b) && b !== undefined) {
        notes.push('עודכן: ' + DIFF_LABELS[key]);
      }
      return;
    }
    const mapA = {}, mapB = {};
    a.forEach(function(x){ mapA[x.id] = x; });
    b.forEach(function(x){ mapB[x.id] = x; });

    b.forEach(function(x) {
      if (!mapA[x.id]) {
        notes.push('נוסף ב' + DIFF_LABELS[key] + ': ' + _label(x, key));
      } else if (JSON.stringify(mapA[x.id]) !== JSON.stringify(x)) {
        notes.push('עודכן ב' + DIFF_LABELS[key] + ': ' + _label(x, key));
      }
    });
    a.forEach(function(x) {
      if (!mapB[x.id]) {
        notes.push('נמחק מ' + DIFF_LABELS[key] + ': ' + _label(x, key));
      }
    });
  });
  return notes.length ? notes : ['אין שינוי בנתונים'];
}
// ═══════════════════ קטיף מטופס גוגל ═══════════════════
const HARVEST_SS_ID = '1WgwanfbRBcWMtJQVL5AfcmJD8VQUB5g0AjDxhXMRW54';
const HARVEST_TAB   = 'Form Responses 1';

function readHarvest() {
  const ss = SpreadsheetApp.openById(HARVEST_SS_ID);
  const sh = ss.getSheetByName(HARVEST_TAB) || ss.getSheets()[0];
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(function(h) { return String(h).trim(); });

  return vals.slice(1).filter(function(r) {
    return r.join('').trim() !== '';
  }).map(function(r, i) {
    const o = { _row: i + 2 };
    head.forEach(function(h, c) {
      if (h) o[h] = r[c];
    });
    if (o['Timestamp'] instanceof Date) o['Timestamp'] = o['Timestamp'].toISOString();
    if (o['תאריך קטיף'] instanceof Date) {
      const d = o['תאריך קטיף'];
      o['תאריך קטיף'] = d.getFullYear() + '-' +
        String(d.getMonth()+1).padStart(2,'0') + '-' +
        String(d.getDate()).padStart(2,'0');
    }
    return o;
  });
}

// ═══════════════════ תיעוד תכנון השקיה ═══════════════════
const IRRIG_SHEET = 'תכנון השקיה';

function saveIrrigation(ss, d) {
  const valves = d.valves || [];
  let sheet = ss.getSheetByName(IRRIG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(IRRIG_SHEET);
    sheet.setRightToLeft(true);
    const head = ['תאריך','נשמר בשעה','מי','התאדות (מ״מ)','תוספת (%)','יעד (מ״מ/יום)',
                  'שעת פתיחה','מרווח (שעות)'];
    valves.forEach(function(v) {
           head.push(v.name + ' — שעת פתיחה', v.name + ' — תוספת %', v.name + ' — רטיבות', v.name + ' — השקיות/יום',
                v.name + ' — כל כמה ימים', v.name + ' — דקות', v.name + ' — קו״ב/יום');
    });
    const r = sheet.getRange(1, 1, 1, head.length);
    r.setValues([head]);
    r.setBackground('#2c6e8f').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  const row = [d.date || '', _fmtStamp(new Date().toISOString()), d.by || '',
               d.evap || 0, d.bump || 0, d.target || 0,
               d.startTime || '', d.gapHours || 0];
  valves.forEach(function(v) {
        row.push(v.start || '', v.bump !== undefined ? v.bump : '', v.wet || '', v.perDay || '', v.everyDays || '',
             Math.round(v.minsPerRun || 0), Number((v.cubePerDay || 0).toFixed(2)));
  });

  sheet.appendRow(row);
  return sheet.getLastRow();
}

function readIrrigation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IRRIG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const vals = sheet.getDataRange().getValues();
  const head = vals[0];
  return vals.slice(1).reverse().map(function(r) {
    const o = {};
    head.forEach(function(h, i) { if (h) o[h] = r[i]; });
    return o;
  });
}
// ═══════════════════ תוכנית עונה ═══════════════════
function saveSeasonPlan(ss, d) {
  const name = 'תוכנית ' + (d.season || '');
  const existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(name);
  sheet.setRightToLeft(true);

  const head = ['שבוע','תאריך','ירק','משתל','שתילים','מיקום','אופן שתילה','דגשים'];
  const hr = sheet.getRange(1, 1, 1, head.length);
  hr.setValues([head]);
  hr.setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const rows = d.rows || [];
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, head.length).setValues(rows);

    // פס צבע לכל שבוע חדש
    let prev = null, shade = false;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] !== prev) { shade = !shade; prev = rows[i][0]; }
      if (shade) sheet.getRange(i + 2, 1, 1, head.length).setBackground('#f6fdf8');
    }
  }

  const widths = [70, 130, 170, 60, 70, 180, 190, 220];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.getRange(1, 1, Math.max(rows.length + 1, 1), head.length).setVerticalAlignment('top');
  if (rows.length) sheet.getRange(2, 6, rows.length, 3).setWrap(true);

  return rows.length;
}
