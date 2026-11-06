// Macro Tracker web app (Google Apps Script)
// Reads and writes the "Macro Tracker" Google Sheet.

const SHEET_ID = '1mXTi8Cc6HF7jy1iWR4cF9vDEgRqN_AGrCbX9eyyGnqA';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Macros')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover');
}

function book_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// Returns settings, today's date, and every Log entry.
function getData() {
  const ss = book_();
  const tz = ss.getSpreadsheetTimeZone();
  const stats = ss.getSheetByName('Stats');
  const settings = {
    target: Number(stats.getRange('B4').getValue()) || 1800,
    protein: Number(stats.getRange('B7').getValue()) || 0,
    maintenance: Number(stats.getRange('B8').getValue()) || 0
  };

  const log = ss.getSheetByName('Log');
  const last = log.getLastRow();
  const entries = [];
  if (last >= 2) {
    const range = log.getRange(2, 1, last - 1, 5);
    const vals = range.getValues();
    const disp = range.getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (v[0] === '' && v[2] === '') continue;
      const date = v[0] instanceof Date
        ? Utilities.formatDate(v[0], tz, 'yyyy-MM-dd')
        : String(disp[i][0]);
      entries.push({
        row: i + 2,
        date: date,
        time: disp[i][1],
        food: String(v[2]),
        kcal: Number(v[3]) || 0,
        protein: Number(v[4]) || 0
      });
    }
  }

  const weightSheet = ss.getSheetByName('Weight');
  const weights = readWeightRows_(weightSheet, tz)
    .filter(w => typeof w.lb === 'number' && w.lb > 0)
    .sort((a, b) => a.date.localeCompare(b.date) || a.row - b.row);

  // Read-only formula block: F5 food days, F7 weigh-ins, F8 trend, F9 adaptive,
  // F10 status, F11 manual maintenance, F12 maintenance in use.
  const block = weightSheet.getRange('F5:F12').getValues().map(r => r[0]);
  const numOrNull = v => (typeof v === 'number' ? v : null);
  const tdee = {
    foodDays: Number(block[0]) || 0,
    weighIns: Number(block[2]) || 0,
    trend: numOrNull(block[3]),
    adaptive: numOrNull(block[4]),
    status: String(block[5]),
    manual: numOrNull(block[6]),
    inUse: numOrNull(block[7])
  };

  return {
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    settings: settings,
    entries: entries,
    weights: weights,
    tdee: tdee
  };
}

// Normalizes a Weight!A cell to 'yyyy-MM-dd', or '' if it doesn't hold a date.
function weightDate_(value, display, tz) {
  const s = value instanceof Date
    ? Utilities.formatDate(value, tz, 'yyyy-MM-dd')
    : String(display).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

// Every Weight row holding a date, in sheet order. lb is whatever B holds.
function readWeightRows_(sheet, tz) {
  const last = lastWeightRow_(sheet);
  if (last < 2) return [];
  const range = sheet.getRange(2, 1, last - 1, 2);
  const vals = range.getValues();
  const disp = range.getDisplayValues();
  const rows = [];
  for (let i = 0; i < vals.length; i++) {
    const date = weightDate_(vals[i][0], disp[i][0], tz);
    if (date) rows.push({ row: i + 2, date: date, lb: vals[i][1] });
  }
  return rows;
}

// Last row with anything in A:B. sheet.getLastRow() isn't enough on its own
// because the E:F block also counts towards it.
function lastWeightRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' || vals[i][1] !== '') return i + 2;
  }
  return 1;
}

// Logs a weigh-in. Overwrites the row for that date if there is one, else appends.
function addWeight(e) {
  const lb = Math.round(Number(e && e.lb) * 10) / 10;
  if (!isFinite(lb) || lb < 50 || lb > 500) throw new Error('Enter a weight from 50 to 500 lb.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = book_();
    const tz = ss.getSpreadsheetTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const date = (e && e.date) ? String(e.date) : today;
    const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const check = m && new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (!m || check.getUTCMonth() !== +m[2] - 1 || check.getUTCDate() !== +m[3]) throw new Error('Enter a valid date.');
    if (date > today) throw new Error('The date cannot be in the future.');

    const sheet = ss.getSheetByName('Weight');
    const existing = readWeightRows_(sheet, tz).filter(w => w.date === date).pop();
    if (existing) {
      sheet.getRange(existing.row, 2).setNumberFormat('0.0').setValue(lb);
    } else {
      const row = lastWeightRow_(sheet) + 1;
      if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
      const r = sheet.getRange(row, 1, 1, 2);
      r.setNumberFormats([['yyyy-mm-dd', '0.0']]);
      r.setValues([[date, lb]]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// Removes one weigh-in, only if the row still holds the expected date.
// Never use deleteRow() here: rows on this tab are shared with the E:F formula
// block, so only A:B may change. Rows below are moved up by value so that the
// open-ended $A$2:$A references in those formulas are left alone.
function deleteWeight(row, date) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = book_();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName('Weight');
    const last = lastWeightRow_(sheet);
    const cell = (row >= 2 && row <= last) ? sheet.getRange(row, 1) : null;
    if (!cell || weightDate_(cell.getValue(), cell.getDisplayValue(), tz) !== date) {
      throw new Error('The log changed. Refresh and try again.');
    }
    if (row < last) {
      const below = sheet.getRange(row + 1, 1, last - row, 2).getValues();
      sheet.getRange(row, 1, below.length, 2).setValues(below);
    }
    sheet.getRange(last, 1, 1, 2).clearContent();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// Appends one row to the Log tab: Date | Time | Food | Calories | Protein
function addEntry(e) {
  const food = String((e && e.food) || '').trim();
  const kcal = Math.round(Number(e && e.kcal));
  const protein = Math.round(Number(e && e.protein) || 0);
  if (!food) throw new Error('Enter a food name.');
  if (!isFinite(kcal) || kcal < 0) throw new Error('Enter calories.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = book_();
    const tz = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const log = ss.getSheetByName('Log');
    const row = log.getLastRow() + 1;
    const r = log.getRange(row, 1, 1, 5);
    r.setNumberFormats([['yyyy-mm-dd', 'hh:mm', '@', '0', '0']]);
    r.setValues([[
      Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(now, tz, 'HH:mm'),
      food,
      kcal,
      protein
    ]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// Removes one Log row, only if it still holds the expected food.
function deleteEntry(row, food) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const log = book_().getSheetByName('Log');
    if (row < 2 || String(log.getRange(row, 3).getValue()) !== food) {
      throw new Error('The log changed. Refresh and try again.');
    }
    log.deleteRow(row);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getData();
}