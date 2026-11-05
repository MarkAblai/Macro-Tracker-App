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

  return {
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    settings: settings,
    entries: entries
  };
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