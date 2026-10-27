
const CFG = {
  sheetName: 'Timesheet',   // '' = active sheet
  firstDataRow: 10,         // first day row
  lastDataRow: 26,          // last spare day row
  cDate: 3,                 // C = day of month
  cIn1: 4,                  // D = Time In
  cOut1: 5,                 // E = Time Out
  cIn2: 7,                  // G = Time In  (2nd shift)
  cOut2: 8,                 // H = Time Out (2nd shift)
  timeFormat: 'h:mm am/pm',
  roundTo15: true           // sheet's rule: nearest 15 minutes
};

/** Menu — always available, survives a deleted drawing. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Time Clock')
    .addItem('Clock In / Out', 'clockInOut')
    .addSeparator()
    .addItem('Undo last punch', 'undoLastPunch')
    .addToUi();
}

/** Assign this to the drawing button. Fills the next empty slot today. */
function clockInOut() {
  const sh = sheet_();
  const row = todaysRow_(sh);
  if (!row) {
    return toast_('Today is not in the timesheet rows. Update the pay-period dates first.');
  }

  const slots = [
    [CFG.cIn1,  'Time In'],
    [CFG.cOut1, 'Time Out'],
    [CFG.cIn2,  'Time In (2nd)'],
    [CFG.cOut2, 'Time Out (2nd)']
  ];

  for (const slot of slots) {
    const cell = sh.getRange(row, slot[0]);
    const v = cell.getValue();
    if (v === '' || v === null) {
      cell.setValue(nowFraction_()).setNumberFormat(CFG.timeFormat);
      return toast_(slot[1] + ' punched at ' + cell.getDisplayValue() + ' (row ' + row + ').');
    }
  }
  toast_('All four punches on row ' + row + ' are full. Use Undo last punch to correct one.');
}

/** Clears the last filled punch on today's row. */
function undoLastPunch() {
  const sh = sheet_();
  const row = todaysRow_(sh);
  if (!row) return toast_('Today is not in the timesheet rows.');

  for (const col of [CFG.cOut2, CFG.cIn2, CFG.cOut1, CFG.cIn1]) {
    const cell = sh.getRange(row, col);
    if (cell.getValue() !== '' && cell.getValue() !== null) {
      const shown = cell.getDisplayValue();
      cell.clearContent();
      return toast_('Cleared ' + shown + ' on row ' + row + '.');
    }
  }
  toast_('Nothing to undo on row ' + row + '.');
}

/* --------------------------------- helpers --------------------------------- */

/** Today's row number, matched on the day-of-month in column C. 0 if absent. */
function todaysRow_(sh) {
  const day = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd'));
  const n = CFG.lastDataRow - CFG.firstDataRow + 1;
  const dates = sh.getRange(CFG.firstDataRow, CFG.cDate, n, 1).getValues();

  for (let i = 0; i < n; i++) {
    if (dayMatches_(dates[i][0], day)) return CFG.firstDataRow + i;
  }
  return 0;
}

function dayMatches_(v, day) {
  if (v === '' || v === null) return false;
  if (v instanceof Date) return v.getDate() === day;
  const m = String(v).match(/\d{1,2}/);
  return !!m && Number(m[0]) === day;
}

/** Now as a day fraction (what a typed "1:00 pm" is), rounded to 15 min. */
function nowFraction_() {
  const d = new Date();
  let mins = d.getHours() * 60 + d.getMinutes();
  if (CFG.roundTo15) {
    mins = Math.round(mins / 15) * 15;
    if (mins >= 1440) mins = 1425;   // never round up into tomorrow
  }
  return mins / 1440;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return (CFG.sheetName && ss.getSheetByName(CFG.sheetName)) || ss.getActiveSheet();
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Time Clock', 5); } catch (e) {}
}
