
const CFG = {
  sheetName: 'Timesheet',              // tab name — never "whatever is active"
  firstDataRow: 10,                    // first day row
  lastDataRow: 26,                     // last spare day row
  cDate: 3,                            // C = day of month
  cIn1: 4,                             // D = Time In
  cOut1: 5,                            // E = Time Out
  cIn2: 7,                             // G = Time In  (2nd shift)
  cOut2: 8,                            // H = Time Out (2nd shift)
  timeFormat: 'h:mm am/pm',
  roundTo15: true,                     // sheet's rule: nearest 15 minutes
  buttonSlots: ['F5', 'L2', 'M2', 'N2', 'K4'],  // first free one wins
  buttonLabel: 'CLOCK IN / OUT',
  statusCell: 'F6'                     // visible feedback (mobile app has no toasts)
};

const BUTTON_NOTE = 'TIMECLOCK-BUTTON';

/* --------------------------------- setup ---------------------------------- */

/** Desktop menu + auto-heal: puts the checkbox back if it ever disappears. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Time Clock')
      .addItem('Clock In / Out', 'clockInOut')
      .addSeparator()
      .addItem('Undo last punch', 'undoLastPunch')
      .addItem('Re-create checkbox', 'setupTimeClock')
      .addToUi();
  } catch (e) { /* mobile / no UI — not an error */ }
  try { ensureButton_(); } catch (e) { /* never block opening the file */ }
}

/** Run once after pasting. Reports exactly where the checkbox landed. */
function setupTimeClock() {
  const a1 = ensureButton_();
  status_('Checkbox ready at ' + a1 + ' on the ' + CFG.sheetName + ' tab.');
  SpreadsheetApp.getUi().alert(
    'Clock button is at ' + a1 + ' on the "' + CFG.sheetName + '" tab.\n\n' +
    'Click it to punch the clock. Status shows in ' + CFG.statusCell + '.');
}

/** Finds or creates the checkbox cell. Throws if the tab is wrong or no slot is free. */
function ensureButton_() {
  const existing = findButton_();
  if (existing) return existing;

  const sh = sheet_();
  const free = CFG.buttonSlots.filter(function (a1) {
    const r = sh.getRange(a1);
    return !r.getMergedRanges().length && r.getValue() === '' && !r.getNote();
  });
  if (!free.length) {
    throw new Error('No free cell among ' + CFG.buttonSlots.join(', ') + ' on tab ' +
      sh.getName() + ' — free one up, or edit CFG.buttonSlots.');
  }
  const slot = free[0];

  const cell = sh.getRange(slot);
  cell.insertCheckboxes().setValue(false);
  cell.setNote(BUTTON_NOTE + ': click to punch the clock. Fills the next empty slot on today\'s row.');
  const label = cell.offset(0, 1);
  if (label.getValue() === '') label.setValue(CFG.buttonLabel).setFontWeight('bold');

  if (String(sh.getRange(slot).getNote()).indexOf(BUTTON_NOTE) !== 0) {
    throw new Error('Checkbox write failed at ' + slot + ' — cell may be protected or merged.');
  }
  return slot;
}

/** A1 of the existing checkbox, or '' if it isn't there. */
function findButton_() {
  const sh = sheet_();
  for (const a1 of CFG.buttonSlots) {
    const note = sh.getRange(a1).getNote();
    if (note && String(note).indexOf(BUTTON_NOTE) === 0) return a1;
  }
  return '';
}

/* -------------------------------- triggers -------------------------------- */

/** Checkbox tap (mobile + desktop). Re-arms itself so it stays clickable forever. */
function onEdit(e) {
  try {
    if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    if (e.range.getSheet().getName() !== sheet_().getName()) return;
    const a1 = findButton_();
    if (!a1 || e.range.getA1Notation() !== a1) return;
    if (String(e.value).toUpperCase() !== 'TRUE') return;
    e.range.setValue(false);       // programmatic writes don't re-fire onEdit
    punch_();
  } catch (err) {
    status_('Error: ' + err.message);
  }
}

/** Menu item and drawing-button entry point. */
function clockInOut() {
  punch_();
}

/** Clears the last filled punch on today's row. */
function undoLastPunch() {
  const sh = sheet_();
  const row = todaysRow_(sh);
  if (!row) return status_('Today is not in the timesheet rows.');

  for (const col of [CFG.cOut2, CFG.cIn2, CFG.cOut1, CFG.cIn1]) {
    const cell = sh.getRange(row, col);
    if (cell.getValue() !== '' && cell.getValue() !== null) {
      const shown = cell.getDisplayValue();
      cell.clearContent();
      return status_('Cleared ' + shown + ' on row ' + row + '.');
    }
  }
  status_('Nothing to undo on row ' + row + '.');
}

/* --------------------------------- engine --------------------------------- */

function punch_() {
  const sh = sheet_();
  const row = todaysRow_(sh);
  if (!row) {
    return status_('Today is not in the timesheet rows — update the pay-period dates.');
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
      return status_(slot[1] + ' punched at ' + cell.getDisplayValue() + ' (row ' + row + ').');
    }
  }
  status_('All four punches on row ' + row + ' are full — use Undo last punch.');
}

/* --------------------------------- helpers -------------------------------- */

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

/** The named tab, or a loud failure — never silently "whatever is active". */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = CFG.sheetName ? ss.getSheetByName(CFG.sheetName) : ss.getActiveSheet();
  if (!sh) {
    throw new Error('Tab "' + CFG.sheetName + '" not found. Tabs are: ' +
      ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  }
  return sh;
}

/** Writes status where the user can see it, then toasts on desktop. */
function status_(msg) {
  try { sheet_().getRange(CFG.statusCell).setValue(msg); } catch (e) {}
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Time Clock', 5); } catch (e) {}
  Logger.log(msg);
}
