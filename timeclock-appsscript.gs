const CFG = {
  sheetName: '',            // '' = whichever sheet is active
  headerRow: 1,             // row that holds the column titles

  // Column headers the script looks for (case-insensitive, any column order).
  // First match wins. If none are found, CFG.fallback positions are used.
  hDate:  ['date', 'day'],
  hIn:    ['time in', 'in', 'clock in', 'time-in'],
  hOut:   ['time out', 'out', 'clock out', 'time-out'],
  hHours: ['hours', 'total', 'duration', 'total hours', 'hrs'],
  hNote:  ['note', 'notes', 'comment', 'comments', 'task'],

  // Used only when a header can't be found (1 = column A).
  fallback: { date: 1, in: 2, out: 3, hours: 4, note: 0 },

  dateFormat:  'yyyy-mm-dd',
  timeFormat:  'hh:mm:ss AM/PM',   // swap to 'HH:mm:ss' if you want 24h
  hoursFormat: '0.00',

  triggerCell: 'G2',        // checkbox "button" cell
  triggerLabel: 'CLOCK',    // label written next to the checkbox
  autoFillDate: true,       // put the date in even if the Date column is blank
  roundHours: false         // true = round each shift to the nearest 15 min
};

/* ============================ PUBLIC ENTRY POINTS ============================ */

/** Main function — assign this to your drawing button. Toggles in/out. */
function clockInOut() {
  punch_('auto');
}

/** Force a time-in on a fresh row (skips the toggle logic). */
function clockIn() {
  punch_('in');
}

/** Force a time-out on the last open row. */
function clockOut() {
  punch_('out');
}

/** Deletes the last row's Time Out (undo a mistaken clock-out). */
function undoLastClockOut() {
  const sh = sheet_(), c = cols_(sh), r = lastRow_(sh, c);
  if (r <= CFG.headerRow) return toast_('Nothing to undo.');
  if (!sh.getRange(r, c.in).getValue()) return toast_('Last row has no Time In.');
  sh.getRange(r, c.out).clearContent();
  if (c.hours) sh.getRange(r, c.hours).clearContent();
  toast_('Removed the clock-out on row ' + r + '.');
}

/** Run once after pasting: builds the checkbox button + the menu. */
function setupTimeClock() {
  const sh = sheet_();
  const cell = sh.getRange(CFG.triggerCell);
  cell.insertCheckboxes().setValue(false);
  const labelCell = cell.offset(0, 1);
  if (!labelCell.getValue()) labelCell.setValue(CFG.triggerLabel);
  labelCell.setFontWeight('bold');

  // Make sure the checkbox edit trigger exists (onEdit is a simple trigger,
  // no install needed — this is just a friendly nudge).
  toast_('Ready. Click the checkbox in ' + CFG.triggerCell + ' to punch in/out. ' +
         'Or use the Time Clock menu.');
}

/** Simple trigger: clicking the checkbox punches the clock and re-arms itself. */
function onEdit(e) {
  try {
    if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    const a1 = e.range.getA1Notation();
    const target = sheet_().getRange(CFG.triggerCell).getA1Notation();
    if (e.range.getSheet().getName() !== sheet_().getName() || a1 !== target) return;
    if (e.value !== 'TRUE') return;          // only on tick
    e.range.setValue(false);                 // re-arm for the next click
    punch_('auto', true);
  } catch (err) {
    toast_('Time clock error: ' + err.message);
  }
}

/** Menu so the script is usable even if the drawing/checkbox gets deleted. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Time Clock')
    .addItem('Clock In / Out (toggle)', 'clockInOut')
    .addItem('Clock In only', 'clockIn')
    .addItem('Clock Out only', 'clockOut')
    .addSeparator()
    .addItem('Undo last clock-out', 'undoLastClockOut')
    .addItem('Re-create the checkbox button', 'setupTimeClock')
    .addToUi();
}

/* ================================ ENGINE ================================ */

/**
 * mode: 'auto' (toggle) | 'in' | 'out'
 * quiet: true = no toast (used by onEdit so the checkbox feels instant)
 */
function punch_(mode, quiet) {
  const sh  = sheet_();
  const c   = cols_(sh);
  const now = new Date();
  const r   = lastRow_(sh, c);

  const openRow  = r > CFG.headerRow ? r : 0;
  const lastIn   = openRow ? sh.getRange(openRow, c.in).getValue() : '';
  const lastOut  = openRow && c.out ? sh.getRange(openRow, c.out).getValue() : '';
  const isOpen   = !!lastIn && !lastOut;

  let action = mode;
  if (mode === 'auto') action = isOpen ? 'out' : 'in';

  if (action === 'in') {
    if (isOpen && mode === 'in') {           // already punched in today
      return quiet ? 0 : toast_('Already clocked in on row ' + openRow + ' — clock out first.');
    }
    const nr = Math.max(r + 1, CFG.headerRow + 1);
    if (c.date && CFG.autoFillDate) {
      sh.getRange(nr, c.date).setValue(now).setNumberFormat(CFG.dateFormat);
    }
    const inCell = sh.getRange(nr, c.in);
    inCell.setValue(now).setNumberFormat(CFG.timeFormat);
    if (c.out) sh.getRange(nr, c.out).clearContent();
    if (c.hours) sh.getRange(nr, c.hours).clearContent();
    sh.getRange(nr, c.in).setBackground('#d9ead3');   // green = on the clock
    if (!quiet) toast_('Clocked IN at ' + fmt_(now, CFG.timeFormat) + ' (row ' + nr + ').');
    return nr;
  }

  // action === 'out'
  if (!openRow || !lastIn || lastOut) {
    return quiet ? 0 : toast_('No open time-in found — clock in first.');
  }
  const outCell = sh.getRange(openRow, c.out);
  outCell.setValue(now).setNumberFormat(CFG.timeFormat);

  if (c.hours) {
    let hrs = (now.getTime() - new Date(lastIn).getTime()) / 3600000;
    if (hrs < 0) hrs += 24;                  // crossed midnight
    if (CFG.roundHours) hrs = Math.round(hrs * 4) / 4;
    sh.getRange(openRow, c.hours).setValue(hrs).setNumberFormat(CFG.hoursFormat);
  }
  sh.getRange(openRow, c.in).setBackground(null);
  if (!quiet) toast_('Clocked OUT at ' + fmt_(now, CFG.timeFormat) + ' (row ' + openRow + ').');
  return openRow;
}

/* ================================ HELPERS ================================ */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CFG.sheetName) {
    const s = ss.getSheetByName(CFG.sheetName);
    if (s) return s;
  }
  return ss.getActiveSheet();
}

/** Resolve the Date / In / Out / Hours / Note columns once per call. */
function cols_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(CFG.headerRow, 1, 1, lastCol).getValues()[0]
                    .map(v => String(v).trim().toLowerCase());

  const find = (names, fb) => {
    for (const n of names) {
      const i = headers.indexOf(n);
      if (i !== -1) return i + 1;
    }
    return fb || 0;
  };

  return {
    date:  find(CFG.hDate,  CFG.fallback.date),
    in:    find(CFG.hIn,    CFG.fallback.in),
    out:   find(CFG.hOut,   CFG.fallback.out),
    hours: find(CFG.hHours, CFG.fallback.hours),
    note:  find(CFG.hNote,  CFG.fallback.note)
  };
}

/** Last row that actually holds data in any tracked column. */
function lastRow_(sh, c) {
  const last = sh.getLastRow();
  if (last <= CFG.headerRow) return CFG.headerRow;
  const width = Math.max(c.date || 0, c.in || 0, c.out || 0, c.hours || 0, c.note || 0, 1);
  const vals = sh.getRange(CFG.headerRow, 1, last - CFG.headerRow, width).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(v => v !== '' && v !== null)) return CFG.headerRow + i;
  }
  return CFG.headerRow;
}

function fmt_(d, pattern) {
  // A bound script inherits the spreadsheet's timezone.
  return Utilities.formatDate(d, Session.getScriptTimeZone(), pattern);
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Time Clock', 4); } catch (e) {}
}
