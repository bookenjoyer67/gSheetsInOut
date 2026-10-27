# gSheetsInOut — one-click time in / time out

Google Apps Script for as bi-weekly timesheet. One click fills the next
empty punch slot on today's row.

## Install

1. Open the timesheet → **Extensions → Apps Script**
2. Replace the contents of the boilerplate function with `timeclock-appsscript.gs` → **Save**
3. Reload

The **Time Clock** menu appears on load (`onOpen`), offering *Clock In / Out* and
*Undo last punch*. Nothing else to authorize or run.

## Button

The button is a manual 4-step:

1. **Insert → Drawing**, draw a box, type `CLOCK IN / OUT`, **Save and Close**
2. Click the drawing → **⋮** (3-dot menu) → **Assign script**
3. Type `clockInOut` → OK
4. Click it — it punches the clock

## Behaviour

Today's row is found by matching the day-of-month in column **C** against today, within
rows 10-26. Punches fill in this order:

| order | cell | meaning |
|-------|------|---------|
| 1 | D | Time In |
| 2 | E | Time Out |
| 3 | G | Time In (2nd shift) |
| 4 | H | Time Out (2nd shift) |

- 2 clicks = normal day, 4 clicks = split shift, 5th click writes nothing and says the
  row is full.
- Values are written as time-of-day fractions with `h:mm am/pm`, so the sheet's
  `E-D` / `H-G` Sub-Total formulas keep working exactly as with hand-typed times.
- Rounded to the nearest 15 minutes (`CFG.roundTo15`, per the sheet's own rule).
  23:5x clamps to 23:45 instead of wrapping to midnight.
- A date not in the sheet (e.g. after the pay period rolls over) writes nothing and
  toasts an error rather than guessing a row.
- `Undo last punch` clears the last filled slot on today's row.

## Cells it writes

Only `D`, `E`, `G`, `H` on today's row in column-C range. Formulas, the How To Use
column, and everything else are untouched.

## Layout assumptions

Rows 10-23 are the pay-period days, row 24 is the `SPARE DATES` label (no day number
in column C, so it is skipped automatically), rows 25-26 are the spare days. If the
template changes, adjust `CFG` at the top of the script.

