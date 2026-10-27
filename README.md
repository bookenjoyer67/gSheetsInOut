# gSheetsInOut — one-click time in / time out

Google Apps Script for a bi-weekly timesheet. One tap on a checkbox fills the next
empty punch slot on today's row.

## Install

1. Open the timesheet → **Extensions → Apps Script**
2. Replace the contents of `Code.gs` with `timeclock-appsscript.gs` → **Save**
3. Run `setupTimeClock` once → approve the permission prompt → reload the sheet

`setupTimeClock` tells you exactly which cell the checkbox landed in. You do not have to
run it more than once: `onOpen` re-creates the checkbox whenever it is missing, so it
cannot quietly vanish again.

## Why it did nothing on mobile

Nothing was broken in the script — the two desktop affordances do not exist in the
Sheets mobile app:

- **Custom menus** (`onOpen` + `createMenu`) render only in the desktop/web version.
- **Drawings and images** cannot run assigned scripts in the mobile app; tapping them
  does nothing.

The one mechanism the app does support is a cell edit firing `onEdit` server-side, which
is why the checkbox is the mobile path.

## Behaviour

Today's row is found by matching the day-of-month in column **C** against today, within
rows 10-26. Punches fill in this order:

| order | cell | meaning |
|-------|------|---------|
| 1 | D | Time In |
| 2 | E | Time Out |
| 3 | G | Time In (2nd shift) |
| 4 | H | Time Out (2nd shift) |

- 2 taps = normal day, 4 taps = split shift, 5th tap writes nothing and reports the row
  is full.
- The checkbox re-arms itself after every tap, so it stays clickable forever.
- Feedback goes to **K3** ("Time In punched at 9:00 AM (row 13)") because the mobile app
  shows no toasts; desktop also gets a toast.
- Values are written as time-of-day fractions with `h:mm am/pm`, so the sheet's `E-D` /
  `H-G` Sub-Total formulas keep working exactly as with hand-typed times.
- Rounded to the nearest 15 minutes (`CFG.roundTo15`, per the sheet's own rule); 23:5x
  clamps to 23:45 instead of wrapping to midnight.
- A date not in the sheet writes nothing and reports an error instead of guessing a row.
- `Time Clock → Undo last punch` (desktop) clears the last filled slot on today's row.

## Cells it writes

`D`, `E`, `G`, `H` on today's row, plus the checkbox cell, its label, and `K3` for
status.

