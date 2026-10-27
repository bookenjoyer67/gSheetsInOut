/**
 * TIME CLOCK  —  automatic time in / time out for this spreadsheet
 * -----------------------------------------------------------------
 * Install:  Extensions > Apps Script, delete the boilerplate, paste this whole file, Save.
 * Then run  setupTimeClock  once and approve the permissions prompt.
 *
 * That gives you three ways to punch in/out:
 *   1. Menu:      "Time Clock" > Clock In / Clock Out   (always available)
 *   2. Checkbox:  cell G2 (or whatever CFG.triggerCell says) — click to toggle
 *   3. Button:    Insert > Drawing, draw a box, type "CLOCK IN / OUT",
 *                 then the 3-dot menu on the drawing > Assign script > clockInOut
 *
 * Nothing else in the sheet is touched: the script only writes the
 * Date / Time In / Time Out / Hours columns it finds by header name.
 */


