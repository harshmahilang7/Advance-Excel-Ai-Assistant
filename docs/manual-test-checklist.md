# Manual test checklist (real Excel)

The automated tests mock Excel, so they cannot prove the add-in works inside it. Run this list in **desktop Excel (Microsoft 365)** before a release, or after changing anything in `src/excel/` or `src/ui/`. It takes about 15 minutes.

Use a throwaway workbook. Record the Excel build (File → Account → About Excel) and the model you used next to any failure.

## Setup

- [ ] `npm run dev`, then `npm run sideload` (or the installed add-in). The **AI Assistant** pane opens.
- [ ] First run shows the key screen. A bad key is rejected with a masked message; a good key continues.
- [ ] Add two keys. Close and reopen the pane: it restores both without asking again.
- [ ] The pane fits at its narrowest width: no overlapping controls, the chat area is readable, the text box and Send are always visible.

## Basic edits

Put this in Sheet1 first: headers `Region, Product, Qty, Price` and 5 rows of data.

- [ ] "Add a Total column that multiplies Qty by Price"
  - A formula column appears next to the data (not over it).
  - The step line says `verified n/n cells`.
  - The first new cell has a cell note explaining the formula (Excel 365 with API 1.18; otherwise the step says the note was skipped).
  - The reply shows cell addresses as links; clicking one selects that range.
  - The pane shows `Worked for … · N steps · N change(s)` and a `Sent to the AI` line.
- [ ] **Undo last** removes the column and the note.
- [ ] "Bold the header row" does not show a plan (one change).
- [ ] "Add a Total column, bold the header row and chart the totals" shows a numbered **Plan** first, then the steps.

## Checks and warnings

- [ ] "Put the text 2024-01-05 in G1 and =1/0 in G2": the step line warns the first cell "read back differently" and the second "shows an error".
- [ ] Ask for something that leaves a `#REF!` in a multi-step task. The model gets a self-check and either fixes it or explains it.
- [ ] "Create a pivot table of total Qty by Region at B2" (a non-empty cell): refused with `DESTINATION_OCCUPIED`, and nothing changes.
- [ ] "Make a chart of Z1:Z10" (empty range): refused with `SOURCE_EMPTY`.

## Preview, undo and stop

- [ ] Turn on **Preview before applying**, ask for a change: nothing changes until **Apply**. **Discard** leaves the sheet untouched.
- [ ] Format a whole column ("make column B bold with a yellow fill"), then **Undo last**: the column's formatting returns to how it was (a header you formatted differently keeps its look).
- [ ] **Stop** during a long request ends it with "Stopped."
- [ ] A request that hits the step limit shows **Continue**, and Continue resumes without redoing finished steps.

## Data tools

- [ ] **Import CSV** with a small file (with a quoted comma and a line break in a cell): a new sheet appears with a bold, frozen header, and existing sheets are untouched. **Undo last** deletes the new sheet.
- [ ] A CSV cell starting with `=` stays as text.
- [ ] **Save template**, then **Clear chat**: the template is on the start screen; clicking it fills the text box; ✕ deletes it.
- [ ] A starter chip under "check your sheet without changing it" runs and changes nothing.

## Keys and limits

- [ ] With two keys, make the first one hit a quota (or use a deliberately invalid first key): the request still finishes using the second.
- [ ] **⋯ → Manage API keys** and **Remove all API keys** work; removing returns to the key screen.
- [ ] Turn off the network: the error is plain English, not a stack trace.

## After the run

- [ ] No console errors in the pane (right-click → Inspect, if available).
- [ ] Note anything that looked wrong, even if it did not fail. UI glitches are the likeliest bugs.
