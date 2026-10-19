# Implementation Plan: Officium (Outlook calendar report macro)

## Overview

Officium summarizes the calendar appointments of one customer over a chosen
period in classic Outlook for Windows. The user picks a period, a customer and
a grouping (none, month, quarter). Officium lists date, start, end, duration,
subject and body text of every matching appointment, sums the durations per
group and overall, and exports the report as CSV and as plain text on the
clipboard. If Excel is installed, it also opens the report in Excel. The model
is VBOffice olReporter; v1 covers appointments only (no tasks, no journal, no
mail sending).

## Architecture Decisions

- **VBA macro in classic Outlook, nothing but Office.** Decided 2026-10-19
  after the web add-in prototype (kept in git history up to `2160434`). The
  user wants no web page, no hosting and no installation beyond Office. New
  Outlook cannot run VBA, so Officium targets classic Outlook only.
- **Development and test happen on different machines.** This PC has only new
  Outlook and no Excel. Code is written here; the user imports and runs it on
  a PC with classic Outlook. Every task therefore ends with a manual check by
  the user, and the self-test macro is the main safety net.
- **Source layout:** plain-text VBA modules in `vba/`, imported via the VBA
  editor (File > Import). `.bas` files are stored with CRLF (`.gitattributes`).
  They are encoded as Windows-1252, because the editor reads them that way
  and UTF-8 umlauts would arrive garbled, and stored byte for byte
  (`-text`), so a download from GitHub works as well as a clone.
  No `.frm`/`.frx`: a UserForm needs the binary `.frx`, which cannot be
  produced or reviewed here.
  - `OfficiumCore.bas`: pure logic (period presets, customer match, business
    rules, grouping, totals, CSV and text formatting). No Outlook objects, so
    the self-test can cover it completely.
  - `Officium.bas`: entry macro `OfficiumReport`, dialogs, calendar access,
    file/clipboard/Excel output.
  - `OfficiumTests.bas`: `OfficiumSelfTest`, assertions over the core module,
    result in a message box and the Immediate window.
- **Dialogs via `InputBox`/`MsgBox`.** Period (preset or free range), customer,
  grouping. A UserForm can follow later if the dialog chain proves clumsy.
- **Calendar access via the Outlook object model:**
  - `Items.Sort "[Start]"`, then `IncludeRecurrences = True`, then `Restrict`
    with a closed range (`[Start] < to AND [End] > from`). Any other order
    returns series only once or loops endlessly.
  - Dates in the `Restrict` string use `Format(d, "ddddd hh:nn")`: the
    system's short date as Outlook expects, with a 24-hour time. The often
    quoted `h:nn AMPM` is ambiguous on German systems, which have no AM/PM
    designators.
  - Times come in local time; no timezone handling needed.
- **Customer match, default until clarified:** an appointment belongs to the
  customer if one of its categories equals the input (case-insensitive) or the
  subject contains it. `Categories` is split on both `,` and `;` because the
  separator follows the Windows list separator.
- **Business rules (defaults, confirm in review):** duration = `End - Start`
  in minutes; skip cancelled meetings (`MeetingStatus` 5 and 7); skip declined
  ones (`ResponseStatus` 4); skip all-day appointments; include free/tentative.
- **Exports:**
  - CSV: semicolon-separated, UTF-8 with BOM via `ADODB.Stream` (part of
    Windows), written next to the user's documents, opened afterwards.
  - Clipboard: Win32 clipboard API via `Declare PtrSafe`. `MSForms.DataObject`
    is avoided because it pastes `??` on Windows 8 and later while Explorer
    windows are open.
  - Excel: late-bound `CreateObject("Excel.Application")`, only if it
    succeeds; durations as Excel time values so `=SUM()` works.
- **Start button:** the user adds `OfficiumReport` to the ribbon or the Quick
  Access Toolbar (File > Options > Customize Ribbon > Macros). Signed with a
  certificate from `SelfCert.exe` (ships with Office) so macro security can
  stay at "signed macros only".

## Task List

Tasks live in `tasks/todo.md`.

### Phase 0: Repository
- [x] Task 0: Own git repository (web add-in prototype archived in history)

### Phase 1: Feasibility (highest risk first)
- [ ] Task 1: Read appointments of a period including recurring series
- [ ] Task 2: Self-test macro, signing and ribbon button

### Checkpoint A
- [ ] Real appointments listed, series correct, self-test runs on the user's PC
- [ ] Human review: go for the dialog and report phase

### Phase 2: Core report
- [ ] Task 3: Customer filter and business rules
- [ ] Task 4: Grouping by month/quarter with totals
- [ ] Task 5: Dialog chain and on-screen summary

### Checkpoint B
- [ ] Report for a real customer and month matches a manual count
- [ ] Self-test green

### Phase 3: Export
- [ ] Task 6: CSV export
- [ ] Task 7: Clipboard text
- [ ] Task 8: Optional Excel output

### Checkpoint C
- [ ] All exports verified with a real report
- [ ] README describes import, signing, button and usage

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Code cannot be run here; slow feedback loop | High | Small tasks, pure core with self-test, precise check steps per task |
| Recurring series missing or endless loop | High | Task 1 first, with a test series in the calendar |
| `Restrict` date string misread under German locale | Med | `ddddd hh:nn` format; Task 1 checks a range across month end |
| Macro security blocks unsigned code | Med | SelfCert signing in Task 2, documented in README |
| Organisation policy disables VBA in Outlook entirely | High | Checked in Task 1 on the target PC; if so, stop and rethink |
| Large bodies or many appointments make Excel output slow | Low | Write a 2D array to the sheet in one assignment |

## Open Questions

- Customer assignment: category, subject keyword, or both? (Default: both.)
- All-day appointments: exclude, or count as a fixed number of hours?
- Count tentative / free appointments?
- Should the body text be shortened (first line only) in the clipboard text?
