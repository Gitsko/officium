# Todo: Officium

All manual checks run on the user's PC with classic Outlook. "Self-test" means
running the macro `OfficiumSelfTest` there.

## Task 0: Own git repository

**Description:** Own repository at `github.com/Gitsko/officium`. The web add-in
prototype (Office.js, unified manifest) was built first and removed when the
project switched to VBA; it stays in history up to `2160434`.

**Acceptance criteria:**
- [x] Repository, identity and signing set up
- [x] Web add-in code removed from `main`

**Estimated scope:** S

## Task 1: Read appointments of a period including recurring series

**Description:** Macro `OfficiumListAppointments` asks for a from/to date and
writes every appointment in that range (start, end, minutes, subject,
categories) to the Immediate window, followed by the count.

**Acceptance criteria:**
- [ ] A recurring series shows up once per occurrence in the range
- [ ] An appointment on the last evening of a month is included, one on the
      first morning of the next month is not
- [ ] Count matches the Outlook calendar view for the same week

**Verification:**
- [ ] Manual check: user creates a test series, runs the macro, compares with the calendar
- [ ] Manual check: VBA is allowed at all on the target PC (no policy block)

**Dependencies:** Task 0

**Files likely touched:** `vba/Officium.bas`, `README.md`

**Notes:** Code written (`dd0abf4`); checks pending on the user's PC.

**Estimated scope:** S

## Task 2: Self-test macro, signing and ribbon button

**Description:** `OfficiumSelfTest` with a tiny assertion helper and one test
for the period presets (this month, last month, this quarter, last quarter).
README explains import, SelfCert signing and the ribbon button.

**Acceptance criteria:**
- [ ] Self-test reports passed/failed counts; a deliberately broken assertion shows as failed
- [ ] Macro runs from a ribbon button with macro security at "signed macros only"

**Verification:**
- [ ] Manual check: self-test output, screenshot of the button

**Dependencies:** Task 1

**Files likely touched:** `vba/OfficiumCore.bas`, `vba/OfficiumTests.bas`, `README.md`

**Notes:** Code and README written; checks pending on the user's PC.

**Estimated scope:** S

## Checkpoint A
- [ ] Real appointments listed, series correct, self-test runs
- [ ] Human review before Phase 2

## Task 3: Customer filter and business rules

**Description:** Pure functions: does an appointment (subject, categories,
all-day flag, meeting and response status) belong to the customer and count?

**Acceptance criteria:**
- [ ] Match by category (case-insensitive, `,` or `;` separated) or subject substring
- [ ] Cancelled, declined and all-day appointments are skipped

**Verification:**
- [ ] Self-test green with cases for each rule

**Dependencies:** Checkpoint A

**Files likely touched:** `vba/OfficiumCore.bas`, `vba/OfficiumTests.bas`

**Notes:** Code and tests written; self-test pending on the user's PC. An
empty customer matches every appointment.

**Estimated scope:** S

## Task 4: Grouping by month/quarter with totals

**Description:** Group keys (`2026-10`, `2026-Q4`), subtotal minutes per group
and a grand total, formatted as hours:minutes.

**Acceptance criteria:**
- [ ] Ranges across a year boundary group correctly (Q4/2026, Q1/2027)
- [ ] Totals above 24 hours are shown as e.g. `37:30`, not wrapped

**Verification:**
- [ ] Self-test green

**Dependencies:** Task 3

**Files likely touched:** `vba/OfficiumCore.bas`, `vba/OfficiumTests.bas`

**Notes:** Code and tests written; self-test pending on the user's PC.

**Estimated scope:** S

## Task 5: Dialog chain and on-screen summary

**Description:** `OfficiumReport` asks for period (preset or free range),
customer and grouping, reads the calendar, applies Tasks 3 and 4 and shows
the totals per group in a message box.

**Acceptance criteria:**
- [ ] Cancel at any prompt ends the macro without an error
- [ ] Invalid dates are rejected with a clear message
- [ ] Totals match a manual count for one real customer and month

**Verification:**
- [ ] Manual check on the user's PC

**Dependencies:** Task 4

**Files likely touched:** `vba/Officium.bas`, `README.md`

**Notes:** Code written; checks pending on the user's PC. An appointment
counts in the period it starts in; the last customer is remembered via
`SaveSetting`.

**Estimated scope:** M

## Checkpoint B
- [ ] Real report matches a manual count
- [ ] Self-test green
- [ ] Human review of the business rules

## Task 6: CSV export

**Description:** Detail rows plus group totals as semicolon-separated CSV,
UTF-8 with BOM, saved to the documents folder and opened.

**Acceptance criteria:**
- [ ] Opens in German Excel with correct columns and umlauts
- [ ] Body text with line breaks, quotes and semicolons is quoted correctly

**Verification:**
- [ ] Self-test green (CSV field quoting)
- [ ] Manual check: open the file

**Dependencies:** Checkpoint B

**Files likely touched:** `vba/OfficiumCore.bas`, `vba/Officium.bas`, `vba/OfficiumTests.bas`

**Notes:** Code and tests written; checks pending on the user's PC. Text
starting with = + - @ gets a leading space so Excel does not read it as a
formula.

**Estimated scope:** S

## Task 7: Clipboard text

**Description:** Plain-text summary (period, customer, group totals, one line
per appointment) on the clipboard via the Win32 API.

**Acceptance criteria:**
- [ ] Pasting into a mail shows umlauts correctly
- [ ] Works while Explorer windows are open

**Verification:**
- [ ] Self-test green (text layout)
- [ ] Manual check: paste into a mail

**Dependencies:** Checkpoint B

**Files likely touched:** `vba/OfficiumCore.bas`, `vba/Officium.bas`, `vba/OfficiumTests.bas`

**Estimated scope:** S

## Task 8: Optional Excel output

**Description:** If `CreateObject("Excel.Application")` succeeds, write the
report to a new workbook (detail sheet with autofilter, summary sheet) and
show Excel. Otherwise skip silently and mention the CSV.

**Acceptance criteria:**
- [ ] Duration column is summable with `=SUM()`
- [ ] Without Excel the macro still finishes with CSV and clipboard

**Verification:**
- [ ] Manual check on a PC with Excel

**Dependencies:** Task 6

**Files likely touched:** `vba/Officium.bas`

**Estimated scope:** S

## Checkpoint C
- [ ] All acceptance criteria met
- [ ] README: import, signing, button, usage
