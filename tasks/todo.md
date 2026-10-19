# Todo: Officium

## Task 0: Own git repository and scaffold

**Description:** `officium` currently sits inside the home dotfiles repo.
Initialise its own repository, scaffold a TypeScript Outlook task-pane add-in
with the current Yeoman Office generator (package name `officium`, display name "Officium"), and add Vitest.

**Acceptance criteria:**
- [x] `git -C ~/projects/officium rev-parse --show-toplevel` points to the project, not `~`
- [x] `npm run build` succeeds; `npx vitest run` runs (manifest identity test); `npm run typecheck` clean
- [x] Remote, identity and signing set up according to the chosen repo target

**Verification:**
- [x] Build: `npm run build`
- [x] Tests: `npx vitest run`

**Dependencies:** None (repo target from Open Questions)

**Files likely touched:** `package.json`, `manifest.json`, `src/**`, `tests/manifest.test.ts`, `tsconfig.json`, `.gitignore`

**Notes:** Scaffolded with `yo office taskpane Officium outlook json --ts` (unified manifest). The `nested-app-auth` template was rejected: it is from 2024 and only targets Excel/Word/PowerPoint with an XML manifest. MSAL/NAA is added in Task 2.

**Estimated scope:** M

## Task 1: Sideload task pane from WSL2 dev server into new Outlook

**Description:** Run the dev server in WSL2 and sideload the add-in into the new
Outlook on Windows (and OWA). The dev certificate must be trusted on the
Windows side.

**Acceptance criteria:**
- [ ] Task pane opens in new Outlook from a calendar context and shows "Hello"
- [ ] Same in Outlook on the web
- [ ] Steps (cert trust, sideload) documented in README

**Verification:**
- [ ] Manual check: screenshot of the open task pane in both clients

**Dependencies:** Task 0

**Files likely touched:** `manifest.*`, `README.md`

**Estimated scope:** S

## Task 2: NAA sign-in and one calendarView call

**Description:** Register an Entra ID app (delegated `Calendars.Read`),
configure NAA per Microsoft Learn, acquire a token silently (popup fallback)
and render the subjects of the next seven days.

**Acceptance criteria:**
- [ ] Token acquired without a separate login in new Outlook
- [ ] Real appointments of the next 7 days listed in the task pane
- [ ] App registration steps (redirect URI, permissions, consent) in README

**Verification:**
- [ ] Manual check: list matches the Outlook calendar view

**Dependencies:** Task 1

**Files likely touched:** `src/auth/msal.ts`, `src/graph/client.ts`, `src/taskpane/taskpane.ts`, `manifest.*`, `README.md`

**Estimated scope:** M

## Task 3: Probe blob download and clipboard in the add-in host

**Description:** Two throwaway buttons: download a tiny CSV blob, copy a string
to the clipboard. Record which works in new Outlook and OWA and pick the
fallback where needed.

**Acceptance criteria:**
- [ ] Result per client recorded in `tasks/plan.md` (Architecture Decisions)
- [ ] Probe code removed afterwards

**Verification:**
- [ ] Manual check in both clients

**Dependencies:** Task 1

**Files likely touched:** `src/taskpane/taskpane.ts`, `tasks/plan.md`

**Estimated scope:** XS

## Checkpoint A: Spike
- [ ] Token obtained, real data visible, export paths known
- [ ] Human review: go/no-go before Phase 2

## Task 4: Date range selection with paged, timezone-correct fetch

**Description:** From/to date inputs (presets: this month, last month, this
quarter). Fetch all instances via `calendarView` with paging, `$select`, the
timezone and text-body `Prefer` headers. Map to an internal `Appointment` type.

**Acceptance criteria:**
- [ ] A full month with >50 appointments loads completely (paging works)
- [ ] An appointment at 23:30 on the last day of the month stays in that month
- [ ] Recurring series appear as individual instances

**Verification:**
- [ ] Tests: `npx vitest run` (mapping, paging with mocked fetch responses)
- [ ] Manual check: count equals Outlook calendar for the same month

**Dependencies:** Checkpoint A

**Files likely touched:** `src/graph/calendarView.ts`, `src/report/types.ts`, `src/taskpane/*`, `tests/graph/*.test.ts`

**Estimated scope:** M

## Task 5: Customer filter and business rules

**Description:** Pure filter function: match category and/or subject keyword;
exclude cancelled, declined and all-day items (per resolved Open Questions).
Category dropdown filled from `/me/outlook/masterCategories`.

**Acceptance criteria:**
- [ ] Filter by category and by keyword works, also combined
- [ ] Cancelled and declined appointments are excluded
- [ ] Table shows date, start, end, duration, subject, body text

**Verification:**
- [ ] Tests: `npx vitest run tests/report/filter.test.ts`
- [ ] Manual check with a real customer category

**Dependencies:** Task 4

**Files likely touched:** `src/report/filter.ts`, `src/taskpane/*`, `tests/report/filter.test.ts`

**Estimated scope:** S

## Task 6: Grouping by month/quarter with totals

**Description:** Pure functions to group by none/month/quarter and sum
durations per group and overall. Render group headers with subtotals.

**Acceptance criteria:**
- [ ] Subtotals per month and quarter plus grand total in hours:minutes
- [ ] Ranges spanning year boundaries group correctly (Q4/2026, Q1/2027)
- [ ] DST change weekends produce correct durations

**Verification:**
- [ ] Tests: `npx vitest run tests/report/group.test.ts`
- [ ] Manual check: one real month matches a manual sum

**Dependencies:** Task 5

**Files likely touched:** `src/report/group.ts`, `src/taskpane/*`, `tests/report/group.test.ts`

**Estimated scope:** S

## Checkpoint B: Core
- [ ] Real report for one customer/month matches a manual count
- [ ] `npx vitest run` green, `npm run build` clean
- [ ] Human review of the business rules

## Task 7: CSV export

**Description:** Serialise the current report (detail rows plus group totals)
as CSV (semicolon, UTF-8 BOM so Excel shows umlauts) and deliver it via the
mechanism chosen in Task 3.

**Acceptance criteria:**
- [ ] File opens in German Excel with correct columns and umlauts
- [ ] Body text with line breaks and semicolons is quoted correctly

**Verification:**
- [ ] Tests: `npx vitest run tests/export/csv.test.ts`
- [ ] Manual check: open the exported file in Excel

**Dependencies:** Checkpoint B

**Files likely touched:** `src/export/csv.ts`, `src/taskpane/*`, `tests/export/csv.test.ts`

**Estimated scope:** S

## Task 8: Clipboard text summary

**Description:** Plain-text summary (period, customer, per-group totals, one
line per appointment) copied to the clipboard, ready for invoices or worklogs.

**Acceptance criteria:**
- [ ] Copy works in new Outlook and OWA (or the Task 3 fallback is shown)
- [ ] Format is stable and readable when pasted into mail or Jira

**Verification:**
- [ ] Tests: `npx vitest run tests/export/text.test.ts`
- [ ] Manual check: paste into a mail

**Dependencies:** Checkpoint B

**Files likely touched:** `src/export/text.ts`, `src/taskpane/*`, `tests/export/text.test.ts`

**Estimated scope:** S

## Task 9: XLSX export

**Description:** Pick the library (SheetJS CDN build vs. `exceljs`), export
detail sheet plus summary sheet with durations as Excel time values so sums
work in Excel.

**Acceptance criteria:**
- [ ] Library decision recorded in `tasks/plan.md`
- [ ] Detail and summary sheet; duration column summable in Excel
- [ ] Bundle size increase noted

**Verification:**
- [ ] Tests: `npx vitest run tests/export/xlsx.test.ts` (read back the generated workbook)
- [ ] Build: `npm run build`
- [ ] Manual check: open in Excel, `=SUM()` over duration works

**Dependencies:** Task 7

**Files likely touched:** `src/export/xlsx.ts`, `package.json`, `src/taskpane/*`, `tests/export/xlsx.test.ts`

**Estimated scope:** M

## Checkpoint C: Complete
- [ ] All acceptance criteria met
- [ ] README: app registration, sideloading, usage
- [ ] Ready for review
