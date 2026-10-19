# Implementation Plan: Officium (Outlook calendar report add-in)

## Overview

Officium is an Outlook add-in for the new Outlook / Outlook on the web (Microsoft 365) that
summarizes calendar appointments for one customer over a chosen period. The
user picks a date range and a customer filter. The add-in lists date, duration,
subject and body text of every matching appointment, sums the durations and
groups them by month or quarter. It exports to CSV, clipboard text and XLSX.
The model is VBOffice olReporter; v1 covers appointments only (no tasks, no
journal, no mail sending).

## Architecture Decisions

- **Office Web Add-in (TypeScript), task pane.** VSTO is ruled out because it
  does not run in the new Outlook. Scaffold with the current Yeoman Office
  generator (`yo office`). The spike decides between the unified JSON manifest
  and the XML manifest: check which one the generator produces today and which
  one NAA in new Outlook requires.
- **Data source: Microsoft Graph `GET /me/calendarView`.** Office.js only
  exposes the current item and cannot enumerate a date range. EWS
  (`makeEwsRequestAsync`) and the legacy Outlook REST API are retired for
  Exchange Online and are not a fallback.
- **Auth: MSAL.js with Nested App Authentication (NAA)** plus an Entra ID app
  registration with delegated `Calendars.Read`. The exact redirect URI scheme
  and the manifest settings are taken from Microsoft Learn ("Nested app
  authentication in Office Add-ins") during the spike, not from memory.
- **Graph request rules:**
  - `calendarView` (not `/events`) so recurring series arrive as instances.
  - Page through `@odata.nextLink`; `$top` raised; `$select` limited to
    `subject,body,start,end,isAllDay,isCancelled,categories,showAs,responseStatus,attendees,organizer`.
  - Header `Prefer: outlook.timezone="W. Europe Standard Time"`. Otherwise
    times arrive in UTC and appointments near midnight slip into the wrong
    month.
  - Header `Prefer: outlook.body-content-type="text"`. The body text is part
    of the report.
- **Domain logic as pure functions** (filter, duration, grouping, summing,
  formatting) in `src/report/`, independent of Office.js and Graph, unit-tested
  with Vitest. Office/Graph code stays a thin adapter.
- **Customer filter, default until clarified:** configurable match on Outlook
  category and/or a subject keyword, applied client-side after fetching.
- **Business rules (defaults, confirm in review):** duration = `end - start`;
  exclude `isCancelled`; exclude appointments the user declined; exclude
  all-day appointments; include `showAs = free` items.
- **XLSX library:** not the npm `xlsx` package (stuck at 0.18.x with known
  CVEs). Use the SheetJS build from the SheetJS CDN tarball or `exceljs`.
  Decide in Task 8 by bundle size and licence.
- **Export order by risk:** CSV (no dependency) → clipboard text → XLSX.
  `<a download>` blob downloads and `navigator.clipboard.writeText` may be
  blocked inside the add-in iframe on some hosts. The spike checks both.

## Task List

Tasks live in `tasks/todo.md`.

### Phase 0: Repository
- [x] Task 0: Own git repository and scaffold

### Phase 1: Feasibility spike (highest risk first)
- [ ] Task 1: Sideload task pane from WSL2 dev server into new Outlook
- [ ] Task 2: NAA sign-in and one `calendarView` call
- [ ] Task 3: Probe blob download and clipboard in the add-in host

### Checkpoint A: Spike
- [ ] Token obtained, real appointments displayed, export paths known
- [ ] Human review: go/no-go for the Graph + NAA approach

### Phase 2: Core report
- [ ] Task 4: Date range selection with paged, timezone-correct fetch
- [ ] Task 5: Customer filter and business rules
- [ ] Task 6: Grouping by month/quarter with totals

### Checkpoint B: Core
- [ ] Report for a real customer and month matches a manual count
- [ ] Vitest suite green

### Phase 3: Export
- [ ] Task 7: CSV export
- [ ] Task 8: Clipboard text summary
- [ ] Task 9: XLSX export

### Checkpoint C: Complete
- [ ] All three exports verified with a real report
- [ ] README describes app registration, sideloading and usage

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Tenant admin consent for `Calendars.Read` not granted | High | Clarify with M365 admin before Task 2; spike stops here if refused |
| NAA unsupported or configured differently in the target client | High | Task 2 is a spike; check the Microsoft Learn support matrix first |
| WSL2 dev certificate not trusted by Outlook/WebView2 on Windows | Med | Task 1 installs the dev CA into the Windows store and proves it |
| Download/clipboard blocked in add-in iframe | Med | Task 3 probes; fallback `Office.context.ui.displayDialogAsync` or a text area to copy from |
| Timezone/DST errors at month boundaries | Med | `Prefer: outlook.timezone` header plus unit tests around DST and month ends |
| Recurring series counted wrongly | Med | Use `calendarView`; test with a real recurring series |
| Large ranges (a year) hit Graph throttling | Low | Paging, `$select`, honour `Retry-After` on 429 |

## Open Questions

- Customer assignment: category, subject keyword, attendee domain, or a mix?
  (Default: category or keyword, configurable.)
- All-day appointments: exclude, or count as a fixed number of hours?
- Count declined / tentative / `showAs = free` appointments?
- Ticket key for branch names and commit prefixes (repo: github.com/Gitsko/officium, private, signed + DCO)?
- Who registers the Entra app and grants consent (own tenant or Netresearch)?
- Should filter presets per customer be saved (roaming settings) in v1?
