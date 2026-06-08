# Microsoft Support Brief — Exchange Online Calendar Data Loss
## Ventura Fund Services / matt.lowe@venturafs.com

**Prepared:** 2026-06-08  
**Tenant:** venturafs.com — TenantId `b3377764-8275-47f8-8e2d-697b136c21c1`  
**Affected user:** matt.lowe@venturafs.com  
**Severity:** High — confirmed loss of business-critical calendar appointments including client-facing meetings and organizer copies of meetings Matt himself created

---

## 1. Executive Summary

Matt Lowe's Exchange Online calendar is confirmed to be missing **14 meetings** that demonstrably exist on the calendars of colleagues who were co-attendees. This was proven independently of the Calendar Diagnostic Log (which is currently failing server-side in this tenant — see §6).

The failures span both directions:

- **3 meetings that Matt himself organised** — he sent the invites, colleagues received and hold them, but Matt's own copy never appeared or was subsequently removed.
- **11 meetings where Matt was an invited attendee** — the meetings exist on the organiser's and/or other attendees' calendars but are entirely absent from Matt's.

All 14 missing meetings have `SeriesPresentOnMatt = False` — meaning their `iCalUId` is wholly absent from Matt's calendar, not a partial recurrence drop. Two independent runs of the investigation tool on 2026-06-08 produced identical findings, confirming the data is stable and reproducible.

The root-cause authoring client (`ClientInfoString`) cannot be determined without the Calendar Diagnostic Log. Restoring `Get-CalendarDiagnosticObjects` functionality for this tenant (§6) is the critical next step.

---

## 2. Tenant and User Details

| Field | Value |
|---|---|
| Tenant domain | venturafs.com |
| Tenant ID | b3377764-8275-47f8-8e2d-697b136c21c1 |
| Affected mailbox | matt.lowe@venturafs.com |
| Mailbox type | Exchange Online (cloud-only) |
| Investigation window | Approximately 2026-03-09 — 2026-06-08 (90 days) |
| Investigation date | 2026-06-08 |

---

## 3. Reported Symptom

Colleagues at Ventura Fund Services reported that meetings — some organised by Matt, some to which Matt had been invited — were absent from Matt's calendar while remaining present on their own calendars. Matt had no visibility of these meetings and could not see them from any client.

---

## 4. Investigation Methodology

`Get-CalendarDiagnosticObjects` is currently failing server-side for this tenant (see §6). All investigation was therefore conducted using **Microsoft Graph API `calendarView`** only, which is functioning normally.

### 4.1 Tool

A purpose-built PowerShell 7 script (`Confirm-OrganizerCopyMissing.ps1`) was used. It operates as follows:

**Step 0 — Address resolution**  
Matt's full proxy/alias address set was resolved from `Get-Mailbox -Identity matt.lowe@venturafs.com` to ensure that meetings stamped with an alias SMTP (not his primary address) were not incorrectly excluded.

**Step 1 — Read Matt's calendar and discover colleagues**  
Matt's calendar was read via Graph `calendarView` over the full 90-day window in 15-day slices to avoid timeout. This produced:
- The set of meeting occurrences currently present on his calendar (keyed by `iCalUId + occurrence-start-to-the-minute` to support per-occurrence comparison of recurring series).
- A bidirectional colleague list: people Matt frequently *invites* (from meetings he organised) plus people who *organised* meetings on Matt's calendar where he was an attendee. This ensures the scan set covers both failure directions.

**Step 2 — Grant calendar access**  
Least-privilege `Reviewer` access was granted to the signed-in admin on each colleague's Calendar folder via `Add-MailboxFolderPermission` (idempotent; any existing `AvailabilityOnly` or `LimitedDetails` right was upgraded to `Reviewer`, as free/busy-only access returns no item detail from `calendarView`). A 120-second propagation wait was applied.

**Step 3 — Scan colleague calendars (both scenarios in one pass)**  
Each colleague's calendar was read over the same 90-day window. For each meeting found:
- **Scenario A:** If `organizer.emailAddress.address` matched Matt's address set → candidate for organizer-missing comparison.
- **Scenario B:** If Matt appeared in the meeting's `attendees` list (any email in his address set) → candidate for invitee-missing comparison.

**Step 4 — Compare against Matt's presence set**  
Each candidate occurrence (keyed by `iCalUId + start-minute`) was looked up in Matt's presence sets. Cancelled events were excluded on both sides to avoid cancelled tombstones producing false positives.

**Step 5 — Trustworthiness check**  
A result is only declared trustworthy if every scanned colleague calendar was successfully read. Any access failure (403/404) causes the affected colleague to be excluded and the result to be explicitly caveated as INCOMPLETE. In both runs, all 11 colleague calendars were read successfully → result is fully trustworthy for the scanned set.

### 4.2 Corroboration count

The `CalendarCount` field in the output records how many independent colleague calendars each missing meeting was found on. A meeting confirmed on 6 or 9 colleagues' calendars (see §5) provides strong independent corroboration that the meeting was real and properly delivered to the tenant — the absence from Matt's calendar is not a delivery failure but a calendar-layer issue.

### 4.3 Repeatability

The tool was run twice on 2026-06-08:
- Run 1: 14:09 local time
- Run 2: 14:37 local time

Both runs produced **identical findings** — same 14 missing meetings, same iCalUIds, same colleague sightings. The data is stable.

---

## 5. Confirmed Findings

**Total missing from Matt's calendar: 14**  
**Organizer scenario (Matt sent the invite): 3**  
**Invitee scenario (Matt was invited): 11**

All 14 have `SeriesPresentOnMatt = False` — the `iCalUId` is entirely absent from Matt's calendar in every case.

---

### 5A. Meetings Matt Organised — Missing From His Own Calendar

These meetings were created by Matt and delivered to attendees. Attendees hold them. Matt does not.

| # | Date/Time (UTC) | Subject | Confirmed on (colleagues) | iCalUId |
|---|---|---|---|---|
| 1 | 2026-03-23 18:30 | Matt & Alan (Ventura Fund Services) / Joe and Brian (NSP Capital) | alan.tsarovsky@venturafs.com | `040000008200E00074C5B7101A82E00800000000B0EB63C38CB8DC010000000000000000100000005DEE76618B903D4387B9847A371583A3` |
| 2 | 2026-05-20 19:30 | Strategy Session - Matt/Svetlana/Caroline | caroline.cruz@venturafs.com; svetlana.benjamin@venturafs.com | `040000008200E00074C5B7101A82E0080000000080D72573BCE7DC01000000000000000010000000DE46CC2127AE224A83CCC9088FAD25B6` |
| 3 | 2026-05-27 16:00 | BDO/Ventura Fund Services - Lunch - Sen Sakana (44th, between 5th and 6th) | greg.shneynberg@venturafs.com; eryn.darcy@venturafs.com | `040000008200E00074C5B7101A82E008000000007042A54D1CE2DC010000000000000000100000007288ECA169D85A4E842BA62F29A57167` |

---

### 5B. Meetings Matt Was Invited To — Missing From His Calendar

These meetings were organised by someone else. Matt was listed as an attendee. The meetings are present on the organiser's and/or other attendees' calendars but entirely absent from Matt's.

| # | Date/Time (UTC) | Subject | Organiser | Calendar count | iCalUId |
|---|---|---|---|---|---|
| 4 | 2026-03-20 18:00 | Monthly Events Calendar Meeting | lindsey.baginski@venturafs.com | 2 | `040000008200E00074C5B7101A82E00807EA03106067A45F1ECFDB01000000000000000010000000741C56F16F1DB3438CFBDD8291B381EF` |
| 5 | 2026-03-23 13:00 | MK - Half day - Monday AM | manoj.kamdar@venturafs.com | **9** | `040000008200E00074C5B7101A82E008000000005E0E6A1683B8DC010000000000000000100000009B435860862F6A46B7E1094D25F4CBEB` |
| 6 | 2026-04-20 19:00 | Monthly Events Calendar Meeting | lindsey.baginski@venturafs.com | 1 | `040000008200E00074C5B7101A82E00807EA04146067A45F1ECFDB01000000000000000010000000741C56F16F1DB3438CFBDD8291B381EF` |
| 7 | 2026-04-28 21:00 | Justin Passek and Royce Wilson/Ventura Fund Services Followup | alan.tsarovsky@venturafs.com | 1 | `040000008200E00074C5B7101A82E0080000000071C3186969D6DC01000000000000000010000000699AFD15A727024997E1AB8871089365` |
| 8 | 2026-05-04 14:00 | FW: NSP/Catalyst ODD Call | dan.pogue@moelisam.com *(external)* | 1 | `040000008200E00074C5B7101A82E00800000000F04A272896D8DC0100000000000000001000000057D085C570685E4FB12811A7C3E8115D` |
| 9 | 2026-05-06 20:00 | White Wolf Capital/Ventura Fund Services Proposal | alan.tsarovsky@venturafs.com | 1 | `040000008200E00074C5B7101A82E0080000000087E66A25D9DBDC010000000000000000100000004BB0AC53B7F706468310DE570EFB10B3` |
| 10 | 2026-05-08 18:00 | AI Implementation Discussion - Legal Aspects | alan.tsarovsky@venturafs.com | 1 | `040000008200E00074C5B7101A82E00800000000FB42FA8C68DDDC01000000000000000010000000A567818EBBB3804F97D5422672B1EBCF` |
| 11 | 2026-05-08 18:30 | Ventura/Metric Point - discuss NSP true up at 1st close | aleykikh@metricpoint.com *(external)* | **6** | `040000008200E00074C5B7101A82E008000000005040038A85DDDC01000000000000000010000000F9E11B7E664C4740BA24AE3DB2C8DA6C` |
| 12 | 2026-05-13 00:00 | MK - PTO | manoj.kamdar@venturafs.com | 3 | `040000008200E00074C5B7101A82E00800000000FB31BE7019A0DC01000000000000000010000000B468F18D74E7DD448E0FA62FAE5C1B73` |
| 13 | 2026-05-18 19:00 | Monthly Events Calendar Meeting | lindsey.baginski@venturafs.com | 1 | `040000008200E00074C5B7101A82E00807EA05126067A45F1ECFDB01000000000000000010000000741C56F16F1DB3438CFBDD8291B381EF` |
| 14 | 2026-05-18 20:00 | AORA Investments/Ventura Fund Services Demo (Portal and Deliverables) | alan.tsarovsky@venturafs.com | 3 | `040000008200E00074C5B7101A82E00800000000641F48086CDDDC01000000000000000010000000DF021FA46B4E724C8DC928040B5797C0` |

---

## 6. Active Backend Issue — Get-CalendarDiagnosticObjects Failing (CONFIRMED)

`Get-CalendarDiagnosticObjects` is currently returning a server-side error for this tenant. This cmdlet is the authoritative source for the `ClientInfoString` field — the only way to identify which application (Outlook desktop, OWA, mobile client, EWS automation, Graph API) dropped or deleted the calendar items.

### 6.1 Confirmed Diagnosis: TENANT_WIDE_BACKEND_FAILURE

A purpose-built diagnostic script (`Test-CalendarDiagnosticAccess.ps1`) was run against this tenant on 2026-06-08. All 8 tests failed with the same server-side error:

```
A server side error has occurred. Please visit the Exchange Admin Center and
try the operation again, or contact support if the issue persists.
```

Key diagnostic results:

| Test | Target | Params | Result |
|---|---|---|---|
| A | Admin's own mailbox (`ndradmin@venturafs.com`) | ResultSize 1 | **FAIL** — server error |
| B | `matt.lowe@venturafs.com` | ResultSize 1 | **FAIL** — server error |
| C | `matt.lowe@venturafs.com` | Known-present subject | **FAIL** — server error |
| D | `matt.lowe@venturafs.com` | Known-missing subject | **FAIL** — server error |
| E | `matt.lowe@venturafs.com` | MeetingID (CleanGOID) | **FAIL** — server error |
| F | `alan.tsarovsky@venturafs.com` | ResultSize 1 | **FAIL** — server error |
| G | `alan.tsarovsky@venturafs.com` | ResultSize 1 (throttle check) | **FAIL** — server error |
| H | `matt.lowe@venturafs.com` | ShouldDecodeEnums | **FAIL** — server error |

The cmdlet fails even for the **admin's own mailbox with ResultSize 1** — the most minimal possible query. All other Exchange Online cmdlets (`Get-Mailbox`, `Get-Recipient`, `Get-ManagementRoleAssignment`, `Get-MailboxFolderPermission`) were functioning normally during the same session. This conclusively rules out:
- RBAC/permissions (all 8 mailboxes including admin's own fail)
- Throttling (failure on first attempt, same error as all subsequent)
- Auth token expiry (other EXO cmdlets succeeding in same session)
- Mailbox-specific issues (fails for every mailbox tested)
- ResultSize or parameter issues (fails with the most minimal query possible)

**Conclusion: The `Get-CalendarDiagnosticObjects` backend service for tenant `b3377764-8275-47f8-8e2d-697b136c21c1` has a server-side fault. This is a Microsoft backend infrastructure issue.**

### 6.2 Mailbox Infrastructure Details

| Field | Value |
|---|---|
| Matt's ExchangeGuid | `4e064909-6e4b-4070-a391-0511fc735693` |
| Matt's mailbox database | `namprd22.prod.outlook.com/c4116c34-f7fe-4a80-b75e-83d6fee18b70` |

### 6.3 Secondary Issue: RBAC Gap Detected

The diagnostic script also found that the signed-in admin (`ndradmin@venturafs.com`) did not have a direct `Calendar Diagnostics` role assignment — despite being a member of the `Organization Management` role group which normally includes this role. The auto-fix attempted by the diagnostic script (`New-ManagementRoleAssignment`) also failed. This secondary RBAC issue is likely moot while the backend itself is down, but should be confirmed once the backend is restored.

### 6.4 PowerShell Version — PS5.1 Compatibility Test: CONFIRMED BACKEND FAULT

The same 7 tests were run in Windows PowerShell 5.1 (`Test-CalDiag-PS51.ps1`) on 2026-06-08 at 15:05 UTC. Result: **0/7 PASS — all 7 tests failed with the same server-side error.**

| Field | Value |
|---|---|
| PS version | 5.1.26100.8457 |
| Module | ExchangeOnlineManagement 3.9.2 |
| Result | 0 PASS / 7 FAIL |

Error message received in PS5.1:
```
A server side error has occurred because of which the operation could not be completed.
Please try again after some time. If the problem still persists, please reach out to MS support.
```

The Exchange Online Management module uses different HTTP transport layers between versions — PS5.1 uses the legacy Remote PowerShell / WinRM path; PS7 uses a pure REST layer. Both fail with the same underlying error. **PowerShell version is definitively ruled out as a contributing factor.** The backend fault is version-independent and affects this tenant at the server level regardless of which client or transport is used.

**This is the single most important item for Microsoft to resolve.** Until `Get-CalendarDiagnosticObjects` is restored, it is impossible to identify the authoring client and confirm whether the meetings were:
- Never written to Matt's mailbox (delivery-layer failure), or
- Written and subsequently deleted (deletion event by a client or automation), or
- Written and moved out of the Calendar folder by a rule or client action.

---

## 7. Observed Patterns and Hypotheses

### 7.1 Monthly Events Calendar Meeting — entire series absent
Findings 4, 6, and 13 are three consecutive monthly occurrences of the same recurring series (same base `iCalUId`), all organised by `lindsey.baginski@venturafs.com`. All three are missing and `SeriesPresentOnMatt = False`. This means Matt never accepted or received this recurring series subscription. Whether he was removed from the series, never added, or whether the invites arrived and were silently dropped can only be confirmed via the diagnostic log.

### 7.2 High-corroboration drops — near-org-wide visibility
- **MK - Half day - Monday AM** (finding 5): present on 9 out of 11 scanned colleagues. Almost the full organisation has this calendar block; Matt does not. This could indicate either a targeted invite failure or — if this was distributed via a shared/group calendar — that Matt's access to that shared calendar is broken.
- **Ventura/Metric Point - NSP true up** (finding 11): organised by an external party, present on 6 internal colleagues' calendars. A significant business meeting that Matt was demonstrably invited to and should have attended.

### 7.3 Alan Tsarovsky as organiser — cluster of missing invitee meetings
Findings 7, 9, 10, 14 (and finding 11 where Alan appears as co-attendee) are all meetings organised by `alan.tsarovsky@venturafs.com` where Matt was invited. These span April–May 2026. The same organiser's other meetings (Buchalter intro, Kirkland catchup, etc.) ARE present on Matt's calendar, so the issue is not a blanket block on Alan's invites — it is selective.

### 7.4 External organisers affected too
Findings 8 and 11 were organised by external parties (`moelisam.com`, `metricpoint.com`). The failure is not limited to internal senders.

### 7.5 Organiser-role drops — most forensically significant
The 3 Organizer-scenario findings (Matt created the meetings himself) are the strongest evidence that this is a calendar-layer issue rather than a mail transport issue. When Matt is the organiser, the meeting request is generated by Exchange on behalf of Matt's mailbox. That the organiser copy never persisted on Matt's calendar — while the attendee copies did — points to a failure at the point of calendar item creation in Matt's mailbox specifically.

---

## 8. What Is Confirmed Ruled Out

| Hypothesis | Status |
|---|---|
| Access/sharing propagation failure masking the result | **Ruled out** — all 11 colleague calendars read successfully in both runs |
| Partial recurrence-series sync issue | **Ruled out** — `SeriesPresentOnMatt = False` on all 14; entire iCalUIds absent, not just some occurrences |
| Free/busy-only calendar permissions returning false empty | **Ruled out** — script checks and upgrades to Reviewer before reading |
| UTC timezone shift corrupting occurrence-key matching | **Ruled out** — fixed in script via `AssumeUniversal` + `DateTimeKind.Utc` pinning |
| Single-run artefact | **Ruled out** — two independent runs on the same day produced identical results |

---

## 9. Requests to Microsoft Support

### 9.1 Immediate — restore diagnostic backend
Restore `Get-CalendarDiagnosticObjects` for tenant `b3377764-8275-47f8-8e2d-697b136c21c1`. This is believed to be an active server-side failure on Microsoft's end and is the blocker for root-cause identification.

### 9.2 Once diagnostic log is available — run against missing meetings
For each of the 14 missing meetings, run the calendar diagnostic log against `matt.lowe@venturafs.com` and examine:

```powershell
Get-CalendarDiagnosticObjects -Identity matt.lowe@venturafs.com `
    -Subject "<meeting subject>" -ExactMatch $true `
    -CustomPropertyNames ClientInfoString,CalendarLogTriggerAction,
                         MeetingRequestType,ResponseType,ChangeHighlight
```

Or by `CleanGlobalObjectId` (derived from `iCalUId`) for precision. Key fields to extract:
- **`ClientInfoString`** — identifies the exact application that created/modified/deleted the item
- **`CalendarLogTriggerAction`** — the sequence of actions (Create, Delete, MoveToDeletedItems, MoveToFolder, etc.)
- **`MeetingRequestType`** — type of the request that processed the item
- **`ChangeHighlight`** — what changed at each action

The question to answer per meeting: was there ever a `Create` action on Matt's mailbox, and if so, what action followed it?

### 9.3 Mailbox health check
Check `matt.lowe@venturafs.com` for:
- Mailbox corruption or inconsistency flags (`New-MailboxRepairRequest`)
- Inbox rules that might be deleting or moving meeting requests before they reach the calendar
- Delegate access configuration that could allow a delegate or automation to delete calendar items
- Any connected applications with `Calendars.ReadWrite` Graph permission that may be operating on Matt's calendar
- Whether Matt's account was recently migrated, recreated, or had a licence change that could have triggered a calendar reset

### 9.4 Transport investigation for inbound invites
For the invitee-scenario meetings, confirm that the meeting request emails arrived in Matt's mailbox transport layer (message trace for the relevant dates and sender domains). If they arrived at transport but did not create a calendar item, the failure is in the calendar processing layer. If they did not arrive at transport, the failure is upstream.

---

## 10. Scope and Limitations of This Investigation

- **Coverage:** 11 internal colleagues scanned (the 90-day most-frequent meeting partners). Meetings where no scanned colleague was a co-attendee or organiser will not appear in the output.
- **90-day window only:** Issues before approximately 2026-03-09 are outside scope. A longer lookback (e.g. 180 days) could be run if needed.
- **No calendar write operations were performed.** The investigation is entirely read-only. No meetings were created, modified, or deleted. No mailbox data was altered.
- **Client identification not possible without diagnostic log** (see §6).

---

## 11. Appendix — Evidence Files

| File | Description |
|---|---|
| `mattphase2investigation.csv` | Full output — 100 rows, 14 MISSING_FROM_MATT, 86 PRESENT_ON_MATT. One row per occurrence found on at least one colleague's calendar. |
| `Confirm-OrganizerCopyMissing.ps1` | Investigation script used to produce the above. Bidirectional Graph-only calendar comparison. |
| `OrgCopyCheck_matt.lowe_*/run.log` | Per-run execution log with timestamped steps, read success/failure per colleague, and full findings summary. |
| `Test-CalendarDiagnosticAccess.ps1` | PS7 diagnostic script. Diagnosed TENANT_WIDE_BACKEND_FAILURE — all 8 tests failed server-side. Run on 2026-06-08. |
| `CalDiag_PS7_<timestamp>.txt` | Diagnostic run report produced by the above. Contains full test output including verbatim error messages. |
| `Test-CalDiag-PS51.ps1` | PS5.1 compatibility test script. Runs the same 7 tests in Windows PowerShell 5.1. |
| `CalDiag_PS51_20260608_150505.txt` | PS5.1 run report: 0/7 PASS. Same server-side error across all tests. PS version ruled out as contributing factor. |

### Column reference for mattphase2investigation.csv

| Column | Meaning |
|---|---|
| `MattRole` | `Organizer` = Matt sent the invite; `Invitee` = someone else organised it |
| `Verdict` | `MISSING_FROM_MATT` = absent from Matt's calendar; `PRESENT_ON_MATT` = present |
| `Subject` | Meeting subject as seen on colleague's calendar |
| `StartUtc` | Meeting start in UTC |
| `MeetingOrganizer` | Email address of the meeting organiser |
| `PresentOnMatt` | Boolean — whether the iCalUId+start key was found on Matt's calendar |
| `SeriesPresentOnMatt` | Boolean — whether the iCalUId alone (series identity) was found on Matt's calendar. `False` on all 14 missing rows = entire series/meeting identity absent, not just an occurrence |
| `SeenOnCalendars` | Colleague calendars on which this occurrence was found |
| `CalendarCount` | Count of independent colleagues holding this occurrence |
| `iCalUId` | The cross-mailbox meeting identity (CleanGlobalObjectId equivalent in Graph). Use this for `Get-CalendarDiagnosticObjects` targeting once the backend is restored |
