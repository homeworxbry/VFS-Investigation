# Code Review & Fixes — Calendar Forensics Scripts

Review of `Invoke-OrganizerCalendarForensics.ps1` (main) and
`Confirm-OrganizerCopyMissing.ps1` (Graph-only companion), addressing the asks in
the handover brief. Both files were rewritten in place with the fixes below.

Scripts could not be executed here (no PowerShell runtime in this environment, and
the live cmdlets require the M365 tenant). Changes were validated by careful static
review and structural checks. **Run both against the tenant to confirm behaviour**,
especially the main script's classification/export path (previously never reached at
runtime) using a known-good meeting per ask #4 of the brief.

---

## A. Companion — `Confirm-OrganizerCopyMissing.ps1`

### A1. (CRITICAL) iCalUId-only correlation collapsed recurring series → missed the symptom
`calendarView` expands a recurring series into per-occurrence instances that all
share **one** `iCalUId`. The old code keyed presence by `iCalUId` alone, so if Matt
was missing 3 of 10 occurrences of a weekly series but kept the other 7, his calendar
still contained that `iCalUId` → every occurrence was reported `PRESENT_ON_ORGANIZER`.
**The exact failure the tool exists to detect (individual missing occurrences) was
silently classified as present.**
**Fix:** correlate by `(iCalUId + occurrence-start-to-the-minute)` via `Get-OccurrenceKey`.
Both the organizer presence set and the attendee-found set are now occurrence-level.
A new `SeriesPresentOnOrg` column flags the strong signal "series exists on organizer
but THIS occurrence is missing".

### A2. (CRITICAL) Access failures masqueraded as a clean bill of health
If Reviewer sharing hadn't propagated, the attendee read returned `403`;
`Invoke-GraphPaged` swallowed it and returned an empty list. That attendee then
contributed zero events, and the summary printed "no missing copies found" — a
forensic "all clear" produced by an **access failure**, not by evidence.
**Fix:** `Invoke-GraphPaged` and `Get-CalendarViewChunked` now return a result object
`{ Ok; Reason; Items/Events }`. Per-attendee read status is tracked (`$readOk` /
`$readFailed`); attendees whose read failed are **excluded and reported**, and the
final verdict is explicitly **caveated as INCOMPLETE** if any read failed. A clean
result is only declared trustworthy when every scanned attendee was read.
Additionally, the organizer's OWN calendar read failing now **aborts** (no baseline →
no valid comparison).

### A3. `Grant-CalendarShareFor` accepted any pre-existing permission, including free/busy-only
A pre-existing `AvailabilityOnly` / `LimitedDetails` right was treated as success, but
a `calendarView` under those returns no item detail (no subject/organizer/uid) →
silently empty.
**Fix:** only the roles that actually permit item read (`Reviewer`, `Author`,
`Editor`, …) count as already-shared; a lower existing right is **upgraded to
Reviewer** via `Set-MailboxFolderPermission`.

### A4. Timezone shift bug in `Get-EventStartUtc`
With `Prefer: outlook.timezone="UTC"`, Graph returns an offset-less dateTime; `[datetime]`
tags it `Unspecified`, so `.ToUniversalTime()` subtracted the host's local offset
(e.g. a 14:30Z meeting became 19:30Z on a UTC-5 host). Cosmetic before, but **load-bearing
now** that the correlation key includes the start instant (A1).
**Fix:** `ConvertTo-Utc` pins `DateTimeKind.Utc` (or parses with
`AssumeUniversal | AdjustToUniversal`, InvariantCulture).

### A5. Organizer match was too strict
`[string]$_.organizer.emailAddress.address -ieq $TargetOrganizer` missed meetings
stamped with a proxy/alias SMTP, dropping them from the comparison entirely.
**Fix:** resolve the organizer's full address set from `Get-Mailbox` (`EmailAddresses`
+ `PrimarySmtpAddress`) and match (trimmed, lower-cased) against that set via
`Test-IsOrganizer`. Falls back to the UPN if the mailbox lookup fails.

### A6. Asymmetric cancellation handling
The attendee side filtered `isCancelled`; the organizer side counted cancelled
tombstones as "present", skewing the comparison.
**Fix:** cancelled events are excluded on **both** sides.

### A7. (the open `Argument types do not match` crash)
Most consistent with a member access on a deserialized Graph node that, under
`-OutputType PSObject`, can surface as an `IDictionary` rather than a `PSCustomObject`
(the crashing path is the only one that walks `attendees`, which the working sibling
never selected).
**Fix:** all nested Graph access now goes through `Get-Prop`, which reads either shape
safely; attendee iteration guards `$null` entries. The defensive accessors plus the
`Ok/Reason` plumbing also make any future failure report *where* it happened instead
of dying with a typeless .NET error.

### A8. Minor hardening
- `Sort-Object { [int]$_.Value } -Descending` for the attendee tally (type-stable).
- `@($freq.GetEnumerator() ...)` result forced to an array so `.Count` is always valid.

---

## B. Main — `Invoke-OrganizerCalendarForensics.ps1`

> The enum/field assumptions called out in the brief were **verified correct** against
> Microsoft docs and CSS-Exchange: `ResponseType=1` = organizer copy;
> `MeetingRequestType` codes (1/65536/131072/262144/524288/1048576); and
> `CalendarLogTriggerAction` is emitted as readable strings without `-ShouldDecodeEnums`.
> The bugs were in the *logic that consumed* those fields.

### B1. Unanchored `-match 'Create'` over-matched → false `PRESENT_OK`
`CalendarLogTriggerAction -match 'Create'` matched any action containing "create".
**Fix:** anchored `Test-IsCreateAction` (`^Create$`).

### B2. Removal detection missed `MoveToFolder`
An appointment moved out of the Calendar folder logs `MoveToFolder`, which the old
`Delete|MoveToDeletedItems|SoftDelete|HardDelete` regex did not treat as removal →
a removed copy read as `PRESENT_OK`.
**Fix:** anchored `Test-IsRemovalAction` (`^(Delete|MoveToDeletedItems|SoftDelete|HardDelete|MoveToFolder)$`).

### B3. Culture-sensitive date reparse drove the "last action" sort
`Get-SortDate` used `[datetime]::Parse([string]$Value)` (current culture, no
`AssumeUniversal`), so on non-US hosts ordering could invert or fail (→ `MinValue`),
making "last action" the wrong object and flipping the verdict.
**Fix:** parse with `InvariantCulture` + `AssumeUniversal|AdjustToUniversal`, and pass
the raw `[datetime]` through untouched.

### B4. Organizer attribution relied on `ResponseType` on appointment rows
`ResponseType` is frequently blank on `IPM.Appointment` rows, so genuinely
organizer-owned meetings were dropped from the "organizer-only" report.
**Fix:** `OrganizerMeeting` is now `ResponseType ∈ {1,Organizer}` **OR** the meeting key
matches a Graph `isOrganizer=true` seed (`$organizerSeedKeys`). This also removes the
old inconsistency where data-less seed rows were hardcoded `OrganizerMeeting=$true`
while real rows could be marked non-organizer.

### B5. Classification: Graph could override an EXO deletion; `MoveToFolder` fell through to WARN
Old branch 1 (`$apptCreate -and $currentlyThere`, where `$currentlyThere` included
`$graphPresent`) let a fuzzy same-subject/same-day Graph match force `PRESENT_OK` over
an EXO-confirmed deletion; and a `MoveToFolder` removal fell through to WARN-level
`AMBIGUOUS` instead of ERROR-level `CREATED_THEN_REMOVED`.
**Fix:** classification is now strictly EXO-authoritative —
`PRESENT_OK` requires `$exoPresent` (pure EXO); a created-but-not-present item is
`CREATED_THEN_REMOVED` when a removal action is seen, or `AMBIGUOUS` when Graph still
shows it (conservative: avoids a false removal claim on an in-mailbox folder move).
Graph is used only to corroborate the no-diagnostic cases.

### B6. §7 seed-coverage deduped by subject-only → false alarms AND suppressed gaps
Subject-only keys (a) raised false `NO_DIAGNOSTIC_DATA` when diag rows carried a
slightly different normalized subject, and (b) collapsed distinct-date instances so a
genuinely uncovered instance was suppressed by a same-subject sibling. It also emitted
`NO_DIAGNOSTIC_DATA` rows with `InCalendarNow=$true` (self-contradictory).
**Fix:** coverage keys are now subject **+ date** (`Get-NormalizedKey`), built from each
GOID group's representative start; a seed still present per Graph is labelled
`IN_CALENDAR_NO_DIAG` (lifecycle outside window) rather than `NO_DIAGNOSTIC_DATA`.

### B7. Consistency / hardening
- Same `Get-Prop` / `ConvertTo-Utc` defensive Graph helpers as the companion;
  fixes the identical timezone bug in `Get-OrganizerCalendarSeed`.
- Removed the dead `Get-SentMeetingRequestSeed` function and `-IncludeSentItems`
  parameter (the Graph mail seed was intentionally dropped to avoid a broad
  `FullAccess` grant; the EXO diagnostic log already carries the request signal).
- Role-group membership check compares role **name strings** (the old `-notin
  (...).Roles` compared a string against role objects and would never match).
- Renamed `ApptDeletedLater` → `ApptRemovedLater` to match the broader removal set.

---

## C. Still open / not addressed here (needs the tenant)

1. **§5 backend blocker** — `Get-CalendarDiagnosticObjects` server-side failure is not a
   code issue; the main script will run unchanged once Microsoft restores it. Pursue the
   support case from the brief.
2. **Ask #3 — `MeetingID` (CleanGlobalObjectId) collection path.** Not added yet; it is
   the right hardening once the backend is healthy (derive `MeetingID` from a working
   subject query's output, per CSS-Exchange `Get-CalendarDiagnosticObjectsSummary`).
   Recommend doing this after a successful end-to-end run validates the current path.
3. **Ask #4 — end-to-end validation** on a known-good meeting
   (`OrganizerMeeting=True, OrganizerApptCreated=True, ExoPresent=True,
   Classification=PRESENT_OK`) must be done against the live tenant before trusting any
   `MISSING_*` rows.
4. **Ask #5 — attendee-mailbox expansion** of the main tool (most definitive proof) — not
   implemented; the companion already provides the attendee-side cross-check.
