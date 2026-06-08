#Requires -Version 7.0
<#
================================================================================
 Confirm-OrganizerCopyMissing.ps1   (Graph-only companion)
================================================================================
 PURPOSE
   Confirm the reported failure -- "invitees received the meeting, but the
   organizer's own calendar copy never appeared" -- WITHOUT using
   Get-CalendarDiagnosticObjects (which is currently failing server-side in this
   tenant). Uses only Microsoft Graph calendarView, which is working.

 METHOD
   1. Read the organizer's calendar (last N days) to (a) learn who he invites and
      (b) build the set of meeting OCCURRENCES currently present on HIS calendar.
   2. Auto-discover his frequent internal attendees from those meetings.
   3. Grant the signed-in admin least-privilege Reviewer on each attendee's
      Calendar (idempotent; upgrades a too-low existing right), wait for sharing
      to propagate, then read each attendee's calendar.
   4. Collect meetings ORGANIZED BY the target that appear on attendees' calendars,
      and check each OCCURRENCE against the organizer's own calendar.
        present on attendee + ABSENT from organizer = MISSING_FROM_ORGANIZER
        present on attendee + present on organizer  = PRESENT_ON_ORGANIZER

 WHY OCCURRENCE-LEVEL (not iCalUId alone)
   calendarView expands a recurring series into per-occurrence instances that all
   share ONE iCalUId. Correlating by iCalUId alone would report a series as
   "present" even if individual occurrences went missing from the organizer --
   exactly the failure we are hunting. We therefore correlate by
   (iCalUId + occurrence-start-minute), and also surface the series-level signal.

 WHAT THIS PROVES (and doesn't)
   Proves the SYMPTOM definitively (organizer copy missing while attendees have it).
   Does NOT reveal the authoring client (ClientInfoString) -- that root-cause detail
   lives only in the calendar diagnostic log, which needs the backend restored.

 TRUSTWORTHINESS GUARD
   A forensic "all clear" is only meaningful if every attendee calendar was
   actually READ. Reads that fail (e.g. sharing not yet propagated -> 403) are
   tracked separately and the final verdict is explicitly caveated if any read
   was incomplete -- access failure must never masquerade as "no problem".

 AUTH / PERMISSIONS
   Interactive delegated admin auth (session-aware; prompts only if needed).
   Requires Graph scopes Calendars.Read.Shared + User.Read.All, and Exchange
   Online admin rights to grant Reviewer on attendee calendars. Sessions are not
   disconnected.

 RUN
   .\Confirm-OrganizerCopyMissing.ps1
   .\Confirm-OrganizerCopyMissing.ps1 -TargetOrganizer matt.lowe@venturafs.com -LookbackDays 90 -MaxAttendees 15
================================================================================
#>

[CmdletBinding()]
param(
    [string]$TargetOrganizer = 'matt.lowe@venturafs.com',
    [int]$LookbackDays       = 90,
    [string]$OutputRoot      = (Get-Location).Path,

    # Only scan attendees in this SMTP domain (defaults to the organizer's domain).
    # Set to '' to scan attendees in any domain (still must be mailboxes you can share).
    [string]$AttendeeDomain  = '',

    # Cap on how many distinct attendees to scan (most-frequent first).
    [int]$MaxAttendees       = 15,

    # Seconds to wait after granting Reviewer before reading attendee calendars,
    # so freshly-granted sharing has time to propagate to Graph.
    [int]$ShareWaitSeconds   = 120
)

$ErrorActionPreference = 'Stop'
if (-not $AttendeeDomain) { $AttendeeDomain = ($TargetOrganizer -split '@')[1] }

# ---- Output + logging ------------------------------------------------------
$stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$tag    = ($TargetOrganizer -split '@')[0]
$OutDir = Join-Path $OutputRoot ("OrgCopyCheck_{0}_{1}" -f $tag, $stamp)
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$LogFile    = Join-Path $OutDir 'run.log'
$ResultCsv  = Join-Path $OutDir 'organizer_copy_findings.csv'

function Write-Log {
    param([Parameter(Mandatory)][string]$Message,
          [ValidateSet('INFO','WARN','ERROR','OK','STEP')][string]$Level='INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    Add-Content -Path $LogFile -Value $line
}

# Diagnostic trap: any unhandled terminating error prints its TYPE, the exact
# script position, and the call stack -- so a bare ".NET" message can never again
# leave us guessing which line failed.
trap {
    Write-Host ("FATAL {0}: {1}" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ForegroundColor Red
    Write-Host ("  at: {0}" -f $_.InvocationInfo.PositionMessage) -ForegroundColor Red
    Write-Host ("  stack:`n{0}" -f $_.ScriptStackTrace) -ForegroundColor DarkRed
    if ($LogFile) {
        Add-Content -Path $LogFile -Value ("FATAL {0}: {1}`n{2}`n{3}" -f $_.Exception.GetType().FullName, $_.Exception.Message, $_.InvocationInfo.PositionMessage, $_.ScriptStackTrace)
    }
    break
}

Write-Log "Organizer-copy check started. Organizer=$TargetOrganizer Lookback=$LookbackDays Domain=$AttendeeDomain" 'STEP'

# ---- Modules ---------------------------------------------------------------
foreach ($m in 'ExchangeOnlineManagement','Microsoft.Graph.Authentication') {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Log "Installing module '$m' for current user..." 'WARN'
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module $m -ErrorAction Stop
    Write-Log "Module '$m' ready." 'OK'
}

# ---- Connections (reuse if present) ----------------------------------------
$exo = $null
try { $exo = Get-ConnectionInformation -ErrorAction Stop } catch { $exo = $null }
if (-not ($exo | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' })) {
    Write-Log "Connecting to Exchange Online..." 'STEP'
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}
Write-Log "Exchange Online ready." 'OK'

$needScopes = @('Calendars.Read.Shared','User.Read.All')
$ctx = $null
try { $ctx = Get-MgContext -ErrorAction Stop } catch { $ctx = $null }
$missingScopes = @()
if ($ctx) { $missingScopes = $needScopes | Where-Object { $_ -notin $ctx.Scopes } }
if (-not $ctx -or $missingScopes.Count -gt 0) {
    Write-Log "Connecting to Microsoft Graph (scopes: $($needScopes -join ', '))..." 'STEP'
    Connect-MgGraph -Scopes $needScopes -NoWelcome -ErrorAction Stop
}
Write-Log "Microsoft Graph ready." 'OK'

# ---- Helpers ---------------------------------------------------------------
$startDt = (Get-Date).AddDays(-$LookbackDays)
$endDt   = (Get-Date).AddDays(1)

# Safe nested-property reader. Invoke-MgGraphRequest -OutputType PSObject can yield
# PSCustomObject OR (for some nodes) IDictionary; this reads either shape without
# triggering ".NET Argument types do not match" on a wrong-typed member access.
function Get-Prop {
    param($Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

# Parse a Graph dateTime that was requested with Prefer: outlook.timezone="UTC".
# Such values carry NO offset, so [datetime] would tag them Unspecified and
# ToUniversalTime() would wrongly shift by the host's local offset. Pin Kind=Utc.
function ConvertTo-Utc {
    param($Value)
    if (-not $Value) { return $null }
    if ($Value -is [datetime]) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    if ([datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return [datetime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    return $null
}

function Get-EventStartUtc {
    param($Event)
    $start = Get-Prop $Event 'start'
    return (ConvertTo-Utc (Get-Prop $start 'dateTime'))
}

# Occurrence key: an iCalUId plus the occurrence start to the minute. This makes a
# recurring series compare per-occurrence instead of collapsing to one identity.
function Get-OccurrenceKey {
    param([string]$ICalUId, $StartUtc)
    $d = if ($StartUtc -is [datetime]) { $StartUtc.ToString('yyyy-MM-ddTHH:mm') } else { 'nostart' }
    return "$ICalUId|$d"
}

# Some Microsoft.Graph.Authentication builds throw a CLIENT-SIDE ".NET Argument
# types do not match" while binding a custom -Headers hashtable. The Prefer header
# only asks Graph to render times in UTC; calendarView already returns UTC by
# default, and ConvertTo-Utc copes with offset-bearing values, so if the header is
# rejected we transparently retry WITHOUT it (once), then stop trying it.
$script:GraphPreferHeaderOk = $true
function Invoke-GraphGet {
    param([Parameter(Mandatory)][string]$Uri, [hashtable]$Headers = @{})
    if ($script:GraphPreferHeaderOk -and $Headers -and $Headers.Count -gt 0) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers -OutputType PSObject -ErrorAction Stop
        } catch {
            $em = $_.Exception.Message
            # Only treat a local binding/type fault as "headers unsupported"; let real
            # HTTP errors (401/403/404/throttling/etc.) propagate to the retry logic.
            if ($em -match 'Argument types do not match|does not match|IDictionary|header') {
                Write-Log ("  -Headers not accepted by this Graph module ({0}); retrying without the Prefer header (times still returned in UTC)." -f $em) 'WARN'
                $script:GraphPreferHeaderOk = $false
            } else {
                throw
            }
        }
    }
    return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

# Paged Graph GET. Returns a result object so callers can tell a genuine empty
# result apart from an access/permission failure (critical for the verdict).
function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri, [hashtable]$Headers = @{}, [int]$MaxAttempts = 4)
    $items  = New-Object System.Collections.Generic.List[object]
    $ok     = $true
    $reason = ''
    $next   = $Uri
    while ($next) {
        $attempt = 0
        $resp    = $null
        $pageOk  = $false
        while ($true) {
            $attempt++
            try { $resp = Invoke-GraphGet -Uri $next -Headers $Headers; $pageOk = $true; break }
            catch {
                $msg = $_.Exception.Message
                if ($msg -match 'Forbidden|Unauthorized|\b401\b|\b403\b|\b404\b|\b400\b|NotFound|denied') {
                    Write-Log ("    permanent error ({0}); not retrying." -f $msg) 'WARN'
                    $ok = $false; $reason = $msg; break
                }
                if ($attempt -ge $MaxAttempts) { Write-Log ("    failed after {0} attempts ({1})." -f $attempt, $msg) 'WARN'; $ok = $false; $reason = $msg; break }
                $wait = @(5,15,30,60)[[math]::Min($attempt-1,3)]
                Write-Log ("    transient ({0}); retry {1}/{2} in {3}s..." -f $msg, $attempt, $MaxAttempts, $wait) 'WARN'
                Start-Sleep -Seconds $wait
            }
        }
        if (-not $pageOk) { break }
        $val = Get-Prop $resp 'value'
        if ($val) { foreach ($v in $val) { $items.Add($v) } }
        $next = [string](Get-Prop $resp '@odata.nextLink')
    }
    return [pscustomobject]@{ Ok = $ok; Reason = $reason; Items = $items }
}

# Pull a mailbox's calendar over the window in 15-day slices (recurrence expansion
# stays light). Returns Ok=$false if ANY slice failed, so the caller never treats a
# partial/blocked read as a complete one.
function Get-CalendarViewChunked {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$Select)
    $hdr     = @{ Prefer = 'outlook.timezone="UTC"' }
    $seen    = @{}
    $events  = New-Object System.Collections.Generic.List[object]
    $ok      = $true
    $reason  = ''
    $cursor  = $startDt
    while ($cursor -lt $endDt) {
        $sliceEnd = $cursor.AddDays(15)
        if ($sliceEnd -gt $endDt) { $sliceEnd = $endDt }
        $sEnc = [uri]::EscapeDataString($cursor.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $eEnc = [uri]::EscapeDataString($sliceEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $uri  = "https://graph.microsoft.com/v1.0/users/$Mailbox/calendarView?startDateTime=$sEnc&endDateTime=$eEnc&" +
                '$top=100&$select=' + $Select
        $page = Invoke-GraphPaged -Uri $uri -Headers $hdr
        if (-not $page.Ok) { $ok = $false; $reason = $page.Reason }
        foreach ($e in $page.Items) {
            $id = [string](Get-Prop $e 'id')
            if ($id -and -not $seen.ContainsKey($id)) { $seen[$id] = $true; $events.Add($e) }
        }
        $cursor = $sliceEnd
    }
    return [pscustomobject]@{ Ok = $ok; Reason = $reason; Events = $events }
}

# Access rights that actually permit reading calendar ITEMS (subject/organizer/uid).
# Free/busy-only rights (AvailabilityOnly, LimitedDetails) and Contributor are NOT
# enough -- a calendarView under those returns nothing useful.
$ItemReadRoles = @('Reviewer','Author','NonEditingAuthor','Editor','PublishingAuthor','PublishingEditor','Owner')

function Grant-CalendarShareFor {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$AdminUpn)
    $folder = "$Mailbox`:\Calendar"
    $cur = $null
    try { $cur = Get-MailboxFolderPermission -Identity $folder -User $AdminUpn -ErrorAction Stop } catch { $cur = $null }
    if ($cur) {
        $rights = @($cur.AccessRights | ForEach-Object { [string]$_ })
        if ($rights | Where-Object { $_ -in $ItemReadRoles }) { return $true }   # already item-readable
        # Existing right is too low (e.g. AvailabilityOnly) -> upgrade in place.
        try {
            Set-MailboxFolderPermission -Identity $folder -User $AdminUpn -AccessRights Reviewer -ErrorAction Stop | Out-Null
            Write-Log ("  upgraded '{0}' from [{1}] to Reviewer." -f $Mailbox, ($rights -join ',')) 'INFO'
            return $true
        } catch {
            Write-Log ("  could not upgrade calendar permission for {0}: {1}" -f $Mailbox, $_.Exception.Message) 'WARN'
            return $false
        }
    }
    try {
        Add-MailboxFolderPermission -Identity $folder -User $AdminUpn -AccessRights Reviewer -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log ("  could not share calendar for {0}: {1}" -f $Mailbox, $_.Exception.Message) 'WARN'
        return $false
    }
}

$AdminUpn = (Get-MgContext).Account

# ---- 0. Resolve the organizer's address aliases (for robust organizer match) ----
# A meeting "organized by Matt" may stamp a proxy/alias SMTP rather than the exact
# UPN we were passed; match against his whole address set, case- and space-insensitively.
$organizerAddrs = @{}
$organizerAddrs[$TargetOrganizer.Trim().ToLowerInvariant()] = $true
try {
    $mbx = Get-Mailbox -Identity $TargetOrganizer -ErrorAction Stop
    foreach ($pa in @($mbx.EmailAddresses)) {
        $s = [string]$pa
        if ($s -match '^smtp:(.+)$') { $organizerAddrs[$Matches[1].Trim().ToLowerInvariant()] = $true }
    }
    if ($mbx.PrimarySmtpAddress) { $organizerAddrs[([string]$mbx.PrimarySmtpAddress).Trim().ToLowerInvariant()] = $true }
    Write-Log ("Organizer address set resolved ({0} addresses)." -f $organizerAddrs.Keys.Count) 'OK'
} catch {
    Write-Log ("Could not resolve organizer proxy addresses ({0}); matching on '{1}' only." -f $_.Exception.Message, $TargetOrganizer) 'WARN'
}
function Test-IsOrganizer {
    param($Event)
    $org  = Get-Prop $Event 'organizer'
    $ea   = Get-Prop $org 'emailAddress'
    $addr = ([string](Get-Prop $ea 'address')).Trim().ToLowerInvariant()
    if (-not $addr) { return $false }
    return $organizerAddrs.ContainsKey($addr)
}

# ---- 1. Organizer's own calendar: presence sets + attendee discovery --------
Write-Log "Reading organizer calendar..." 'STEP'
$orgSelect = 'id,iCalUId,subject,start,end,organizer,isOrganizer,isCancelled,attendees'
$orgRead   = Get-CalendarViewChunked -Mailbox $TargetOrganizer -Select $orgSelect
$orgEvents = $orgRead.Events
if (-not $orgRead.Ok) {
    Write-Log ("Could not fully read the organizer's OWN calendar ({0}). Without it there is no baseline to compare against -- aborting." -f $orgRead.Reason) 'ERROR'
    return
}

# Presence keyed two ways: series-level (iCalUId) and occurrence-level (iCalUId|start).
# Cancelled tombstones are excluded so a cancelled meeting is not counted as "present".
$organizerSeries     = @{}
$organizerOccurrence = @{}
foreach ($e in $orgEvents) {
    if (Get-Prop $e 'isCancelled') { continue }
    $uid = [string](Get-Prop $e 'iCalUId')
    if (-not $uid) { continue }
    $organizerSeries[$uid] = $true
    $organizerOccurrence[(Get-OccurrenceKey -ICalUId $uid -StartUtc (Get-EventStartUtc $e))] = $true
}
Write-Log ("  organizer calendar holds {0} events ({1} distinct series)." -f (@($orgEvents).Count), $organizerSeries.Keys.Count) 'INFO'

# Tally attendees from meetings the organizer owns.
$freq = @{}
foreach ($e in $orgEvents) {
    if (-not (Get-Prop $e 'isOrganizer')) { continue }
    foreach ($a in @(Get-Prop $e 'attendees')) {
        if ($null -eq $a) { continue }
        $ea   = Get-Prop $a 'emailAddress'
        $addr = ([string](Get-Prop $ea 'address')).Trim()
        $type = [string](Get-Prop $a 'type')
        if (-not $addr) { continue }
        if ($type -eq 'resource') { continue }                                   # skip rooms/equipment
        if ($organizerAddrs.ContainsKey($addr.ToLowerInvariant())) { continue }  # skip organizer himself
        if ($AttendeeDomain -and ($addr -notlike "*@$AttendeeDomain")) { continue }
        if ($freq.ContainsKey($addr)) { $freq[$addr]++ } else { $freq[$addr] = 1 }
    }
}
$attendees = @($freq.GetEnumerator() | Sort-Object { [int]$_.Value } -Descending | Select-Object -First $MaxAttendees -ExpandProperty Name)
Write-Log ("  discovered {0} distinct attendees; scanning top {1}." -f $freq.Keys.Count, $attendees.Count) 'OK'
if ($attendees.Count -eq 0) { Write-Log "No internal attendees discovered; nothing to compare. Exiting." 'WARN'; return }

# ---- 2. Grant Reviewer on each attendee, then wait for propagation ---------
Write-Log "Granting least-privilege Reviewer on attendee calendars..." 'STEP'
$granted = @()
foreach ($a in $attendees) {
    if (Grant-CalendarShareFor -Mailbox $a -AdminUpn $AdminUpn) { $granted += $a; Write-Log "  shared: $a" 'INFO' }
}
if (@($granted).Count -eq 0) { Write-Log "No attendee calendars could be shared. Exiting." 'WARN'; return }
Write-Log ("Waiting {0}s for sharing to propagate to Graph..." -f $ShareWaitSeconds) 'STEP'
Start-Sleep -Seconds $ShareWaitSeconds

# ---- 3. Scan attendee calendars for organizer-owned meeting OCCURRENCES -----
Write-Log "Scanning attendee calendars for meetings organized by $TargetOrganizer..." 'STEP'
$attSelect   = 'id,iCalUId,subject,start,end,organizer,isOrganizer,isCancelled'
$found       = @{}   # occurrenceKey -> record
$readOk      = @()
$readFailed  = @()
foreach ($a in $granted) {
    Write-Log "  reading $a ..." 'INFO'
    $read = Get-CalendarViewChunked -Mailbox $a -Select $attSelect
    if (-not $read.Ok) {
        $readFailed += $a
        Write-Log ("  read for {0} was INCOMPLETE ({1}) -- its meetings are NOT counted; verdict will be caveated." -f $a, $read.Reason) 'WARN'
        continue
    }
    $readOk += $a
    foreach ($e in $read.Events) {
        if (Get-Prop $e 'isCancelled') { continue }
        if (-not (Test-IsOrganizer $e)) { continue }
        $uid = [string](Get-Prop $e 'iCalUId')
        if (-not $uid) { continue }
        $startUtc = Get-EventStartUtc $e
        $key = Get-OccurrenceKey -ICalUId $uid -StartUtc $startUtc
        if (-not $found.ContainsKey($key)) {
            $found[$key] = [pscustomobject]@{
                iCalUId         = $uid
                Subject         = [string](Get-Prop $e 'subject')
                StartUtc        = $startUtc
                SeenOnAttendees = New-Object System.Collections.Generic.List[string]
            }
        }
        if ($a -notin $found[$key].SeenOnAttendees) { $found[$key].SeenOnAttendees.Add($a) }
    }
}
Write-Log ("  found {0} distinct {1}-organized occurrences across {2} readable attendee calendar(s)." -f $found.Keys.Count, $tag, @($readOk).Count) 'OK'
if (@($readFailed).Count -gt 0) {
    Write-Log ("  {0} attendee calendar(s) could NOT be fully read: {1}" -f @($readFailed).Count, ($readFailed -join ', ')) 'WARN'
}

# ---- 4. Compare each attendee-seen occurrence against the organizer ---------
$results = New-Object System.Collections.Generic.List[object]
foreach ($key in $found.Keys) {
    $rec            = $found[$key]
    $presentOcc     = $organizerOccurrence.ContainsKey($key)
    $presentSeries  = $organizerSeries.ContainsKey($rec.iCalUId)
    $startStr = ''
    if ($rec.StartUtc -is [datetime]) { $startStr = $rec.StartUtc.ToString('u') }
    $verdict = if ($presentOcc) { 'PRESENT_ON_ORGANIZER' } else { 'MISSING_FROM_ORGANIZER' }
    $results.Add([pscustomobject]@{
        Subject              = $rec.Subject
        StartUtc             = $startStr
        Verdict              = $verdict
        PresentOnOrganizer   = $presentOcc
        SeriesPresentOnOrg   = $presentSeries     # series exists but THIS occurrence missing = strong signal
        AttendeesWhoHaveIt   = (@($rec.SeenOnAttendees) -join '; ')
        AttendeeCount        = @($rec.SeenOnAttendees).Count
        iCalUId              = $rec.iCalUId
    })
}

$results = $results | Sort-Object Verdict, StartUtc
$results | Export-Csv -Path $ResultCsv -NoTypeInformation -Encoding UTF8

# ---- 5. Summary ------------------------------------------------------------
Write-Log '================ FINDINGS ================' 'STEP'
$missing = @($results | Where-Object { $_.Verdict -eq 'MISSING_FROM_ORGANIZER' })
$present = @($results | Where-Object { $_.Verdict -eq 'PRESENT_ON_ORGANIZER' })
Write-Log ("PRESENT_ON_ORGANIZER   : {0}" -f $present.Count) 'OK'
$missLvl = if ($missing.Count) { 'ERROR' } else { 'OK' }
Write-Log ("MISSING_FROM_ORGANIZER : {0}" -f $missing.Count) $missLvl
foreach ($m in $missing) {
    $seriesNote = if ($m.SeriesPresentOnOrg) { ' (series present on organizer, but THIS occurrence missing)' } else { '' }
    Write-Log ("  MISSING: '{0}' [{1}] - on {2} attendee(s): {3}{4}" -f $m.Subject, $m.StartUtc, $m.AttendeeCount, $m.AttendeesWhoHaveIt, $seriesNote) 'ERROR'
}
Write-Log "Findings CSV: $ResultCsv" 'INFO'

if ($missing.Count) {
    Write-Log "MISSING_FROM_ORGANIZER rows are the proof: occurrence is on attendees' calendars but absent from the organizer's. Root-cause client still needs the diagnostic log once the backend is restored." 'STEP'
} else {
    Write-Log "No missing organizer copies found among occurrences that were successfully compared." 'OK'
}

# Trustworthiness caveat: an "all clear" is only valid if every attendee was read.
if (@($readFailed).Count -gt 0) {
    Write-Log ("CAVEAT: {0} attendee calendar(s) could not be fully read ({1}). Their meetings were NOT compared, so a 'no missing' result is INCOMPLETE. Re-run later (sharing may still be propagating) or raise -ShareWaitSeconds." -f @($readFailed).Count, ($readFailed -join ', ')) 'WARN'
} elseif ($missing.Count -eq 0) {
    Write-Log "All scanned attendee calendars were read successfully, so this clean result is trustworthy for the scanned set. Widen -MaxAttendees or scan a specific attendee on a known-bad meeting for broader coverage." 'INFO'
}

Write-Log "Sessions left connected. Re-run later to pick up attendees whose sharing hadn't propagated yet." 'OK'
Write-Log "Done." 'STEP'
