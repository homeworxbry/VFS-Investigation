#Requires -Version 7.0
<#
================================================================================
 Confirm-OrganizerCopyMissing.ps1   (Graph-only companion)
================================================================================
 PURPOSE
   Confirm the reported failure -- meetings that should appear on matt.lowe's
   calendar are missing -- WITHOUT using Get-CalendarDiagnosticObjects (which
   is currently failing server-side in this tenant).  Uses only Microsoft Graph
   calendarView, which is working.

 TWO SCENARIOS COVERED
   A. Matt as ORGANIZER
      Meetings Matt organised that appear on colleagues' calendars but are
      absent from his own.
   B. Matt as INVITEE
      Meetings organised by someone else where Matt appears as an attendee on
      at least one colleague's calendar, but the occurrence is absent from
      Matt's calendar.

 METHOD
   1. Read Matt's calendar (last N days) to build:
        - The set of meeting occurrences currently ON his calendar.
        - A bidirectional colleague list: people he frequently invites PLUS
          the organizers of meetings on his calendar where someone else ran it.
          Using both directions avoids biasing the scan set toward only one
          role.
   2. Grant the signed-in admin least-privilege Reviewer on each colleague's
      Calendar (idempotent; upgrades a too-low existing right), wait for
      sharing to propagate, then read each colleague's calendar.
   3. On each colleague's calendar collect:
        Scenario A: occurrences where organizer = Matt.
        Scenario B: occurrences organised by anyone where Matt appears in the
                    attendees list.
   4. Compare each occurrence against Matt's own calendar.
        On colleague + ABSENT from Matt  =  MISSING_FROM_MATT
        On colleague + present on Matt   =  PRESENT_ON_MATT

 WHY OCCURRENCE-LEVEL
   calendarView expands recurring series into per-occurrence instances that all
   share ONE iCalUId.  Correlating by iCalUId alone would report a series as
   "present" even when individual occurrences are missing.  We correlate by
   (iCalUId + occurrence-start-to-the-minute).

 OUTPUT
   mattphase2investigation.csv  --  written directly to $OutputRoot (the folder
                                    the script is run from).  One row per
                                    occurrence found on at least one colleague's
                                    calendar.  Key columns:
                                      MattRole         Organizer | Invitee
                                      Verdict          MISSING_FROM_MATT | PRESENT_ON_MATT
                                      Subject
                                      StartUtc
                                      MeetingOrganizer who set the meeting up
                                      SeriesPresentOnMatt  series exists but THIS occurrence missing
                                      SeenOnCalendars  whose calendar(s) it was found on
   A timestamped run subfolder (OrgCopyCheck_*) holds the run log alongside
   this file.

 COVERAGE CAVEAT
   We can only see meetings that are visible through the scanned colleague set.
   If a meeting had no scanned colleague as a co-attendee or organizer it will
   not appear in the output.  Widen -MaxColleagues or add specific addresses
   with -ExtraColleagues to improve coverage.

 TRUSTWORTHINESS GUARD
   A forensic "all clear" is only meaningful if every colleague calendar was
   actually READ.  Access failures are tracked and the final verdict is
   explicitly caveated when any read was incomplete -- an access failure must
   never masquerade as "no problem found".

 AUTH / PERMISSIONS
   Interactive delegated admin auth (session-aware; prompts only if needed).
   Requires Graph scopes Calendars.Read.Shared + User.Read.All, and Exchange
   Online admin rights to grant Reviewer on colleague calendars.
   Sessions are NOT disconnected.

 RUN
   .\Confirm-OrganizerCopyMissing.ps1
   .\Confirm-OrganizerCopyMissing.ps1 -TargetOrganizer matt.lowe@venturafs.com `
       -LookbackDays 90 -MaxColleagues 20
   .\Confirm-OrganizerCopyMissing.ps1 -ExtraColleagues @('alice@venturafs.com','bob@venturafs.com')
================================================================================
#>

[CmdletBinding()]
param(
    [string]$TargetOrganizer   = 'matt.lowe@venturafs.com',
    [int]$LookbackDays         = 90,
    [string]$OutputRoot        = (Get-Location).Path,

    # Only auto-discover colleagues in this SMTP domain (defaults to the
    # target's domain).  Set to '' to include any domain.
    [string]$ColleagueDomain   = '',

    # Cap on auto-discovered colleagues to scan (most-frequent first,
    # bidirectional: people he invites + people who invited him).
    [int]$MaxColleagues        = 20,

    # Additional specific addresses to always include in the scan, regardless
    # of the auto-discovered set.
    [string[]]$ExtraColleagues = @(),

    # Seconds to wait after granting Reviewer before reading colleague
    # calendars, so freshly-granted sharing has time to propagate to Graph.
    [int]$ShareWaitSeconds     = 120
)

$ErrorActionPreference = 'Stop'
if (-not $ColleagueDomain) { $ColleagueDomain = ($TargetOrganizer -split '@')[1] }

# ---- Output + logging -------------------------------------------------------
$stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
$tag       = ($TargetOrganizer -split '@')[0]
$OutDir    = Join-Path $OutputRoot ("OrgCopyCheck_{0}_{1}" -f $tag, $stamp)
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$LogFile   = Join-Path $OutDir 'run.log'
$Phase2Csv = Join-Path $OutputRoot 'mattphase2investigation.csv'

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

# Diagnostic trap: unhandled terminating error prints type, exact script
# position, and call stack so a bare .NET message can never leave us guessing.
trap {
    Write-Host ("FATAL {0}: {1}" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ForegroundColor Red
    Write-Host ("  at: {0}" -f $_.InvocationInfo.PositionMessage) -ForegroundColor Red
    Write-Host ("  stack:`n{0}" -f $_.ScriptStackTrace) -ForegroundColor DarkRed
    if ($LogFile) {
        Add-Content -Path $LogFile -Value ("FATAL {0}: {1}`n{2}`n{3}" -f $_.Exception.GetType().FullName, $_.Exception.Message, $_.InvocationInfo.PositionMessage, $_.ScriptStackTrace)
    }
    break
}

# Count by iteration -- avoids PS 7.5/.NET 9 bug where @(List[object]).Count
# throws "Argument types do not match".  Dictionaries use native .Count (safe).
function Get-Count {
    param($Collection)
    if ($null -eq $Collection) { return 0 }
    if ($Collection -is [System.Collections.IDictionary]) { return $Collection.Count }
    if ($Collection -is [string]) { return 1 }
    if ($Collection -is [System.Collections.IEnumerable]) {
        $n = 0; foreach ($x in $Collection) { $n++ }; return $n
    }
    return 1
}

Write-Log "Phase-2 investigation started.  Target=$TargetOrganizer  Lookback=$LookbackDays  Domain=$ColleagueDomain" 'STEP'

# ---- Modules ----------------------------------------------------------------
foreach ($m in 'ExchangeOnlineManagement','Microsoft.Graph.Authentication') {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Log "Installing module '$m' for current user..." 'WARN'
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module $m -ErrorAction Stop
    Write-Log "Module '$m' ready." 'OK'
}

# ---- Connections (reuse if present) -----------------------------------------
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

# ---- Core helpers -----------------------------------------------------------
$startDt = (Get-Date).AddDays(-$LookbackDays)
$endDt   = (Get-Date).AddDays(1)

# Safe nested-property reader -- handles both PSCustomObject and IDictionary,
# which Invoke-MgGraphRequest -OutputType PSObject can return for nested nodes.
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

# Parse a Graph dateTime returned under Prefer: outlook.timezone="UTC".
# No UTC offset is present in the string, so a bare [datetime] cast tags the
# value Unspecified and ToUniversalTime() shifts by the host's local offset.
# Pin Kind=Utc explicitly to prevent that.
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
    return (ConvertTo-Utc (Get-Prop (Get-Prop $Event 'start') 'dateTime'))
}

# Correlate per-occurrence: iCalUId alone collapses a whole recurring series,
# masking individually missing occurrences.
function Get-OccurrenceKey {
    param([string]$ICalUId, $StartUtc)
    $d = if ($StartUtc -is [datetime]) { $StartUtc.ToString('yyyy-MM-ddTHH:mm') } else { 'nostart' }
    return "$ICalUId|$d"
}

# Some Microsoft.Graph.Authentication builds reject a custom -Headers hashtable
# client-side.  The Prefer header only asks for UTC rendering; ConvertTo-Utc
# handles offset-bearing values anyway.  Retry without the header once if it
# fails locally, then stop trying it.
$script:GraphPreferHeaderOk = $true
function Invoke-GraphGet {
    param([Parameter(Mandatory)][string]$Uri, [hashtable]$Headers = @{})
    if ($script:GraphPreferHeaderOk -and $Headers -and $Headers.Count -gt 0) {
        try { return Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers -OutputType PSObject -ErrorAction Stop }
        catch {
            $em = $_.Exception.Message
            if ($em -match 'Argument types do not match|does not match|IDictionary|header') {
                Write-Log "  -Headers rejected by this Graph module build; retrying without Prefer header." 'WARN'
                $script:GraphPreferHeaderOk = $false
            } else { throw }
        }
    }
    return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

# Paged Graph GET.  Returns {Ok; Reason; Items} so callers distinguish a
# genuine empty result from an access/permission failure.
function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri, [hashtable]$Headers = @{}, [int]$MaxAttempts = 4)
    $items  = New-Object System.Collections.Generic.List[object]
    $ok     = $true; $reason = ''
    $next   = $Uri
    while ($next) {
        $attempt = 0; $resp = $null; $pageOk = $false
        while ($true) {
            $attempt++
            try { $resp = Invoke-GraphGet -Uri $next -Headers $Headers; $pageOk = $true; break }
            catch {
                $msg = $_.Exception.Message
                if ($msg -match 'Forbidden|Unauthorized|\b401\b|\b403\b|\b404\b|\b400\b|NotFound|denied') {
                    Write-Log ("    permanent error ({0}); not retrying." -f $msg) 'WARN'
                    $ok = $false; $reason = $msg; break
                }
                if ($attempt -ge $MaxAttempts) {
                    Write-Log ("    failed after {0} attempts ({1})." -f $attempt, $msg) 'WARN'
                    $ok = $false; $reason = $msg; break
                }
                $wait = @(5,15,30,60)[[math]::Min($attempt-1,3)]
                Write-Log ("    transient; retry {0}/{1} in {2}s..." -f $attempt, $MaxAttempts, $wait) 'WARN'
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

# Read a mailbox's calendarView in 15-day slices.
# Returns {Ok; Reason; Events}.  Ok=$false if ANY slice failed, so a partial
# read is never silently treated as complete.
function Get-CalendarViewChunked {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$Select)
    $hdr    = @{ Prefer = 'outlook.timezone="UTC"' }
    $seen   = @{}; $events = New-Object System.Collections.Generic.List[object]
    $ok     = $true; $reason = ''
    $cursor = $startDt
    while ($cursor -lt $endDt) {
        $sliceEnd = $cursor.AddDays(15)
        if ($sliceEnd -gt $endDt) { $sliceEnd = $endDt }
        $sEnc = [uri]::EscapeDataString($cursor.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $eEnc = [uri]::EscapeDataString($sliceEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $uri  = "https://graph.microsoft.com/v1.0/users/$Mailbox/calendarView" +
                "?startDateTime=$sEnc&endDateTime=$eEnc&`$top=100&`$select=$Select"
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

# Access rights that actually permit reading calendar ITEMS (subject/uid/etc.).
# Free/busy-only rights (AvailabilityOnly, LimitedDetails) return nothing useful.
$ItemReadRoles = @('Reviewer','Author','NonEditingAuthor','Editor',
                   'PublishingAuthor','PublishingEditor','Owner')

function Grant-CalendarShareFor {
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$AdminUpn)
    $folder = "$Mailbox`:\Calendar"
    $cur = $null
    try { $cur = Get-MailboxFolderPermission -Identity $folder -User $AdminUpn -ErrorAction Stop } catch { $cur = $null }
    if ($cur) {
        $rights = @($cur.AccessRights | ForEach-Object { [string]$_ })
        if ($rights | Where-Object { $_ -in $ItemReadRoles }) { return $true }
        try {
            Set-MailboxFolderPermission -Identity $folder -User $AdminUpn -AccessRights Reviewer -ErrorAction Stop | Out-Null
            Write-Log ("  upgraded '{0}' from [{1}] to Reviewer." -f $Mailbox, ($rights -join ',')) 'INFO'
            return $true
        } catch {
            Write-Log ("  could not upgrade permission for {0}: {1}" -f $Mailbox, $_.Exception.Message) 'WARN'
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

# ---- 0. Resolve Matt's full address set (proxy/alias robust matching) -------
# A meeting "organized by Matt" or "attended by Matt" may carry a proxy address
# rather than his primary SMTP; match against his whole alias set.
$mattAddrs = @{}
$mattAddrs[$TargetOrganizer.Trim().ToLowerInvariant()] = $true
try {
    $mbx = Get-Mailbox -Identity $TargetOrganizer -ErrorAction Stop
    foreach ($pa in @($mbx.EmailAddresses)) {
        $s = [string]$pa
        if ($s -match '^smtp:(.+)$') { $mattAddrs[$Matches[1].Trim().ToLowerInvariant()] = $true }
    }
    if ($mbx.PrimarySmtpAddress) { $mattAddrs[([string]$mbx.PrimarySmtpAddress).Trim().ToLowerInvariant()] = $true }
    Write-Log ("Matt's address set resolved ({0} address(es))." -f $mattAddrs.Keys.Count) 'OK'
} catch {
    Write-Log ("Could not resolve proxy addresses ({0}); matching on '{1}' only." -f $_.Exception.Message, $TargetOrganizer) 'WARN'
}

function Test-IsMatt {
    param([string]$Address)
    return $mattAddrs.ContainsKey($Address.Trim().ToLowerInvariant())
}

function Test-IsOrganizerMatt {
    param($Event)
    $addr = [string](Get-Prop (Get-Prop (Get-Prop $Event 'organizer') 'emailAddress') 'address')
    return (Test-IsMatt $addr)
}

# Returns true if Matt appears anywhere in the event's attendees list.
function Test-MattIsAttendee {
    param($Event)
    foreach ($a in @(Get-Prop $Event 'attendees')) {
        if ($null -eq $a) { continue }
        $addr = [string](Get-Prop (Get-Prop $a 'emailAddress') 'address')
        if ($addr -and (Test-IsMatt $addr)) { return $true }
    }
    return $false
}

# ---- 1. Read Matt's calendar: build presence sets + discover colleagues -----
Write-Log "Reading Matt's calendar..." 'STEP'
$mattSelect = 'id,iCalUId,subject,start,end,organizer,isOrganizer,isCancelled,attendees'
$mattRead   = Get-CalendarViewChunked -Mailbox $TargetOrganizer -Select $mattSelect
$mattEvents = $mattRead.Events
if (-not $mattRead.Ok) {
    Write-Log ("Could not fully read Matt's own calendar ({0}). No baseline to compare against -- aborting." -f $mattRead.Reason) 'ERROR'
    return
}

# Presence sets (cancelled tombstones excluded from both sides of the comparison).
$mattSeries     = @{}
$mattOccurrence = @{}
foreach ($e in $mattEvents) {
    if (Get-Prop $e 'isCancelled') { continue }
    $uid = [string](Get-Prop $e 'iCalUId')
    if (-not $uid) { continue }
    $mattSeries[$uid] = $true
    $mattOccurrence[(Get-OccurrenceKey -ICalUId $uid -StartUtc (Get-EventStartUtc $e))] = $true
}
$mattEventCount  = Get-Count $mattEvents
$mattSeriesCount = Get-Count $mattSeries
Write-Log "  Matt's calendar: $mattEventCount event(s), $mattSeriesCount distinct series." 'INFO'

# Bidirectional colleague discovery:
#   Scenario-A direction: attendees on meetings Matt organised (people he invites).
#   Scenario-B direction: organizers of meetings where Matt is an attendee (people
#     who invite him).
# Counting frequency in both directions and taking the top-N ensures the scan set
# covers both halves of the investigation.
$freq = @{}
foreach ($e in $mattEvents) {
    if (Get-Prop $e 'isCancelled') { continue }

    if (Get-Prop $e 'isOrganizer') {
        # Matt ran this meeting -- collect his invitees.
        foreach ($a in @(Get-Prop $e 'attendees')) {
            if ($null -eq $a) { continue }
            $addr = ([string](Get-Prop (Get-Prop $a 'emailAddress') 'address')).Trim()
            $type = [string](Get-Prop $a 'type')
            if (-not $addr -or $type -eq 'resource') { continue }
            if (Test-IsMatt $addr) { continue }
            if ($ColleagueDomain -and ($addr -notlike "*@$ColleagueDomain")) { continue }
            if ($freq.ContainsKey($addr)) { $freq[$addr]++ } else { $freq[$addr] = 1 }
        }
    } else {
        # Someone else ran this meeting -- collect the organizer.
        $orgAddr = ([string](Get-Prop (Get-Prop (Get-Prop $e 'organizer') 'emailAddress') 'address')).Trim()
        if (-not $orgAddr -or (Test-IsMatt $orgAddr)) { continue }
        if ($ColleagueDomain -and ($orgAddr -notlike "*@$ColleagueDomain")) { continue }
        if ($freq.ContainsKey($orgAddr)) { $freq[$orgAddr]++ } else { $freq[$orgAddr] = 1 }
    }
}

$colleagues = [System.Collections.Generic.List[string]]::new()
foreach ($e in ($freq.GetEnumerator() | Sort-Object { [int]$_.Value } -Descending | Select-Object -First $MaxColleagues)) {
    $colleagues.Add($e.Name)
}
# Always include any explicitly requested addresses.
foreach ($ex in $ExtraColleagues) {
    $exN = $ex.Trim()
    if ($exN -and $exN -notin $colleagues) { $colleagues.Add($exN) }
}

$freqCount = Get-Count $freq
$scanCount = Get-Count $colleagues
Write-Log "  discovered $freqCount distinct colleague(s) (bidirectional); scanning $scanCount." 'OK'
if ($scanCount -eq 0) { Write-Log "No colleagues discovered; nothing to compare. Exiting." 'WARN'; return }

# ---- 2. Grant Reviewer on each colleague, then wait for propagation ----------
Write-Log "Granting least-privilege Reviewer on colleague calendars..." 'STEP'
$granted = [System.Collections.Generic.List[string]]::new()
foreach ($c in $colleagues) {
    if (Grant-CalendarShareFor -Mailbox $c -AdminUpn $AdminUpn) {
        $granted.Add($c); Write-Log "  shared: $c" 'INFO'
    }
}
if ((Get-Count $granted) -eq 0) { Write-Log "No colleague calendars could be shared. Exiting." 'WARN'; return }
Write-Log ("Waiting {0}s for sharing to propagate to Graph..." -f $ShareWaitSeconds) 'STEP'
Start-Sleep -Seconds $ShareWaitSeconds

# ---- 3. Scan colleague calendars -- both scenarios in one pass ---------------
Write-Log "Scanning colleague calendars (A: Matt as organizer; B: Matt as invitee)..." 'STEP'
$collSelect  = 'id,iCalUId,subject,start,end,organizer,isOrganizer,isCancelled,attendees'
$foundOrg    = @{}   # occurrenceKey -> record  (Matt organised)
$foundInv    = @{}   # occurrenceKey -> record  (Matt was invited by someone else)
$readOk      = [System.Collections.Generic.List[string]]::new()
$readFailed  = [System.Collections.Generic.List[string]]::new()

foreach ($c in $granted) {
    Write-Log "  reading $c ..." 'INFO'
    $read = Get-CalendarViewChunked -Mailbox $c -Select $collSelect
    if (-not $read.Ok) {
        $readFailed.Add($c)
        Write-Log ("  INCOMPLETE read for {0} ({1}) -- excluded from comparison; verdict will be caveated." -f $c, $read.Reason) 'WARN'
        continue
    }
    $readOk.Add($c)

    foreach ($e in $read.Events) {
        if (Get-Prop $e 'isCancelled') { continue }
        $uid = [string](Get-Prop $e 'iCalUId')
        if (-not $uid) { continue }
        $startUtc = Get-EventStartUtc $e
        $key      = Get-OccurrenceKey -ICalUId $uid -StartUtc $startUtc
        $subject  = [string](Get-Prop $e 'subject')

        if (Test-IsOrganizerMatt $e) {
            # Scenario A: Matt organised this occurrence.
            if (-not $foundOrg.ContainsKey($key)) {
                $foundOrg[$key] = [pscustomobject]@{
                    iCalUId          = $uid
                    Subject          = $subject
                    StartUtc         = $startUtc
                    MeetingOrganizer = $TargetOrganizer
                    SeenOnCalendars  = [System.Collections.Generic.List[string]]::new()
                }
            }
            if ($c -notin $foundOrg[$key].SeenOnCalendars) { $foundOrg[$key].SeenOnCalendars.Add($c) }

        } elseif (Test-MattIsAttendee $e) {
            # Scenario B: someone else organised; Matt is listed as an attendee.
            $meetOrg = [string](Get-Prop (Get-Prop (Get-Prop $e 'organizer') 'emailAddress') 'address')
            if (-not $foundInv.ContainsKey($key)) {
                $foundInv[$key] = [pscustomobject]@{
                    iCalUId          = $uid
                    Subject          = $subject
                    StartUtc         = $startUtc
                    MeetingOrganizer = $meetOrg
                    SeenOnCalendars  = [System.Collections.Generic.List[string]]::new()
                }
            }
            if ($c -notin $foundInv[$key].SeenOnCalendars) { $foundInv[$key].SeenOnCalendars.Add($c) }
        }
    }
}

$orgFoundCt      = Get-Count $foundOrg
$invFoundCt      = Get-Count $foundInv
$readOkCt        = Get-Count $readOk
$readFailedCt    = Get-Count $readFailed
Write-Log "  Scenario A (Matt as organizer): $orgFoundCt occurrence(s) found on colleague calendars." 'INFO'
Write-Log "  Scenario B (Matt as invitee):   $invFoundCt occurrence(s) found on colleague calendars." 'INFO'
Write-Log "  Calendars read successfully: $readOkCt" 'OK'
if ($readFailedCt -gt 0) {
    Write-Log "  $readFailedCt calendar(s) could NOT be fully read: $($readFailed -join ', ')" 'WARN'
}

# ---- 4. Compare each found occurrence against Matt's presence sets -----------
$results = New-Object System.Collections.Generic.List[object]

foreach ($key in $foundOrg.Keys) {
    $rec           = $foundOrg[$key]
    $presentOcc    = $mattOccurrence.ContainsKey($key)
    $presentSeries = $mattSeries.ContainsKey($rec.iCalUId)
    $startStr      = if ($rec.StartUtc -is [datetime]) { $rec.StartUtc.ToString('u') } else { '' }
    $results.Add([pscustomobject]@{
        MattRole            = 'Organizer'
        Verdict             = if ($presentOcc) { 'PRESENT_ON_MATT' } else { 'MISSING_FROM_MATT' }
        Subject             = $rec.Subject
        StartUtc            = $startStr
        MeetingOrganizer    = $rec.MeetingOrganizer
        PresentOnMatt       = $presentOcc
        SeriesPresentOnMatt = $presentSeries
        SeenOnCalendars     = ($rec.SeenOnCalendars -join '; ')
        CalendarCount       = (Get-Count $rec.SeenOnCalendars)
        iCalUId             = $rec.iCalUId
    })
}

foreach ($key in $foundInv.Keys) {
    $rec           = $foundInv[$key]
    $presentOcc    = $mattOccurrence.ContainsKey($key)
    $presentSeries = $mattSeries.ContainsKey($rec.iCalUId)
    $startStr      = if ($rec.StartUtc -is [datetime]) { $rec.StartUtc.ToString('u') } else { '' }
    $results.Add([pscustomobject]@{
        MattRole            = 'Invitee'
        Verdict             = if ($presentOcc) { 'PRESENT_ON_MATT' } else { 'MISSING_FROM_MATT' }
        Subject             = $rec.Subject
        StartUtc            = $startStr
        MeetingOrganizer    = $rec.MeetingOrganizer
        PresentOnMatt       = $presentOcc
        SeriesPresentOnMatt = $presentSeries
        SeenOnCalendars     = ($rec.SeenOnCalendars -join '; ')
        CalendarCount       = (Get-Count $rec.SeenOnCalendars)
        iCalUId             = $rec.iCalUId
    })
}

# Sort: missing rows first (MISSING_ < PRESENT_ alphabetically), then by role
# and date, so the most actionable findings are at the top of the CSV.
$results = $results | Sort-Object Verdict, MattRole, StartUtc
$results | Export-Csv -Path $Phase2Csv -NoTypeInformation -Encoding UTF8

# ---- 5. Summary -------------------------------------------------------------
Write-Log '================ FINDINGS ================' 'STEP'

$missingOrg = @($results | Where-Object { $_.MattRole -eq 'Organizer' -and $_.Verdict -eq 'MISSING_FROM_MATT' })
$presentOrg = @($results | Where-Object { $_.MattRole -eq 'Organizer' -and $_.Verdict -eq 'PRESENT_ON_MATT'  })
$missingInv = @($results | Where-Object { $_.MattRole -eq 'Invitee'   -and $_.Verdict -eq 'MISSING_FROM_MATT' })
$presentInv = @($results | Where-Object { $_.MattRole -eq 'Invitee'   -and $_.Verdict -eq 'PRESENT_ON_MATT'  })

$moC = Get-Count $missingOrg; $poC = Get-Count $presentOrg
$miC = Get-Count $missingInv; $piC = Get-Count $presentInv

$orgLevel = if ($moC -gt 0) { 'ERROR' } else { 'OK' }
$invLevel = if ($miC -gt 0) { 'ERROR' } else { 'OK' }
Write-Log ("Organizer scenario  --  PRESENT_ON_MATT: {0}   MISSING_FROM_MATT: {1}" -f $poC, $moC) $orgLevel
Write-Log ("Invitee scenario    --  PRESENT_ON_MATT: {0}   MISSING_FROM_MATT: {1}" -f $piC, $miC) $invLevel

if ($moC -gt 0) {
    Write-Log "--- Matt ORGANISED these meetings; they are missing from his own calendar ---" 'ERROR'
    foreach ($m in $missingOrg) {
        $sn = if ($m.SeriesPresentOnMatt) { ' [series present -- THIS occurrence missing]' } else { '' }
        Write-Log "  MISSING (org): '$($m.Subject)' [$($m.StartUtc)]  seen on: $($m.SeenOnCalendars)$sn" 'ERROR'
    }
}
if ($miC -gt 0) {
    Write-Log "--- Matt was INVITED to these meetings; they are missing from his own calendar ---" 'ERROR'
    foreach ($m in $missingInv) {
        $sn = if ($m.SeriesPresentOnMatt) { ' [series present -- THIS occurrence missing]' } else { '' }
        Write-Log "  MISSING (inv): '$($m.Subject)' [$($m.StartUtc)]  org: $($m.MeetingOrganizer)  seen on: $($m.SeenOnCalendars)$sn" 'ERROR'
    }
}

Write-Log "Output: $Phase2Csv" 'INFO'
Write-Log "Run log: $LogFile" 'INFO'

if ($readFailedCt -gt 0) {
    $caveatMsg = "CAVEAT: {0} colleague calendar(s) could not be fully read ({1}). Their meetings were NOT compared -- this result is INCOMPLETE. Re-run (sharing may still be propagating) or raise -ShareWaitSeconds." -f $readFailedCt, ($readFailed -join ', ')
    Write-Log $caveatMsg 'WARN'
} elseif (($moC + $miC) -eq 0) {
    Write-Log ("All $readOkCt colleague calendar(s) were read successfully. " +
               "Clean result is trustworthy for the scanned set. " +
               "Widen -MaxColleagues or use -ExtraColleagues to broaden coverage.") 'INFO'
}

Write-Log "Sessions left connected. Re-run later to pick up calendars whose sharing hadn't propagated yet." 'OK'
Write-Log "Done." 'STEP'
