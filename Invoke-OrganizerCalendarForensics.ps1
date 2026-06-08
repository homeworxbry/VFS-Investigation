#Requires -Version 7.0
<#
================================================================================
 Invoke-OrganizerCalendarForensics.ps1
================================================================================
 PURPOSE
   Forensic verification that every meeting ORGANIZED by a target mailbox in the
   last N days actually produced the organizer's own calendar appointment.

   Scenario being proven: the organizer (default matt.lowe@venturafs.com) created
   meetings, invitees received them, but the organizer's OWN calendar copy never
   appeared. This script confirms, per meeting, whether the organizer calendar
   item was (a) created and still present, (b) never created, or (c) created and
   later removed -- and records WHICH client authored each action.

 DATA SOURCES (unioned for determinism)
   1. Exchange Online  : Get-CalendarDiagnosticObjects  -> authoritative item lifecycle
   2. Microsoft Graph  : /calendarView (organizer = target) -> current calendar state

 WHAT THIS SCRIPT WRITES TO THE TENANT
   * Read-only everywhere EXCEPT two least-privilege administrative actions:
     (1) if your account cannot run Get-CalendarDiagnosticObjects, it creates a
     dedicated least-privilege Exchange role group and adds your account to it;
     (2) it grants your account Reviewer on the target's Calendar so the Graph
     overlay can read it. Nothing else is modified -- no mail, calendar, or
     message data is ever altered or deleted.

 AUTH MODEL
   * Interactive (delegated) admin auth. Checks for an existing session first and
     only prompts to sign in if you are not connected. No app registration or
     certificate is required. Sessions are NOT disconnected at the end so you can
     re-run with changes.

 REQUIREMENTS
   * PowerShell 7+
   * Modules: ExchangeOnlineManagement (v3+), Microsoft.Graph.Authentication
     (the script will install them for the current user if missing)
   * Your account must be Global Administrator / Exchange Administrator so the
     in-script role-group provisioning can succeed.

 HOW TO RUN (review the CONFIG block first, then):
     .\Invoke-OrganizerCalendarForensics.ps1
   or override defaults, e.g.:
     .\Invoke-OrganizerCalendarForensics.ps1 -TargetMailbox matt.lowe@venturafs.com -LookbackDays 90
================================================================================
#>

[CmdletBinding()]
param(
    # The organizer mailbox under investigation.
    [string]$TargetMailbox = 'matt.lowe@venturafs.com',

    # How many days back to look. Diagnostic retention is finite (~90 days), so
    # 90 is the practical maximum for complete data.
    [int]$LookbackDays = 90,

    # Where reports are written. A timestamped subfolder is created automatically.
    [string]$OutputRoot = (Get-Location).Path,

    # Name of the dedicated role group the script provisions if needed.
    [string]$ForensicRoleGroup = 'Forensics-CalendarDiagnostics',

    # Date-window size (days) for the chunked diagnostic sweep. Lower this (e.g. 3 or 1)
    # if a sweep slice returns a server-side error on a very busy mailbox.
    [int]$SweepChunkDays = 7
)

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------
# Output folder + transcript
# ----------------------------------------------------------------------------
$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$mailboxTag = ($TargetMailbox -split '@')[0]
$OutDir     = Join-Path $OutputRoot ("CalForensics_{0}_{1}" -f $mailboxTag, $stamp)
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$LogFile     = Join-Path $OutDir 'run.log'
$SummaryCsv  = Join-Path $OutDir 'meeting_summary.csv'
$TimelineTxt = Join-Path $OutDir 'meeting_timelines.txt'
$RawCsv      = Join-Path $OutDir 'raw_diagnostic_objects.csv'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK','STEP')][string]$Level = 'INFO'
    )
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

# Count a collection by iterating, not via the PowerShell-extended .Count member,
# which on this tenant's host throws "Argument types do not match" for certain
# objects returned by Invoke-MgGraphRequest. Dictionaries expose a native .Count
# that is unaffected. Robust for null/scalar/array/list inputs.
function Get-Count {
    param($Collection)
    if ($null -eq $Collection) { return 0 }
    if ($Collection -is [System.Collections.IDictionary]) { return $Collection.Count }
    if ($Collection -is [string]) { return 1 }
    if ($Collection -is [System.Collections.IEnumerable]) {
        $n = 0
        foreach ($x in $Collection) { $n++ }
        return $n
    }
    return 1
}

Write-Log "Forensic run started. Target=$TargetMailbox  Lookback=$LookbackDays days  Output=$OutDir" 'STEP'

# ============================================================================
# 1. MODULES
# ============================================================================
function Import-RequiredModule {
    param([Parameter(Mandatory)][string]$Name)
    if (Get-Module -ListAvailable -Name $Name) {
        Import-Module $Name -ErrorAction Stop
        Write-Log "Module '$Name' imported." 'OK'
        return
    }
    Write-Log "Module '$Name' not found. Installing for current user..." 'WARN'
    if (-not (Get-PackageSource -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Default -ErrorAction SilentlyContinue
    }
    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Import-Module $Name -ErrorAction Stop
    Write-Log "Module '$Name' installed and imported." 'OK'
}

Import-RequiredModule -Name 'ExchangeOnlineManagement'
Import-RequiredModule -Name 'Microsoft.Graph.Authentication'

# ============================================================================
# 2. CONNECTIONS  (check existing session first; only prompt if needed)
# ============================================================================
function Connect-EXOIfNeeded {
    $conn = $null
    try { $conn = Get-ConnectionInformation -ErrorAction Stop } catch { $conn = $null }
    $live = $conn | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' }
    if ($live) {
        Write-Log ("Reusing existing Exchange Online session as {0}." -f ($live.UserPrincipalName | Select-Object -First 1)) 'OK'
        return
    }
    Write-Log "No active Exchange Online session. Prompting for sign-in..." 'STEP'
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Log "Connected to Exchange Online." 'OK'
}

function Connect-GraphIfNeeded {
    param([string[]]$RequiredScopes)
    $ctx = $null
    try { $ctx = Get-MgContext -ErrorAction Stop } catch { $ctx = $null }
    $missing = @()
    if ($ctx) { $missing = $RequiredScopes | Where-Object { $_ -notin $ctx.Scopes } }
    if ($ctx -and $missing.Count -eq 0) {
        Write-Log ("Reusing existing Microsoft Graph session as {0}." -f $ctx.Account) 'OK'
        return
    }
    if ($ctx -and $missing.Count -gt 0) {
        Write-Log ("Existing Graph session missing scopes: {0}. Re-consenting..." -f ($missing -join ', ')) 'WARN'
    } else {
        Write-Log "No active Microsoft Graph session. Prompting for sign-in..." 'STEP'
    }
    Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph." 'OK'
}

Connect-EXOIfNeeded
Connect-GraphIfNeeded -RequiredScopes @('Calendars.Read','Calendars.Read.Shared','User.Read.All')

# Gather tenant / connected identity from the live sessions.
$exoInfo   = Get-ConnectionInformation | Select-Object -First 1
$graphCtx  = Get-MgContext
$AdminUpn  = $exoInfo.UserPrincipalName
$TenantId  = $graphCtx.TenantId
Write-Log "Operating as $AdminUpn  (TenantId $TenantId)." 'INFO'

# ============================================================================
# 3. ENSURE PERMISSION TO RUN Get-CalendarDiagnosticObjects
#    Probe first; provision a least-privilege role group only if denied.
# ============================================================================
function Test-CalDiagAccess {
    param([Parameter(Mandatory)][string]$Identity)
    try {
        $probeSubject = [guid]::NewGuid().ToString()   # guaranteed-empty result
        Get-CalendarDiagnosticObjects -Identity $Identity -Subject $probeSubject `
            -ExactMatch $true -MaxResults 1 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        # Empty result for a random subject is success; only treat auth/recognition
        # failures as "no access".
        if ($msg -match 'isn''t recognized|not recognized|access is denied|not authorized|cannot be found.*role|operation.*not permitted') {
            return $false
        }
        # Any other error here means the cmdlet IS available (we reached it) but the
        # target/parameters had an issue -> we have access.
        return $true
    }
}

function Initialize-CalDiagPermission {
    Write-Log "Verifying access to Get-CalendarDiagnosticObjects..." 'STEP'
    if (Test-CalDiagAccess -Identity $TargetMailbox) {
        Write-Log "Account already has access to calendar diagnostic data." 'OK'
        return
    }

    Write-Log "Access denied. Discovering which management role grants the cmdlet..." 'WARN'
    $roleEntries = Get-ManagementRoleEntry "*\Get-CalendarDiagnosticObjects" -ErrorAction Stop
    $roleName = ($roleEntries | Select-Object -ExpandProperty Role -Unique | Select-Object -First 1)
    if (-not $roleName) {
        throw "Could not locate any management role containing Get-CalendarDiagnosticObjects."
    }
    Write-Log "Cmdlet is provided by management role '$roleName'." 'INFO'

    $existing = Get-RoleGroup -Identity $ForensicRoleGroup -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Log "Creating least-privilege role group '$ForensicRoleGroup' with role '$roleName'." 'STEP'
        New-RoleGroup -Name $ForensicRoleGroup -Roles $roleName -Members $AdminUpn -ErrorAction Stop | Out-Null
    } else {
        Write-Log "Role group '$ForensicRoleGroup' exists. Ensuring role + membership." 'STEP'
        # Roles is a collection of role objects; compare by name string.
        $existingRoleNames = @((Get-RoleGroup $ForensicRoleGroup).Roles | ForEach-Object { [string]$_ })
        if ($roleName -notin $existingRoleNames) {
            New-ManagementRoleAssignment -SecurityGroup $ForensicRoleGroup -Role $roleName -ErrorAction Stop | Out-Null
        }
        try {
            Add-RoleGroupMember -Identity $ForensicRoleGroup -Member $AdminUpn -ErrorAction Stop
        } catch {
            if ($_.Exception.Message -notmatch 'already a member') { throw }
        }
    }

    # RBAC changes are not instant and the current session caches roles. Retry with
    # backoff; refresh the EXO session once if it still hasn't taken effect.
    Write-Log "Waiting for RBAC propagation (can take a few minutes)..." 'WARN'
    $refreshed = $false
    for ($i = 1; $i -le 10; $i++) {
        Start-Sleep -Seconds 30
        if (Test-CalDiagAccess -Identity $TargetMailbox) {
            Write-Log "Access confirmed after provisioning (attempt $i)." 'OK'
            return
        }
        Write-Log "Still propagating (attempt $i/10)..." 'INFO'
        if ($i -eq 5 -and -not $refreshed) {
            Write-Log "Refreshing Exchange Online session to pick up the new role..." 'STEP'
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop   # refresh, not disconnect
            $refreshed = $true
        }
    }
    throw "RBAC role was provisioned but access did not take effect within the wait window. Re-run the script in a few minutes."
}

Initialize-CalDiagPermission

# ============================================================================
# 4. SEED SOURCES (Graph)
# ============================================================================
$startDt   = (Get-Date).AddDays(-$LookbackDays)
$endDt     = (Get-Date).AddDays(1)            # +1 day to include anything created today

# Safe nested-property reader. Invoke-MgGraphRequest -OutputType PSObject can yield
# PSCustomObject OR (for some nodes) IDictionary; read either shape without tripping
# the ".NET Argument types do not match" error on a wrong-typed member access.
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

# Parse a value to UTC. Graph dateTime requested with Prefer: outlook.timezone="UTC"
# carries NO offset, so a plain [datetime] cast would tag it Unspecified and
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

function Invoke-GraphPaged {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers = @{},
        [int]$MaxAttempts = 4
    )
    $items = New-Object System.Collections.Generic.List[object]
    $next  = $Uri
    while ($next) {
        $attempt = 0
        $resp    = $null
        $pageOk  = $false
        while ($true) {
            $attempt++
            try {
                $resp = Invoke-GraphGet -Uri $next -Headers $Headers
                $pageOk = $true
                break
            } catch {
                $m = $_.Exception.Message
                # Permanent errors (auth / not found / bad request) will never succeed on retry.
                if ($m -match 'Forbidden|Unauthorized|\b401\b|\b403\b|\b404\b|\b400\b|NotFound|denied') {
                    Write-Log ("  Graph request returned a permanent error ({0}). Not retrying." -f $m) 'WARN'
                    break
                }
                if ($attempt -ge $MaxAttempts) {
                    Write-Log ("  Graph request failed after {0} attempts ({1}). Returning partial results." -f $attempt, $m) 'WARN'
                    break
                }
                $wait = @(5,15,30,60)[[math]::Min($attempt - 1, 3)]
                Write-Log ("  Graph request issue (attempt {0}/{1}): {2}. Retrying in {3}s..." -f $attempt, $MaxAttempts, $m, $wait) 'WARN'
                Start-Sleep -Seconds $wait
            }
        }
        if (-not $pageOk) { break }
        $val = Get-Prop $resp 'value'
        if ($val) { foreach ($v in $val) { $items.Add($v) } }
        $next = [string](Get-Prop $resp '@odata.nextLink')
    }
    return ,$items
}

# 4a. calendarView -- meetings currently in the organizer's calendar that he organized.
#     Queried in 15-day slices so recurrence expansion never makes one request heavy
#     enough to hit the Graph 300s timeout. Results deduped by event id.
function Get-OrganizerCalendarSeed {
    Write-Log "Seed: reading organizer calendar via Graph /calendarView (chunked)..." 'STEP'
    $hdr      = @{ Prefer = 'outlook.timezone="UTC"' }
    $seen     = @{}
    $events   = New-Object System.Collections.Generic.List[object]
    $sliceLen = 15
    $cursor   = $startDt
    while ($cursor -lt $endDt) {
        $sliceEnd = $cursor.AddDays($sliceLen)
        if ($sliceEnd -gt $endDt) { $sliceEnd = $endDt }
        $sEnc = [uri]::EscapeDataString($cursor.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $eEnc = [uri]::EscapeDataString($sliceEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
        $uri  = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/calendarView?startDateTime=$sEnc&endDateTime=$eEnc&" +
                '$top=100&$select=id,iCalUId,subject,start,end,organizer,isOrganizer,isCancelled'
        Write-Log ("  slice {0} -> {1}" -f $cursor.ToString('yyyy-MM-dd'), $sliceEnd.ToString('yyyy-MM-dd')) 'INFO'
        $page = Invoke-GraphPaged -Uri $uri -Headers $hdr
        foreach ($e in $page) {
            $id = [string](Get-Prop $e 'id')
            if ($id -and -not $seen.ContainsKey($id)) { $seen[$id] = $true; $events.Add($e) }
        }
        $cursor = $sliceEnd
    }
    $seed = foreach ($e in ($events | Where-Object { (Get-Prop $_ 'isOrganizer') -eq $true })) {
        $start = Get-Prop $e 'start'
        [pscustomobject]@{
            Source     = 'CalendarView'
            Subject    = [string](Get-Prop $e 'subject')
            StartUtc   = (ConvertTo-Utc (Get-Prop $start 'dateTime'))
            iCalUId    = [string](Get-Prop $e 'iCalUId')
            IsCancelled= (Get-Prop $e 'isCancelled')
        }
    }
    $seedCount = Get-Count $seed
    Write-Log "  -> $seedCount organizer events currently in calendar." 'INFO'
    return @($seed)
}

# Least-privilege self-provisioning: grant the signed-in admin Reviewer access to the
# target's Calendar folder so the delegated Graph calendar overlay can read it.
# Best-effort and idempotent; the EXO diagnostic data is authoritative regardless.
function Grant-CalendarShare {
    $folder = "$TargetMailbox`:\Calendar"
    try {
        $cur = Get-MailboxFolderPermission -Identity $folder -User $AdminUpn -ErrorAction Stop
        if ($cur) { Write-Log "Admin already has '$($cur.AccessRights)' on target Calendar." 'OK'; return }
    } catch {
        # No existing permission entry -> add one.
    }
    try {
        Add-MailboxFolderPermission -Identity $folder -User $AdminUpn -AccessRights Reviewer -ErrorAction Stop | Out-Null
        Write-Log "Granted Reviewer on target Calendar to $AdminUpn (Graph sharing may take a few minutes to honor)." 'OK'
    } catch {
        Write-Log ("Could not provision calendar sharing ({0}). Graph overlay may 403; EXO data is authoritative." -f $_.Exception.Message) 'WARN'
    }
}
Grant-CalendarShare

$calSeed = Get-OrganizerCalendarSeed
# Sent-items signal is taken from the EXO diagnostic log (IPM.Schedule.Meeting.Request),
# which needs no extra mailbox permission. A Graph mail seed (which would require a broad
# FullAccess grant) is intentionally NOT used.

# Subject normalizer (strip RE:/FW:, collapse whitespace, lowercase).
function Get-SubjectKey {
    param([string]$Subject)
    if (-not $Subject) { return '' }
    $s = ($Subject -replace '^(RE:|FW:|FWD:)\s*','').Trim().ToLowerInvariant()
    return ($s -replace '\s+',' ')
}

# Build a "currently present in calendar" lookup (normalized subject + start date).
# $StartUtc is intentionally untyped so a $null start does not cause a binding error.
function Get-NormalizedKey {
    param([string]$Subject, $StartUtc)
    $s = Get-SubjectKey -Subject $Subject
    $d = if ($StartUtc -is [datetime]) { $StartUtc.ToString('yyyy-MM-dd') } else { 'nodate' }
    return "$s|$d"
}
$presentInCalendar = @{}
# Organizer-attribution backstop: every calendar-seed event came from a Graph query
# filtered to isOrganizer=true, so a meeting whose key is in this set IS organizer-owned
# even when the diagnostic rows don't carry ResponseType on the appointment item.
$organizerSeedKeys = @{}
foreach ($c in $calSeed) {
    $nk = Get-NormalizedKey -Subject $c.Subject -StartUtc $c.StartUtc
    $organizerSeedKeys[$nk] = $true
    if ($c.IsCancelled) { continue }
    $presentInCalendar[$nk] = $true
}

# ============================================================================
# 5. DIAGNOSTIC COLLECTION (Exchange Online)
#    A) Per-subject queries seeded from the Graph organizer list = reliable backbone
#       for meetings currently/recently in his calendar.
#    B) A CHUNKED, retried date-range sweep = best-effort catch for organizer copies
#       that are MISSING from his calendar (their sent meeting-request still logs).
#       Bare mailbox-wide sweeps fail server-side, so we slice the window small.
# ============================================================================
$byGoid = @{}
function Add-DiagObjects {
    param([object[]]$Objects)
    foreach ($o in $Objects) {
        $goid = [string]$o.CleanGlobalObjectId
        if (-not $goid) { continue }
        if (-not $byGoid.ContainsKey($goid)) { $byGoid[$goid] = New-Object System.Collections.Generic.List[object] }
        $byGoid[$goid].Add($o)
    }
}

function Invoke-DiagQuery {
    param([hashtable]$Params, [int]$MaxAttempts = 3)
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Get-CalendarDiagnosticObjects @Params -ErrorAction Stop
        } catch {
            $m = $_.Exception.Message
            if ($attempt -ge $MaxAttempts) { throw }
            if ($m -match 'server side error|try again|timed out|timeout|throttl') {
                $wait = @(10,30,60)[[math]::Min($attempt - 1, 2)]
                Write-Log ("    transient diagnostic error (attempt {0}/{1}); retrying in {2}s..." -f $attempt, $MaxAttempts, $wait) 'WARN'
                Start-Sleep -Seconds $wait
            } else {
                throw   # non-transient: let caller handle
            }
        }
    }
}

# ---- A) Per-subject backbone ----
$seedSubjects = @($calSeed | Where-Object { $_.Subject } |
                  Select-Object -ExpandProperty Subject -Unique)
$seedSubjectCount = Get-Count $seedSubjects
Write-Log "Querying diagnostics for $seedSubjectCount distinct organizer subjects..." 'STEP'
foreach ($subj in $seedSubjects) {
    try {
        $hit = Invoke-DiagQuery -Params @{
            Identity = $TargetMailbox; Subject = $subj; ExactMatch = $true
        }
        if ($hit) { Add-DiagObjects -Objects $hit }
    } catch {
        Write-Log ("  Subject query failed for '{0}': {1}" -f $subj, $_.Exception.Message) 'WARN'
    }
}

# ---- B) Chunked date-range sweep to catch MISSING organizer copies (best-effort) ----
Write-Log "Chunked diagnostic sweep (catches meetings absent from his calendar)..." 'STEP'
$sweepOk    = $true
$chunkStart = $startDt
while ($chunkStart -lt $endDt) {
    $chunkEnd = $chunkStart.AddDays($SweepChunkDays)
    if ($chunkEnd -gt $endDt) { $chunkEnd = $endDt }
    Write-Log ("  sweep slice {0} -> {1}" -f $chunkStart.ToString('yyyy-MM-dd'), $chunkEnd.ToString('yyyy-MM-dd')) 'INFO'
    try {
        $chunk = Invoke-DiagQuery -Params @{
            Identity = $TargetMailbox; StartDate = $chunkStart; EndDate = $chunkEnd
        }
        if ($chunk) { Add-DiagObjects -Objects $chunk }
    } catch {
        $sweepOk = $false
        Write-Log ("  sweep slice failed ({0}). Reduce -SweepChunkDays and re-run if missing-copy coverage matters." -f $_.Exception.Message) 'WARN'
    }
    $chunkStart = $chunkEnd
}
if (-not $sweepOk) {
    Write-Log "One or more sweep slices failed: detection of meetings entirely absent from his mailbox may be incomplete. Present/removed meetings are unaffected." 'WARN'
}

$goidCount = $byGoid.Count
Write-Log "Total distinct meetings (by CleanGlobalObjectId): $goidCount" 'OK'

# ============================================================================
# 6. CLASSIFY EACH MEETING
# ============================================================================
$mrtMap = @{
    '1'='MeetingRequest'; '65536'='FullUpdate'; '131072'='InformationalUpdate';
    '262144'='SilentUpdate'; '524288'='Outdated'; '1048576'='ForwardedToDelegate'
}

# CalendarLogTriggerAction values are emitted as readable strings by default.
# Use ANCHORED matches: 'Create' must be exactly Create (not "CreateXxx"), and the
# removal set covers all ways an item leaves the Calendar folder, including a plain
# folder move (MoveToFolder), which an unanchored 'Delete' regex would miss.
function Test-IsCreateAction  { param([string]$A) return ($A -match '^Create$') }
function Test-IsRemovalAction { param([string]$A) return ($A -match '^(Delete|MoveToDeletedItems|SoftDelete|HardDelete|MoveToFolder)$') }

# Safe date helpers. Parse with InvariantCulture + AssumeUniversal so a string round-trip
# of OriginalLastModifiedTime can never reorder events by the host's locale/offset.
function Get-SortDate {
    param($Value)
    if ($Value -is [datetime]) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    if ([datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return [datetime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    return [datetime]::MinValue
}
function Format-Utc {
    param($Value)
    $d = ConvertTo-Utc $Value
    if ($d -is [datetime]) { return $d.ToString('yyyy-MM-dd HH:mm:ss') }
    return [string]$Value
}
# Participant resolution is off (it triggers the server-side error), so ResponsibleUserName
# arrives as a LegacyExchangeDN. Pull the trailing cn= segment for readability.
function Format-Responsible {
    param($Value)
    $s = [string]$Value
    if (-not $s) { return '' }
    if ($s -match 'cn=([^/]+)$') { return $matches[1] }
    return $s
}

$results   = New-Object System.Collections.Generic.List[object]
$timelines = New-Object System.Text.StringBuilder

foreach ($goid in $byGoid.Keys) {
    # Sort on the raw datetime property (not a string reparse) for stable ordering.
    $objs     = $byGoid[$goid] | Sort-Object { Get-SortDate $_.OriginalLastModifiedTime }
    $first    = $objs | Select-Object -First 1
    $subject  = ($objs | Where-Object NormalizedSubject | Select-Object -First 1 -ExpandProperty NormalizedSubject)
    if (-not $subject) { $subject = $first.Subject }

    # Determine a representative start date for matching against the calendar seed.
    $startProp = ($objs | Where-Object { $_.StartTime } | Select-Object -First 1 -ExpandProperty StartTime)
    $startUtc  = ConvertTo-Utc $startProp

    # Appointment lifecycle (organizer's own Calendar item).
    $apptObjs    = $objs | Where-Object { ([string]$_.ItemClass) -like 'IPM.Appointment*' }
    $apptCreate  = $apptObjs | Where-Object { Test-IsCreateAction ([string]$_.CalendarLogTriggerAction) }
    $apptRemoval = $apptObjs | Where-Object { Test-IsRemovalAction ([string]$_.CalendarLogTriggerAction) }
    $reqObjs     = $objs | Where-Object { ([string]$_.ItemClass) -like 'IPM.Schedule.Meeting*' }
    $reqCount    = Get-Count $reqObjs

    # EXO-authoritative current presence: appointment was created AND its most recent
    # appointment action is not a removal. Does not depend on Graph access.
    $exoPresent = $false
    if ($apptCreate) {
        $lastAppt   = $apptObjs | Sort-Object { Get-SortDate $_.OriginalLastModifiedTime } | Select-Object -Last 1
        $lastAction = [string]$lastAppt.CalendarLogTriggerAction
        if (-not (Test-IsRemovalAction $lastAction)) { $exoPresent = $true }
    }

    # Graph corroboration (only meaningful if the calendar overlay succeeded).
    $key          = Get-NormalizedKey -Subject $subject -StartUtc $startUtc
    $graphPresent = $presentInCalendar.ContainsKey($key)

    # Authoritative answer is EXO; Graph only adds confidence to the "no diag" cases.
    $currentlyThere = $exoPresent -or $graphPresent

    # Organizer attribution: ResponseType "1"/"Organizer" on ANY row, OR the meeting
    # key matches a Graph isOrganizer=true seed (ResponseType is often blank on the
    # appointment row, so the seed backstop prevents false negatives).
    $isOrganizer = [bool]($objs | Where-Object { "$($_.ResponseType)" -match '^(1|Organizer)$' }) -or
                   $organizerSeedKeys.ContainsKey($key)

    # ---- Classification (EXO-authoritative; Graph never overrides an EXO deletion) ----
    $classification = 'AMBIGUOUS'
    if ($apptCreate -and $exoPresent) {
        $classification = 'PRESENT_OK'
    }
    elseif ($apptCreate -and -not $exoPresent) {
        # Created then no longer present per the log's last appointment action.
        if ($graphPresent) {
            # Log says removed but Graph still sees it (e.g. moved between folders, or a
            # same-subject/same-day sibling matched) -> don't over-claim removal.
            $classification = 'AMBIGUOUS'
        } elseif ($apptRemoval) {
            $classification = 'CREATED_THEN_REMOVED'
        } else {
            $classification = 'AMBIGUOUS'
        }
    }
    elseif (-not $apptCreate -and ($reqCount -gt 0)) {
        # Meeting-request activity exists for the organizer mailbox but no appointment
        # was ever created in his Calendar -> the reported failure signature.
        $classification = 'MISSING_NEVER_CREATED'
    }
    elseif (-not $apptCreate -and $reqCount -eq 0 -and $graphPresent) {
        $classification = 'IN_CALENDAR_NO_DIAG'  # present per Graph, lifecycle outside window
    }

    # The client that authored the appointment create (root-cause signal).
    $authoringClient = ($apptCreate | Select-Object -First 1 -ExpandProperty ClientInfoString -ErrorAction SilentlyContinue)
    if (-not $authoringClient) {
        $authoringClient = ($reqObjs | Select-Object -First 1 -ExpandProperty ClientInfoString -ErrorAction SilentlyContinue)
    }

    $startUtcStr = ''
    if ($startUtc -is [datetime]) { $startUtcStr = $startUtc.ToString('u') }
    $results.Add([pscustomobject]@{
        Subject              = $subject
        StartUtc             = $startUtcStr
        Classification       = $classification
        OrganizerMeeting     = $isOrganizer
        OrganizerApptCreated = [bool]$apptCreate
        ApptRemovedLater     = [bool]$apptRemoval
        InCalendarNow        = $currentlyThere
        ExoPresent           = $exoPresent
        GraphPresent         = $graphPresent
        AuthoringClient      = $authoringClient
        EventCount           = (Get-Count $objs)
        CleanGlobalObjectId  = $goid
    })

    # ---- Rich per-meeting timeline ----
    [void]$timelines.AppendLine('================================================================')
    [void]$timelines.AppendLine("MEETING : $subject")
    [void]$timelines.AppendLine("START   : $($startUtc)")
    [void]$timelines.AppendLine("GOID    : $goid")
    [void]$timelines.AppendLine("VERDICT : $classification")
    [void]$timelines.AppendLine("AUTHOR  : $authoringClient")
    [void]$timelines.AppendLine('----------------------------------------------------------------')
    [void]$timelines.AppendLine("{0,-22} {1,-14} {2,-30} {3,-22} {4}" -f 'LastModified(UTC)','Action','ItemClass','ReqType','Client / ResponsibleUser')
    foreach ($o in $objs) {
        $t   = Format-Utc $o.OriginalLastModifiedTime
        $mrt = [string]$o.MeetingRequestType
        if ($mrtMap.ContainsKey($mrt)) { $mrt = $mrtMap[$mrt] }
        $who = Format-Responsible $o.ResponsibleUserName
        $cli = $o.ClientInfoString
        [void]$timelines.AppendLine("{0,-22} {1,-14} {2,-30} {3,-22} {4}" -f `
            $t, ([string]$o.CalendarLogTriggerAction), ([string]$o.ItemClass), $mrt, "$cli  |  $who")
    }
    [void]$timelines.AppendLine('')
}

# ============================================================================
# 7. SEED COVERAGE CHECK -- any organized meeting with NO diagnostic data at all?
#    (e.g. organizer copy never written AND nothing logged in his mailbox)
#    Keyed by subject+DATE (not subject alone) so distinct instances of a recurring
#    subject are not collapsed -- otherwise a genuinely uncovered instance would be
#    silently suppressed by a same-subject sibling that DID log diagnostics.
# ============================================================================
$diagCoverageKeys = @{}
foreach ($g in $byGoid.Keys) {
    $grpSubj = ($byGoid[$g] | Where-Object NormalizedSubject | Select-Object -First 1 -ExpandProperty NormalizedSubject)
    if (-not $grpSubj) { $grpSubj = ($byGoid[$g] | Select-Object -First 1).Subject }
    $grpStartProp = ($byGoid[$g] | Where-Object { $_.StartTime } | Select-Object -First 1 -ExpandProperty StartTime)
    $grpStartUtc  = ConvertTo-Utc $grpStartProp
    $diagCoverageKeys[(Get-NormalizedKey -Subject $grpSubj -StartUtc $grpStartUtc)] = $true
}

foreach ($s in $calSeed) {
    if (-not $s.Subject) { continue }
    $k = Get-NormalizedKey -Subject $s.Subject -StartUtc $s.StartUtc
    if ($diagCoverageKeys.ContainsKey($k)) { continue }   # already covered by a meeting row

    $seedStartStr = ''
    if ($s.StartUtc -is [datetime]) { $seedStartStr = $s.StartUtc.ToString('u') }
    $graphHas = $presentInCalendar.ContainsKey($k)
    # If Graph still shows it but no diagnostic rows fell in the window, that's
    # "in calendar, lifecycle outside window" -- not a true no-data gap.
    $seedClass = if ($graphHas) { 'IN_CALENDAR_NO_DIAG' } else { 'NO_DIAGNOSTIC_DATA' }
    $results.Add([pscustomobject]@{
        Subject              = $s.Subject
        StartUtc             = $seedStartStr
        Classification       = $seedClass
        OrganizerMeeting     = $true    # seeded from his own calendar (isOrganizer=true)
        OrganizerApptCreated = $false
        ApptRemovedLater     = $false
        InCalendarNow        = $graphHas
        ExoPresent           = $false
        GraphPresent         = $graphHas
        AuthoringClient      = ''
        EventCount           = 0
        CleanGlobalObjectId  = ''
    })
    $diagCoverageKeys[$k] = $true   # avoid duplicate rows
}

# ============================================================================
# 8. EXPORT
# ============================================================================
$results = $results | Sort-Object Classification, Subject
$results | Export-Csv -Path $SummaryCsv -NoTypeInformation -Encoding UTF8
Set-Content -Path $TimelineTxt -Value $timelines.ToString() -Encoding UTF8
$allRaw = foreach ($g in $byGoid.Keys) { $byGoid[$g] }
$allRaw | Select-Object CleanGlobalObjectId, NormalizedSubject, OriginalLastModifiedTime,
    CalendarLogTriggerAction, ItemClass, ClientInfoString, ResponsibleUserName,
    MeetingRequestType, AppointmentState, StartTime, EndTime |
    Export-Csv -Path $RawCsv -NoTypeInformation -Encoding UTF8

# Optional Excel if ImportExcel is present (no auto-install; CSV is always written).
if (Get-Module -ListAvailable -Name ImportExcel) {
    try {
        Import-Module ImportExcel -ErrorAction Stop
        $xlsx = Join-Path $OutDir 'forensic_report.xlsx'
        $results | Export-Excel -Path $xlsx -WorksheetName 'Summary' -AutoSize -FreezeTopRow -BoldTopRow
        Write-Log "Excel report written: $xlsx" 'OK'
    } catch { Write-Log "Excel export skipped: $($_.Exception.Message)" 'WARN' }
}

# ============================================================================
# 9. CONSOLE SUMMARY  (sessions intentionally left CONNECTED for re-runs)
# ============================================================================
Write-Log '================ RESULTS (all meetings in mailbox) ================' 'STEP'
$results | Group-Object Classification | Sort-Object Name | ForEach-Object {
    $lvl = if ($_.Name -match 'MISSING|NO_DIAGNOSTIC|REMOVED') { 'ERROR' } elseif ($_.Name -eq 'PRESENT_OK') { 'OK' } else { 'WARN' }
    $gName = $_.Name; $gCount = $_.Count
    Write-Log ("{0,-22} : {1}" -f $gName, $gCount) $lvl
}
Write-Log '---------------- Organizer meetings only (matt.lowe-organized) ----------------' 'STEP'
$orgRows = $results | Where-Object { $_.OrganizerMeeting -eq $true }
if ((Get-Count $orgRows) -eq 0) {
    Write-Log "No meetings were positively identified as organizer-owned. If the Graph overlay 403'd and the organizer copies never wrote, expand scope to an attendee mailbox." 'WARN'
} else {
    $orgRows | Group-Object Classification | Sort-Object Name | ForEach-Object {
        $lvl = if ($_.Name -match 'MISSING|NO_DIAGNOSTIC|REMOVED') { 'ERROR' } elseif ($_.Name -eq 'PRESENT_OK') { 'OK' } else { 'WARN' }
        $gName = $_.Name; $gCount = $_.Count
        Write-Log ("{0,-22} : {1}" -f $gName, $gCount) $lvl
    }
}
Write-Log "Summary CSV : $SummaryCsv"  'INFO'
Write-Log "Timelines   : $TimelineTxt" 'INFO'
Write-Log "Raw objects : $RawCsv"      'INFO'
Write-Log "Sessions left connected (EXO + Graph). Re-run anytime; provisioning is idempotent." 'OK'
Write-Log "Forensic run complete." 'STEP'
