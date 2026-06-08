#Requires -Version 7.0
<#
================================================================================
 Test-CalendarDiagnosticAccess.ps1
================================================================================
 PURPOSE
   Systematically tests every known reason Get-CalendarDiagnosticObjects might
   be failing for a target mailbox and produces a written report.  Where the
   cause is a missing RBAC role, the script self-assigns it (authorized by
   tenant admin) and re-tests to confirm the fix.

 PHASES
   1  Module & session health          -- is the right module version loaded?
   2  RBAC audit + auto-fix            -- does the admin hold Calendar Diagnostics?
   3  Target mailbox health            -- holds, rules, database (for Microsoft)
   4  Functional tests (A-H)           -- does the cmdlet actually work?
        A  Known-present meeting on target  (positive control)
        B  Known-missing meeting on target  (key diagnostic question)
        C  Known-missing meeting on colleague  (cross-mailbox / tenant check)
        D  Admin's own mailbox  (baseline: does cmdlet work for anyone?)
        E  MeetingID (CleanGlobalObjectId) query for missing meeting on target
        F  MeetingID query on colleague  (confirms meeting has log data in tenant)
        G  Minimal ResultSize 1  (throttle-avoidance check)
        H  Second missing meeting on target  (repeat-failure pattern check)
   5  Error pattern analysis           -- classify the failure
   6  Final report + verdict           -- written to file + console

 KEY FINDING INTERPRETATION
   Tests B + E answer: "was this meeting EVER processed by Matt's mailbox?"
     0 rows  →  meeting never arrived at / was processed by the mailbox.
               Likely a transport or meeting-request delivery failure.
     rows found  →  meeting DID touch the mailbox; look at
               CalendarLogTriggerAction and ClientInfoString for what removed it.

   Tests C + F answer: "is the cmdlet and meeting data reachable at all?"
     If C/F pass but B/E fail  →  cmdlet works; issue is mailbox-specific.
     If all fail  →  tenant-wide / backend / RBAC issue.

 AUTO-FIX SCOPE
   Only Management Role Assignments are modified.  If the Calendar Diagnostics
   RBAC role is absent for the running admin, one direct assignment is created
   and the functional tests are re-run.  No mailbox data, calendar items,
   permissions, rules, or other configuration is altered.

 OUTPUT
   CalDiagAccess_<timestamp>.txt   human-readable full report
   CalDiagAccess_<timestamp>.csv   machine-readable test results (one row per test)
   Both written to $OutputRoot (the folder the script is run from).

 RUN
   .\Test-CalendarDiagnosticAccess.ps1
   .\Test-CalendarDiagnosticAccess.ps1 -TargetMailbox matt.lowe@venturafs.com
================================================================================
#>

[CmdletBinding()]
param(
    [string]$TargetMailbox     = 'matt.lowe@venturafs.com',
    [string]$ColleagueMailbox  = 'alan.tsarovsky@venturafs.com',
    [string]$OutputRoot        = (Get-Location).Path,

    # Known-PRESENT meeting on the target mailbox (positive control).
    # Must be a meeting confirmed present in mattphase2investigation.csv.
    [string]$KnownPresentSubject = "Managers' Meeting",

    # Known-MISSING meeting -- subject and iCalUId confirmed by Graph investigation.
    [string]$KnownMissingSubject  = 'Matt & Alan (Ventura Fund Services) / Joe and Brian (NSP Capital)',
    [string]$KnownMissingICalUId  = '040000008200E00074C5B7101A82E00800000000B0EB63C38CB8DC010000000000000000100000005DEE76618B903D4387B9847A371583A3',

    # Second missing meeting (6-colleague corroboration -- strongest invitee signal).
    [string]$KnownMissing2Subject = 'Ventura/Metric Point - discuss NSP true up at 1st close',
    [string]$KnownMissing2ICalUId = '040000008200E00074C5B7101A82E008000000005040038A85DDDC01000000000000000010000000F9E11B7E664C4740BA24AE3DB2C8DA6C',

    # When true, the script assigns the Calendar Diagnostics role to the
    # running admin if the role is found to be missing.  Authorized by admin.
    [bool]$AllowAutoFix = $true
)

$ErrorActionPreference = 'Stop'

# ---- Output -----------------------------------------------------------------
$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$ReportTxt  = Join-Path $OutputRoot "CalDiagAccess_$stamp.txt"
$ReportCsv  = Join-Path $OutputRoot "CalDiagAccess_$stamp.csv"
$testResults = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message,
          [ValidateSet('INFO','WARN','PASS','FAIL','FIX','STEP','ERROR')][string]$Level = 'INFO')
    if ($Message -eq '') { Write-Host ''; Add-Content -Path $ReportTxt -Value ''; return }
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "$ts [$Level] $Message"
    switch ($Level) {
        'PASS'  { Write-Host $line -ForegroundColor Green }
        'FAIL'  { Write-Host $line -ForegroundColor Red }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'FIX'   { Write-Host $line -ForegroundColor Cyan }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    Add-Content -Path $ReportTxt -Value $line
}

function Add-TestResult {
    param([string]$Phase, [string]$Test, [string]$Status,
          [string]$Detail, [string]$Recommendation = '')
    $obj = [pscustomobject]@{
        Phase          = $Phase
        Test           = $Test
        Status         = $Status
        Detail         = $Detail
        Recommendation = $Recommendation
    }
    $testResults.Add($obj)
    $lvl = switch ($Status) {
        'PASS' { 'PASS' } 'FAIL' { 'FAIL' } 'WARN' { 'WARN' }
        'FIX'  { 'FIX'  } 'INFO' { 'INFO' }
        default { 'INFO' }
    }
    Write-Log ("  [{0}] {1} : {2}" -f $Status, $Test, $Detail) $lvl
}

trap {
    Write-Log ("FATAL {0}: {1}" -f $_.Exception.GetType().FullName, $_.Exception.Message) 'ERROR'
    Write-Log ("  at: {0}" -f $_.InvocationInfo.PositionMessage) 'ERROR'
    if ($testResults.Count -gt 0) { $testResults | Export-Csv $ReportCsv -NoTypeInformation -Encoding UTF8 }
    break
}

# Convert a Graph iCalUId to a CleanGlobalObjectId by zeroing the occurrence-
# date bytes (hex chars 32-39 = bytes 16-19).  For non-recurring meetings
# those bytes are already zero and the value passes through unchanged.
function ConvertTo-CleanGoid {
    param([Parameter(Mandatory)][string]$ICalUId)
    if ($ICalUId.Length -lt 48) { return $ICalUId }
    return $ICalUId.Substring(0, 32) + '00000000' + $ICalUId.Substring(40)
}

# Run a Get-CalendarDiagnosticObjects call and return a result object.
# Never throws -- all failures are captured and classified.
function Invoke-DiagTest {
    param(
        [string]$Phase,
        [string]$TestId,
        [string]$TestName,
        [scriptblock]$Query
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $raw = & $Query
        $sw.Stop()
        $count = 0; foreach ($x in @($raw)) { if ($null -ne $x) { $count++ } }
        Add-TestResult $Phase $TestName 'PASS' "$count row(s) in $($sw.ElapsedMilliseconds)ms"
        return [pscustomobject]@{ Ok = $true; Rows = $raw; Count = $count; ErrorMessage = '' }
    } catch {
        $sw.Stop()
        $msg   = $_.Exception.Message
        $inner = $_.Exception.InnerException?.Message
        $full  = if ($inner -and $inner -ne $msg) { "$msg  [inner: $inner]" } else { $msg }

        $rec = switch -Regex ($full) {
            'insufficient.access|access.denied|not.assigned|permission|Forbidden|\b403\b|role|unauthorized|\b401\b' {
                'RBAC/permission issue. The running admin may lack the Calendar Diagnostics management role. See Phase 2.'
            }
            'throttl|ExceededBudget|\b429\b|exceeded.the.allowed|TooManyRequests' {
                'Request was throttled. Wait a few minutes and retry with -ResultSize reduced.'
            }
            'timeout|timed.out|ServiceUnavailable|\b503\b|\b502\b|InternalServerError|\b500\b|server.error|backend' {
                'Microsoft backend / service error. Not a client-side issue -- escalate to Microsoft support.'
            }
            'token.expired|authentication|re.authenticate|re.connect|session.has.expired' {
                'Auth token expired. Run: Connect-ExchangeOnline'
            }
            'mailbox.not.found|MailboxNotFound|\b404\b|object.not.found|does.not.exist' {
                'Mailbox not found or inaccessible under this admin account.'
            }
            'corrupt|damage|inconsist' {
                'Mailbox or diagnostic store may be corrupted. Provide ExchangeGuid to Microsoft support.'
            }
            default {
                "Unclassified error -- review full detail above."
            }
        }
        Add-TestResult $Phase $TestName 'FAIL' $full $rec
        return [pscustomobject]@{ Ok = $false; Rows = $null; Count = 0; ErrorMessage = $full }
    }
}

# ============================================================================
Write-Log ("=" * 72) 'STEP'
Write-Log "  Test-CalendarDiagnosticAccess  --  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 'STEP'
Write-Log ("=" * 72) 'STEP'
Write-Log "Target mailbox : $TargetMailbox" 'INFO'
Write-Log "Colleague      : $ColleagueMailbox" 'INFO'
Write-Log "Report         : $ReportTxt" 'INFO'

# ============================================================================
# PHASE 1 -- Module & session health
# ============================================================================
Write-Log "" 'INFO'
Write-Log "--- PHASE 1: Module & Session Health ---" 'STEP'

$exoMod = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
          Sort-Object Version -Descending | Select-Object -First 1
if ($exoMod) {
    $modVer = $exoMod.Version.ToString()
    $modStatus = if ([version]$modVer -ge [version]'3.0.0') { 'PASS' } else { 'WARN' }
    Add-TestResult 'Phase1' 'P1-01' 'EXO Module Version' $modStatus `
        "ExchangeOnlineManagement $modVer installed" `
        $(if ($modStatus -eq 'WARN') { 'Upgrade: Install-Module ExchangeOnlineManagement -Force.  v3.x has fixes relevant to Get-CalendarDiagnosticObjects.' } else { '' })
} else {
    Add-TestResult 'Phase1' 'P1-01' 'EXO Module Version' 'FAIL' 'ExchangeOnlineManagement NOT installed' `
        'Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force'
    Write-Log "FATAL: EXO module missing. Cannot continue." 'ERROR'
    return
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Session check / connect
$conn = $null
try { $conn = Get-ConnectionInformation -ErrorAction Stop } catch {}
$live = $conn | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' }
if ($live) {
    $AdminUpn = ($live | Select-Object -First 1).UserPrincipalName
    Add-TestResult 'Phase1' 'P1-02' 'EXO Session' 'PASS' "Active session as $AdminUpn"
} else {
    Write-Log "No active EXO session -- connecting..." 'WARN'
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $conn2 = Get-ConnectionInformation -ErrorAction Stop
        $AdminUpn = ($conn2 | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' } |
                    Select-Object -First 1).UserPrincipalName
        Add-TestResult 'Phase1' 'P1-02' 'EXO Session' 'PASS' "Connected (new session) as $AdminUpn"
    } catch {
        Add-TestResult 'Phase1' 'P1-02' 'EXO Session' 'FAIL' $_.Exception.Message 'Run Connect-ExchangeOnline manually first.'
        Write-Log "Cannot proceed without an EXO session." 'ERROR'
        $testResults | Export-Csv $ReportCsv -NoTypeInformation -Encoding UTF8
        return
    }
}

# Basic EXO API health (not calendar-specific)
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Add-TestResult 'Phase1' 'P1-03' 'EXO API Baseline' 'PASS' 'Get-OrganizationConfig succeeded (EXO APIs responsive)'
} catch {
    Add-TestResult 'Phase1' 'P1-03' 'EXO API Baseline' 'FAIL' $_.Exception.Message `
        'Core EXO API is failing -- likely a wider service issue. Check admin.microsoft.com service health.'
}

# ============================================================================
# PHASE 2 -- RBAC Audit + Auto-Fix
# ============================================================================
Write-Log "" 'INFO'
Write-Log "--- PHASE 2: RBAC -- Calendar Diagnostics Role ---" 'STEP'

$CalDiagRole = 'Calendar Diagnostics'

# Check direct user assignment
$directHasRole = $false
try {
    $da = @(Get-ManagementRoleAssignment -Role $CalDiagRole -RoleAssignee $AdminUpn -ErrorAction SilentlyContinue)
    if ($da.Count -gt 0) { $directHasRole = $true }
} catch { Write-Log "  Direct-assignment query failed: $($_.Exception.Message)" 'WARN' }

# Check via role group membership
$groupHasRole   = $false
$adminGroups    = @()
$groupWithRole  = ''
if (-not $directHasRole) {
    try {
        # Get all role groups this admin is a member of
        $allGroups = @(Get-ManagementRoleAssignment -RoleAssignee $AdminUpn -ErrorAction Stop |
                      Where-Object { $_.RoleAssigneeType -in @('RoleGroup','SecurityGroup','USG') } |
                      Select-Object -ExpandProperty RoleAssigneeName -Unique)
        $adminGroups = $allGroups
        foreach ($rg in $allGroups) {
            $rgMatch = @(Get-ManagementRoleAssignment -Role $CalDiagRole -RoleAssignee $rg -ErrorAction SilentlyContinue)
            if ($rgMatch.Count -gt 0) { $groupHasRole = $true; $groupWithRole = $rg; break }
        }
    } catch { Write-Log "  Role-group membership query failed: $($_.Exception.Message)" 'WARN' }
}

$groupList = if ($adminGroups.Count -gt 0) { $adminGroups -join '; ' } else { '(none detected)' }
Write-Log "  Admin role groups detected: $groupList" 'INFO'

$hasRole = $directHasRole -or $groupHasRole
$roleVia = if ($directHasRole) { 'direct assignment' } elseif ($groupHasRole) { "role group: $groupWithRole" } else { 'NOT FOUND' }

if ($hasRole) {
    Add-TestResult 'Phase2' 'P2-01' 'Calendar Diagnostics Role' 'PASS' "Assigned via $roleVia"
} else {
    Add-TestResult 'Phase2' 'P2-01' 'Calendar Diagnostics Role' 'FAIL' `
        "Role NOT present for $AdminUpn.  Role groups: $groupList" `
        "Fix: New-ManagementRoleAssignment -Role 'Calendar Diagnostics' -User $AdminUpn"

    if ($AllowAutoFix) {
        Write-Log "  AUTO-FIX: assigning Calendar Diagnostics role directly to $AdminUpn ..." 'FIX'
        $fixAssignName = "CalDiag-AutoFix-$($AdminUpn.Split('@')[0])-$stamp"
        try {
            New-ManagementRoleAssignment -Name $fixAssignName -Role $CalDiagRole -User $AdminUpn -ErrorAction Stop | Out-Null
            Add-TestResult 'Phase2' 'P2-02' 'RBAC Auto-Fix' 'FIX' `
                "Role assigned as '$fixAssignName'.  Waiting 90s for propagation before re-testing."
            Write-Log "  Waiting 90s for RBAC propagation..." 'FIX'
            Start-Sleep -Seconds 90
            $hasRole = $true  # flag so Phase 4 proceeds with optimism and re-verifies
        } catch {
            Add-TestResult 'Phase2' 'P2-02' 'RBAC Auto-Fix' 'FAIL' $_.Exception.Message `
                "Manual fix: in EAC Permissions > Admin Roles, add '$CalDiagRole' to your role group, or run the New-ManagementRoleAssignment above."
        }
    } else {
        Add-TestResult 'Phase2' 'P2-02' 'RBAC Auto-Fix' 'INFO' 'AllowAutoFix=$false -- skipped.'
    }
}

# Enumerate all roles the admin currently has (informational, useful for Microsoft)
try {
    $adminRoles = @(Get-ManagementRoleAssignment -GetEffectiveUsers -ErrorAction Stop |
                   Where-Object { $_.EffectiveUserName -ieq $AdminUpn } |
                   Select-Object -ExpandProperty Role -Unique | Sort-Object)
    $roleList = $adminRoles -join '; '
    Add-TestResult 'Phase2' 'P2-03' 'Effective Roles (informational)' 'INFO' $roleList
} catch {
    Add-TestResult 'Phase2' 'P2-03' 'Effective Roles (informational)' 'WARN' `
        "Could not enumerate (may time out on large tenants): $($_.Exception.Message)"
}

# ============================================================================
# PHASE 3 -- Target Mailbox Health
# ============================================================================
Write-Log "" 'INFO'
Write-Log "--- PHASE 3: Target Mailbox Health ($TargetMailbox) ---" 'STEP'

$mbx = $null
try {
    $mbx = Get-Mailbox -Identity $TargetMailbox -ErrorAction Stop
    Add-TestResult 'Phase3' 'P3-01' 'Mailbox Reachable' 'PASS' `
        "DisplayName='$($mbx.DisplayName)'  Database='$($mbx.Database)'  ExchangeGuid=$($mbx.ExchangeGuid)"
    Write-Log "  Database (for Microsoft support): $($mbx.Database)" 'INFO'
    Write-Log "  ExchangeGuid: $($mbx.ExchangeGuid)" 'INFO'
} catch {
    Add-TestResult 'Phase3' 'P3-01' 'Mailbox Reachable' 'FAIL' $_.Exception.Message `
        "Verify $TargetMailbox exists and the admin has appropriate EXO access."
}

if ($mbx) {
    # Litigation hold
    $lhStatus = if ($mbx.LitigationHoldEnabled) { 'WARN' } else { 'PASS' }
    Add-TestResult 'Phase3' 'P3-02' 'Litigation Hold' $lhStatus `
        "LitigationHoldEnabled=$($mbx.LitigationHoldEnabled)" `
        $(if ($mbx.LitigationHoldEnabled) { 'Litigation hold active. May affect diagnostic log accessibility in some configurations. Note for Microsoft.' } else { '' })

    # In-place holds
    $holdCount = 0; foreach ($h in $mbx.InPlaceHolds) { $holdCount++ }
    $holdStatus = if ($holdCount -gt 0) { 'WARN' } else { 'PASS' }
    Add-TestResult 'Phase3' 'P3-03' 'In-Place / Compliance Holds' $holdStatus `
        "Count: $holdCount  Values: $($mbx.InPlaceHolds -join ', ')" `
        $(if ($holdCount -gt 0) { 'Compliance holds present. Provide to Microsoft in support case.' } else { '' })

    # Mailbox type
    Add-TestResult 'Phase3' 'P3-04' 'Mailbox Type' 'INFO' `
        "RecipientTypeDetails=$($mbx.RecipientTypeDetails)  MailboxPlan=$($mbx.MailboxPlan)"

    # Audit enabled
    $auditStatus = if ($mbx.AuditEnabled) { 'PASS' } else { 'WARN' }
    Add-TestResult 'Phase3' 'P3-05' 'Audit Logging' $auditStatus `
        "AuditEnabled=$($mbx.AuditEnabled)  AuditLogAgeLimit=$($mbx.AuditLogAgeLimit)" `
        $(if (-not $mbx.AuditEnabled) { 'Mailbox audit logging is off. Enable for future forensic visibility: Set-Mailbox -Identity $TargetMailbox -AuditEnabled $true' } else { '' })

    # Inbox rules that could drop or move meeting requests
    try {
        $allRules = @(Get-InboxRule -Mailbox $TargetMailbox -ErrorAction Stop)
        $suspectRules = @($allRules | Where-Object {
            $_.Enabled -and
            ($_.DeleteMessage -or $_.MoveToFolder -or $_.PermanentDelete -or $_.RedirectTo -or $_.ForwardTo) -and
            ($_.SubjectContainsWords -or $_.SubjectOrBodyContainsWords -or
             $_.MessageTypeMatches -or $_.From -or $_.SentTo)
        })
        $ruleStatus = if ($suspectRules.Count -gt 0) { 'WARN' } else { 'PASS' }
        $ruleDetail = "Total rules: $($allRules.Count)  Potentially relevant (delete/move/redirect + filter): $($suspectRules.Count)"
        $ruleRec    = ''
        if ($suspectRules.Count -gt 0) {
            $ruleNames = ($suspectRules | ForEach-Object { "'$($_.Name)' (delete=$($_.DeleteMessage) move=$($_.MoveToFolder -ne $null))" }) -join ';  '
            $ruleDetail += "  Rules: $ruleNames"
            $ruleRec = 'Review these inbox rules -- one may be discarding meeting request emails before they process to the calendar.'
        }
        Add-TestResult 'Phase3' 'P3-06' 'Inbox Rules (suspect)' $ruleStatus $ruleDetail $ruleRec
    } catch {
        Add-TestResult 'Phase3' 'P3-06' 'Inbox Rules' 'WARN' "Could not retrieve: $($_.Exception.Message)"
    }

    # Delegate access (someone else may have write access to Matt's calendar)
    try {
        $delegates = @(Get-MailboxFolderPermission -Identity "$($TargetMailbox):\Calendar" -ErrorAction Stop |
                      Where-Object { $_.User -ne 'Default' -and $_.User -ne 'Anonymous' -and
                                     $_.AccessRights -match 'Editor|Owner|Author' })
        $delStatus = if ($delegates.Count -gt 0) { 'WARN' } else { 'PASS' }
        $delDetail = "Write-capable calendar delegates: $($delegates.Count)"
        if ($delegates.Count -gt 0) {
            $delDetail += "  -- " + ($delegates | ForEach-Object { "$($_.User) [$($_.AccessRights -join ',')]" } | Select-Object -First 10 | Join-String -Separator '; ')
        }
        Add-TestResult 'Phase3' 'P3-07' 'Calendar Write Delegates' $delStatus $delDetail `
            $(if ($delegates.Count -gt 0) { 'Delegates with write access could delete calendar items. Verify these are expected.' } else { '' })
    } catch {
        Add-TestResult 'Phase3' 'P3-07' 'Calendar Write Delegates' 'WARN' "Could not retrieve: $($_.Exception.Message)"
    }
}

# ============================================================================
# PHASE 4 -- Functional Tests
# ============================================================================
Write-Log "" 'INFO'
Write-Log "--- PHASE 4: Functional Tests ---" 'STEP'

$cleanGoid1 = ConvertTo-CleanGoid -ICalUId $KnownMissingICalUId
$cleanGoid2 = ConvertTo-CleanGoid -ICalUId $KnownMissing2ICalUId
Write-Log "  CleanGOID (meeting 1): $cleanGoid1" 'INFO'
Write-Log "  CleanGOID (meeting 2): $cleanGoid2" 'INFO'

# Test A -- Known-present meeting on target (positive control)
Write-Log "" 'INFO'
Write-Log "  Test A: Known-present meeting on target (positive control)" 'INFO'
$testA = Invoke-DiagTest 'Phase4' 'P4-A' "A: Known-present '$KnownPresentSubject' on target" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -Subject $KnownPresentSubject -ExactMatch $false -ResultSize 10 -ErrorAction Stop
}

# Test B -- Known-missing meeting on target (the key question)
Write-Log "" 'INFO'
Write-Log "  Test B: Known-MISSING meeting on target (key diagnostic question)" 'INFO'
$testB = Invoke-DiagTest 'Phase4' 'P4-B' "B: Known-missing '$KnownMissingSubject' on target" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -Subject $KnownMissingSubject -ExactMatch $true -ResultSize 50 -ErrorAction Stop
}
if ($testB.Ok) {
    if ($testB.Count -gt 0) {
        Write-Log "  *** Test B: $($testB.Count) log row(s) found -- meeting DID touch Matt's mailbox. ***" 'WARN'
        Write-Log "  Printing action log (first 20 rows):" 'INFO'
        foreach ($row in @($testB.Rows) | Select-Object -First 20) {
            $action  = try { $row.CalendarLogTriggerAction } catch { '?' }
            $client  = try { $row.ClientInfoString }         catch { '?' }
            $resType = try { $row.ResponseType }             catch { '?' }
            $ts2     = try { $row.OriginalLastModifiedTime } catch { '?' }
            Write-Log "    [$ts2]  Action=$action  ResponseType=$resType  Client=$client" 'INFO'
        }
        Add-TestResult 'Phase4' 'P4-B-DETAIL' 'B: Meeting HAS log entries on target' 'WARN' `
            "Meeting was processed by Matt's mailbox ($($testB.Count) log rows). See Action/ClientInfoString above -- this identifies what removed it." `
            'Filter for Delete/MoveToDeletedItems/HardDelete in CalendarLogTriggerAction to find the removal event and the ClientInfoString.'
    } else {
        Write-Log "  *** Test B: 0 rows -- NO trace of this meeting on Matt's mailbox. ***" 'WARN'
        Add-TestResult 'Phase4' 'P4-B-DETAIL' 'B: No log entries on target' 'WARN' `
            "0 rows returned. The meeting was NEVER processed by Matt's mailbox." `
            'Failure is at the delivery/transport layer, not a client-side deletion. Raise with Microsoft: meeting request never arrived at this mailbox.'
    }
}

# Test C -- Known-missing meeting on colleague (cross-mailbox check)
Write-Log "" 'INFO'
Write-Log "  Test C: Known-missing meeting on COLLEAGUE ($ColleagueMailbox) -- cross-tenant check" 'INFO'
$testC = Invoke-DiagTest 'Phase4' 'P4-C' "C: Known-missing meeting on $ColleagueMailbox (should exist there)" {
    Get-CalendarDiagnosticObjects -Identity $ColleagueMailbox `
        -Subject $KnownMissingSubject -ExactMatch $true -ResultSize 10 -ErrorAction Stop
}
if ($testC.Ok -and $testC.Count -gt 0) {
    Write-Log "  Test C: $($testC.Count) rows on colleague -- meeting has diagnostic data in tenant. Cmdlet works cross-mailbox." 'PASS'
} elseif ($testC.Ok -and $testC.Count -eq 0) {
    Write-Log "  Test C: 0 rows on colleague -- meeting may not have a log entry on that mailbox either (normal if colleague only RSVP'd)" 'INFO'
}

# Test D -- Admin's own mailbox baseline
Write-Log "" 'INFO'
Write-Log "  Test D: Admin's own mailbox baseline (does cmdlet work for ANY mailbox?)" 'INFO'
$testD = Invoke-DiagTest 'Phase4' 'P4-D' "D: Admin mailbox $AdminUpn (baseline)" {
    Get-CalendarDiagnosticObjects -Identity $AdminUpn -ResultSize 1 -ErrorAction Stop
}

# Test E -- MeetingID (CleanGOID) on target
Write-Log "" 'INFO'
Write-Log "  Test E: MeetingID query on target (precise -- no subject matching ambiguity)" 'INFO'
$testE = Invoke-DiagTest 'Phase4' 'P4-E' "E: MeetingID (CleanGOID) for missing meeting on target" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -MeetingID $cleanGoid1 -ResultSize 50 -ErrorAction Stop
}
if ($testE.Ok) {
    if ($testE.Count -gt 0) {
        Write-Log "  *** Test E: $($testE.Count) rows by MeetingID on target -- definitive confirmation meeting touched mailbox. ***" 'WARN'
        foreach ($row in @($testE.Rows) | Select-Object -First 20) {
            $action = try { $row.CalendarLogTriggerAction } catch { '?' }
            $client = try { $row.ClientInfoString }         catch { '?' }
            $ts2    = try { $row.OriginalLastModifiedTime } catch { '?' }
            Write-Log "    [$ts2]  Action=$action  Client=$client" 'INFO'
        }
        Add-TestResult 'Phase4' 'P4-E-DETAIL' 'E: MeetingID found on target' 'WARN' `
            "$($testE.Count) rows. Meeting was delivered and processed. ClientInfoString above is the authoring client." ''
    } else {
        Add-TestResult 'Phase4' 'P4-E-DETAIL' 'E: MeetingID NOT found on target' 'WARN' `
            '0 rows via MeetingID. Corroborates Test B: meeting never touched this mailbox.' `
            'Delivery-layer failure. Provide CleanGOID to Microsoft for transport-level investigation.'
    }
}

# Test F -- MeetingID on colleague (confirms meeting has log data in tenant at all)
Write-Log "" 'INFO'
Write-Log "  Test F: MeetingID query on colleague ($ColleagueMailbox) -- confirms meeting data exists in tenant" 'INFO'
$testF = Invoke-DiagTest 'Phase4' 'P4-F' "F: MeetingID for missing meeting on $ColleagueMailbox" {
    Get-CalendarDiagnosticObjects -Identity $ColleagueMailbox `
        -MeetingID $cleanGoid1 -ResultSize 10 -ErrorAction Stop
}
if ($testF.Ok -and $testF.Count -gt 0) {
    Write-Log "  Test F: $($testF.Count) rows -- meeting log data exists in tenant (on colleague). Cmdlet+MeetingID path works." 'PASS'
}

# Test G -- Minimal ResultSize 1 (throttle diagnostic)
Write-Log "" 'INFO'
Write-Log "  Test G: Minimal ResultSize 1 on target (throttle-avoidance check)" 'INFO'
$testG = Invoke-DiagTest 'Phase4' 'P4-G' "G: Minimal ResultSize 1 on target" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox -ResultSize 1 -ErrorAction Stop
}

# Test H -- Second missing meeting on target (repeat-pattern check)
Write-Log "" 'INFO'
Write-Log "  Test H: Second known-missing meeting on target (Ventura/Metric Point -- 6-colleague corroboration)" 'INFO'
$testH = Invoke-DiagTest 'Phase4' 'P4-H' "H: Second missing '$KnownMissing2Subject' on target (MeetingID)" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -MeetingID $cleanGoid2 -ResultSize 50 -ErrorAction Stop
}
if ($testH.Ok) {
    if ($testH.Count -gt 0) {
        Write-Log "  Test H: $($testH.Count) rows -- second meeting also found on Matt's mailbox log." 'WARN'
        Add-TestResult 'Phase4' 'P4-H-DETAIL' 'H: Second missing meeting HAS log on target' 'WARN' `
            "Both meeting 1 (Test E) and meeting 2 (Test H) have log entries. The meeting data arrived and was subsequently removed by a client." ''
    } else {
        Add-TestResult 'Phase4' 'P4-H-DETAIL' 'H: Second missing meeting NOT on target log' 'WARN' `
            "Second meeting also has 0 log rows on target. Pattern: multiple meetings never arrived at this mailbox." `
            'Systemic delivery failure. This affects meetings from multiple organizers and is not a single-meeting anomaly.'
    }
}

# ============================================================================
# PHASE 5 -- Pattern Analysis & Verdict
# ============================================================================
Write-Log "" 'INFO'
Write-Log "--- PHASE 5: Pattern Analysis & Verdict ---" 'STEP'

$p4Fails  = @($testResults | Where-Object { $_.Phase -eq 'Phase4' -and $_.Status -eq 'FAIL' })
$p4Passes = @($testResults | Where-Object { $_.Phase -eq 'Phase4' -and $_.Status -eq 'PASS' })
$p4FailCt = 0; foreach ($x in $p4Fails)  { $p4FailCt++ }
$p4PassCt = 0; foreach ($x in $p4Passes) { $p4PassCt++ }

Write-Log "  Phase 4 results: $p4PassCt PASS, $p4FailCt FAIL" 'INFO'

# Determine verdict
$verdict = 'UNKNOWN'
$verdictDetail = ''
$verdictRec    = ''

if ($p4FailCt -eq 0) {
    # Everything passed -- interpret B/E results
    if (($testB.Count + $testE.Count) -gt 0) {
        $verdict = 'CMDLET_WORKS_MEETINGS_WERE_DELETED'
        $verdictDetail = 'Get-CalendarDiagnosticObjects is functioning. The missing meetings have log entries on the target mailbox, meaning they were delivered and then removed. Check CalendarLogTriggerAction and ClientInfoString in the Test B/E output above.'
        $verdictRec    = 'Identify the deletion ClientInfoString. If it is an automation/EWS/Graph app, revoke its access or investigate the app. If it is Outlook, check for client-side rules or an old disconnected device syncing.'
    } else {
        $verdict = 'CMDLET_WORKS_MEETINGS_NEVER_ARRIVED'
        $verdictDetail = 'Get-CalendarDiagnosticObjects is functioning. The missing meetings have NO log entries on the target mailbox -- they were never delivered to this mailbox.'
        $verdictRec    = 'This is a transport / meeting-request delivery failure. Provide the iCalUIds and CleanGOIDs to Microsoft for a server-side transport trace. Also check: mail flow rules, safe sender / block lists, MX record, and inbound connector configuration.'
    }
} elseif ($p4Fails | Where-Object { $_.Detail -match 'insufficient.access|denied|permission|Forbidden|\b403\b|role|unauthorized|\b401\b' }) {
    if ($AllowAutoFix -and $hasRole) {
        $verdict = 'RBAC_FIXED_RETEST_NEEDED'
        $verdictDetail = 'RBAC role was missing and has been auto-assigned. The functional tests may have run before propagation completed.'
        $verdictRec    = 'Wait 5 minutes and re-run this script. If still failing, verify the role assignment: Get-ManagementRoleAssignment -Role "Calendar Diagnostics" -GetEffectiveUsers | Where-Object EffectiveUserName -ieq $AdminUpn'
    } else {
        $verdict = 'RBAC_FAILURE'
        $verdictDetail = 'The running admin does not have the Calendar Diagnostics management role. This is the most likely cause of the cmdlet failure.'
        $verdictRec    = "Assign the role: New-ManagementRoleAssignment -Role 'Calendar Diagnostics' -User $AdminUpn. Then wait 5 minutes and re-run."
    }
} elseif ($p4Fails | Where-Object { $_.Detail -match 'throttl|ExceededBudget|\b429\b' }) {
    $verdict = 'THROTTLED'
    $verdictDetail = 'The cmdlet calls are being throttled by Exchange Online.'
    $verdictRec    = 'Wait 10-15 minutes and re-run. Reduce -ResultSize if queries are broad. If persistent, raise with Microsoft -- tenant may have a degraded budget.'
} elseif ($p4Fails | Where-Object { $_.Detail -match 'token.expired|authentication|re.authenticate|\b401\b' }) {
    $verdict = 'AUTH_TOKEN_EXPIRED'
    $verdictDetail = 'The EXO auth token expired mid-run.'
    $verdictRec    = 'Run: Connect-ExchangeOnline, then re-run this script.'
} elseif (
    ($p4Fails | Where-Object { $_.Test -match 'target' }) -and
    ($p4Passes | Where-Object { $_.Test -match 'baseline|Admin mailbox' })
) {
    $verdict = 'MAILBOX_SPECIFIC_FAILURE'
    $verdictDetail = 'The cmdlet works on the admin mailbox but fails on the target. The diagnostic store for this specific mailbox may be corrupted or unavailable.'
    $verdictRec    = "Provide to Microsoft: TargetMailbox=$TargetMailbox, ExchangeGuid=$($mbx?.ExchangeGuid), Database=$($mbx?.Database). Request a diagnostic store repair for this mailbox."
} elseif ($p4FailCt -gt 0 -and $p4PassCt -eq 0) {
    $verdict = 'TENANT_WIDE_BACKEND_FAILURE'
    $verdictDetail = 'All functional tests failed. This is a tenant-wide or database-level Microsoft backend issue.'
    $verdictRec    = "Escalate to Microsoft. Tenant: venturafs.com / b3377764-8275-47f8-8e2d-697b136c21c1. Database: $($mbx?.Database). Request investigation of the Calendar Diagnostics service for this tenant."
} else {
    $verdict = 'PARTIAL_FAILURE'
    $verdictDetail = "Some tests passed ($p4PassCt) and some failed ($p4FailCt). Review individual test details above."
    $verdictRec    = 'Investigate the specific failing tests. The passing tests narrow down which layer is working.'
}

Add-TestResult 'Phase5' 'P5-01' 'Diagnostic Verdict' $verdict $verdictDetail $verdictRec

# ============================================================================
# PHASE 6 -- Final Report
# ============================================================================
Write-Log "" 'INFO'
$divider = "=" * 72
$lines = @(
    ""
    $divider
    "DIAGNOSTIC VERDICT: $verdict"
    $divider
    ""
    "DETAIL: $verdictDetail"
    ""
    "RECOMMENDATION: $verdictRec"
    ""
)
if ($mbx) {
    $lines += @(
        "MAILBOX DETAILS FOR MICROSOFT SUPPORT:"
        "  Mailbox       : $TargetMailbox"
        "  Database      : $($mbx.Database)"
        "  ExchangeGuid  : $($mbx.ExchangeGuid)"
        "  RecipType     : $($mbx.RecipientTypeDetails)"
        "  LitigHold     : $($mbx.LitigationHoldEnabled)"
        "  InPlaceHolds  : $($mbx.InPlaceHolds -join ', ')"
        ""
    )
}
$lines += @(
    "CLEAN GLOBAL OBJECT IDs (for Get-CalendarDiagnosticObjects -MeetingID):"
    "  Missing meeting 1 (Matt & Alan / NSP Capital 2026-03-23):"
    "    $cleanGoid1"
    ""
    "  Missing meeting 2 (Metric Point NSP true-up 2026-05-08):"
    "    $cleanGoid2"
    ""
    "PHASE 4 SUMMARY:"
    "  PASS: $p4PassCt   FAIL: $p4FailCt"
    ""
    "OUTPUT FILES:"
    "  Report : $ReportTxt"
    "  CSV    : $ReportCsv"
    $divider
)
foreach ($l in $lines) {
    Add-Content -Path $ReportTxt -Value $l
}

Write-Log ($divider) 'STEP'
Write-Log ("VERDICT: $verdict") 'STEP'
Write-Log ("DETAIL:  $verdictDetail") 'INFO'
Write-Log ("ACTION:  $verdictRec") 'INFO'
Write-Log ($divider) 'STEP'
Write-Log "" 'INFO'
Write-Log "Report : $ReportTxt" 'INFO'
Write-Log "CSV    : $ReportCsv" 'INFO'
Write-Log "Done." 'STEP'

$testResults | Export-Csv -Path $ReportCsv -NoTypeInformation -Encoding UTF8
