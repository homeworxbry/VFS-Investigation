#Requires -Version 5.1
<#
================================================================================
 Test-CalDiag-PS51.ps1
================================================================================
 PURPOSE
   Runs the same Get-CalendarDiagnosticObjects tests as the main diagnostic
   script but in a Windows PowerShell 5.1 process.  Used to determine whether
   the tenant-wide backend failure seen under PS7 is a PS-version artefact or
   a genuine Microsoft backend issue.

 EXPECTED OUTCOMES
   - If all tests pass here but failed in PS7:
       The failure is a PowerShell 7 / ExchangeOnlineManagement REST-layer
       incompatibility with Get-CalendarDiagnosticObjects.  Use PS5.1 until
       Microsoft fixes the module behaviour, and include this finding in the
       support case.

   - If all tests fail here with the same server-side error as PS7:
       The failure is a Microsoft backend issue, not a PS version issue.
       This rules out the PS7 hypothesis and strengthens the case for a
       tenant-level backend service fault.

   - If tests fail here with a DIFFERENT error than PS7:
       Both versions are broken but for different reasons.  Report both
       error messages to Microsoft.

 RUN (must be run in Windows PowerShell 5.1, NOT PS7)
   Verify first:  $PSVersionTable.PSVersion   # should show 5.x
   Then:          .\Test-CalDiag-PS51.ps1

 NOTE
   This script deliberately has no PS7 dependencies.  Do NOT add null-
   conditional (?.) operators, ternary expressions, or other PS7 syntax.
================================================================================
#>

[CmdletBinding()]
param(
    [string]$TargetMailbox    = 'matt.lowe@venturafs.com',
    [string]$ColleagueMailbox = 'alan.tsarovsky@venturafs.com',
    [string]$OutputRoot       = (Get-Location).Path,

    # Known-present meeting subject (positive control).
    [string]$KnownPresentSubject = "Managers' Meeting",

    # Known-missing meeting.
    [string]$KnownMissingSubject = 'Matt & Alan (Ventura Fund Services) / Joe and Brian (NSP Capital)',
    [string]$KnownMissingGoid    = '040000008200E00074C5B7101A82E00800000000B0EB63C38CB8DC010000000000000000100000005DEE76618B903D4387B9847A371583A3'
)

$ErrorActionPreference = 'Stop'

$stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
$ReportTxt = Join-Path $OutputRoot "CalDiag_PS51_$stamp.txt"
$pass = 0; $fail = 0

function Write-R {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "{0} [PS51/{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Msg
    switch ($Level) {
        'PASS'  { Write-Host $line -ForegroundColor Green }
        'FAIL'  { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    Add-Content -Path $ReportTxt -Value $line
}

function Test-Diag {
    param([string]$Label, [scriptblock]$Query)
    Write-R "TEST: $Label" 'STEP'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $result = & $Query
        $sw.Stop()
        $n = 0
        if ($null -ne $result) { foreach ($x in $result) { $n++ } }
        Write-R "  PASS -- $n row(s) in $($sw.ElapsedMilliseconds)ms" 'PASS'
        $script:pass++
        return $result
    } catch {
        $sw.Stop()
        $msg = $_.Exception.Message
        $inner = ''
        if ($_.Exception.InnerException) { $inner = $_.Exception.InnerException.Message }
        Write-R "  FAIL -- $msg" 'FAIL'
        if ($inner -and $inner -ne $msg) { Write-R "  Inner: $inner" 'FAIL' }
        $script:fail++
        return $null
    }
}

# ---- Banner -----------------------------------------------------------------
Write-R ("=" * 60) 'STEP'
Write-R "  Test-CalDiag-PS51  --  PS version: $($PSVersionTable.PSVersion)" 'STEP'
Write-R ("=" * 60) 'STEP'

# ---- PS version guard -------------------------------------------------------
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-R "WARNING: Running in PS $($PSVersionTable.PSVersion). This script is intended for PS5.1." 'WARN'
    Write-R "Results will NOT be comparable to the PS7 tests. Run in Windows PowerShell 5.1." 'WARN'
}

# ---- Module -----------------------------------------------------------------
$mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
       Sort-Object Version -Descending | Select-Object -First 1
if ($mod) {
    Write-R "ExchangeOnlineManagement $($mod.Version) available" 'INFO'
} else {
    Write-R "ExchangeOnlineManagement NOT installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser" 'FAIL'
    return
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Write-R "Module loaded." 'INFO'

# ---- Session ----------------------------------------------------------------
$conn = $null
try { $conn = Get-ConnectionInformation -ErrorAction Stop } catch {}
$live = $conn | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' }
if (-not $live) {
    Write-R "No active session -- connecting..." 'WARN'
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    $conn  = Get-ConnectionInformation -ErrorAction Stop
    $live  = $conn | Where-Object { $_.State -eq 'Connected' -or $_.TokenStatus -eq 'Active' }
}
$AdminUpn = ($live | Select-Object -First 1).UserPrincipalName
Write-R "Session active as: $AdminUpn" 'INFO'

# ---- Tests ------------------------------------------------------------------

# T1: Admin's own mailbox, ResultSize 1 (purest baseline)
Test-Diag "T1 -- Admin own mailbox ($AdminUpn), ResultSize 1" {
    Get-CalendarDiagnosticObjects -Identity $AdminUpn -ResultSize 1 -ErrorAction Stop
}

# T2: Target mailbox, ResultSize 1 (isolate target vs. admin)
Test-Diag "T2 -- Target mailbox ($TargetMailbox), ResultSize 1" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox -ResultSize 1 -ErrorAction Stop
}

# T3: Target, known-present meeting (positive control)
Test-Diag "T3 -- Target, known-present meeting: '$KnownPresentSubject'" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -Subject $KnownPresentSubject -ExactMatch $false `
        -ResultSize 10 -ErrorAction Stop
}

# T4: Target, known-missing meeting by subject
$t4 = Test-Diag "T4 -- Target, known-MISSING meeting: '$KnownMissingSubject'" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -Subject $KnownMissingSubject -ExactMatch $true `
        -ResultSize 50 -ErrorAction Stop
}
if ($t4 -ne $null) {
    $n = 0; foreach ($x in $t4) { $n++ }
    if ($n -gt 0) {
        Write-R "  *** T4: $n rows found -- meeting DID touch Matt's mailbox ***" 'WARN'
        foreach ($row in $t4 | Select-Object -First 20) {
            $action = 'unknown'; $client = 'unknown'
            try { $action = $row.CalendarLogTriggerAction } catch {}
            try { $client = $row.ClientInfoString }         catch {}
            Write-R "    Action=$action  Client=$client" 'INFO'
        }
    } else {
        Write-R "  T4: 0 rows -- meeting never touched Matt's mailbox." 'WARN'
    }
}

# T5: Target, known-missing meeting by CleanGlobalObjectId (no subject match)
$t5 = Test-Diag "T5 -- Target, known-missing meeting by MeetingID (CleanGOID)" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -MeetingID $KnownMissingGoid -ResultSize 50 -ErrorAction Stop
}
if ($t5 -ne $null) {
    $n = 0; foreach ($x in $t5) { $n++ }
    if ($n -gt 0) {
        Write-R "  *** T5: $n rows by MeetingID -- definitive: meeting was processed by mailbox ***" 'WARN'
        foreach ($row in $t5 | Select-Object -First 20) {
            $action = 'unknown'; $client = 'unknown'
            try { $action = $row.CalendarLogTriggerAction } catch {}
            try { $client = $row.ClientInfoString }         catch {}
            Write-R "    Action=$action  Client=$client" 'INFO'
        }
    } else {
        Write-R "  T5: 0 rows by MeetingID -- meeting never touched this mailbox." 'WARN'
    }
}

# T6: Colleague mailbox (alan.tsarovsky) -- cross-mailbox baseline
Test-Diag "T6 -- Colleague mailbox ($ColleagueMailbox), ResultSize 1" {
    Get-CalendarDiagnosticObjects -Identity $ColleagueMailbox -ResultSize 1 -ErrorAction Stop
}

# T7: With -ShouldDecodeEnums (different parameter path through the cmdlet)
Test-Diag "T7 -- Target, ShouldDecodeEnums, ResultSize 1" {
    Get-CalendarDiagnosticObjects -Identity $TargetMailbox `
        -ShouldDecodeEnums $true -ResultSize 1 -ErrorAction Stop
}

# ---- Summary ----------------------------------------------------------------
$divider = "=" * 60
$lines = @(
    ""
    $divider
    "PS VERSION: $($PSVersionTable.PSVersion)"
    "MODULE:     ExchangeOnlineManagement $($mod.Version)"
    "RESULTS:    $pass PASS   $fail FAIL"
    ""
)

if ($fail -eq 0) {
    $lines += "CONCLUSION: Get-CalendarDiagnosticObjects WORKS in PS5.1."
    $lines += "The PS7 failure is a PS-version / module runtime issue."
    $lines += "Use PS5.1 for diagnostic work until Microsoft resolves the module behaviour."
    $lines += "Include both the PS5.1 success and PS7 failure screenshots in the support case."
} elseif ($pass -eq 0) {
    $lines += "CONCLUSION: Get-CalendarDiagnosticObjects FAILS in PS5.1 with the same error."
    $lines += "This rules out PS7 as the cause. The failure is a Microsoft backend issue"
    $lines += "affecting this tenant regardless of PowerShell version."
    $lines += "Escalate to Microsoft with both this report and the PS7 report."
} else {
    $lines += "CONCLUSION: MIXED -- some tests passed, some failed."
    $lines += "Review individual test results above for the pattern."
}

$lines += ""
$lines += "Report: $ReportTxt"
$lines += $divider

foreach ($l in $lines) { Add-Content -Path $ReportTxt -Value $l }
Write-R "" 'INFO'
Write-R $divider 'STEP'
if ($fail -eq 0) {
    Write-R "RESULT: $pass/$($pass+$fail) PASS -- cmdlet WORKS in PS5.1" 'PASS'
} elseif ($pass -eq 0) {
    Write-R "RESULT: 0/$($pass+$fail) PASS -- cmdlet FAILS in PS5.1 with same error (backend confirmed)" 'FAIL'
} else {
    Write-R "RESULT: $pass/$($pass+$fail) PASS -- mixed results, review above" 'WARN'
}
Write-R $divider 'STEP'
Write-R "Report: $ReportTxt" 'INFO'
Write-R "Done." 'STEP'
