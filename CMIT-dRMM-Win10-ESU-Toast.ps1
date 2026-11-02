<#
.SYNOPSIS
    Windows 10 ESU license check (Year 1 / 2 / 3) with end-user toast notification.

.DESCRIPTION
    Datto RMM component - runs as SYSTEM.
      1. Exits if the device is not Windows 10 (or is LTSC/LTSB). Bypassed when usrTestDate is set.
      2. Works out which ESU year is required for today's date (or usrTestDate).
      3. Checks the local licensing store for that year's ESU add-on, and (by default) every
         prior year as well. Each must be installed AND activated (LicenseStatus = 1).
      4. If anything is missing, shows a toast (with optional hero image) to every logged-on
         user via a short-lived scheduled task (BUILTIN\Users, limited rights).
      5. Optionally writes a status string to a Datto UDF.

    ESU coverage windows (Patch Tuesday based):
      Year 1: 2025-10-15 - 2026-10-13   f520e45e-7413-4a34-a497-d2765967d094
      Year 2: 2026-10-14 - 2027-10-12   1043add5-23b1-4afb-9a0f-64343c8f3f8d
      Year 3: 2027-10-13 - 2028-10-10   83d49986-add3-41d7-ba33-87c7bfb5c0fb
    After 2028-10-10 the ESU program has ended and every Windows 10 device is flagged.

.NOTES
    Datto RMM component variables (all optional):
      usrToastTitle         String   Toast heading
      usrToastBody          String   Toast message text (overrides the default message)
      usrSupportContact     String   Name used in the default message       (default: CMIT Solutions)
      usrAppDisplayName     String   Sender name shown on the toast         (default: IT Support)
      usrHeroImgURL         String   https:// URL for the "hero" branding image. PNG @ 364x180.
                                     (default: https://storage.centrastage.net/dattoHero.png)
      usrLearnMoreUrl       String   https:// URL; adds a "Learn more" button when set
      usrLeadDays           String   Start checking for the NEXT ESU year this many days early (default: 0)
      usrUDFNumber          String   1-30, writes status to that UDF (blank = disabled)
      usrRequirePriorYears  Boolean  Require all earlier ESU years too (default: true)
      usrDebug              Boolean  Evaluate and print everything, make NO changes, send NO toast
      usrForceToast         Boolean  Send the toast even when the device is compliant (testing)
      usrTestDate           String   yyyy-MM-dd, evaluate as if today were this date (testing).
                                     Also bypasses the Windows 10 check so the toast can be previewed
                                     on any Windows machine.

    Suggested test sequence:
      1. usrDebug=true                                  -> verify detection and the generated toast XML
      2. usrDebug=true,  usrTestDate=2027-11-01         -> verify Year 3 logic
      3. usrDebug=false, usrForceToast=true             -> see the real toast on a test machine
      4. usrDebug=false, usrForceToast=false            -> production

    Exit codes: 0 = ran successfully (compliant, non-compliant or not applicable), 1 = script error.
#>

$ErrorActionPreference = 'Stop'

function Get-Var {
    param([string]$Name, [string]$Default = '')
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}
function Get-Bool {
    param([string]$Name, [bool]$Default = $false)
    $v = Get-Var $Name
    if (-not $v) { return $Default }
    return $v -match '^(true|1|yes)$'
}

# ---------------------------------------------------------------- variables
$SupportContact    = Get-Var  'usrSupportContact' 'CMIT Solutions'
$AppDisplayName    = Get-Var  'usrAppDisplayName' 'IT Support'
$ToastTitle        = Get-Var  'usrToastTitle'     'Action needed: Windows 10 security updates'
$ToastBody         = Get-Var  'usrToastBody'      ("This computer runs Windows 10, which needs an Extended Security Update license to keep " +
                                                   "receiving security fixes. It does not have one for the current period. " +
                                                   "Please contact $SupportContact to discuss your options.")
$HeroImgURL        = Get-Var  'usrHeroImgURL'     'https://storage.centrastage.net/dattoHero.png'
$LearnMoreUrl      = Get-Var  'usrLearnMoreUrl'
$UDFNumber         = Get-Var  'usrUDFNumber'
$RequirePriorYears = Get-Bool 'usrRequirePriorYears' $true
$DebugMode         = Get-Bool 'usrDebug'
$ForceToast        = Get-Bool 'usrForceToast'
$TestDate          = Get-Var  'usrTestDate'
$LeadDays          = 0
[void][int]::TryParse((Get-Var 'usrLeadDays' '0'), [ref]$LeadDays)

$EsuYears = @(
    [pscustomobject]@{ Year = 1; Start = [datetime]'2025-10-15'; End = [datetime]'2026-10-13'; ActivationId = 'f520e45e-7413-4a34-a497-d2765967d094' }
    [pscustomobject]@{ Year = 2; Start = [datetime]'2026-10-14'; End = [datetime]'2027-10-12'; ActivationId = '1043add5-23b1-4afb-9a0f-64343c8f3f8d' }
    [pscustomobject]@{ Year = 3; Start = [datetime]'2027-10-13'; End = [datetime]'2028-10-10'; ActivationId = '83d49986-add3-41d7-ba33-87c7bfb5c0fb' }
)

$AppId    = 'CMIT.ITNotifications'
$WorkDir  = Join-Path $env:ProgramData 'CMIT\ESUToast'
$xmlPath  = Join-Path $WorkDir 'toast.xml'
$ps1Path  = Join-Path $WorkDir 'Show-Toast.ps1'
$vbsPath  = Join-Path $WorkDir 'Launch.vbs'
$heroPath = Join-Path $WorkDir 'hero.png'
$TaskName = 'CMIT-ESU-Toast'

function Write-Dbg([string]$Msg) { if ($DebugMode) { Write-Host "[DEBUG] $Msg" } }

function Set-DattoUDF {
    param([string]$Number, [string]$Value)
    if ($Number -notmatch '^([1-9]|[12][0-9]|30)$') { return }
    if ($Value.Length -gt 255) { $Value = $Value.Substring(0, 255) }
    if ($DebugMode) { Write-Dbg "Would set UDF $Number to: $Value"; return }
    $key = 'HKLM:\SOFTWARE\CentraStage'
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    New-ItemProperty -Path $key -Name "Custom$Number" -Value $Value -PropertyType String -Force | Out-Null
    Write-Host "UDF $Number set to: $Value"
}

function Get-EsuState {
    param([string]$ActivationId)
    $p = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ID='$ActivationId'" -ErrorAction SilentlyContinue
    if (-not $p -or -not $p.PartialProductKey) { return [pscustomobject]@{ State = 'NotInstalled'; Detail = 'no key installed' } }
    if ($p.LicenseStatus -ne 1) { return [pscustomobject]@{ State = 'NotActivated'; Detail = "key *****-$($p.PartialProductKey) installed, NOT activated (LicenseStatus=$($p.LicenseStatus))" } }
    return [pscustomobject]@{ State = 'Licensed'; Detail = "key *****-$($p.PartialProductKey) installed and activated" }
}

function Get-HeroImage {
    param([string]$Url, [string]$OutFile)
    if (-not $Url) { return $false }
    if ($Url -notmatch '^https://') { Write-Host 'usrHeroImgURL ignored - must start with https://'; return $false }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -TimeoutSec 30
        $bytes = [IO.File]::ReadAllBytes($OutFile)
        # PNG signature check: 89 50 4E 47
        if ($bytes.Length -lt 8 -or $bytes[0] -ne 0x89 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x4E -or $bytes[3] -ne 0x47) {
            Write-Host 'Hero image downloaded but is not a PNG - skipping image.'
            Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
            return $false
        }
        Write-Host "Hero image downloaded ($([math]::Round($bytes.Length/1KB,1)) KB)."
        return $true
    } catch {
        Write-Host "Hero image download failed ($($_.Exception.Message)) - toast will be sent without it."
        return $false
    }
}

try {
    if ($DebugMode) { Write-Host '===== DEBUG MODE: no changes will be made and no toast will be sent =====' }
    Write-Dbg "Variables: Title='$ToastTitle' | Hero='$HeroImgURL' | LearnMore='$LearnMoreUrl' | LeadDays=$LeadDays | UDF='$UDFNumber' | RequirePriorYears=$RequirePriorYears | ForceToast=$ForceToast | TestDate='$TestDate'"
    Write-Dbg "Body: $ToastBody"

    # ------------------------------------------------------------ date
    if ($TestDate) {
        $Today = [datetime]::ParseExact($TestDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).Date
        Write-Host "TEST DATE: evaluating as of $($Today.ToString('yyyy-MM-dd'))"
    } else {
        $Today = (Get-Date).Date
    }

    # ------------------------------------------------------------ OS applicability
    $os    = Get-CimInstance -ClassName Win32_OperatingSystem
    $build = [int]$os.BuildNumber
    Write-Host "OS: $($os.Caption) (build $build)"

    $isWin10 = ($os.Caption -match 'Windows 10') -and ($build -lt 22000)
    $isLtsc  = $os.Caption -match 'LTSC|LTSB'
    if (-not $isWin10 -or $isLtsc) {
        $reason = if (-not $isWin10) { 'Not Windows 10' } else { 'Windows 10 LTSC/LTSB (separate lifecycle)' }
        if ($TestDate) {
            Write-Host "TEST DATE: $reason - OS check bypassed so the notification can be previewed."
        } else {
            Write-Host "$reason - nothing to do."
            exit 0
        }
    } elseif ($build -ne 19045) {
        Write-Host "WARNING: build $build is not Windows 10 22H2 (19045). ESU requires 22H2 - device must be upgraded before ESU can apply."
    }

    # ------------------------------------------------------------ which ESU year is required?
    $required = $EsuYears | Where-Object { $Today -ge $_.Start -and $Today -le $_.End } | Select-Object -First 1
    $upcoming = $EsuYears | Where-Object { $_.Start -gt $Today -and ($_.Start - $Today).TotalDays -le $LeadDays } | Select-Object -First 1
    if ($upcoming) {
        Write-Host "usrLeadDays=$LeadDays : ESU Year $($upcoming.Year) starts $($upcoming.Start.ToString('yyyy-MM-dd')) - checking for it now."
        $required = $upcoming
    }
    $programEnded = $Today -gt $EsuYears[-1].End

    # ------------------------------------------------------------ report all years
    $states = @{}
    foreach ($y in $EsuYears) {
        $states[$y.Year] = Get-EsuState -ActivationId $y.ActivationId
        Write-Host ("ESU Year {0}: {1}" -f $y.Year, $states[$y.Year].Detail)
    }

    # ------------------------------------------------------------ evaluate
    if ($programEnded) {
        $compliant = $false
        $summary   = 'ESU program ended ' + $EsuYears[-1].End.ToString('yyyy-MM-dd') + ' - no coverage possible'
    }
    elseif (-not $required) {
        Write-Host "Date is before the ESU Year 1 window ($($EsuYears[0].Start.ToString('yyyy-MM-dd'))). Nothing to check."
        Set-DattoUDF -Number $UDFNumber -Value "Win10 ESU: not yet required ($($Today.ToString('yyyy-MM-dd')))"
        exit 0
    }
    else {
        $toCheck = if ($RequirePriorYears) { $EsuYears | Where-Object { $_.Year -le $required.Year } } else { @($required) }
        Write-Host ("Required for this date: Year {0}{1}" -f $required.Year, $(if ($RequirePriorYears -and $required.Year -gt 1) { ' (plus all prior years)' } else { '' }))
        $missing = @($toCheck | Where-Object { $states[$_.Year].State -ne 'Licensed' } |
                     ForEach-Object { "Y$($_.Year) $(if ($states[$_.Year].State -eq 'NotActivated') {'not activated'} else {'not installed'})" })
        $compliant = $missing.Count -eq 0
        $summary   = if ($compliant) { "Year $($required.Year) compliant" } else { "Year $($required.Year) NON-compliant - " + ($missing -join ', ') }
    }

    Write-Host "Result: $summary"
    Set-DattoUDF -Number $UDFNumber -Value ("Win10 ESU: $summary ($($Today.ToString('yyyy-MM-dd')))")

    $sendToast = (-not $compliant) -or $ForceToast
    if (-not $sendToast) {
        Write-Host 'Compliant - no notification needed.'
        exit 0
    }
    if ($compliant) { Write-Host 'usrForceToast is set - notification will be sent anyway.' }

    # ------------------------------------------------------------ anyone logged on?
    $explorer = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue)
    Write-Host "Interactive sessions found: $($explorer.Count)"

    # ------------------------------------------------------------ build toast XML
    function New-ToastXml([bool]$IncludeHero) {
        $t = [Security.SecurityElement]::Escape($ToastTitle)
        $m = [Security.SecurityElement]::Escape($ToastBody)
        $hero = ''
        if ($IncludeHero) {
            $src  = [Security.SecurityElement]::Escape('file:///' + ($heroPath -replace '\\', '/'))
            $hero = "<image placement=`"hero`" src=`"$src`"/>"
        }
        $learnMore = ''
        if ($LearnMoreUrl) {
            if ($LearnMoreUrl -match '^https://') {
                $u = [Security.SecurityElement]::Escape($LearnMoreUrl)
                $learnMore = "<action activationType=`"protocol`" content=`"Learn more`" arguments=`"$u`"/>"
            } else {
                Write-Host 'usrLearnMoreUrl ignored - must start with https://'
            }
        }
        return @"
<toast scenario="reminder">
  <visual>
    <binding template="ToastGeneric">
      $hero
      <text>$t</text>
      <text>$m</text>
    </binding>
  </visual>
  <actions>
    $learnMore
    <action activationType="system" arguments="dismiss" content="Dismiss"/>
  </actions>
</toast>
"@
    }

    if ($DebugMode) {
        # Validate the hero image URL without touching ProgramData
        $tmpHero = Join-Path $env:TEMP ("esu-hero-test-{0}.png" -f [guid]::NewGuid())
        $heroOk  = Get-HeroImage -Url $HeroImgURL -OutFile $tmpHero
        Remove-Item $tmpHero -Force -ErrorAction SilentlyContinue
        Write-Dbg 'Toast XML that would be sent:'
        (New-ToastXml -IncludeHero $heroOk) -split "`r?`n" | ForEach-Object { Write-Dbg "  $_" }
        if ($explorer.Count -eq 0) { Write-Dbg 'No user logged on - a live run would skip the toast.' }
        Write-Host '===== DEBUG MODE complete: toast NOT sent ====='
        exit 0
    }

    if ($explorer.Count -eq 0) {
        Write-Host 'No interactive user logged on - toast skipped.'
        exit 0
    }

    # ------------------------------------------------------------ stage payload
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    # Lock folder down: SYSTEM + Administrators full, Users read/execute only
    & icacls.exe $WorkDir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-32-545:(OI)(CI)RX' | Out-Null

    $appKey = "HKLM:\SOFTWARE\Classes\AppUserModelId\$AppId"
    New-Item -Path $appKey -Force | Out-Null
    New-ItemProperty -Path $appKey -Name 'DisplayName'    -Value $AppDisplayName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $appKey -Name 'ShowInSettings' -Value 0 -PropertyType DWord -Force | Out-Null

    Remove-Item $heroPath -Force -ErrorAction SilentlyContinue
    $heroOk = Get-HeroImage -Url $HeroImgURL -OutFile $heroPath
    Set-Content -Path $xmlPath -Value (New-ToastXml -IncludeHero $heroOk) -Encoding UTF8 -Force

    $userScript = @'
$ErrorActionPreference = 'Stop'
try {
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml([IO.File]::ReadAllText('__XMLPATH__'))
    $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('__APPID__').Show($toast)
} catch { }
'@
    $userScript = $userScript.Replace('__XMLPATH__', $xmlPath).Replace('__APPID__', $AppId)
    Set-Content -Path $ps1Path -Value $userScript -Encoding UTF8 -Force

    # VBS wrapper so no PowerShell window flashes on the user's screen
    $vbs = 'Set sh = CreateObject("WScript.Shell")' + "`r`n" +
           ('sh.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File ""{0}""", 0, False' -f $ps1Path)
    Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII -Force

    # ------------------------------------------------------------ deliver via temporary task
    $usersName = (New-Object Security.Principal.SecurityIdentifier 'S-1-5-32-545').Translate([Security.Principal.NTAccount]).Value
    $action    = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\wscript.exe" -Argument "//B //NoLogo `"$vbsPath`""
    $principal = New-ScheduledTaskPrincipal -GroupId $usersName -RunLevel Limited
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances Parallel

    Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings -Force | Out-Null
    Start-ScheduledTask -TaskName $TaskName
    Start-Sleep -Seconds 15
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false

    Write-Host "Toast sent to $($explorer.Count) interactive session(s)."
    exit 0
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    if (-not $DebugMode) {
        try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    }
    exit 1
}
