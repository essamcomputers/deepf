# DF_RemoveOffice_SetLibreDefaults.ps1
# One-file Deep Freeze Custom App script (no uploaded EXEs):
# - Downloads Office Deployment Tool (ODT) inside the script
# - Uses ODT to remove Click-to-Run Office (<Remove All="TRUE" />)
# - Removes MSI Office (best-effort)
# - Sets LibreOffice default file associations (machine defaults) + forces Adult to inherit
#
# Log: C:\Windows\Temp\DF_RemoveOffice_SetLibreDefaults.log
#
# NOTES:
# - Script does NOT restart the machine.
# - Some uninstallers may request a reboot; script does not initiate it.

$ErrorActionPreference = "Stop"
$LogPath = "C:\Windows\Temp\DF_RemoveOffice_SetLibreDefaults.log"

function Log($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $line
}

function Run-Exe([string]$FilePath, [string]$Arguments, [int]$TimeoutSec = 7200) {
    Log "Running: `"$FilePath`" $Arguments"
    $p = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden
    if (-not $p.WaitForExit($TimeoutSec * 1000)) {
        try { $p.Kill() } catch {}
        throw "Timeout: $FilePath $Arguments"
    }
    Log "ExitCode: $($p.ExitCode)"
    return $p.ExitCode
}

function Stop-OfficeProcesses {
    $procs = @("winword","excel","powerpnt","outlook","onenote","msaccess","visio","project","officeclicktorun","teams","groove")
    foreach ($name in $procs) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Log "Stopping process: $($_.Name) (PID $($_.Id))"
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    }
}

function Ensure-LibreOfficePresent {
    $paths = @(
        "C:\Program Files\LibreOffice\program\soffice.exe",
        "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            Log "LibreOffice found: $p"
            return $true
        }
    }
    throw "LibreOffice not found (soffice.exe). Install LibreOffice first."
}

function Ensure-Tls12 {
    try {
        # Some estates default to TLS 1.0/1.1; enforce TLS 1.2 for Microsoft downloads
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Log "TLS set to 1.2"
    } catch {
        Log "Could not set TLS 1.2: $($_.Exception.Message)"
    }
}

function Download-File($Url, $OutFile) {
    Log "Downloading: $Url -> $OutFile"
    Ensure-Tls12
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
    } catch {
        # Fallback for older PS / locked-down IE components
        Log "Invoke-WebRequest failed, trying BITS: $($_.Exception.Message)"
        Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
    }
}

function Get-ODT-DownloadUrl {
    # Microsoft Download Center entry for ODT (id=49117)
    # We scrape the page for the actual download link, because the direct file name can change.
    $page = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
    Log "Fetching ODT download page: $page"
    Ensure-Tls12
    $html = Invoke-WebRequest -Uri $page -UseBasicParsing -ErrorAction Stop

    # Look for a download link ending in .exe that contains "officedeploymenttool" or similar patterns.
    $links = @()
    foreach ($l in $html.Links) {
        if ($l.href -and $l.href -match "\.exe($|\?)") {
            $links += $l.href
        }
    }
    $links = $links | Select-Object -Unique

    # Prefer links that look like ODT
    $preferred = $links | Where-Object { $_ -match "(?i)officedeploymenttool|odt|office" } | Select-Object -First 1
    if ($preferred) {
        Log "Found ODT EXE link (preferred): $preferred"
        return $preferred
    }

    $anyExe = $links | Select-Object -First 1
    if ($anyExe) {
        Log "Found EXE link (fallback): $anyExe"
        return $anyExe
    }

    throw "Could not locate ODT download EXE link on the Microsoft download page."
}

function Get-ODT-SetupExePath {
    # Downloads ODT EXE, extracts it to a temp folder, and returns the path to setup.exe
    $workRoot = Join-Path $env:TEMP ("ODT_" + ([Guid]::NewGuid().ToString("N")))
    New-Item -Path $workRoot -ItemType Directory -Force | Out-Null

    $odtStub = Join-Path $workRoot "ODT.exe"

    $url = Get-ODT-DownloadUrl
    Download-File -Url $url -OutFile $odtStub

    # ODT EXE supports /extract:<path> (commonly)
    Log "Extracting ODT to: $workRoot"
    $exit = Run-Exe -FilePath $odtStub -Arguments "/extract:`"$workRoot`" /quiet" -TimeoutSec 600
    if ($exit -ne 0) {
        # Some ODT stubs use /extract without /quiet
        Log "First extract attempt returned $exit. Trying alternate extract args..."
        $exit2 = Run-Exe -FilePath $odtStub -Arguments "/extract:`"$workRoot`"" -TimeoutSec 600
        if ($exit2 -ne 0) {
            throw "ODT extraction failed (exit codes: $exit, $exit2)."
        }
    }

    $setup = Join-Path $workRoot "setup.exe"
    if (-not (Test-Path $setup)) {
        # Sometimes it extracts into a subfolder; search for setup.exe
        $setup = Get-ChildItem -Path $workRoot -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object { $_.FullName }
    }

    if (-not $setup -or -not (Test-Path $setup)) {
        throw "setup.exe not found after ODT extraction."
    }

    Log "ODT setup.exe located at: $setup"
    return $setup
}

function Uninstall-Office-C2R-WithODT {
    # Reliable Click-to-Run removal using ODT setup.exe + config XML
    $setup = Get-ODT-SetupExePath

    $xmlPath = Join-Path $env:TEMP "ODT_Uninstall_AllOffice.xml"
    $xml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
    $xml | Out-File -FilePath $xmlPath -Encoding UTF8 -Force
    Log "ODT uninstall config written: $xmlPath"

    Log "Starting ODT uninstall (Click-to-Run removal)..."
    $exit = Run-Exe -FilePath $setup -Arguments "/configure `"$xmlPath`"" -TimeoutSec 14400
    if ($exit -ne 0) {
        Log "ODT uninstall returned non-zero exit code: $exit"
    } else {
        Log "ODT uninstall completed with exit code 0."
    }
}

function Uninstall-Office-FromUninstallStrings {
    # MSI/perpetual Office removal via registry uninstall strings (best-effort)
    $roots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $targets = @()
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $p = Get-ItemProperty $_.PSPath -ErrorAction Stop
                $dn = $p.DisplayName
                if ([string]::IsNullOrWhiteSpace($dn)) { return }

                if ($dn -match "Microsoft 365" -or
                    $dn -match "Microsoft Office" -or
                    $dn -match "Office( )?(1[0-9]|20[0-9]{2})" -or
                    $dn -match "Visio" -or
                    $dn -match "Project") {

                    $targets += [PSCustomObject]@{
                        DisplayName = $dn
                        QuietUninstallString = $p.QuietUninstallString
                        UninstallString = $p.UninstallString
                    }
                }
            } catch {}
        }
    }

    $targets = $targets | Sort-Object DisplayName -Unique
    if ($targets.Count -eq 0) {
        Log "No Office/Visio/Project entries found in Uninstall registry keys."
        return
    }

    Log "Found $($targets.Count) MSI/ARP uninstall entries."
    foreach ($t in $targets) {
        Log "ARP Target: $($t.DisplayName)"

        $cmd = $null
        if ($t.QuietUninstallString) { $cmd = $t.QuietUninstallString }
        elseif ($t.UninstallString)  { $cmd = $t.UninstallString }

        if (-not $cmd) {
            Log "No uninstall string for $($t.DisplayName). Skipping."
            continue
        }

        if ($cmd -match "msiexec(\.exe)?") {
            $cmd2 = $cmd
            $cmd2 = $cmd2 -replace "\s/ I\s", " /X " -replace "\s/I\s", " /X "
            if ($cmd2 -notmatch "/qn") { $cmd2 += " /qn" }
            if ($cmd2 -notmatch "/norestart") { $cmd2 += " /norestart" }

            Log "Executing MSI uninstall: $cmd2"
            Run-Exe -FilePath "cmd.exe" -Arguments "/c $cmd2" -TimeoutSec 14400 | Out-Null
        }
        else {
            $silent = $cmd
            if ($silent -notmatch "(?i)/quiet|/qn|/s|/silent") { $silent += " /quiet" }
            Log "Executing uninstall (best-effort): $silent"
            Run-Exe -FilePath "cmd.exe" -Arguments "/c $silent" -TimeoutSec 14400 | Out-Null
        }
    }
}

function Import-DefaultAppAssociations-LibreOffice {
    $dism = Join-Path $env:SystemRoot "System32\dism.exe"
    if (-not (Test-Path $dism)) { throw "DISM not found." }

    $xmlPath = Join-Path $env:TEMP "DefaultAppAssociations-LibreOffice.xml"
    $xml = @"
<?xml version="1.0" encoding="UTF-8"?>
<DefaultAssociations>
  <Association Identifier=".doc"  ProgId="LibreOffice.WriterDocument.1" ApplicationName="LibreOffice Writer" />
  <Association Identifier=".docx" ProgId="LibreOffice.WriterDocument.1" ApplicationName="LibreOffice Writer" />
  <Association Identifier=".rtf"  ProgId="LibreOffice.WriterDocument.1" ApplicationName="LibreOffice Writer" />

  <Association Identifier=".xls"  ProgId="LibreOffice.CalcDocument.1" ApplicationName="LibreOffice Calc" />
  <Association Identifier=".xlsx" ProgId="LibreOffice.CalcDocument.1" ApplicationName="LibreOffice Calc" />
  <Association Identifier=".csv"  ProgId="LibreOffice.CalcDocument.1" ApplicationName="LibreOffice Calc" />

  <Association Identifier=".ppt"  ProgId="LibreOffice.ImpressDocument.1" ApplicationName="LibreOffice Impress" />
  <Association Identifier=".pptx" ProgId="LibreOffice.ImpressDocument.1" ApplicationName="LibreOffice Impress" />

  <Association Identifier=".odt"  ProgId="LibreOffice.WriterDocument.1" ApplicationName="LibreOffice Writer" />
  <Association Identifier=".ods"  ProgId="LibreOffice.CalcDocument.1" ApplicationName="LibreOffice Calc" />
  <Association Identifier=".odp"  ProgId="LibreOffice.ImpressDocument.1" ApplicationName="LibreOffice Impress" />
</DefaultAssociations>
"@
    $xml | Out-File -FilePath $xmlPath -Encoding UTF8 -Force
    Log "Created default associations XML: $xmlPath"

    Run-Exe -FilePath $dism -Arguments "/Online /Import-DefaultAppAssociations:`"$xmlPath`"" -TimeoutSec 1800 | Out-Null
    Log "Imported machine default app associations (LibreOffice)."
}

function Clear-UserChoiceOverridesForProfile([string]$ProfilePath) {
    $ntuser = Join-Path $ProfilePath "NTUSER.DAT"
    if (-not (Test-Path $ntuser)) {
        Log "Hive not found: $ntuser (skipping per-user override cleanup)"
        return
    }

    $mountName = "TempHive_$([Guid]::NewGuid().ToString('N'))"
    $mountPath = "HKU\$mountName"

    Log "Loading hive: $ntuser"
    Run-Exe -FilePath "reg.exe" -Arguments "load $mountPath `"$ntuser`"" -TimeoutSec 60 | Out-Null

    try {
        $exts = @(".doc",".docx",".rtf",".xls",".xlsx",".csv",".ppt",".pptx",".odt",".ods",".odp")
        foreach ($ext in $exts) {
            $userChoice = "$mountPath\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\UserChoice"
            if (Test-Path $userChoice) {
                Log "Deleting UserChoice override: $userChoice"
                Run-Exe -FilePath "reg.exe" -Arguments "delete `"$userChoice`" /f" -TimeoutSec 60 | Out-Null
            }

            $openWith = "$mountPath\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\OpenWithProgids"
            if (Test-Path $openWith) {
                Log "Deleting OpenWithProgids: $openWith"
                Run-Exe -FilePath "reg.exe" -Arguments "delete `"$openWith`" /f" -TimeoutSec 60 | Out-Null
            }
        }
        Log "Cleared per-user overrides under: $ProfilePath"
    }
    finally {
        Log "Unloading hive: $mountPath"
        Run-Exe -FilePath "reg.exe" -Arguments "unload $mountPath" -TimeoutSec 60 | Out-Null
    }
}

# ---------------- MAIN ----------------
Log "==== START Remove Office (ODT download) + Set LibreOffice Defaults ===="
Log "Log file: $LogPath"

try {
    Ensure-LibreOfficePresent
    Stop-OfficeProcesses

    # Reliable Click-to-Run removal without uploading EXEs: download ODT within script
    Uninstall-Office-C2R-WithODT  # ODT is Microsoft-recommended for C2R removal :contentReference[oaicite:1]{index=1}

    # MSI/perpetual Office cleanup (if any)
    Uninstall-Office-FromUninstallStrings

    # Default associations to LibreOffice
    Import-DefaultAppAssociations-LibreOffice

    # Force Adult to inherit defaults
    $adultProfile = "C:\Users\Adult"
    Clear-UserChoiceOverridesForProfile -ProfilePath $adultProfile

    Log "Completed successfully."
    Log "NOTE: Adult should log off/on (or reboot) for defaults to fully apply."
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    Log $_.Exception.ToString()
    exit 1
}

Log "==== END ===="
exit 0
