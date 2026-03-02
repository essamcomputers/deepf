# DF_RemoveOffice_SetLibreDefaults.ps1
# Uninstall Office (Click-to-Run + MSI best-effort) and set LibreOffice defaults for Adult user.
# Log: C:\Windows\Temp\DF_RemoveOffice_SetLibreDefaults.log

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

function Stop-OfficeProcesses {
    $procs = @("winword","excel","powerpnt","outlook","onenote","msaccess","lync","teams","groove","visio","project")
    foreach ($name in $procs) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Log "Stopping process: $($_.Name) (PID $($_.Id))"
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    }
}

function Get-C2R-ProductReleaseIds {
    # Common key: HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -> ProductReleaseIds
    $key = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    $ids = @()

    if (Test-Path $key) {
        try {
            $v = (Get-ItemProperty $key -ErrorAction Stop).ProductReleaseIds
            if ($v) {
                $ids += ($v -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
        } catch {}
    }

    # Also check WOW6432Node (occasionally relevant)
    $key2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $key2) {
        try {
            $v2 = (Get-ItemProperty $key2 -ErrorAction Stop).ProductReleaseIds
            if ($v2) {
                $ids += ($v2 -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
        } catch {}
    }

    $ids = $ids | Select-Object -Unique
    return $ids
}

function Uninstall-Office-C2R-BestEffort {
    $c2rExe = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    if (-not (Test-Path $c2rExe)) {
        Log "OfficeClickToRun.exe not found. Skipping C2R scenario uninstall."
        return
    }

    $products = Get-C2R-ProductReleaseIds
    if (-not $products -or $products.Count -eq 0) {
        Log "No ProductReleaseIds found. Attempting ARP-based uninstall strings later."
        return
    }

    Log "C2R ProductReleaseIds detected: $($products -join ', ')"

    # Build uninstall calls: OfficeClickToRun.exe scenario=install scenariosubtype=ARP ...
    # NOTE: culture/version values are not always required, but commonly used.
    foreach ($prod in $products) {
        try {
            $args = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$prod " +
                    "culture=en-us version.16=16.0 DisplayLevel=False ForceAppShutdown=True"
            Run-Exe -FilePath $c2rExe -Arguments $args -TimeoutSec 7200 | Out-Null
        } catch {
            Log "C2R uninstall attempt failed for product '$prod': $($_.Exception.Message)"
        }
    }
}

function Uninstall-Office-FromUninstallStrings {
    # Avoid Win32_Product.
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

    Log "Found $($targets.Count) uninstall entries (ARP)."

    foreach ($t in $targets) {
        Log "ARP Target: $($t.DisplayName)"
        $cmd = $null
        if ($t.QuietUninstallString) { $cmd = $t.QuietUninstallString }
        elseif ($t.UninstallString)  { $cmd = $t.UninstallString }

        if (-not $cmd) {
            Log "No uninstall string for $($t.DisplayName). Skipping."
            continue
        }

        # If it's msiexec, enforce silent flags
        if ($cmd -match "msiexec(\.exe)?") {
            $cmd2 = $cmd
            $cmd2 = $cmd2 -replace "\s/ I\s", " /X " -replace "\s/I\s", " /X "
            if ($cmd2 -notmatch "/qn") { $cmd2 += " /qn" }
            if ($cmd2 -notmatch "/norestart") { $cmd2 += " /norestart" }
            Log "Executing MSI uninstall: $cmd2"
            Run-Exe -FilePath "cmd.exe" -Arguments "/c $cmd2" -TimeoutSec 7200 | Out-Null
        }
        else {
            # Best effort non-msi: try quiet
            $silent = $cmd
            if ($silent -notmatch "(?i)/quiet|/qn|/s|/silent") { $silent += " /quiet" }
            Log "Executing non-MSI uninstall (best-effort): $silent"
            Run-Exe -FilePath "cmd.exe" -Arguments "/c $silent" -TimeoutSec 7200 | Out-Null
        }
    }
}

function Import-DefaultAppAssociations-LibreOffice {
    $dism = Join-Path $env:SystemRoot "System32\dism.exe"
    if (-not (Test-Path $dism)) { throw "DISM not found." }

    $xmlPath = Join-Path $env:TEMP "DefaultAppAssociations-LibreOffice.xml"

    # Common LibreOffice ProgIDs (usually correct when LO installed normally)
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
        Log "Hive not found: $ntuser (skipping)"
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
Log "==== START Remove Office + Set LibreOffice Defaults ===="

try {
    Ensure-LibreOfficePresent
    Stop-OfficeProcesses

    # 1) Try Click-to-Run scenario uninstall using built-in OfficeClickToRun.exe
    Uninstall-Office-C2R-BestEffort

    # 2) MSI + other ARP uninstall strings (also catches Visio/Project)
    Uninstall-Office-FromUninstallStrings

    # 3) Set machine defaults to LibreOffice (new users + users without overrides)
    Import-DefaultAppAssociations-LibreOffice

    # 4) Force Adult user to inherit (clear UserChoice overrides)
    # If your Adult profile folder is different, adjust here.
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
