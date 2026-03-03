<#
Uses EXACT uninstall string provided by user.
Runs silently under SYSTEM account (DeepFreeze compatible).
#>

$ErrorActionPreference = "SilentlyContinue"

# --- Logging ---
$LogDir = "C:\ProgramData\OfficeRemoval"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Log = "$LogDir\Removal_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)

Function Log {
    param([string]$msg)
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg" | Out-File -FilePath $Log -Append -Encoding UTF8
}

Log "=== Starting Office Removal (using user‑provided uninstall string) ==="

# --- EXACT STRING FOUND IN REGISTRY ---
$Exe  = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"

$Args = @(
    "scenario=install"
    "scenariosubtype=ARP"
    "sourcetype=None"
    "productstoremove=Standard2019Volume.16_en-us_x-none"
    "culture=en-us"
    "version.16=16.0"
    "DisplayLevel=False"
    "ForceAppShutdown=True"
    "UpdatePromptUser=False"
)

$ArgString = $Args -join " "

Log "Uninstall command:"
Log "`"$Exe`" $ArgString"

# --- Stop Office Apps ---
$OfficeApps = "WINWORD","EXCEL","OUTLOOK","POWERPNT","MSACCESS","MSPUB","VISIO","WINPROJ","ONENOTE","LYNC"

foreach ($app in $OfficeApps) {
    Get-Process $app -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
Log "Stopped running Office apps."

# --- Run uninstall ---
try {
    $p = Start-Process -FilePath $Exe -ArgumentList $ArgString -PassThru -Wait -WindowStyle Hidden
    Log "Office uninstall completed. ExitCode = $($p.ExitCode)"
}
catch {
    Log "ERROR: $($_.Exception.Message)"
}

Log "=== Script finished ==="
exit 0
