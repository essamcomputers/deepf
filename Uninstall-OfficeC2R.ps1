<#
.SYNOPSIS
  Silent uninstall of Microsoft Office Click‑to‑Run.
  Works under SYSTEM account (DeepFreeze), no user session required, no UI.

.DESCRIPTION
  - Reads the real ClickToRun uninstall string from registry
  - Appends silent flags
  - Executes uninstall silently
  - Logs results to C:\ProgramData\OfficeRemoval\
#>

$ErrorActionPreference = "SilentlyContinue"

# ---------------------------
# Create Log Directory
# ---------------------------
$LogPath = "C:\ProgramData\OfficeRemoval"
if (!(Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }
$LogFile = "$LogPath\Removal_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)

Function Write-Log {
    param([string]$Message)
    $stamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    "$stamp  $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

Write-Log "===== Starting Office C2R Removal Script ====="

# ---------------------------
# Registry paths to check
# ---------------------------
$RegPaths = @(
  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$C2R = $null

foreach ($Path in $RegPaths) {
    $apps = Get-ItemProperty -Path "$Path\*" -ErrorAction SilentlyContinue | `
        Where-Object {
            $_.UninstallString -and
            $_.UninstallString -match "OfficeClickToRun\.exe"
        }

    if ($apps) {
        $C2R = $apps[0]
        break
    }
}

if (-not $C2R) {
    Write-Log "No Click‑to‑Run Office uninstall entry found."
    exit 0
}

Write-Log "Found Office uninstall entry: $($C2R.DisplayName)"
Write-Log "Original uninstall string: $($C2R.UninstallString)"

# ---------------------------
# Extract EXE + Arguments
# ---------------------------
$u = $C2R.UninstallString.Trim()

if ($u.StartsWith('"')) {
    $exe = $u.Split('"')[1]
    $args = $u.Substring($exe.Length + 2).Trim()
} else {
    $parts = $u.Split(" ", 2)
    $exe = $parts[0]
    $args = $parts[1]
}

# ---------------------------
# Add silent flags
# ---------------------------
if ($args -notmatch "DisplayLevel")     { $args += " DisplayLevel=False" }
if ($args -notmatch "ForceAppShutdown") { $args += " ForceAppShutdown=True" }
if ($args -notmatch "UpdatePromptUser") { $args += " UpdatePromptUser=False" }

Write-Log "Modified uninstall args: $args"

# ---------------------------
# Kill Office apps (just in case)
# ---------------------------
$procs = "WINWORD","EXCEL","POWERPNT","OUTLOOK","MSACCESS","MSPUB","ONENOTE","VISIO","WINPROJ","LYNC"
foreach ($p in $procs) { 
    Get-Process $p -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
Write-Log "Stopped running Office processes."

# ---------------------------
# Execute uninstall
# ---------------------------
Write-Log "Starting Office uninstall…"

try {
    $process = Start-Process -FilePath $exe -ArgumentList $args -PassThru -Wait -WindowStyle Hidden
    Write-Log "Uninstall completed with ExitCode: $($process.ExitCode)"
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}

Write-Log "===== Office removal complete ====="
exit 0
``
