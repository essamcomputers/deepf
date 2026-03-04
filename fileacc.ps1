#requires -RunAsAdministrator
<#
Deep Freeze deployable script:
- Targets user profile named "Adult"
- Sets Microsoft Office file associations to LibreOffice for that user
- Works even if Adult is not logged in by loading NTUSER.DAT
Log: C:\Temp\DeepFreeze-LibreOffice-FileAssoc.log
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -------- Logging --------
$LogDir  = "C:\Temp"
$LogFile = Join-Path $LogDir "DeepFreeze-LibreOffice-FileAssoc.log"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$ts  $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

Write-Log "==== START: LibreOffice file association script ===="

# -------- Helper: Admin check --------
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (-not (Test-IsAdmin)) {
    Write-Log "ERROR: Script is not running elevated."
    throw "Script must run as Administrator/SYSTEM."
}

# -------- Helper: Locate LibreOffice command --------
function Find-LibreOfficeCommand {
    $candidates = @(
        "C:\Program Files\LibreOffice\program\soffice.exe",
        "C:\Program Files\LibreOffice\program\soffice.com",
        "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "C:\Program Files (x86)\LibreOffice\program\soffice.com"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

$SofficePath = Find-LibreOfficeCommand
if (-not $SofficePath) {
    Write-Log "WARNING: LibreOffice not found in default locations. Associations may still work if LibreOffice is registered, but open-commands could fail."
} else {
    Write-Log "LibreOffice detected at: $SofficePath"
}

# -------- Extensions to target --------
# Add/remove as you like.
$TargetExtensions = @(
    ".doc",".docx",".dot",".dotx",".rtf",
    ".xls",".xlsx",".xlt",".xltx",".csv",
    ".ppt",".pptx",".pot",".potx"
)

# -------- Known LibreOffice ProgID fallbacks --------
# These are commonly present when LibreOffice is installed.
$FallbackProgIdByExt = @{
    ".doc"  = "LibreOffice.WriterDocument.1"
    ".docx" = "LibreOffice.WriterDocument.1"
    ".dot"  = "LibreOffice.WriterDocument.1"
    ".dotx" = "LibreOffice.WriterDocument.1"
    ".rtf"  = "LibreOffice.WriterDocument.1"

    ".xls"  = "LibreOffice.CalcDocument.1"
    ".xlsx" = "LibreOffice.CalcDocument.1"
    ".xlt"  = "LibreOffice.CalcDocument.1"
    ".xltx" = "LibreOffice.CalcDocument.1"
    ".csv"  = "LibreOffice.CalcDocument.1"

    ".ppt"  = "LibreOffice.ImpressDocument.1"
    ".pptx" = "LibreOffice.ImpressDocument.1"
    ".pot"  = "LibreOffice.ImpressDocument.1"
    ".potx" = "LibreOffice.ImpressDocument.1"
}

# -------- Helper: choose best LibreOffice ProgID for an extension --------
function Get-LibreProgIdForExtension {
    param([string]$Ext)

    # Try to discover progids already registered for this ext, prefer ones that open with LibreOffice
    $progIds = New-Object System.Collections.Generic.List[string]

    try {
        $hkcrExt = "Registry::HKEY_CLASSES_ROOT\$Ext"
        if (Test-Path $hkcrExt) {
            # Current default ProgID
            $def = (Get-ItemProperty -Path $hkcrExt -ErrorAction SilentlyContinue)."(default)"
            if ($def) { $progIds.Add($def) | Out-Null }

            # OpenWithProgids list
            $owp = Join-Path $hkcrExt "OpenWithProgids"
            if (Test-Path $owp) {
                $names = (Get-Item $owp).GetValueNames()
                foreach ($n in $names) { if ($n) { $progIds.Add($n) | Out-Null } }
            }
        }
    } catch {}

    # Also try a few common LibreOffice ProgIDs directly (in case they are not listed under OpenWithProgids)
    $common = @("LibreOffice.WriterDocument.1","LibreOffice.CalcDocument.1","LibreOffice.ImpressDocument.1")
    foreach ($c in $common) { $progIds.Add($c) | Out-Null }

    # De-dup
    $progIds = $progIds | Select-Object -Unique

    # Prefer progids whose open command points to soffice
    foreach ($pid in $progIds) {
        try {
            $cmdKey = "Registry::HKEY_CLASSES_ROOT\$pid\shell\open\command"
            if (Test-Path $cmdKey) {
                $cmd = (Get-ItemProperty -Path $cmdKey -ErrorAction SilentlyContinue)."(default)"
                if ($cmd -and ($cmd -match "soffice" -or $cmd -match "LibreOffice")) {
                    return $pid
                }
            }
        } catch {}
    }

    # Fallback mapping
    if ($FallbackProgIdByExt.ContainsKey($Ext)) { return $FallbackProgIdByExt[$Ext] }

    return $null
}

# -------- Find Adult profile + SID + NTUSER.DAT --------
$AdultProfile = Get-CimInstance Win32_UserProfile |
    Where-Object { $_.LocalPath -and ($_.LocalPath -match "\\Adult$") -and (-not $_.Special) } |
    Select-Object -First 1

if (-not $AdultProfile) {
    Write-Log "ERROR: Could not find a local profile path ending in \Adult"
    throw "Adult profile not found."
}

$AdultSid  = $AdultProfile.SID
$AdultPath = $AdultProfile.LocalPath
$NtUserDat = Join-Path $AdultPath "NTUSER.DAT"

Write-Log "Adult profile: $AdultPath"
Write-Log "Adult SID:     $AdultSid"

if (-not (Test-Path $NtUserDat)) {
    Write-Log "ERROR: NTUSER.DAT not found at $NtUserDat"
    throw "NTUSER.DAT not found for Adult."
}

# -------- Load hive (if not already loaded) --------
$HiveName = "TempAdultHive"
$HiveRoot = "Registry::HKEY_USERS\$HiveName"

function IsHiveLoaded($name) {
    return Test-Path ("Registry::HKEY_USERS\$name")
}

$loadedByScript = $false
if (-not (IsHiveLoaded $HiveName)) {
    Write-Log "Loading Adult hive into HKU\$HiveName ..."
    & reg.exe load "HKU\$HiveName" "$NtUserDat" | Out-Null
    $loadedByScript = $true
} else {
    Write-Log "Hive HKU\$HiveName already loaded."
}

try {
    # -------- Apply associations --------
    foreach ($ext in $TargetExtensions) {
        $progId = Get-LibreProgIdForExtension -Ext $ext

        if (-not $progId) {
            Write-Log "SKIP: $ext (could not determine LibreOffice ProgID)"
            continue
        }

        Write-Log "Setting $ext => $progId"

        # 1) Machine-level default (helps if UserChoice removed)
        $hklmExtKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$ext"
        if (-not (Test-Path $hklmExtKey)) { New-Item -Path $hklmExtKey -Force | Out-Null }
        New-ItemProperty -Path $hklmExtKey -Name "(default)" -Value $progId -PropertyType String -Force | Out-Null

        # 2) Per-user classes for Adult
        $adultClassesExtKey = "$HiveRoot\Software\Classes\$ext"
        if (-not (Test-Path $adultClassesExtKey)) { New-Item -Path $adultClassesExtKey -Force | Out-Null }
        New-ItemProperty -Path $adultClassesExtKey -Name "(default)" -Value $progId -PropertyType String -Force | Out-Null

        # 3) Encourage Windows OpenWithProgids
        $adultOpenWithProgids = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\OpenWithProgids"
        if (-not (Test-Path $adultOpenWithProgids)) { New-Item -Path $adultOpenWithProgids -Force | Out-Null }
        # value type doesn't matter much here; set DWORD 0 like Windows does
        New-ItemProperty -Path $adultOpenWithProgids -Name $progId -Value 0 -PropertyType DWord -Force | Out-Null

        # 4) Remove UserChoice override (this is the key step)
        $adultUserChoice = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\UserChoice"
        if (Test-Path $adultUserChoice) {
            Remove-Item -Path $adultUserChoice -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Removed Adult UserChoice for $ext"
        }

        # 5) Remove toast entries (optional but helps stop Windows nagging)
        $toastKey = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts"
        if (Test-Path $toastKey) {
            # Remove any toast values related to this ext or progid
            try {
                $valNames = (Get-Item $toastKey).GetValueNames()
                foreach ($vn in $valNames) {
                    if ($vn -match [regex]::Escape($ext) -or $vn -match [regex]::Escape($progId)) {
                        Remove-ItemProperty -Path $toastKey -Name $vn -ErrorAction SilentlyContinue
                    }
                }
            } catch {}
        }
    }

    Write-Log "Done applying registry changes."

    # -------- Optional: restart Explorer for next login effect --------
    # We won't kill explorer here because Adult may not be logged in.
    # Changes will apply on next sign-in, or after default app cache refresh.

} finally {
    # -------- Unload hive --------
    if ($loadedByScript) {
        Write-Log "Unloading Adult hive HKU\$HiveName ..."
        & reg.exe unload "HKU\$HiveName" | Out-Null
    }
    Write-Log "==== END: LibreOffice file association script ===="
}
