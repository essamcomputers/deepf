<#
.SYNOPSIS
  Create LibreOffice shortcuts on the Desktop of a specific user profile (default: "Adult").

.NOTES
  - Tested for standard LibreOffice MSI installs under Program Files / Program Files (x86).
  - For Store/MSIX installs, individual EXEs may not exist; script falls back to soffice.exe with arguments.
  - Run as Administrator or SYSTEM (e.g., via RMM/Deep Freeze).
#>

param(
    [string]$ProfileName = 'Adult'
)

# --- Helper: Write log line ---
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp][$Level] $Message"
}

# --- Resolve Desktop path for the target profile ---
$desktopPath = Join-Path -Path "$env:SystemDrive\Users\$ProfileName" -ChildPath 'Desktop'
if (-not (Test-Path -LiteralPath $desktopPath)) {
    Write-Log "Desktop path not found: $desktopPath" 'ERROR'
    throw "User profile '$ProfileName' not found or Desktop directory missing."
}

# --- Potential LibreOffice program folders (x64 then x86) ---
$loCandidates = @(
    "$env:ProgramFiles\LibreOffice\program",
    "${env:ProgramFiles(x86)}\LibreOffice\program"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $loCandidates) {
    Write-Log "LibreOffice 'program' folder not found in Program Files." 'WARN'
    # You can optionally exit here if you want to enforce presence
    # throw "LibreOffice not found."
}

# Pick the first valid path (prefer x64)
$loProg = $loCandidates | Select-Object -First 1

# Fallback to a likely default in case of odd environments
if (-not $loProg) {
    $loProg = "$env:ProgramFiles\LibreOffice\program"
}

# --- Define applications: try direct EXE first, then fallback to soffice.exe + argument ---
$apps = @(
    @{
        Name = 'LibreOffice Writer'
        Exe  = 'swriter.exe'
        FallbackArg = '-writer'
    },
    @{
        Name = 'LibreOffice Calc'
        Exe  = 'scalc.exe'
        FallbackArg = '-calc'
    },
    @{
        Name = 'LibreOffice Impress'
        Exe  = 'simpress.exe'
        FallbackArg = '-impress'
    },
    @{
        Name = 'LibreOffice Draw'
        Exe  = 'sdraw.exe'
        FallbackArg = '-draw'
    },
    @{
        Name = 'LibreOffice Base'
        Exe  = 'sbase.exe'
        FallbackArg = '-base'
    },
    @{
        Name = 'LibreOffice Math'
        Exe  = 'smath.exe'
        FallbackArg = '-math'
    }
)

# --- CreateShortcut function using WScript.Shell COM ---
function New-DesktopShortcut {
    param(
        [Parameter(Mandatory)]
        [string]$ShortcutPath,

        [Parameter(Mandatory)]
        [string]$TargetPath,

        [string]$Arguments,
        [string]$IconLocation,
        [string]$Description = ''
    )

    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($ShortcutPath)
    $sc.TargetPath   = $TargetPath
    if ($Arguments)  { $sc.Arguments = $Arguments }
    if ($IconLocation){ $sc.IconLocation = $IconLocation }
    if ($Description){ $sc.Description = $Description }
    # Start in the LibreOffice program directory if possible
    try {
        $sc.WorkingDirectory = Split-Path -Path $TargetPath -Parent
    } catch { }

    $sc.Save()
}

# Ensure desktop path exists
if (-not (Test-Path -LiteralPath $desktopPath)) {
    New-Item -Path $desktopPath -ItemType Directory -Force | Out-Null
}

# Determine soffice path for fallback
$sofficePath = Join-Path -Path $loProg -ChildPath 'soffice.exe'
$hasSoffice  = Test-Path -LiteralPath $sofficePath

foreach ($app in $apps) {
    $name = $app.Name
    $exe  = $app.Exe
    $fallbackArg = $app.FallbackArg

    $directExePath = Join-Path -Path $loProg -ChildPath $exe
    $shortcutFile  = Join-Path -Path $desktopPath -ChildPath ("$name.lnk")

    $targetPath = $null
    $arguments  = $null
    $iconPath   = $null

    if (Test-Path -LiteralPath $directExePath) {
        # Preferred: use the app-specific launcher if present
        $targetPath = $directExePath
        $iconPath   = $directExePath
        Write-Log "Using direct EXE for $name: $directExePath"
    }
    elseif ($hasSoffice) {
        # Fallback: soffice.exe with a mode argument (e.g. -writer)
        $targetPath = $sofficePath
        $arguments  = $fallbackArg
        $iconPath   = $sofficePath
        Write-Log "Falling back to soffice.exe for $name with arg $fallbackArg"
    }
    else {
        Write-Log "Neither $exe nor soffice.exe found. Skipping $name." 'WARN'
        continue
    }

    try {
        # Overwrite existing shortcut
        if (Test-Path -LiteralPath $shortcutFile) {
            Remove-Item -LiteralPath $shortcutFile -Force -ErrorAction SilentlyContinue
        }

        New-DesktopShortcut -ShortcutPath $shortcutFile `
                            -TargetPath $targetPath `
                            -Arguments $arguments `
                            -IconLocation $iconPath `
                            -Description $name

        Write-Log "Created shortcut: $shortcutFile"
    }
    catch {
        Write-Log "Failed to create shortcut for $name: $($_.Exception.Message)" 'ERROR'
    }
}

Write-Log "All done. Shortcuts processed for profile '$ProfileName'."
