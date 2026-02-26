# Create-LibreOffice-Shortcuts-Adult.ps1

$loDir    = "C:\Program Files\LibreOffice\program"
$desktop  = "C:\Users\adult\Desktop"

$apps = @(
    @{ Exe="swriter.exe";  Name="WORD"       },
    @{ Exe="scalc.exe";    Name="EXCEL"      },
    @{ Exe="simpress.exe"; Name="POWERPOINT" }
)

# Basic checks
if (-not (Test-Path $desktop)) {
    Write-Error "Adult desktop not found: $desktop"
    exit 1
}
if (-not (Test-Path $loDir)) {
    Write-Error "LibreOffice folder not found: $loDir"
    exit 2
}

$wsh = New-Object -ComObject WScript.Shell

foreach ($a in $apps) {
    $exePath = Join-Path $loDir $a.Exe

    if (-not (Test-Path $exePath)) {
        Write-Warning "Missing: $exePath (skipping)"
        continue
    }

    $lnkPath = Join-Path $desktop ($a.Name + ".lnk")

    $sc = $wsh.CreateShortcut($lnkPath)
    $sc.TargetPath       = $exePath
    $sc.WorkingDirectory = $loDir
    $sc.IconLocation     = "$exePath,0"
    $sc.Save()

    Write-Host "Created: $lnkPath -> $exePath"
}

exit 0
