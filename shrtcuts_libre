$Desktop = "C:\Users\adult\Desktop"
$LibreOffice = "C:\Program Files\LibreOffice\program"

$WshShell = New-Object -ComObject WScript.Shell

$Shortcuts = @{
    "Word"       = "swriter.exe"
    "Excel"      = "scalc.exe"
    "PowerPoint" = "simpress.exe"
}

foreach ($Name in $Shortcuts.Keys) {

    $Target = Join-Path $LibreOffice $Shortcuts[$Name]
    $ShortcutPath = Join-Path $Desktop "$Name.lnk"

    if (Test-Path $Target) {
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $Target
        $Shortcut.WorkingDirectory = $LibreOffice
        $Shortcut.IconLocation = "$Target,0"
        $Shortcut.Save()

        Write-Host "Created: $ShortcutPath"
    }
    else {
        Write-Warning "LibreOffice executable not found: $Target"
    }
}
