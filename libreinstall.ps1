#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

Write-Host "======================================="
Write-Host " LibreOffice Latest Version Installer"
Write-Host "======================================="
Write-Host ""

# Official LibreOffice stable repository
$baseUrl = "https://download.documentfoundation.org/libreoffice/stable/"

Write-Host "Checking for latest LibreOffice version..."

# Get available versions
$page = Invoke-WebRequest -Uri $baseUrl -UseBasicParsing

$versions = $page.Links |
    ForEach-Object { $_.href } |
    Where-Object { $_ -match '^\d+\.\d+\.\d+/$' } |
    ForEach-Object { $_.TrimEnd('/') } |
    Sort-Object { [version]$_ } -Descending

if (-not $versions) {
    throw "Could not determine the latest LibreOffice version."
}

$latestVersion = $versions[0]

Write-Host "Latest version: $latestVersion"

# Detect CPU architecture
$cpuArchitecture = (Get-CimInstance Win32_Processor | Select-Object -First 1).Architecture

switch ($cpuArchitecture) {

    12 {
        # ARM64
        $folder = "aarch64"
        $fileArch = "aarch64"
    }

    9 {
        # x64
        $folder = "x86_64"
        $fileArch = "x86-64"
    }

    default {
        # 32-bit x86
        $folder = "x86"
        $fileArch = "x86"
    }
}

$fileName = "LibreOffice_${latestVersion}_Win_${fileArch}.msi"

$downloadUrl = "$baseUrl$latestVersion/win/$folder/$fileName"

$tempFolder = "$env:TEMP\LibreOfficeInstaller"
$installer = Join-Path $tempFolder $fileName

# Create temp directory
if (-not (Test-Path $tempFolder)) {
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
}

Write-Host ""
Write-Host "Downloading:"
Write-Host $downloadUrl
Write-Host ""

Invoke-WebRequest `
    -Uri $downloadUrl `
    -OutFile $installer `
    -UseBasicParsing

if (-not (Test-Path $installer)) {
    throw "LibreOffice installer failed to download."
}

Write-Host "Download complete."
Write-Host ""
Write-Host "Installing LibreOffice silently..."

$process = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList "/i `"$installer`" /qn /norestart" `
    -Wait `
    -PassThru

switch ($process.ExitCode) {

    0 {
        Write-Host ""
        Write-Host "LibreOffice installed successfully."
    }

    3010 {
        Write-Host ""
        Write-Host "LibreOffice installed successfully."
        Write-Host "A reboot is required."
    }

    default {
        throw "LibreOffice installation failed. MSI exit code: $($process.ExitCode)"
    }
}

# Remove downloaded installer
Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Installation complete."
