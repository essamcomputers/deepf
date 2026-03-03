# =========================================================
# Create ODT folder, download Office Deployment Tool,
# extract it, and create an uninstall XML file
# =========================================================

# Define folder path
$BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ODTFolder = Join-Path $BasePath "ODT"

# Create the folder if it doesn't exist
if (!(Test-Path $ODTFolder)) {
    New-Item -Path $ODTFolder -ItemType Directory | Out-Null
}

# Paths inside the folder
$ODTExe = Join-Path $ODTFolder "setup.exe"
$ConfigXML = Join-Path $ODTFolder "RemoveOffice.xml"

# Download Office Deployment Tool EXE
$ODTDownloadUrl = "https://download.microsoft.com/download/2/e/8/2e8a6c9c-a6e2-4c7b-a3e3-6d5cf202e6e2/officedeploymenttool_16227-20258.exe"
$TempFile = Join-Path $env:TEMP "odt_download.exe"

Write-Host "Downloading Office Deployment Tool..."
Invoke-WebRequest -Uri $ODTDownloadUrl -OutFile $TempFile

# Extract ODT into the folder
Write-Host "Extracting ODT files..."
Start-Process -FilePath $TempFile -ArgumentList "/quiet /extract:`"$ODTFolder`"" -Wait

# Create uninstall XML config
$XMLContent = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@

Set-Content -Path $ConfigXML -Value $XMLContent -Encoding UTF8

Write-Host "ODT folder prepared successfully at: $ODTFolder"
Write-Host "Setup.exe and RemoveOffice.xml are ready."
