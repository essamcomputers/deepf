# =========================================================
# Prepare ODT: Download latest ODT EXE, extract to .\ODT,
# and create a silent uninstall XML (RemoveOffice.xml).
# Works behind most proxies and PS versions.
# =========================================================

# 0) Hardening: ensure TLS 1.2+ and a desktop UA
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor `
                                                  [Net.SecurityProtocolType]::Tls13
} catch { }
$Headers = @{
    'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0 Safari/537.36'
}

# 1) Resolve script path (must be run as a .ps1 file)
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptPath) {
    Write-Host "ERROR: Run this script from a saved .ps1 file, not pasted into console." -ForegroundColor Red
    exit 1
}

# 2) Ensure ODT folder
$ODTFolder = "c:\ODT"
if (!(Test-Path $ODTFolder)) { New-Item -Path $ODTFolder -ItemType Directory | Out-Null }

# 3) Get Microsoft Download Center page HTML (static page that always points to current ODT)
$DownloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
Write-Host "Fetching Microsoft ODT page..." -ForegroundColor Cyan
try {
    $pageHtml = (Invoke-WebRequest -Uri $DownloadPage -Headers $Headers -UseBasicParsing -ErrorAction Stop).Content
} catch {
    Write-Host "ERROR: Cannot reach Microsoft Download Center: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 4) Extract the real EXE URL from the HTML (download.microsoft.com...officedeploymenttool_*.exe)
$regex = 'https://download\.microsoft\.com/[^\"]*?/officedeploymenttool_[0-9\-]+\.exe'
$exeUrl = [Text.RegularExpressions.Regex]::Match($pageHtml, $regex).Value

if (-not $exeUrl) {
    Write-Host "ERROR: Could not extract the actual ODT EXE link from the page HTML." -ForegroundColor Red
    Write-Host "Tip: Your network may rewrite the page. Try running from a different network or whitelist microsoft.com & download.microsoft.com." -ForegroundColor Yellow
    exit 1
}

Write-Host "Resolved ODT EXE: $exeUrl" -ForegroundColor Green

# 5) Download the EXE
$TempFile = Join-Path $env:TEMP "odt_download.exe"
Write-Host "Downloading ODT installer..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $exeUrl -OutFile $TempFile -Headers $Headers -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Host "ERROR: Download failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Quick sanity check: the file should be > 2 MB (ODT is approx 3-4 MB)
if ((Get-Item $TempFile).Length -lt 2000000) {
    Write-Host "ERROR: Downloaded file is too small; looks like HTML or blocked content." -ForegroundColor Red
    exit 1
}

# 6) Extract the ODT payload (setup.exe etc.) into .\ODT
Write-Host "Extracting ODT files..." -ForegroundColor Cyan
try {
    Start-Process -FilePath $TempFile -ArgumentList "/quiet /extract:`"$ODTFolder`"" -Wait
} catch {
    Write-Host "ERROR: Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 7) Create a silent uninstall configuration XML
$ConfigXML = Join-Path $ODTFolder "RemoveOffice.xml"
$XMLContent = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
Set-Content -Path $ConfigXML -Value $XMLContent -Encoding UTF8

# 8) Final checks & output
$setup = Join-Path $ODTFolder "setup.exe"
if (!(Test-Path $setup)) {
    Write-Host "ERROR: setup.exe not found after extraction." -ForegroundColor Red
    exit 1
}

Write-Host "===================================================" -ForegroundColor DarkCyan
Write-Host "ODT folder prepared: $ODTFolder" -ForegroundColor Green
Write-Host "Ready files:" -ForegroundColor Green
Write-Host " - $setup" -ForegroundColor Green
Write-Host " - $ConfigXML" -ForegroundColor Green
Write-Host "===================================================" -ForegroundColor DarkCyan
exit 0
