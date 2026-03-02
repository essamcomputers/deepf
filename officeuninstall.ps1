# ======================================================
# Silent Office Uninstall + Logging + LibreOffice Associations
# ======================================================

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $ScriptPath "OfficeRemoval.log"

Function Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "[$timestamp] $Message"
}

Log "=== Starting Office Removal Script ==="

# --------------------------------------------
# 1. CREATE XML CONFIG
# --------------------------------------------
$ConfigXML = Join-Path $ScriptPath "RemoveOffice.xml"

$xmlContent = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@

Log "Creating XML configuration file..."
Set-Content -Path $ConfigXML -Value $xmlContent -Encoding UTF8

# --------------------------------------------
# 2. DOWNLOAD ODT IF MISSING
# --------------------------------------------
$ODTExe = Join-Path $ScriptPath "setup.exe"

if (!(Test-Path $ODTExe)) {
    Log "ODT not found. Downloading Office Deployment Tool..."
    $ODTUrl = "https://download.microsoft.com/download/2/e/8/2e8a6c9c-a6e2-4c7b-a3e3-6d5cf202e6e2/officedeploymenttool_16227-20258.exe"
    $TempODT = Join-Path $env:TEMP "odt.exe"

    try {
        Invoke-WebRequest -Uri $ODTUrl -OutFile $TempODT -ErrorAction Stop
        Log "ODT Downloaded. Extracting..."
        Start-Process -FilePath $TempODT -ArgumentList "/quiet /extract:`"$ScriptPath`"" -Wait
        Log "ODT extracted successfully."
    }
    catch {
        Log "ERROR: Failed to download or extract ODT. $_"
        exit 1
    }
}
else {
    Log "ODT already present."
}

# --------------------------------------------
# 3. UNINSTALL OFFICE SILENTLY
# --------------------------------------------
Log "Starting silent uninstall using ODT..."
try {
    Start-Process -FilePath $ODTExe -ArgumentList "/configure `"$ConfigXML`"" -Wait -NoNewWindow
    Log "ODT uninstall completed."
}
catch {
    Log "ERROR: Uninstall process failed. $_"
    exit 1
}

# --------------------------------------------
# 4. ASSOCIATE OFFICE FILES WITH LIBREOFFICE
# --------------------------------------------
Log "Associating Office file types with LibreOffice..."

# LibreOffice default paths (adjust if needed)
$LibrePath = "C:\Program Files\LibreOffice\program"
$Writer = Join-Path $LibrePath "soffice.exe"
$Calc  = Join-Path $LibrePath "scalc.exe"
$Impress = Join-Path $LibrePath "simpress.exe"

# Ensure LibreOffice exists
if (!(Test-Path $Writer)) {
    Log "WARNING: LibreOffice not found. Skipping file associations."
}
else {
    # Word types → Writer
    $wordExt = @(".doc", ".docx", ".dotx", ".rtf")
    foreach ($ext in $wordExt) {
        ftype "LibreOfficeWriter=$Writer `"%1`""
        assoc "$ext=LibreOfficeWriter"
    }

    # Excel types → Calc
    $excelExt = @(".xls", ".xlsx", ".xlsm", ".csv")
    foreach ($ext in $excelExt) {
        ftype "LibreOfficeCalc=$Calc `"%1`""
        assoc "$ext=LibreOfficeCalc"
    }

    # PowerPoint types → Impress
    $pptExt = @(".ppt", ".pptx", ".ppsx")
    foreach ($ext in $pptExt) {
        ftype "LibreOfficeImpress=$Impress `"%1`""
        assoc "$ext=LibreOfficeImpress"
    }

    Log "File associations updated to LibreOffice."
}

Log "=== Script completed successfully ==="
exit 0
