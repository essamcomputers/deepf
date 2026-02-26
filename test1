# DF_Test_WriteFile.ps1
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$user  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$hostn = $env:COMPUTERNAME

$dir  = "C:\ProgramData\DeepFreezeTest"
$file = Join-Path $dir "df_test.log"

New-Item -Path $dir -ItemType Directory -Force | Out-Null
Add-Content -Path $file -Value "$stamp | Ran OK | User=$user | PC=$hostn"
exit 0
