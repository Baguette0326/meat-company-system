$ProjectRoot = $PSScriptRoot
$Desktop = [Environment]::GetFolderPath("Desktop")
$LauncherPath = Join-Path $Desktop "Daily Report App.bat"

$Content = @"
@echo off
cd /d "$ProjectRoot"
call "$ProjectRoot\run_daily_report_app.bat"
"@

Set-Content -LiteralPath $LauncherPath -Value $Content -Encoding Default
Write-Host "Created: $LauncherPath"
