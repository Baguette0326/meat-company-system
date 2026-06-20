$ProjectRoot = $PSScriptRoot
$Desktop = [Environment]::GetFolderPath("Desktop")
$ChineseName = -join @(
    [char]0x92B7, # 銷
    [char]0x552E, # 售
    [char]0x5831, # 報
    [char]0x8868, # 表
    [char]0x7CFB, # 系
    [char]0x7D71  # 統
)
$ShortcutPath = Join-Path $Desktop ($ChineseName + ".lnk")
$ExePath = Join-Path $ProjectRoot ($ChineseName + ".exe")
$PythonwPath = Join-Path $ProjectRoot "runtime\pythonw.exe"
$ScriptPath = Join-Path $ProjectRoot "scripts\daily_report_app.py"
if (Test-Path -LiteralPath $ExePath) {
    $TargetPath = $ExePath
    $Arguments = ""
} elseif (Test-Path -LiteralPath $PythonwPath) {
    $TargetPath = $PythonwPath
    $Arguments = '"' + $ScriptPath + '"'
} else {
    $TargetPath = Join-Path $ProjectRoot "run_daily_report_app.bat"
    $Arguments = ""
}

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.Arguments = $Arguments
$Shortcut.WorkingDirectory = $ProjectRoot
$Shortcut.Description = "Sales report system"
$Shortcut.IconLocation = Join-Path $ProjectRoot "assets\daily-report-app.ico"
$Shortcut.Save()

Write-Host "Created: $ShortcutPath"
