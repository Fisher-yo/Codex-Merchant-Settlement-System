$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$launcherPath = Join-Path $projectRoot '启动周结算操作台.cmd'

if (-not (Test-Path -LiteralPath $launcherPath)) {
  throw "找不到启动文件：$launcherPath"
}

$desktopPath = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktopPath '周结算操作台.lnk'

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcherPath
$shortcut.WorkingDirectory = $projectRoot
$shortcut.WindowStyle = 1
$shortcut.Description = '商家周结算操作台'
$shortcut.Save()

Write-Host "已创建桌面快捷方式：$shortcutPath"
