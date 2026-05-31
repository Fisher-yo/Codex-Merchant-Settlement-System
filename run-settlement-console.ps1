$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$logPath = Join-Path $projectRoot 'startup-error-log.txt'
$consoleScript = Join-Path $projectRoot 'workspace\04_商家周结算\周结算操作台.ps1'

function Write-StartupLog {
  param([string]$Message)

  $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  "[$time] $Message" | Out-File -LiteralPath $logPath -Encoding UTF8 -Append
}

try {
  if (-not (Test-Path -LiteralPath $consoleScript)) {
    throw "找不到周结算操作台脚本：$consoleScript"
  }

  Write-StartupLog "开始启动周结算操作台：$consoleScript"
  & $consoleScript
  Write-StartupLog "周结算操作台已关闭。"
} catch {
  $message = $_.Exception.Message
  Write-StartupLog "启动失败：$message"
  Write-Host ""
  Write-Host "周结算操作台启动失败：" -ForegroundColor Red
  Write-Host $message
  Write-Host ""
  Write-Host "错误日志：$logPath"
  Write-Host ""
  if ($env:SETTLEMENT_CONSOLE_LAUNCHED_BY_EXE -eq '1') {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("周结算操作台启动失败：`r`n$message`r`n`r`n错误日志：$logPath", '启动失败') | Out-Null
  } else {
    Read-Host "按回车键关闭窗口"
  }
  exit 1
}
