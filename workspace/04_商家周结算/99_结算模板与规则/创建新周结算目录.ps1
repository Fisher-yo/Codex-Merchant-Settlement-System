param(
  [string]$WeekCode = "",

  [Parameter(Mandatory = $true)]
  [string]$Period,

  [string]$Year = "2026"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$settlementRoot = Split-Path -Parent $scriptDir
$templateDir = Join-Path $settlementRoot "00_每周结算工作包模板"
$yearDir = Join-Path $settlementRoot $Year
$targetName = if ([string]::IsNullOrWhiteSpace($WeekCode)) { $Period } else { "${WeekCode}_${Period}" }
$targetDir = Join-Path $yearDir $targetName

if (-not (Test-Path -LiteralPath $templateDir)) {
  throw "找不到模板目录：$templateDir"
}

if (Test-Path -LiteralPath $targetDir) {
  throw "目标目录已存在：$targetDir"
}

New-Item -ItemType Directory -Force -Path $yearDir | Out-Null
Copy-Item -LiteralPath $templateDir -Destination $targetDir -Recurse

Write-Host "已创建新周结算目录：$targetDir"
Write-Host ""
Write-Host "下一步："
Write-Host "1. 将后台总对账单放入 01_后台总对账单"
Write-Host "2. 拆分后的商家对账单放入 02_商家拆分对账单"
Write-Host "3. 用 03_商家确认记录 维护商家确认进度"
