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
Write-Host "1. 在操作台导入后台总对账单"
Write-Host "2. 如有漏单，填写 04_异常漏单处理\手工补充订单_模板.csv"
Write-Host "3. 点击“生成结算结果”"
