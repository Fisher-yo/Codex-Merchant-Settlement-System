$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$settlementRoot = Join-Path $projectRoot 'workspace/04_商家周结算'
$toolDir = Join-Path $settlementRoot '99_结算模板与规则'
$createScript = Join-Path $toolDir '创建新周结算目录.ps1'
$splitScript = Join-Path $toolDir '拆分总对账单.ps1'
$supplierMapFile = Join-Path $toolDir '供应商代码映射表_模板.csv'
$logPath = Join-Path $projectRoot 'startup-error-log.txt'

function Write-StartupLog {
  param([string]$Message)

  $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  "[$time] $Message" | Out-File -LiteralPath $logPath -Encoding UTF8 -Append
}

function Read-RequiredText {
  param(
    [string]$Prompt,
    [string]$Default = ''
  )

  $suffix = if ([string]::IsNullOrWhiteSpace($Default)) { '' } else { " [$Default]" }
  $value = Read-Host "$Prompt$suffix"
  if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
  return $value.Trim()
}

function Convert-ToDateText {
  param([datetime]$Date)
  return $Date.ToString('yyyy.MM.dd')
}

function Get-WeekDir {
  param(
    [string]$Year,
    [string]$Period
  )

  return Join-Path (Join-Path $settlementRoot $Year) $Period
}

function Open-Path {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "路径不存在：$Path" -ForegroundColor Yellow
    return
  }

  if ($IsMacOS) {
    & open $Path
  } elseif ($IsLinux) {
    & xdg-open $Path
  } else {
    Invoke-Item -LiteralPath $Path
  }
}

function ConvertFrom-TerminalPath {
  param([string]$Path)

  $normalized = $Path.Trim().Trim('"').Trim("'")
  if ($IsMacOS -or $IsLinux) {
    $normalized = $normalized.Replace('\ ', ' ')
  }
  return $normalized
}

function Import-InputFile {
  param(
    [string]$SourcePath,
    [string]$WeekDir
  )

  if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "找不到文件：$SourcePath"
  }
  if (-not (Test-Path -LiteralPath $WeekDir)) {
    throw "请先准备本周结算目录：$WeekDir"
  }

  $inputDir = Join-Path $WeekDir '01_后台总对账单'
  New-Item -ItemType Directory -Force -Path $inputDir | Out-Null

  $target = Join-Path $inputDir (Split-Path -Leaf $SourcePath)
  $sourceFull = [IO.Path]::GetFullPath($SourcePath)
  $targetFull = [IO.Path]::GetFullPath($target)
  if ($sourceFull -ne $targetFull) {
    Copy-Item -LiteralPath $SourcePath -Destination $target -Force
  }

  return $target
}

function Show-State {
  param(
    [string]$Year,
    [string]$Period,
    [string]$WeekDir,
    [string]$InputFile
  )

  Clear-Host
  Write-Host '商家周结算操作台（macOS 版）' -ForegroundColor Cyan
  Write-Host ''
  Write-Host "年份：$Year"
  Write-Host "对账日期：$Period"
  Write-Host "本周目录：$WeekDir"
  Write-Host "后台总对账单：$(if ([string]::IsNullOrWhiteSpace($InputFile)) { '未导入' } else { $InputFile })"
  Write-Host ''
}

function Invoke-SettlementMenu {
  if (-not (Test-Path -LiteralPath $settlementRoot)) {
    throw "找不到工作目录：$settlementRoot"
  }
  if (-not (Test-Path -LiteralPath $createScript)) {
    throw "找不到建目录脚本：$createScript"
  }
  if (-not (Test-Path -LiteralPath $splitScript)) {
    throw "找不到拆分脚本：$splitScript"
  }

  $year = (Get-Date).ToString('yyyy')
  $period = Convert-ToDateText -Date (Get-Date).AddDays(-7)
  $weekDir = Get-WeekDir -Year $year -Period $period
  $inputFile = ''

  while ($true) {
    Show-State -Year $year -Period $period -WeekDir $weekDir -InputFile $inputFile
    Write-Host '1. 修改年份 / 对账日期'
    Write-Host '2. 准备本周结算目录'
    Write-Host '3. 导入后台总对账单（xlsx）'
    Write-Host '4. 生成结算结果'
    Write-Host '5. 打开本周目录'
    Write-Host '6. 打开商家对账单目录'
    Write-Host '7. 打开财务汇总目录'
    Write-Host '8. 打开确认台账目录'
    Write-Host '0. 退出'
    Write-Host ''

    $choice = Read-Host '请选择'
    try {
      switch ($choice) {
        '1' {
          $year = Read-RequiredText -Prompt '年份' -Default $year
          $exportDateText = Read-RequiredText -Prompt '后台导出日期，例如 2026-06-01' -Default (Get-Date).ToString('yyyy-MM-dd')
          $exportDate = [datetime]::MinValue
          if ([datetime]::TryParse($exportDateText, [ref]$exportDate)) {
            $period = Read-RequiredText -Prompt '对账日期' -Default (Convert-ToDateText -Date $exportDate.AddDays(-7))
          } else {
            $period = Read-RequiredText -Prompt '对账日期' -Default $period
          }
          $weekDir = Get-WeekDir -Year $year -Period $period
        }
        '2' {
          if (Test-Path -LiteralPath $weekDir) {
            Write-Host "本周结算目录已存在：$weekDir" -ForegroundColor Yellow
          } else {
            & $createScript -Period $period -Year $year
          }
          Read-Host '按回车继续'
        }
        '3' {
          $source = Read-RequiredText -Prompt '请把后台总对账单 xlsx 文件路径拖到这里后回车'
          $source = ConvertFrom-TerminalPath -Path $source
          $inputFile = Import-InputFile -SourcePath $source -WeekDir $weekDir
          Write-Host "已导入：$inputFile" -ForegroundColor Green
          Read-Host '按回车继续'
        }
        '4' {
          if ([string]::IsNullOrWhiteSpace($inputFile)) {
            $inputFile = Read-RequiredText -Prompt '后台总对账单路径'
            $inputFile = ConvertFrom-TerminalPath -Path $inputFile
          }

          $splitParams = @{
            InputFile = $inputFile
            OutputDir = $weekDir
            Period = $period
          }
          if (Test-Path -LiteralPath $supplierMapFile) {
            $splitParams.SupplierMapFile = $supplierMapFile
          }

          $supplementFile = Join-Path (Join-Path $weekDir '04_异常漏单处理') '手工补充订单_模板.csv'
          if (Test-Path -LiteralPath $supplementFile) {
            $splitParams.SupplementFile = $supplementFile
          }

          & $splitScript @splitParams
          Read-Host '按回车继续'
        }
        '5' { Open-Path -Path $weekDir; Read-Host '按回车继续' }
        '6' { Open-Path -Path (Join-Path $weekDir '02_商家拆分对账单'); Read-Host '按回车继续' }
        '7' { Open-Path -Path (Join-Path $weekDir '05_财务汇总'); Read-Host '按回车继续' }
        '8' { Open-Path -Path (Join-Path $weekDir '03_商家确认记录'); Read-Host '按回车继续' }
        '0' { return }
        default {
          Write-Host '请输入 0-8 之间的选项。' -ForegroundColor Yellow
          Read-Host '按回车继续'
        }
      }
    } catch {
      Write-StartupLog "操作失败：$($_.Exception.Message)"
      Write-Host "操作失败：$($_.Exception.Message)" -ForegroundColor Red
      Read-Host '按回车继续'
    }
  }
}

try {
  Write-StartupLog '开始启动 macOS 终端操作台。'
  Invoke-SettlementMenu
  Write-StartupLog 'macOS 终端操作台已关闭。'
} catch {
  Write-StartupLog "启动失败：$($_.Exception.Message)"
  Write-Host "启动失败：$($_.Exception.Message)" -ForegroundColor Red
  Write-Host "错误日志：$logPath"
  Read-Host '按回车关闭窗口'
  exit 1
}
