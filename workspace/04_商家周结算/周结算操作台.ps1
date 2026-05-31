Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

$settlementRoot = Split-Path -Parent $PSCommandPath
$toolDir = Join-Path $settlementRoot '99_结算模板与规则'
$createScript = Join-Path $toolDir '创建新周结算目录.ps1'
$splitScript = Join-Path $toolDir '拆分总对账单.ps1'
$supplierMapFile = Join-Path $toolDir '供应商代码映射表_模板.csv'

function Add-Log {
  param([string]$Message)
  $time = Get-Date -Format 'HH:mm:ss'
  $logBox.AppendText("[$time] $Message`r`n")
}

function Get-CurrentPeriodName {
  $period = $periodText.Text.Trim()
  if ([string]::IsNullOrWhiteSpace($period)) {
    return ''
  }
  return $period
}

function Convert-ToDateText {
  param([datetime]$Date)
  return $Date.ToString('yyyy.MM.dd')
}

function Update-StatementDateFromExportDate {
  $exportDate = [datetime]::MinValue
  if ([datetime]::TryParse($exportDateText.Text.Trim(), [ref]$exportDate)) {
    $periodText.Text = Convert-ToDateText -Date $exportDate.AddDays(-7)
  }
}

function Get-CurrentWeekDir {
  $year = $yearText.Text.Trim()
  $periodName = Get-CurrentPeriodName
  if ([string]::IsNullOrWhiteSpace($year) -or [string]::IsNullOrWhiteSpace($periodName)) {
    return ''
  }
  return Join-Path (Join-Path $settlementRoot $year) $periodName
}

function Set-CurrentWeekDir {
  $dir = Get-CurrentWeekDir
  $currentDirText.Text = $dir
  return $dir
}

function Invoke-WithBusyCursor {
  param([scriptblock]$Action)
  $oldCursor = $form.Cursor
  $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
  $form.Enabled = $false
  try {
    & $Action
  } finally {
    $form.Enabled = $true
    $form.Cursor = $oldCursor
  }
}

function Open-Directory {
  param([string]$Path)
  if (Test-Path -LiteralPath $Path) {
    Invoke-Item -LiteralPath $Path
  } else {
    Add-Log "目录不存在：$Path"
    [System.Windows.Forms.MessageBox]::Show("目录不存在：`r`n$Path", '提示') | Out-Null
  }
}

function Import-InputFile {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "找不到文件：$Path"
  }

  $weekDir = Set-CurrentWeekDir
  if (-not (Test-Path -LiteralPath $weekDir)) {
    throw "请先新建或确认本周结算目录。"
  }

  $inputDir = Join-Path $weekDir '01_后台总对账单'
  New-Item -ItemType Directory -Force -Path $inputDir | Out-Null

  $target = Join-Path $inputDir (Split-Path -Leaf $Path)
  $sourceFull = [IO.Path]::GetFullPath($Path)
  $targetFull = [IO.Path]::GetFullPath($target)

  if ($sourceFull -ne $targetFull) {
    Copy-Item -LiteralPath $Path -Destination $target -Force
    Add-Log "已导入后台总对账单：$target"
  } else {
    Add-Log "后台总对账单已在本周目录中：$target"
  }

  $inputFileText.Text = $target
}

function Generate-Settlement {
  $weekDir = $currentDirText.Text.Trim()
  $inputFile = $inputFileText.Text.Trim()
  $periodName = Get-CurrentPeriodName

  if ([string]::IsNullOrWhiteSpace($weekDir) -or -not (Test-Path -LiteralPath $weekDir)) {
    throw "请先新建或选择本周结算目录。"
  }
  if ([string]::IsNullOrWhiteSpace($inputFile) -or -not (Test-Path -LiteralPath $inputFile)) {
    throw "请先导入后台总对账单。"
  }
  if ([string]::IsNullOrWhiteSpace($periodName)) {
    throw "请填写或确认对账日期。"
  }

  $supplementFile = Join-Path $weekDir '04_异常漏单处理\手工补充订单_模板.csv'
  $splitParams = @{
    InputFile = $inputFile
    OutputDir = $weekDir
    Period = $periodName
  }

  if (Test-Path -LiteralPath $supplierMapFile) {
    $splitParams.SupplierMapFile = $supplierMapFile
  }
  if (Test-Path -LiteralPath $supplementFile) {
    $splitParams.SupplementFile = $supplementFile
  }

  Add-Log "开始生成结算结果..."
  $output = & $splitScript @splitParams 2>&1 | Out-String
  if (-not [string]::IsNullOrWhiteSpace($output)) {
    $logBox.AppendText($output.TrimEnd() + "`r`n")
  }
  Add-Log "结算结果已生成。"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '周结算操作台'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(920, 680)
$form.MinimumSize = New-Object System.Drawing.Size(860, 620)
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.Padding = New-Object System.Windows.Forms.Padding(16)
$main.ColumnCount = 1
$main.RowCount = 5
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 92))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 82))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$form.Controls.Add($main)

$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Dock = 'Fill'
$title = New-Object System.Windows.Forms.Label
$title.Text = '商家周结算'
$title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(0, 0)
$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = '按“后台导出日期 - 7 天”生成对账日期，拖入后台总对账单，一键生成商家对账单、财务汇总表和跟进台账。'
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(2, 44)
$titlePanel.Controls.Add($title)
$titlePanel.Controls.Add($subtitle)
$main.Controls.Add($titlePanel, 0, 0)

$periodPanel = New-Object System.Windows.Forms.TableLayoutPanel
$periodPanel.Dock = 'Fill'
$periodPanel.ColumnCount = 7
$periodPanel.RowCount = 2
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 48))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 72))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 72))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = '年份'
$yearLabel.TextAlign = 'MiddleLeft'
$yearLabel.Dock = 'Fill'
$yearText = New-Object System.Windows.Forms.TextBox
$yearText.Text = '2026'
$yearText.Dock = 'Fill'

$exportDateLabel = New-Object System.Windows.Forms.Label
$exportDateLabel.Text = '导出日期'
$exportDateLabel.TextAlign = 'MiddleLeft'
$exportDateLabel.Dock = 'Fill'
$exportDateText = New-Object System.Windows.Forms.TextBox
$exportDateText.Text = (Get-Date).ToString('yyyy-MM-dd')
$exportDateText.Dock = 'Fill'

$periodLabel = New-Object System.Windows.Forms.Label
$periodLabel.Text = '对账日期'
$periodLabel.TextAlign = 'MiddleLeft'
$periodLabel.Dock = 'Fill'
$periodText = New-Object System.Windows.Forms.TextBox
$periodText.Text = Convert-ToDateText -Date (Get-Date).AddDays(-7)
$periodText.Dock = 'Fill'

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = '新建本周结算'
$createButton.Dock = 'Fill'

$currentDirText = New-Object System.Windows.Forms.TextBox
$currentDirText.Dock = 'Fill'
$currentDirText.ReadOnly = $true

$periodPanel.Controls.Add($yearLabel, 0, 0)
$periodPanel.Controls.Add($yearText, 1, 0)
$periodPanel.Controls.Add($exportDateLabel, 2, 0)
$periodPanel.Controls.Add($exportDateText, 3, 0)
$periodPanel.Controls.Add($periodLabel, 4, 0)
$periodPanel.Controls.Add($periodText, 5, 0)
$periodPanel.Controls.Add($createButton, 6, 0)
$periodPanel.SetColumnSpan($currentDirText, 7)
$periodPanel.Controls.Add($currentDirText, 0, 1)
$main.Controls.Add($periodPanel, 0, 1)

$dropPanel = New-Object System.Windows.Forms.Panel
$dropPanel.Dock = 'Fill'
$dropPanel.BorderStyle = 'FixedSingle'
$dropPanel.AllowDrop = $true
$dropTitle = New-Object System.Windows.Forms.Label
$dropTitle.Text = '拖入后台总对账单'
$dropTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 14, [System.Drawing.FontStyle]::Bold)
$dropTitle.AutoSize = $true
$dropTitle.Location = New-Object System.Drawing.Point(24, 20)
$dropTip = New-Object System.Windows.Forms.Label
$dropTip.Text = '支持 .xlsx 文件。拖入后会复制到本周目录的 01_后台总对账单。'
$dropTip.AutoSize = $true
$dropTip.Location = New-Object System.Drawing.Point(26, 56)
$inputFileText = New-Object System.Windows.Forms.TextBox
$inputFileText.ReadOnly = $true
$inputFileText.Anchor = 'Left,Right,Bottom'
$inputFileText.Location = New-Object System.Drawing.Point(24, 96)
$inputFileText.Width = 700
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = '选择文件'
$browseButton.Anchor = 'Right,Bottom'
$browseButton.Location = New-Object System.Drawing.Point(742, 94)
$browseButton.Size = New-Object System.Drawing.Size(104, 28)
$dropPanel.Controls.Add($dropTitle)
$dropPanel.Controls.Add($dropTip)
$dropPanel.Controls.Add($inputFileText)
$dropPanel.Controls.Add($browseButton)
$main.Controls.Add($dropPanel, 0, 2)

$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.Dock = 'Fill'
$actionPanel.FlowDirection = 'LeftToRight'
$actionPanel.WrapContents = $false

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = '生成结算结果'
$generateButton.Size = New-Object System.Drawing.Size(130, 36)
$openWeekButton = New-Object System.Windows.Forms.Button
$openWeekButton.Text = '打开本周目录'
$openWeekButton.Size = New-Object System.Drawing.Size(120, 36)
$openSplitButton = New-Object System.Windows.Forms.Button
$openSplitButton.Text = '商家对账单'
$openSplitButton.Size = New-Object System.Drawing.Size(110, 36)
$openSummaryButton = New-Object System.Windows.Forms.Button
$openSummaryButton.Text = '财务汇总'
$openSummaryButton.Size = New-Object System.Drawing.Size(96, 36)
$openTrackButton = New-Object System.Windows.Forms.Button
$openTrackButton.Text = '确认台账'
$openTrackButton.Size = New-Object System.Drawing.Size(96, 36)
$openExceptionButton = New-Object System.Windows.Forms.Button
$openExceptionButton.Text = '异常处理'
$openExceptionButton.Size = New-Object System.Drawing.Size(96, 36)

$actionPanel.Controls.Add($generateButton)
$actionPanel.Controls.Add($openWeekButton)
$actionPanel.Controls.Add($openSplitButton)
$actionPanel.Controls.Add($openSummaryButton)
$actionPanel.Controls.Add($openTrackButton)
$actionPanel.Controls.Add($openExceptionButton)
$main.Controls.Add($actionPanel, 0, 3)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Dock = 'Fill'
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true
$main.Controls.Add($logBox, 0, 4)

$createButton.Add_Click({
  try {
    Invoke-WithBusyCursor {
      $weekDir = Set-CurrentWeekDir
      if (Test-Path -LiteralPath $weekDir) {
        Add-Log "本周结算目录已存在：$weekDir"
        return
      }

      & $createScript -Period $periodText.Text.Trim() -Year $yearText.Text.Trim() 2>&1 | ForEach-Object {
        Add-Log ([string]$_)
      }
      Add-Log "已准备本周结算目录：$weekDir"
    }
  } catch {
    Add-Log "失败：$($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '新建失败') | Out-Null
  }
})

$browseButton.Add_Click({
  try {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = 'Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*'
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      Invoke-WithBusyCursor { Import-InputFile -Path $dialog.FileName }
    }
  } catch {
    Add-Log "失败：$($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '导入失败') | Out-Null
  }
})

$dropPanel.Add_DragEnter({
  if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
    $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
  }
})

$dropPanel.Add_DragDrop({
  try {
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
      Invoke-WithBusyCursor { Import-InputFile -Path $files[0] }
    }
  } catch {
    Add-Log "失败：$($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '导入失败') | Out-Null
  }
})

$generateButton.Add_Click({
  try {
    Invoke-WithBusyCursor { Generate-Settlement }
    [System.Windows.Forms.MessageBox]::Show('结算结果已生成。', '完成') | Out-Null
  } catch {
    Add-Log "失败：$($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '生成失败') | Out-Null
  }
})

$openWeekButton.Add_Click({ Open-Directory -Path $currentDirText.Text.Trim() })
$openSplitButton.Add_Click({ Open-Directory -Path (Join-Path $currentDirText.Text.Trim() '02_商家拆分对账单') })
$openSummaryButton.Add_Click({ Open-Directory -Path (Join-Path $currentDirText.Text.Trim() '05_财务汇总') })
$openTrackButton.Add_Click({ Open-Directory -Path (Join-Path $currentDirText.Text.Trim() '03_商家确认记录') })
$openExceptionButton.Add_Click({ Open-Directory -Path (Join-Path $currentDirText.Text.Trim() '99_异常待处理') })

$yearText.Add_TextChanged({ Set-CurrentWeekDir | Out-Null })
$exportDateText.Add_TextChanged({
  Update-StatementDateFromExportDate
  Set-CurrentWeekDir | Out-Null
})
$periodText.Add_TextChanged({ Set-CurrentWeekDir | Out-Null })

Set-CurrentWeekDir | Out-Null
Add-Log "操作台已打开。先确认导出日期和对账日期，再新建本次结算。"

[void]$form.ShowDialog()
