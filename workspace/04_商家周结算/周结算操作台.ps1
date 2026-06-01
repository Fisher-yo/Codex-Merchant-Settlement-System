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

function New-UiColor {
  param([int]$R, [int]$G, [int]$B)
  return [System.Drawing.Color]::FromArgb($R, $G, $B)
}

function Set-ControlMargin {
  param(
    [System.Windows.Forms.Control]$Control,
    [int]$Left = 0,
    [int]$Top = 0,
    [int]$Right = 0,
    [int]$Bottom = 0
  )
  $Control.Margin = New-Object System.Windows.Forms.Padding($Left, $Top, $Right, $Bottom)
}

function Set-ModernButton {
  param(
    [System.Windows.Forms.Button]$Button,
    [System.Drawing.Color]$BackColor,
    [System.Drawing.Color]$ForeColor,
    [System.Drawing.Color]$BorderColor,
    [switch]$Strong
  )

  $Button.FlatStyle = 'Flat'
  $Button.BackColor = $BackColor
  $Button.ForeColor = $ForeColor
  $Button.FlatAppearance.BorderColor = $BorderColor
  $Button.FlatAppearance.BorderSize = 1
  $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
  $Button.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', $(if ($Strong) { 10 } else { 9 }), $(if ($Strong) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }))
  $Button.UseVisualStyleBackColor = $false
  $Button.Add_MouseEnter({
    $this.FlatAppearance.BorderColor = $colorPrimary
    $this.BackColor = Blend-UiColor -A $this.BackColor -B $colorPrimary -Percent 8
  })
  $Button.Add_MouseLeave({
    $this.BackColor = $this.Tag.BackColor
    $this.FlatAppearance.BorderColor = $this.Tag.BorderColor
  })
  $Button.Tag = [pscustomobject]@{
    BackColor = $BackColor
    BorderColor = $BorderColor
  }
}

function Blend-UiColor {
  param(
    [System.Drawing.Color]$A,
    [System.Drawing.Color]$B,
    [int]$Percent
  )

  $ratio = [Math]::Max(0, [Math]::Min(100, $Percent)) / 100
  return [System.Drawing.Color]::FromArgb(
    [int]($A.R + (($B.R - $A.R) * $ratio)),
    [int]($A.G + (($B.G - $A.G) * $ratio)),
    [int]($A.B + (($B.B - $A.B) * $ratio))
  )
}

function Set-AppTextBox {
  param(
    [System.Windows.Forms.TextBox]$TextBox,
    [switch]$ReadOnly
  )

  $TextBox.BorderStyle = 'FixedSingle'
  $TextBox.BackColor = $(if ($ReadOnly) { $colorFieldMuted } else { $colorField })
  $TextBox.ForeColor = $(if ($ReadOnly) { $colorMuted } else { $colorInk })
  $TextBox.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)
}

function Add-PanelBorderPaint {
  param(
    [System.Windows.Forms.Control]$Panel,
    [System.Drawing.Color]$BorderColor = $colorLine
  )

  $Panel.Tag = [pscustomobject]@{
    BorderColor = $BorderColor
  }
  $Panel.Add_Paint({
    param($sender, $event)
    $rect = New-Object System.Drawing.Rectangle(0, 0, ($sender.Width - 1), ($sender.Height - 1))
    $pen = New-Object System.Drawing.Pen($sender.Tag.BorderColor, 1)
    $event.Graphics.DrawRectangle($pen, $rect)
    $pen.Dispose()
  })
}

$colorPage = New-UiColor 246 248 251
$colorCard = New-UiColor 255 255 255
$colorCardAlt = New-UiColor 255 255 255
$colorInk = New-UiColor 31 41 55
$colorMuted = New-UiColor 100 116 139
$colorLine = New-UiColor 218 226 236
$colorPrimary = New-UiColor 37 99 235
$colorPrimaryDark = New-UiColor 29 78 216
$colorPrimarySoft = New-UiColor 239 246 255
$colorAccent = New-UiColor 37 99 235
$colorCyan = New-UiColor 37 99 235
$colorCyanDim = New-UiColor 191 219 254
$colorGreen = New-UiColor 22 163 74
$colorField = New-UiColor 255 255 255
$colorFieldMuted = New-UiColor 248 250 252
$colorLog = New-UiColor 248 250 252

$form = New-Object System.Windows.Forms.Form
$form.Text = '周结算操作台'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1000, 720)
$form.MinimumSize = New-Object System.Drawing.Size(920, 660)
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.BackColor = $colorPage
$form.ForeColor = $colorInk
$iconPath = Join-Path (Split-Path -Parent $settlementRoot) 'assets\settlement-console.ico'
if (Test-Path -LiteralPath $iconPath) {
  $form.Icon = New-Object System.Drawing.Icon($iconPath)
}

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.Padding = New-Object System.Windows.Forms.Padding(22)
$main.BackColor = $colorPage
$main.ColumnCount = 1
$main.RowCount = 5
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 110))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 106))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 160))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 66))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$form.Controls.Add($main)

$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Dock = 'Fill'
$titlePanel.BackColor = $colorPage
$titlePanel.Add_Paint({
  param($sender, $event)
  $linePen = New-Object System.Drawing.Pen($colorLine, 1)
  $event.Graphics.DrawLine($linePen, 0, ($sender.Height - 10), $sender.Width, ($sender.Height - 10))
  $linePen.Dispose()
})
$title = New-Object System.Windows.Forms.Label
$title.Text = '商家周结算操作台'
$title.ForeColor = $colorInk
$title.BackColor = [System.Drawing.Color]::Transparent
$title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 22, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(2, 8)
$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = '确认周期、导入后台总对账单，然后生成商家对账单、财务汇总表和确认台账。'
$subtitle.ForeColor = $colorMuted
$subtitle.BackColor = [System.Drawing.Color]::Transparent
$subtitle.AutoSize = $true
$subtitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$subtitle.Location = New-Object System.Drawing.Point(4, 54)
$badge = New-Object System.Windows.Forms.Label
$badge.Text = '本地周结算工具'
$badge.ForeColor = $colorPrimaryDark
$badge.BackColor = $colorPrimarySoft
$badge.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$badge.AutoSize = $true
$badge.Padding = New-Object System.Windows.Forms.Padding(9, 4, 9, 4)
$badge.Location = New-Object System.Drawing.Point(5, 80)
$statusBadge = New-Object System.Windows.Forms.Label
$statusBadge.Text = 'Windows'
$statusBadge.ForeColor = $colorMuted
$statusBadge.BackColor = $colorPage
$statusBadge.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$statusBadge.AutoSize = $true
$statusBadge.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$statusBadge.Anchor = 'Top,Right'
$statusBadge.Location = New-Object System.Drawing.Point(830, 18)
$titlePanel.Controls.Add($title)
$titlePanel.Controls.Add($subtitle)
$titlePanel.Controls.Add($badge)
$titlePanel.Controls.Add($statusBadge)
$titlePanel.Add_SizeChanged({
  $statusBadge.Left = [Math]::Max(620, $titlePanel.ClientSize.Width - $statusBadge.Width - 22)
})
$main.Controls.Add($titlePanel, 0, 0)

$periodPanel = New-Object System.Windows.Forms.TableLayoutPanel
$periodPanel.Dock = 'Fill'
$periodPanel.BackColor = $colorCard
$periodPanel.Padding = New-Object System.Windows.Forms.Padding(18, 14, 18, 12)
$periodPanel.ColumnCount = 7
$periodPanel.RowCount = 2
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 48))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 72))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 72))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
$periodPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
Add-PanelBorderPaint -Panel $periodPanel -BorderColor $colorLine

$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = '年份'
$yearLabel.TextAlign = 'MiddleLeft'
$yearLabel.Dock = 'Fill'
$yearLabel.ForeColor = $colorMuted
$yearText = New-Object System.Windows.Forms.TextBox
$yearText.Text = '2026'
$yearText.Dock = 'Fill'
Set-AppTextBox -TextBox $yearText

$exportDateLabel = New-Object System.Windows.Forms.Label
$exportDateLabel.Text = '导出日期'
$exportDateLabel.TextAlign = 'MiddleLeft'
$exportDateLabel.Dock = 'Fill'
$exportDateLabel.ForeColor = $colorMuted
$exportDateText = New-Object System.Windows.Forms.TextBox
$exportDateText.Text = (Get-Date).ToString('yyyy-MM-dd')
$exportDateText.Dock = 'Fill'
Set-AppTextBox -TextBox $exportDateText

$periodLabel = New-Object System.Windows.Forms.Label
$periodLabel.Text = '对账日期'
$periodLabel.TextAlign = 'MiddleLeft'
$periodLabel.Dock = 'Fill'
$periodLabel.ForeColor = $colorMuted
$periodText = New-Object System.Windows.Forms.TextBox
$periodText.Text = Convert-ToDateText -Date (Get-Date).AddDays(-7)
$periodText.Dock = 'Fill'
Set-AppTextBox -TextBox $periodText

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = '准备本周目录'
$createButton.Dock = 'Fill'
Set-ModernButton -Button $createButton -BackColor $colorPrimary -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorPrimaryDark -Strong

$currentDirText = New-Object System.Windows.Forms.TextBox
$currentDirText.Dock = 'Fill'
$currentDirText.ReadOnly = $true
Set-AppTextBox -TextBox $currentDirText -ReadOnly

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
$dropPanel.BorderStyle = 'None'
$dropPanel.AllowDrop = $true
$dropPanel.BackColor = $colorCardAlt
Add-PanelBorderPaint -Panel $dropPanel -BorderColor $colorLine
$dropTitle = New-Object System.Windows.Forms.Label
$dropTitle.Text = '拖入后台总对账单'
$dropTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 14, [System.Drawing.FontStyle]::Bold)
$dropTitle.ForeColor = $colorInk
$dropTitle.AutoSize = $true
$dropTitle.Location = New-Object System.Drawing.Point(28, 24)
$dropTip = New-Object System.Windows.Forms.Label
$dropTip.Text = '支持 .xlsx 文件，导入后自动复制到本周工作包的 01_后台总对账单。'
$dropTip.ForeColor = $colorMuted
$dropTip.AutoSize = $true
$dropTip.Location = New-Object System.Drawing.Point(30, 62)
$inputFileText = New-Object System.Windows.Forms.TextBox
$inputFileText.ReadOnly = $true
$inputFileText.Anchor = 'Left,Right,Bottom'
$inputFileText.Location = New-Object System.Drawing.Point(28, 112)
$inputFileText.Width = 760
Set-AppTextBox -TextBox $inputFileText -ReadOnly
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = '选择文件'
$browseButton.Anchor = 'Right,Bottom'
$browseButton.Location = New-Object System.Drawing.Point(812, 108)
$browseButton.Size = New-Object System.Drawing.Size(112, 34)
Set-ModernButton -Button $browseButton -BackColor ([System.Drawing.Color]::White) -ForeColor $colorPrimaryDark -BorderColor $colorLine
$dropPanel.Controls.Add($dropTitle)
$dropPanel.Controls.Add($dropTip)
$dropPanel.Controls.Add($inputFileText)
$dropPanel.Controls.Add($browseButton)
$dropPanel.Add_SizeChanged({
  $browseButton.Left = [Math]::Max(560, $dropPanel.ClientSize.Width - $browseButton.Width - 28)
  $inputFileText.Width = [Math]::Max(280, $browseButton.Left - $inputFileText.Left - 18)
})
$main.Controls.Add($dropPanel, 0, 2)

$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.Dock = 'Fill'
$actionPanel.FlowDirection = 'LeftToRight'
$actionPanel.WrapContents = $false
$actionPanel.BackColor = $colorPage
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 8)

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = '生成结算结果'
$generateButton.Size = New-Object System.Drawing.Size(148, 38)
Set-ModernButton -Button $generateButton -BackColor $colorPrimary -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorPrimaryDark -Strong
$openWeekButton = New-Object System.Windows.Forms.Button
$openWeekButton.Text = '打开本周目录'
$openWeekButton.Size = New-Object System.Drawing.Size(126, 38)
Set-ModernButton -Button $openWeekButton -BackColor ([System.Drawing.Color]::White) -ForeColor $colorInk -BorderColor $colorLine
$openSplitButton = New-Object System.Windows.Forms.Button
$openSplitButton.Text = '商家对账单'
$openSplitButton.Size = New-Object System.Drawing.Size(118, 38)
Set-ModernButton -Button $openSplitButton -BackColor ([System.Drawing.Color]::White) -ForeColor $colorInk -BorderColor $colorLine
$openSummaryButton = New-Object System.Windows.Forms.Button
$openSummaryButton.Text = '财务汇总'
$openSummaryButton.Size = New-Object System.Drawing.Size(104, 38)
Set-ModernButton -Button $openSummaryButton -BackColor ([System.Drawing.Color]::White) -ForeColor $colorInk -BorderColor $colorLine
$openTrackButton = New-Object System.Windows.Forms.Button
$openTrackButton.Text = '确认台账'
$openTrackButton.Size = New-Object System.Drawing.Size(104, 38)
Set-ModernButton -Button $openTrackButton -BackColor ([System.Drawing.Color]::White) -ForeColor $colorInk -BorderColor $colorLine

Set-ControlMargin -Control $generateButton -Right 10
Set-ControlMargin -Control $openWeekButton -Right 8
Set-ControlMargin -Control $openSplitButton -Right 8
Set-ControlMargin -Control $openSummaryButton -Right 8
Set-ControlMargin -Control $openTrackButton -Right 8

$actionPanel.Controls.Add($generateButton)
$actionPanel.Controls.Add($openWeekButton)
$actionPanel.Controls.Add($openSplitButton)
$actionPanel.Controls.Add($openSummaryButton)
$actionPanel.Controls.Add($openTrackButton)
$main.Controls.Add($actionPanel, 0, 3)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Dock = 'Fill'
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true
$logBox.BorderStyle = 'FixedSingle'
$logBox.BackColor = $colorLog
$logBox.ForeColor = $colorInk
$logBox.Font = New-Object System.Drawing.Font('Consolas', 10)
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

$yearText.Add_TextChanged({ Set-CurrentWeekDir | Out-Null })
$exportDateText.Add_TextChanged({
  Update-StatementDateFromExportDate
  Set-CurrentWeekDir | Out-Null
})
$periodText.Add_TextChanged({ Set-CurrentWeekDir | Out-Null })

Set-CurrentWeekDir | Out-Null
Add-Log "操作台已打开。先确认导出日期和对账日期，再新建本次结算。"

[void]$form.ShowDialog()
