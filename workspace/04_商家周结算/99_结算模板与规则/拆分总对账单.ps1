param(
  [Parameter(Mandatory = $true)]
  [string]$InputFile,

  [Parameter(Mandatory = $true)]
  [string]$OutputDir,

  [string]$Period = "",

  [string]$SupplierColumn = "供应商代号",

  [string]$AmountColumn = "预计结账金额",

  [string]$SupplierMapFile = "",

  [string]$SupplementFile = ""
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipText {
  param($Zip, [string]$Name)
  $entry = $Zip.GetEntry($Name)
  if (-not $entry) {
    $entry = $Zip.GetEntry($Name.Replace('/', '\'))
  }
  if (-not $entry) { return $null }
  $reader = New-Object IO.StreamReader($entry.Open())
  try {
    return $reader.ReadToEnd()
  } finally {
    $reader.Close()
  }
}

function Convert-ColumnNameToIndex {
  param([string]$CellRef)
  $letters = $CellRef -replace '[0-9]', ''
  $num = 0
  foreach ($ch in $letters.ToCharArray()) {
    $num = $num * 26 + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $num
}

function ConvertTo-SafeFileName {
  param([string]$Name)
  $safe = $Name
  foreach ($ch in [IO.Path]::GetInvalidFileNameChars()) {
    $safe = $safe.Replace([string]$ch, '_')
  }
  if ([string]::IsNullOrWhiteSpace($safe)) { return "未填写供应商代号" }
  return $safe.Trim()
}

function Write-CsvUtf8Bom {
  param(
    [Parameter(Mandatory = $true)]
    [array]$Rows,
    [Parameter(Mandatory = $true)]
    [string]$Path
  )
  $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function ConvertTo-XmlText {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [Security.SecurityElement]::Escape($Text)
}

function ConvertTo-ExcelColumnName {
  param([int]$Index)
  $name = ''
  while ($Index -gt 0) {
    $mod = ($Index - 1) % 26
    $name = [char](65 + $mod) + $name
    $Index = [math]::Floor(($Index - $mod) / 26)
  }
  return $name
}

function Get-TextDisplayWidth {
  param([string]$Text)
  if ($null -eq $Text) { return 0 }

  $width = 0
  foreach ($ch in $Text.ToCharArray()) {
    if ([int][char]$ch -gt 127) {
      $width += 2
    } else {
      $width += 1
    }
  }
  return $width
}

function Get-SettlementPeriodText {
  param([string]$PeriodText)
  if ([string]::IsNullOrWhiteSpace($PeriodText)) { return '' }

  $trimmed = $PeriodText.Trim()
  if ($trimmed -match '\d{4}[.-]\d{1,2}[.-]\d{1,2}\s*-\s*\d{4}[.-]\d{1,2}[.-]\d{1,2}') {
    return $trimmed
  }

  if ($trimmed -match '^(\d{4})[.-](\d{1,2})[.-](\d{1,2})$') {
    $startDate = [datetime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3])
    $endDate = $startDate.AddDays(6)
    return "$($startDate.ToString('yyyy.MM.dd'))-$($endDate.ToString('yyyy.MM.dd'))"
  }

  return $trimmed
}

function ConvertTo-DecimalValue {
  param([string]$Text)
  $normalized = ([string]$Text).Trim().Replace(',', '').Replace('￥', '').Replace('¥', '')
  $num = 0
  if ([decimal]::TryParse($normalized, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
    return $num
  }
  if ([decimal]::TryParse($normalized, [ref]$num)) {
    return $num
  }
  return 0
}

function New-WorksheetXml {
  param(
    [Parameter(Mandatory = $true)]
    [array]$Rows
  )

  $columns = @($Rows[0].PSObject.Properties.Name)

  $columnWidths = @()
  for ($c = 0; $c -lt $columns.Count; $c++) {
    $maxLen = [Math]::Max(10, (Get-TextDisplayWidth -Text ([string]$columns[$c])))
    foreach ($row in $Rows) {
      $value = [string]$row.PSObject.Properties[$columns[$c]].Value
      foreach ($part in ($value -split "(`r`n|`n|`r)")) {
        $partWidth = Get-TextDisplayWidth -Text $part
        if ($partWidth -gt $maxLen) { $maxLen = $partWidth }
      }
    }
    $columnWidths += [Math]::Min(36, [Math]::Max(10, $maxLen + 2))
  }

  $sheet = New-Object Text.StringBuilder
  [void]$sheet.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sheet.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$sheet.AppendLine('  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
  [void]$sheet.AppendLine('  <cols>')
  for ($c = 0; $c -lt $columns.Count; $c++) {
    $idx = $c + 1
    [void]$sheet.AppendLine("    <col min=""$idx"" max=""$idx"" width=""$($columnWidths[$c])"" customWidth=""1""/>")
  }
  [void]$sheet.AppendLine('  </cols>')
  [void]$sheet.AppendLine('  <sheetData>')

  [void]$sheet.AppendLine('    <row r="1" ht="22" customHeight="1">')
  for ($c = 0; $c -lt $columns.Count; $c++) {
    $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))1"
    [void]$sheet.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""1""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $columns[$c])</t></is></c>")
  }
  [void]$sheet.AppendLine('    </row>')

  for ($r = 0; $r -lt $Rows.Count; $r++) {
    $rowIndex = $r + 2
    $maxLines = 1
    $cellXml = New-Object Text.StringBuilder
    for ($c = 0; $c -lt $columns.Count; $c++) {
      $value = [string]$Rows[$r].PSObject.Properties[$columns[$c]].Value
      $estimatedLines = 1
      foreach ($part in ($value -split "(`r`n|`n|`r)")) {
        $partWidth = Get-TextDisplayWidth -Text $part
        $partLines = [Math]::Max(1, [Math]::Ceiling($partWidth / [Math]::Max(1, $columnWidths[$c] - 1)))
        if ($partLines -gt $estimatedLines) { $estimatedLines = $partLines }
      }
      if ($estimatedLines -gt $maxLines) { $maxLines = $estimatedLines }
      $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))$rowIndex"
      [void]$cellXml.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""0""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $value)</t></is></c>")
    }
    $height = [Math]::Min(220, [Math]::Max(22, 18 * $maxLines))
    [void]$sheet.AppendLine("    <row r=""$rowIndex"" ht=""$height"" customHeight=""1"">")
    [void]$sheet.Append($cellXml.ToString())
    [void]$sheet.AppendLine('    </row>')
  }

  [void]$sheet.AppendLine('  </sheetData>')
  [void]$sheet.AppendLine('  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
  [void]$sheet.AppendLine('</worksheet>')
  return $sheet.ToString()
}

function New-SettlementAnalysisWorksheetXml {
  param(
    [Parameter(Mandatory = $true)]
    [array]$SettlementRows,
    [Parameter(Mandatory = $true)]
    [array]$PivotRows
  )

  $columns = @('结算周期', '商家代码', '商家名称', '总结算金额')
  $pivotColumns = @('商品编号', 'SKU ID（规格）', '单价', '订单数量汇总', '总结算金额汇总')

  $columnWidths = @()
  $maxColumnCount = [Math]::Max($columns.Count, $pivotColumns.Count)
  for ($c = 0; $c -lt $maxColumnCount; $c++) {
    $maxLen = 10

    if ($c -lt $columns.Count) {
      $summaryHeaderWidth = Get-TextDisplayWidth -Text ([string]$columns[$c])
      if ($summaryHeaderWidth -gt $maxLen) { $maxLen = $summaryHeaderWidth }
      foreach ($row in $SettlementRows) {
        $value = [string]$row.PSObject.Properties[$columns[$c]].Value
        $valueWidth = Get-TextDisplayWidth -Text $value
        if ($valueWidth -gt $maxLen) { $maxLen = $valueWidth }
      }
    }

    if ($c -lt $pivotColumns.Count) {
      $pivotHeaderWidth = Get-TextDisplayWidth -Text ([string]$pivotColumns[$c])
      if ($pivotHeaderWidth -gt $maxLen) { $maxLen = $pivotHeaderWidth }
      foreach ($row in $PivotRows) {
        $value = [string]$row.PSObject.Properties[$pivotColumns[$c]].Value
        $valueWidth = Get-TextDisplayWidth -Text $value
        if ($valueWidth -gt $maxLen) { $maxLen = $valueWidth }
      }
    }

    $columnWidths += [Math]::Min(80, [Math]::Max(12, $maxLen + 3))
  }

  $fixedRowHeight = 24

  $sheet = New-Object Text.StringBuilder
  [void]$sheet.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sheet.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$sheet.AppendLine('  <sheetViews><sheetView workbookViewId="0"/></sheetViews>')
  [void]$sheet.AppendLine('  <cols>')
  for ($c = 0; $c -lt $columnWidths.Count; $c++) {
    $idx = $c + 1
    [void]$sheet.AppendLine("    <col min=""$idx"" max=""$idx"" width=""$($columnWidths[$c])"" customWidth=""1""/>")
  }
  [void]$sheet.AppendLine('  </cols>')
  [void]$sheet.AppendLine('  <sheetData>')

  [void]$sheet.AppendLine("    <row r=""1"" ht=""$fixedRowHeight"" customHeight=""1"">")
  [void]$sheet.AppendLine('      <c r="A1" t="inlineStr" s="1"><is><t xml:space="preserve">结算汇总</t></is></c>')
  [void]$sheet.AppendLine('    </row>')

  [void]$sheet.AppendLine("    <row r=""2"" ht=""$fixedRowHeight"" customHeight=""1"">")
  for ($c = 0; $c -lt $columns.Count; $c++) {
    $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))2"
    [void]$sheet.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""1""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $columns[$c])</t></is></c>")
  }
  [void]$sheet.AppendLine('    </row>')

  for ($r = 0; $r -lt $SettlementRows.Count; $r++) {
    $rowIndex = $r + 3
    [void]$sheet.AppendLine("    <row r=""$rowIndex"" ht=""$fixedRowHeight"" customHeight=""1"">")
    for ($c = 0; $c -lt $columns.Count; $c++) {
      $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))$rowIndex"
      $value = [string]$SettlementRows[$r].PSObject.Properties[$columns[$c]].Value
      [void]$sheet.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""0""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $value)</t></is></c>")
    }
    [void]$sheet.AppendLine('    </row>')
  }

  $pivotTitleRow = $SettlementRows.Count + 5
  [void]$sheet.AppendLine("    <row r=""$pivotTitleRow"" ht=""$fixedRowHeight"" customHeight=""1"">")
  [void]$sheet.AppendLine("      <c r=""A$pivotTitleRow"" t=""inlineStr"" s=""1""><is><t xml:space=""preserve"">商品数据透视</t></is></c>")
  [void]$sheet.AppendLine('    </row>')

  $pivotHeaderRow = $pivotTitleRow + 1
  [void]$sheet.AppendLine("    <row r=""$pivotHeaderRow"" ht=""$fixedRowHeight"" customHeight=""1"">")
  for ($c = 0; $c -lt $pivotColumns.Count; $c++) {
    $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))$pivotHeaderRow"
    [void]$sheet.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""1""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $pivotColumns[$c])</t></is></c>")
  }
  [void]$sheet.AppendLine('    </row>')

  for ($r = 0; $r -lt $PivotRows.Count; $r++) {
    $rowIndex = $pivotHeaderRow + $r + 1
    [void]$sheet.AppendLine("    <row r=""$rowIndex"" ht=""$fixedRowHeight"" customHeight=""1"">")
    for ($c = 0; $c -lt $pivotColumns.Count; $c++) {
      $cellRef = "$(ConvertTo-ExcelColumnName ($c + 1))$rowIndex"
      $value = [string]$PivotRows[$r].PSObject.Properties[$pivotColumns[$c]].Value
      [void]$sheet.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""0""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $value)</t></is></c>")
    }
    [void]$sheet.AppendLine('    </row>')
  }

  [void]$sheet.AppendLine('  </sheetData>')
  [void]$sheet.AppendLine('  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
  [void]$sheet.AppendLine('</worksheet>')
  return $sheet.ToString()
}

function Write-FormattedXlsx {
  param(
    [Parameter(Mandatory = $true)]
    [array]$Rows,
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$SheetName = "对账单",
    [array]$AdditionalSheets = @()
  )

  if ($Rows.Count -eq 0) { return $false }

  $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("settlement_xlsx_" + [guid]::NewGuid().ToString('N'))
  try {
    New-Item -ItemType Directory -Force -Path $tempRoot, (Join-Path $tempRoot '_rels'), (Join-Path $tempRoot 'xl'), (Join-Path $tempRoot 'xl\_rels'), (Join-Path $tempRoot 'xl\worksheets') | Out-Null

    $sheetSpecs = @([PSCustomObject]@{ Name = $SheetName; Rows = $Rows; WorksheetXml = $null })
    foreach ($sheetSpec in $AdditionalSheets) {
      if ($sheetSpec.WorksheetXml) {
        $sheetSpecs += [PSCustomObject]@{ Name = [string]$sheetSpec.Name; Rows = $null; WorksheetXml = [string]$sheetSpec.WorksheetXml }
      } elseif ($sheetSpec.Rows -and $sheetSpec.Rows.Count -gt 0) {
        $sheetSpecs += [PSCustomObject]@{ Name = [string]$sheetSpec.Name; Rows = $sheetSpec.Rows; WorksheetXml = $null }
      }
    }

    $sheetContentTypes = New-Object Text.StringBuilder
    $workbookSheets = New-Object Text.StringBuilder
    $workbookRelationships = New-Object Text.StringBuilder
    for ($i = 0; $i -lt $sheetSpecs.Count; $i++) {
      $sheetNo = $i + 1
      [void]$sheetContentTypes.AppendLine("  <Override PartName=""/xl/worksheets/sheet$sheetNo.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
      [void]$workbookSheets.AppendLine("    <sheet name=""$(ConvertTo-XmlText $sheetSpecs[$i].Name)"" sheetId=""$sheetNo"" r:id=""rId$sheetNo""/>")
      [void]$workbookRelationships.AppendLine("  <Relationship Id=""rId$sheetNo"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sheetNo.xml""/>")
    }
    $stylesRelationshipId = "rId$($sheetSpecs.Count + 1)"

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
SHEET_CONTENT_TYPES
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@.Replace('SHEET_CONTENT_TYPES', $sheetContentTypes.ToString().TrimEnd())

    $packageRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    $workbookXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
WORKBOOK_SHEETS
  </sheets>
</workbook>
'@.Replace('WORKBOOK_SHEETS', $workbookSheets.ToString().TrimEnd())

    $workbookRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
WORKBOOK_RELATIONSHIPS
  <Relationship Id="STYLES_RELATIONSHIP_ID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@.Replace('WORKBOOK_RELATIONSHIPS', $workbookRelationships.ToString().TrimEnd()).Replace('STYLES_RELATIONSHIP_ID', $stylesRelationshipId)

    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Microsoft YaHei UI"/></font>
    <font><b/><sz val="11"/><name val="Microsoft YaHei UI"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

    [IO.File]::WriteAllText((Join-Path $tempRoot '[Content_Types].xml'), $contentTypes, [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText((Join-Path $tempRoot '_rels\.rels'), $packageRels, [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText((Join-Path $tempRoot 'xl\workbook.xml'), $workbookXml, [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText((Join-Path $tempRoot 'xl\_rels\workbook.xml.rels'), $workbookRels, [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText((Join-Path $tempRoot 'xl\styles.xml'), $stylesXml, [Text.Encoding]::UTF8)
    for ($i = 0; $i -lt $sheetSpecs.Count; $i++) {
      $sheetNo = $i + 1
      if ($sheetSpecs[$i].WorksheetXml) {
        $worksheetXml = [string]$sheetSpecs[$i].WorksheetXml
      } else {
        $worksheetXml = New-WorksheetXml -Rows $sheetSpecs[$i].Rows
      }
      [IO.File]::WriteAllText((Join-Path $tempRoot "xl\worksheets\sheet$sheetNo.xml"), $worksheetXml, [Text.Encoding]::UTF8)
    }

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    [IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $Path)
    return $true
  } catch {
    Write-Warning "无法生成格式化 Excel 文件：$Path。原因：$($_.Exception.Message)"
    return $false
  } finally {
    if (Test-Path -LiteralPath $tempRoot) {
      Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
  }
}

function Import-CsvFlexible {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @() }

  $resolvedPath = (Resolve-Path -LiteralPath $Path)
  $bytes = [IO.File]::ReadAllBytes($resolvedPath)
  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $text = [Text.UTF8Encoding]::new($true).GetString($bytes)
  } else {
    $utf8Text = [Text.UTF8Encoding]::new($false, $false).GetString($bytes)
    if ($utf8Text.Contains([char]0xFFFD)) {
      $text = [Text.Encoding]::GetEncoding(936).GetString($bytes)
    } else {
      $text = $utf8Text
    }
  }

  return $text | ConvertFrom-Csv
}

function Add-RowToGroup {
  param(
    [hashtable]$Groups,
    [string]$Supplier,
    [object]$Row
  )
  if (-not $Groups.ContainsKey($Supplier)) { $Groups[$Supplier] = @() }
  $Groups[$Supplier] += $Row
}

if (-not (Test-Path -LiteralPath $InputFile)) {
  throw "找不到输入文件：$InputFile"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$zip = [System.IO.Compression.ZipFile]::OpenRead($InputFile)
try {
  [xml]$sharedXml = Get-ZipText $zip 'xl/sharedStrings.xml'
  $sharedStrings = @()
  if ($sharedXml) {
    foreach ($si in $sharedXml.sst.si) {
      if ($si.t) {
        $sharedStrings += [string]$si.t
      } elseif ($si.r) {
        $sharedStrings += (($si.r | ForEach-Object { $_.t }) -join '')
      } else {
        $sharedStrings += ''
      }
    }
  }

  [xml]$sheet = Get-ZipText $zip 'xl/worksheets/sheet1.xml'
  if (-not $sheet) { throw "未找到第一个工作表：xl/worksheets/sheet1.xml" }

  $tableRows = @()
  foreach ($row in $sheet.worksheet.sheetData.row) {
    $values = @{}
    foreach ($cell in $row.c) {
      $idx = Convert-ColumnNameToIndex $cell.r
      $value = [string]$cell.v
      if ($cell.t -eq 's' -and $value -ne '') {
        $value = $sharedStrings[[int]$value]
      } elseif ($cell.t -eq 'inlineStr') {
        $value = [string]$cell.is.t
      }
      $values[$idx] = $value
    }
    if ($values.Keys.Count -gt 0) {
      $max = ($values.Keys | Measure-Object -Maximum).Maximum
      $line = for ($i = 1; $i -le $max; $i++) {
        if ($values.ContainsKey($i)) { $values[$i] } else { '' }
      }
      $tableRows += ,$line
    }
  }
} finally {
  $zip.Dispose()
}

if ($tableRows.Count -lt 2) {
  throw "表格没有可拆分的数据。"
}

$headers = $tableRows[0]
$supplierIndex = [array]::IndexOf($headers, $SupplierColumn)
$amountIndex = [array]::IndexOf($headers, $AmountColumn)
$productIndex = [array]::IndexOf($headers, '商品编号')
$skuIndex = [array]::IndexOf($headers, 'SKU ID（规格）')
$quantityIndex = [array]::IndexOf($headers, '订单数量')
$unitPriceIndex = [array]::IndexOf($headers, '进货单价')

if ($supplierIndex -lt 0) {
  throw "找不到供应商字段：$SupplierColumn。当前表头：$($headers -join ', ')"
}

if ($amountIndex -lt 0) {
  foreach ($candidate in @('预计结算金额', '预计结账金额')) {
    $candidateIndex = [array]::IndexOf($headers, $candidate)
    if ($candidateIndex -ge 0) {
      $AmountColumn = $candidate
      $amountIndex = $candidateIndex
      break
    }
  }
}

if ($amountIndex -lt 0) {
  Write-Warning "找不到金额字段：$AmountColumn / 预计结算金额 / 预计结账金额，将只拆分不汇总金额。"
}

if ($productIndex -lt 0 -or $skuIndex -lt 0 -or $quantityIndex -lt 0) {
  Write-Warning "找不到商品透视所需字段：商品编号、SKU ID（规格）、订单数量。商家对账单将不生成商品数据透视页。"
}

if ($unitPriceIndex -lt 0) {
  Write-Warning "找不到单价字段：进货单价。商品数据透视页的单价列将留空。"
}

$supplierNameMap = @{}
if (-not [string]::IsNullOrWhiteSpace($SupplierMapFile) -and (Test-Path -LiteralPath $SupplierMapFile)) {
  foreach ($row in (Import-CsvFlexible -Path $SupplierMapFile)) {
    $code = [string]$row.供应商代号
    $name = [string]$row.商家名称
    if (-not [string]::IsNullOrWhiteSpace($code)) {
      $supplierNameMap[$code.Trim()] = $name.Trim()
    }
  }
}

$groups = @{}
$skippedRows = @()
for ($i = 1; $i -lt $tableRows.Count; $i++) {
  $line = $tableRows[$i]
  $supplier = if ($supplierIndex -lt $line.Count) { $line[$supplierIndex] } else { '' }
  $obj = [ordered]@{}
  for ($c = 0; $c -lt $headers.Count; $c++) {
    $key = $headers[$c]
    if ([string]::IsNullOrWhiteSpace($key)) { $key = "字段$($c + 1)" }
    $obj[$key] = if ($c -lt $line.Count) { $line[$c] } else { '' }
  }
  $obj['补充类型'] = '后台订单'
  $obj['补充说明'] = ''
  $rowObject = [PSCustomObject]$obj

  if ([string]::IsNullOrWhiteSpace($supplier)) {
    $skippedRows += $rowObject
    continue
  }

  Add-RowToGroup -Groups $groups -Supplier $supplier -Row $rowObject
}

$supplementIncludedCount = 0
$supplementSkippedCount = 0
if (-not [string]::IsNullOrWhiteSpace($SupplementFile) -and (Test-Path -LiteralPath $SupplementFile)) {
  foreach ($supplement in (Import-CsvFlexible -Path $SupplementFile)) {
    $include = ([string]$supplement.是否计入本周结算).Trim()
    if ($include -in @('否', '不计入', 'no', 'NO', '0')) {
      $supplementSkippedCount += 1
      continue
    }

    $supplier = ([string]$supplement.供应商代号).Trim()
    $obj = [ordered]@{}
    foreach ($header in $headers) {
      if ([string]::IsNullOrWhiteSpace($header)) { continue }
      if ($supplement.PSObject.Properties.Name -contains $header) {
        $obj[$header] = [string]$supplement.PSObject.Properties[$header].Value
      } else {
        $obj[$header] = ''
      }
    }

    if ([string]::IsNullOrWhiteSpace($obj[$SupplierColumn])) { $obj[$SupplierColumn] = $supplier }
    if ([string]::IsNullOrWhiteSpace($obj[$AmountColumn]) -and ($supplement.PSObject.Properties.Name -contains $AmountColumn)) {
      $obj[$AmountColumn] = [string]$supplement.PSObject.Properties[$AmountColumn].Value
    }

    $supplementType = ([string]$supplement.补充类型).Trim()
    $reason = ([string]$supplement.补充原因).Trim()
    $remark = ([string]$supplement.备注).Trim()
    $obj['补充类型'] = if ([string]::IsNullOrWhiteSpace($supplementType)) { '手工补充' } else { $supplementType }
    $obj['补充说明'] = (($reason, $remark) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '；'

    $rowObject = [PSCustomObject]$obj
    if ([string]::IsNullOrWhiteSpace($supplier)) {
      $skippedRows += $rowObject
      $supplementSkippedCount += 1
      continue
    }

    Add-RowToGroup -Groups $groups -Supplier $supplier -Row $rowObject
    $supplementIncludedCount += 1
  }
}

$splitDir = Join-Path $OutputDir '02_商家拆分对账单'
$splitXlsxDir = Join-Path $splitDir 'xlsx格式'
$splitCsvDir = Join-Path $splitDir 'csv格式'
$summaryDir = Join-Path $OutputDir '05_财务汇总'
$trackingDir = Join-Path $OutputDir '03_商家确认记录'
$exceptionDir = Join-Path $OutputDir '99_异常待处理'
New-Item -ItemType Directory -Force -Path $splitDir, $splitXlsxDir, $splitCsvDir, $summaryDir, $trackingDir, $exceptionDir | Out-Null

Get-ChildItem -LiteralPath $splitDir -File -Filter '*.xlsx' | ForEach-Object {
  Move-Item -LiteralPath $_.FullName -Destination (Join-Path $splitXlsxDir $_.Name) -Force
}
Get-ChildItem -LiteralPath $splitDir -File -Filter '*.csv' | ForEach-Object {
  Move-Item -LiteralPath $_.FullName -Destination (Join-Path $splitCsvDir $_.Name) -Force
}

$summaryRows = @()
$trackingRows = @()
$settlementPeriod = Get-SettlementPeriodText -PeriodText $Period

foreach ($supplier in ($groups.Keys | Sort-Object)) {
  $safeSupplier = ConvertTo-SafeFileName $supplier
  $supplierName = if ($supplierNameMap.ContainsKey($supplier)) { $supplierNameMap[$supplier] } else { '' }
  $prefix = if ([string]::IsNullOrWhiteSpace($settlementPeriod)) { '' } else { "${settlementPeriod}_" }
  $csvFileName = "${prefix}${safeSupplier}_对账单.csv"
  $xlsxFileName = "${prefix}${safeSupplier}_对账单.xlsx"
  $csvFilePath = Join-Path $splitCsvDir $csvFileName
  $xlsxFilePath = Join-Path $splitXlsxDir $xlsxFileName
  $rows = $groups[$supplier]
  Write-CsvUtf8Bom -Rows $rows -Path $csvFilePath

  $totalAmount = 0
  if ($amountIndex -ge 0) {
    foreach ($row in $rows) {
      $raw = [string]$row.$AmountColumn
      $totalAmount += ConvertTo-DecimalValue -Text $raw
    }
  }

  $settlementRows = @(
    [PSCustomObject][ordered]@{
      结算周期 = $settlementPeriod
      商家代码 = $supplier
      商家名称 = $supplierName
      总结算金额 = $totalAmount
    }
  )

  $pivotRows = @()
  if ($productIndex -ge 0 -and $skuIndex -ge 0 -and $quantityIndex -ge 0) {
    $pivot = @{}
    foreach ($row in $rows) {
      $productId = ([string]$row.'商品编号').Trim()
      $skuId = ([string]$row.'SKU ID（规格）').Trim()
      $key = "$productId`t$skuId"
      if (-not $pivot.ContainsKey($key)) {
        $pivot[$key] = [PSCustomObject][ordered]@{
          商品编号 = $productId
          'SKU ID（规格）' = $skuId
          单价 = ''
          订单数量汇总 = [decimal]0
          总结算金额汇总 = [decimal]0
        }
      }
      if ($unitPriceIndex -ge 0 -and [string]::IsNullOrWhiteSpace([string]$pivot[$key].单价)) {
        $pivot[$key].单价 = [string]$row.'进货单价'
      }
      $pivot[$key].订单数量汇总 += ConvertTo-DecimalValue -Text ([string]$row.'订单数量')
      if ($amountIndex -ge 0) {
        $pivot[$key].总结算金额汇总 += ConvertTo-DecimalValue -Text ([string]$row.$AmountColumn)
      }
    }
    $pivotRows = @($pivot.Values | Sort-Object 商品编号, 'SKU ID（规格）')
  }

  $analysisSheetXml = New-SettlementAnalysisWorksheetXml -SettlementRows $settlementRows -PivotRows $pivotRows
  $extraSheets = @(
    [PSCustomObject]@{ Name = '结算与商品汇总'; WorksheetXml = $analysisSheetXml }
  )

  $hasFormattedXlsx = Write-FormattedXlsx -Rows $rows -Path $xlsxFilePath -AdditionalSheets $extraSheets
  $fileName = if ($hasFormattedXlsx) { "xlsx格式\$xlsxFileName" } else { "csv格式\$csvFileName" }

  $summaryRows += [PSCustomObject][ordered]@{
    结算周期 = $settlementPeriod
    商家代码 = $supplier
    商家名称 = $supplierName
    总结算金额 = $totalAmount
    对账确认状态 = '待确认'
    发票状态 = '未收票'
    付款状态 = '待付款'
    备注 = ''
  }

  $trackingRows += [PSCustomObject][ordered]@{
    结算周期 = $settlementPeriod
    商家代码 = $supplier
    商家名称 = $supplierName
    是否已发送 = '否'
    是否已确认 = '否'
    异常类型 = ''
    异常说明 = ''
    是否涉及漏单 = '否'
    是否已补传后台 = '否'
    备注 = ''
  }
}

$summaryPath = Join-Path $summaryDir "${settlementPeriod}_财务结算汇总表.csv"
$trackingPath = Join-Path $trackingDir "${settlementPeriod}_商家对账跟进台账.csv"
Write-CsvUtf8Bom -Rows $summaryRows -Path $summaryPath
Write-CsvUtf8Bom -Rows $trackingRows -Path $trackingPath
Write-FormattedXlsx -Rows $summaryRows -Path (Join-Path $summaryDir "${settlementPeriod}_财务结算汇总表.xlsx") | Out-Null
Write-FormattedXlsx -Rows $trackingRows -Path (Join-Path $trackingDir "${settlementPeriod}_商家对账跟进台账.xlsx") | Out-Null

if ($skippedRows.Count -gt 0) {
  $skippedPath = Join-Path $exceptionDir "${settlementPeriod}_未拆分空供应商或统计行.csv"
  Write-CsvUtf8Bom -Rows $skippedRows -Path $skippedPath
  Write-FormattedXlsx -Rows $skippedRows -Path (Join-Path $exceptionDir "${settlementPeriod}_未拆分空供应商或统计行.xlsx") | Out-Null
}

Write-Host "拆分完成。"
Write-Host "供应商数量：$($groups.Keys.Count)"
Write-Host "订单行数：$(($groups.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum)"
Write-Host "跳过空供应商或统计行：$($skippedRows.Count)"
Write-Host "计入本周的手工补充行：$supplementIncludedCount"
Write-Host "未计入或跳过的手工补充行：$supplementSkippedCount"
Write-Host "商家对账单目录：$splitDir"
Write-Host "格式化 Excel 目录：$splitXlsxDir"
Write-Host "原始 CSV 目录：$splitCsvDir"
Write-Host "财务汇总表：$summaryPath"
Write-Host "跟进台账：$trackingPath"
