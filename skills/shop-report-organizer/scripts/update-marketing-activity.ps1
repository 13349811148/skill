param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$PlanPath,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [Parameter(Mandatory = $true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$sheetName = [string]$plan.schema.sheet_name
$dataStartRow = [int]$plan.schema.data_start_row
$maxColumn = [int]$plan.schema.max_column
$changes = [System.Collections.Generic.List[object]]::new()
$issues = [System.Collections.Generic.List[object]]::new()
$excel = $null

Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $workbook = $excel.Workbooks.Open($OutputPath, 0, $false)
  try {
    $worksheet = $workbook.Worksheets.Item($sheetName)
    foreach ($write in @($plan.writes)) {
      $rowNumber = [int]$write.row
      if ([bool]$write.new_row -and [int]$write.template_row -ge $dataStartRow) {
        $templateRow = [int]$write.template_row
        $sourceFormat = $worksheet.Range(
          $worksheet.Cells.Item($templateRow, 1),
          $worksheet.Cells.Item($templateRow, $maxColumn)
        )
        $targetFormat = $worksheet.Range(
          $worksheet.Cells.Item($rowNumber, 1),
          $worksheet.Cells.Item($rowNumber, $maxColumn)
        )
        $sourceFormat.Copy()
        $targetFormat.PasteSpecial(-4122) | Out-Null
        $excel.CutCopyMode = 0
      }

      foreach ($entry in @($write.values)) {
        $column = [int]$entry.column
        $cell = $worksheet.Cells.Item($rowNumber, $column)
        $current = $cell.Value2
        if ($null -eq $current -or ([string]$current).Trim().Length -eq 0) {
          if ([bool]$entry.as_text) {
            $cell.NumberFormat = '@'
            $cell.Value2 = [string]$entry.value
          } elseif ([bool]$entry.as_number) {
            $cell.Value2 = [double]$entry.value
          } else {
            $cell.Value2 = [string]$entry.value
          }
          $changes.Add([PSCustomObject]@{
            product_id = [string]$write.product_id
            style_id = [string]$write.style_id
            sku = [string]$write.sku
            row = $rowNumber
            new_row = [bool]$write.new_row
            field = [string]$entry.field
            kind = 'value'
          })
        } elseif ([Convert]::ToString($current) -ne [Convert]::ToString($entry.value)) {
          $issues.Add([PSCustomObject]@{
            product_id = [string]$write.product_id
            style_id = [string]$write.style_id
            sku = [string]$write.sku
            row = $rowNumber
            field = [string]$entry.field
            message = "The cell already had a value at write time; the original value was preserved."
          })
        }
      }

      foreach ($formula in @($plan.formula_columns)) {
        $column = [int]$formula.column
        $cell = $worksheet.Cells.Item($rowNumber, $column)
        if ([bool]$cell.HasFormula -or ($null -ne $cell.Value2 -and ([string]$cell.Value2).Trim().Length -gt 0)) {
          continue
        }
        $formulaR1C1 = $null
        for ($sourceRow = $rowNumber - 1; $sourceRow -ge $dataStartRow; $sourceRow--) {
          $sourceCell = $worksheet.Cells.Item($sourceRow, $column)
          if ([bool]$sourceCell.HasFormula) {
            $formulaR1C1 = [string]$sourceCell.FormulaR1C1
            break
          }
        }
        if (($null -eq $formulaR1C1 -or $formulaR1C1.Length -eq 0) -and
            $null -ne $formula.formula_r1c1) {
          $formulaR1C1 = [string]$formula.formula_r1c1
        }
        if ($null -ne $formulaR1C1 -and $formulaR1C1.Length -gt 0) {
          $cell.FormulaR1C1 = $formulaR1C1
          $changes.Add([PSCustomObject]@{
            product_id = [string]$write.product_id
            style_id = [string]$write.style_id
            sku = [string]$write.sku
            row = $rowNumber
            new_row = [bool]$write.new_row
            field = [string]$formula.field
            kind = 'formula'
          })
        }
      }
    }
    $workbook.Save() | Out-Null
  } finally {
    $workbook.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  }
} finally {
  if ($null -ne $excel) {
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
}

$resultJson = [PSCustomObject]@{
  status = if ($changes.Count -gt 0) { 'updated' } else { 'no_change' }
  changes = @($changes)
  issues = @($issues)
} | ConvertTo-Json -Compress -Depth 8
[System.IO.File]::WriteAllText(
  $ResultPath,
  $resultJson,
  [System.Text.UTF8Encoding]::new($false)
)
