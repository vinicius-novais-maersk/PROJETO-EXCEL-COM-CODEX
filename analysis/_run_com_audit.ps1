$ErrorActionPreference='Continue'
$path = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\_tmp_WK19_goal_com_audit.xlsm'
$out = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_goal_com_audit.md'
$jsonOut = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_goal_com_audit.json'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$result = [ordered]@{ workbook=$path; timestamp=(Get-Date).ToString('s'); sheets=@(); pivots=@(); charts=@(); errors=@() }
try {
  $wb = $excel.Workbooks.Open($path, 0, $true, 5, '', '', $true, 1, '', $false, $false, 0, $false, $false, 0)
  try {
    $result.name = $wb.Name
    $result.readOnly = $wb.ReadOnly
    $result.worksheets = $wb.Worksheets.Count
    foreach ($ws in @($wb.Worksheets)) {
      $sheetInfo = [ordered]@{ name=$ws.Name; usedRange=$ws.UsedRange.Address($false,$false); formulaErrorCount=0; formulaErrorSamples=@(); pivotCount=0; chartCount=0 }
      try {
        $errRange = $ws.UsedRange.SpecialCells(-4123, 16)
        $cnt = 0
        foreach ($area in @($errRange.Areas)) { $cnt += $area.Cells.Count }
        $sheetInfo.formulaErrorCount = $cnt
        $sample = @()
        foreach ($cell in @($errRange.Cells)) {
          if ($sample.Count -ge 15) { break }
          $sample += [ordered]@{ cell=$cell.Address($false,$false); text=[string]$cell.Text; formula=[string]$cell.Formula }
        }
        $sheetInfo.formulaErrorSamples = $sample
      } catch {}
      try { $sheetInfo.pivotCount = $ws.PivotTables().Count } catch {}
      try { $sheetInfo.chartCount = $ws.ChartObjects().Count } catch {}
      $result.sheets += $sheetInfo
      try {
        foreach ($pt in @($ws.PivotTables())) {
          $pages=@(); $rows=@(); $cols=@(); $datas=@()
          try { foreach ($pf in @($pt.PageFields)) { $pages += ([string]$pf.Name + '=' + [string]$pf.CurrentPage) } } catch {}
          try { foreach ($pf in @($pt.RowFields)) { $rows += [string]$pf.Name } } catch {}
          try { foreach ($pf in @($pt.ColumnFields)) { $cols += [string]$pf.Name } } catch {}
          try { foreach ($df in @($pt.DataFields)) { $datas += ([string]$df.Name + ' / ' + [string]$df.SourceName + ' / func=' + [string]$df.Function) } } catch {}
          $result.pivots += [ordered]@{ sheet=$ws.Name; name=$pt.Name; tableRange=$pt.TableRange2.Address($false,$false); sourceData=[string]$pt.SourceData; pageFields=$pages; rowFields=$rows; columnFields=$cols; dataFields=$datas; refreshDate=[string]$pt.RefreshDate; version=[string]$pt.Version }
        }
      } catch {}
      try {
        foreach ($co in @($ws.ChartObjects())) {
          $chart = $co.Chart
          $sers=@(); $issues=@()
          try {
            foreach ($ser in @($chart.SeriesCollection())) {
              $formula = [string]$ser.Formula
              if ($formula -match '#REF!') { $issues += 'series_formula_ref' }
              $sers += [ordered]@{ name=[string]$ser.Name; formula=$formula }
            }
          } catch { $issues += 'series_read_failed' }
          $result.charts += [ordered]@{ sheet=$ws.Name; name=$co.Name; chartType=[string]$chart.ChartType; seriesCount=$sers.Count; series=$sers; issues=$issues }
        }
      } catch {}
    }
    $result.summary = [ordered]@{
      sheets=$result.sheets.Count
      formulaErrorSheets=($result.sheets | Where-Object { $_.formulaErrorCount -gt 0 }).Count
      formulaErrorCount=(($result.sheets | Measure-Object -Property formulaErrorCount -Sum).Sum)
      pivots=$result.pivots.Count
      charts=$result.charts.Count
      chartIssues=($result.charts | Where-Object { $_.issues.Count -gt 0 }).Count
    }
  } finally { $wb.Close($false) }
} catch {
  $result.openError = $_.Exception.Message
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonOut -Encoding UTF8
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Excel COM audit - Base_DSU2026 - TbM - WK19.xlsm')
$lines.Add('')
$lines.Add(('- Audited copy: `{0}`' -f $path))
$lines.Add(('- Timestamp: `{0}`' -f $result.timestamp))
if ($result.openError) { $lines.Add(('- Open error: `{0}`' -f $result.openError)) }
$lines.Add('')
$lines.Add('## Summary')
$lines.Add('')
$lines.Add('| Check | Result |')
$lines.Add('|---|---:|')
foreach ($k in $result.summary.Keys) { $lines.Add(('| {0} | {1} |' -f $k, $result.summary[$k])) }
$lines.Add('')
$lines.Add('## Formula errors by sheet')
$lines.Add('')
$lines.Add('| Sheet | Used range | Formula error count | Samples |')
$lines.Add('|---|---|---:|---|')
foreach ($s in $result.sheets) {
  $samples = ($s.formulaErrorSamples | ForEach-Object { ('{0}={1}' -f $_.cell, $_.text) }) -join '; '
  $lines.Add(('| {0} | {1} | {2} | {3} |' -f $s.name, $s.usedRange, $s.formulaErrorCount, $samples))
}
$lines.Add('')
$lines.Add('## Pivot tables')
$lines.Add('')
$lines.Add('| Sheet | Pivot | Range | SourceData | Page fields | Row fields | Column fields | Data fields |')
$lines.Add('|---|---|---|---|---|---|---|---|')
foreach ($p in $result.pivots) { $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |' -f $p.sheet, $p.name, $p.tableRange, $p.sourceData, (($p.pageFields) -join ', '), (($p.rowFields) -join ', '), (($p.columnFields) -join ', '), (($p.dataFields) -join ', '))) }
$lines.Add('')
$lines.Add('## Charts')
$lines.Add('')
$lines.Add('| Sheet | Chart | Series | Issues |')
$lines.Add('|---|---|---:|---|')
foreach ($c in $result.charts) { $lines.Add(('| {0} | {1} | {2} | {3} |' -f $c.sheet, $c.name, $c.seriesCount, (($c.issues) -join ', '))) }
$lines | Set-Content -LiteralPath $out -Encoding UTF8
Write-Output $out
Write-Output ($result.summary | ConvertTo-Json -Compress)
