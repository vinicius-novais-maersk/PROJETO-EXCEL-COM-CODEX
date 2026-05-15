$ErrorActionPreference='Continue'
$path = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\_tmp_WK19_goal_com_audit.xlsm'
$out = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_goal_com_objects_audit.md'
$jsonOut = 'C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_goal_com_objects_audit.json'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$result = [ordered]@{ workbook=$path; timestamp=(Get-Date).ToString('s'); pivots=@(); charts=@(); sheetObjects=@() }
try {
  $wb = $excel.Workbooks.Open($path, 0, $true, 5, '', '', $true, 1, '', $false, $false, 0, $false, $false, 0)
  try {
    $result.name = $wb.Name; $result.readOnly = $wb.ReadOnly; $result.worksheets = $wb.Worksheets.Count
    foreach ($ws in @($wb.Worksheets)) {
      $pivotCount=0; $chartCount=0
      try { $pivotCount = $ws.PivotTables().Count } catch {}
      try { $chartCount = $ws.ChartObjects().Count } catch {}
      $result.sheetObjects += [ordered]@{ sheet=$ws.Name; pivots=$pivotCount; charts=$chartCount }
      if ($pivotCount -gt 0) {
        foreach ($pt in @($ws.PivotTables())) {
          $pages=@(); $rows=@(); $cols=@(); $datas=@(); $filters=@()
          try { foreach ($pf in @($pt.PageFields)) { $pages += ([string]$pf.Name + '=' + [string]$pf.CurrentPage) } } catch {}
          try { foreach ($pf in @($pt.RowFields)) { $rows += [string]$pf.Name } } catch {}
          try { foreach ($pf in @($pt.ColumnFields)) { $cols += [string]$pf.Name } } catch {}
          try { foreach ($df in @($pt.DataFields)) { $datas += ([string]$df.Name + ' / ' + [string]$df.SourceName + ' / func=' + [string]$df.Function) } } catch {}
          $result.pivots += [ordered]@{ sheet=$ws.Name; name=$pt.Name; tableRange=$pt.TableRange2.Address($false,$false); sourceData=[string]$pt.SourceData; pageFields=$pages; rowFields=$rows; columnFields=$cols; dataFields=$datas }
        }
      }
      if ($chartCount -gt 0) {
        foreach ($co in @($ws.ChartObjects())) {
          $chart = $co.Chart; $issues=@(); $seriesFormulas=@()
          try {
            $sc = $chart.SeriesCollection()
            foreach ($ser in @($sc)) {
              $formula = [string]$ser.Formula
              if ($formula -match '#REF!') { $issues += 'series_formula_ref' }
              $seriesFormulas += $formula
            }
          } catch { $issues += 'series_read_failed' }
          $result.charts += [ordered]@{ sheet=$ws.Name; name=$co.Name; chartType=[string]$chart.ChartType; seriesCount=$seriesFormulas.Count; issues=$issues; formulas=$seriesFormulas }
        }
      }
    }
    $result.summary = [ordered]@{ sheets=$result.sheetObjects.Count; pivots=$result.pivots.Count; charts=$result.charts.Count; chartIssues=($result.charts | Where-Object { $_.issues.Count -gt 0 }).Count }
  } finally { $wb.Close($false) }
} catch { $result.openError = $_.Exception.Message }
finally { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonOut -Encoding UTF8
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Excel COM object audit - Base_DSU2026 - TbM - WK19.xlsm'); $lines.Add('')
$lines.Add(('- Audited copy: `{0}`' -f $path)); $lines.Add(('- Timestamp: `{0}`' -f $result.timestamp)); $lines.Add('')
if ($result.openError) { $lines.Add(('- Open error: `{0}`' -f $result.openError)); $lines.Add('') }
$lines.Add('## Summary'); $lines.Add(''); $lines.Add('| Check | Result |'); $lines.Add('|---|---:|')
foreach ($k in $result.summary.Keys) { $lines.Add(('| {0} | {1} |' -f $k, $result.summary[$k])) }
$lines.Add(''); $lines.Add('## Pivot tables'); $lines.Add(''); $lines.Add('| Sheet | Pivot | Range | SourceData | Page fields | Row fields | Column fields | Data fields |'); $lines.Add('|---|---|---|---|---|---|---|---|')
foreach ($p in $result.pivots) { $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |' -f $p.sheet, $p.name, $p.tableRange, $p.sourceData, (($p.pageFields) -join ', '), (($p.rowFields) -join ', '), (($p.columnFields) -join ', '), (($p.dataFields) -join ', '))) }
$lines.Add(''); $lines.Add('## Charts'); $lines.Add(''); $lines.Add('| Sheet | Chart | Series | Issues |'); $lines.Add('|---|---|---:|---|')
foreach ($c in $result.charts) { $lines.Add(('| {0} | {1} | {2} | {3} |' -f $c.sheet, $c.name, $c.seriesCount, (($c.issues) -join ', '))) }
$lines | Set-Content -LiteralPath $out -Encoding UTF8
Write-Output $out
Write-Output ($result.summary | ConvertTo-Json -Compress)
