$ErrorActionPreference='Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false
$excel.DisplayAlerts=$false
try {
  $wb=$excel.Workbooks.Add()
  $formula='let W=Excel.Workbook(File.Contents("E:\New folder\Economia\8vo\gore\Cash data request\ref docs\reporte20261006085201.xls"),null,true), T=W{[Name="resumen_marco"]}[Data], R=Table.FromRows(List.Transform(Table.ToRows(Table.FirstN(T,35)),each {Text.Combine(List.Transform(List.Positions(_),(i)=>Text.From(i)&":"&(if _{i}=null then "" else Text.From(_{i})))," | ")} ),{"Row"}) in R'
  $null=$wb.Queries.Add('Probe',$formula)
  $ws=$wb.Worksheets.Item(1)
  $lo=$ws.ListObjects.Add(0,'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Probe;Extended Properties="";',$null,1,$ws.Range('A1'))
  $lo.QueryTable.CommandType=2
  $lo.QueryTable.CommandText='SELECT * FROM [Probe]'
  $lo.QueryTable.BackgroundQuery=$false
  $ok=$lo.QueryTable.Refresh($false)
  $rows=@()
  for($r=1;$r -le [Math]::Min(45,$ws.UsedRange.Rows.Count);$r++) {
    $row=@(); for($c=1;$c -le $ws.UsedRange.Columns.Count;$c++) {$row+= $ws.Cells.Item($r,$c).Text}
    $rows+= ,$row
  }
  $rows | ConvertTo-Json -Depth 4
} finally { if($wb){$wb.Close($false)}; $excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
