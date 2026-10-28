param([switch]$SkipSecondRefresh)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$path=Join-Path $root 'outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx'
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
try {
 $wb=$excel.Workbooks.Open($path,0,$false)
 $definitions=@{
 qSeguimiento='Excel.CurrentWorkbook(){[Name="tSeguimiento"]}[Content]'
 qAlex='Excel.CurrentWorkbook(){[Name="tAlex"]}[Content]'
 qFuentes=[IO.File]::ReadAllText((Join-Path $root 'work/qFuentes.m'))
 qProyectos=[IO.File]::ReadAllText((Join-Path $root 'work/qProyectos.m'))
 qControl=[IO.File]::ReadAllText((Join-Path $root 'work/qControl.m'))
 qPDF='qFuentes[PDF]'
 qEjecucion='qFuentes[Ejecucion]'
 qPMI='qFuentes[PMI]'
 }
 foreach($k in @('qSeguimiento','qAlex','qFuentes','qPDF','qEjecucion','qPMI','qProyectos','qControl')) {
  $null=$wb.Queries.Add($k,$definitions[$k])
 }
 $targets=@(@('qProyectos','Proyectos','tProyectos'),@('qPDF','Certificacion PDF','tPDF'),@('qEjecucion','Ejecucion SIAF','tEjecucion'),@('qPMI','PMI','tPMI'),@('qControl','Control','tControl'))
 foreach($target in $targets) {
  $qn=$target[0];$ws=$wb.Worksheets.Item($target[1]);$old=$ws.ListObjects.Item($target[2]);$old.Unlist()
  $ws.Range('A5:AG2000').ClearContents()
  $conn='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location='+$qn+';Extended Properties="";'
  $lo=$ws.ListObjects.Add(0,$conn,$null,1,$ws.Range('A5'));$lo.Name=$target[2]
  $lo.QueryTable.CommandType=2;$lo.QueryTable.CommandText='SELECT * FROM ['+$qn+']';$lo.QueryTable.BackgroundQuery=$false
  $lo.QueryTable.AdjustColumnWidth=$false;$lo.QueryTable.PreserveFormatting=$true
  Write-Output ('REFRESH '+$qn)
  $null=$lo.QueryTable.Refresh($false)
  $lo.TableStyle='TableStyleMedium2';$lo.HeaderRowRange.WrapText=$true;$lo.HeaderRowRange.RowHeight=44
  Write-Output ('ROWS '+$qn+' '+$lo.ListRows.Count)
 }
 $excel.CalculateFull()
 $snap=@{}
 foreach($target in $targets){$lo=$wb.Worksheets.Item($target[1]).ListObjects.Item($target[2]);$snap[$target[0]]=$lo.ListRows.Count}
 if(-not $SkipSecondRefresh){
  $sw=$wb.Worksheets.Item('Seguimiento');$before=$sw.Range('H6').Value2;$sw.Range('H6').Value2='PRUEBA CONSERVACION - TEMPORAL'
  Write-Output 'SECOND REFRESH'
  $wb.RefreshAll();$excel.CalculateUntilAsyncQueriesDone();$excel.CalculateFull()
  if($sw.Range('H6').Value2 -ne 'PRUEBA CONSERVACION - TEMPORAL'){throw 'Se perdió anotación de seguimiento'}
  foreach($target in $targets){$lo=$wb.Worksheets.Item($target[1]).ListObjects.Item($target[2]);if($snap[$target[0]] -ne $lo.ListRows.Count){throw 'Filas cambiaron al actualizar dos veces'}}
  $sw.Range('H6').Value2=$before
  $null=$wb.Worksheets.Item('Proyectos').ListObjects.Item('tProyectos').QueryTable.Refresh($false)
 }
 $ws=$wb.Worksheets.Item('Proyectos');$ws.Range('F6:U2000').NumberFormat='#,##0.00';$ws.Range('P6:P2000').NumberFormat='0.0%';$ws.Range('A6:A2000').NumberFormat='0';$ws.Range('B6:C2000').NumberFormat='@'
 $ws.Range('D6:E2000').WrapText=$true;$ws.Range('A6:AG455').Rows.AutoFit() | Out-Null
 $wb.Worksheets.Item('Resumen').Range('A7:A13').NumberFormat='@'
 foreach($sn in @('PMI','Clasificacion Alex','Certificacion PDF')) { $wb.Worksheets.Item($sn).UsedRange.Rows.AutoFit() | Out-Null }
 $wb.Worksheets.Item('Control').Range('F6:F100').WrapText=$true
 $wb.Worksheets.Item('Control').Range('A6:F100').RowHeight=30
 # Rebind summary formulas after converting static tables into query tables.
 $rs=$wb.Worksheets.Item('Resumen')
 for($r=7;$r -le 13;$r++) {
  $rs.Range("C$r").Formula="=COUNTIF(tProyectos[UE],A$r)"
  $rs.Range("D$r").Formula="=SUMIF(tProyectos[UE],A$r,tProyectos[PIM certificado])"
  $rs.Range("E$r").Formula="=SUMIF(tProyectos[UE],A$r,tProyectos[Certificacion])"
 }
 $rs.Range('H7').Formula='=SUMIF(tProyectos[UE],A7,tProyectos[Devengado])'
 $excel.CalculateFull()
 Write-Output ('SUMMARY PIM '+$rs.Range('D15').Value2+' CERT '+$rs.Range('E15').Value2+' SALDO '+$rs.Range('F15').Value2)
 if([Math]::Abs($rs.Range('D15').Value2-987989928) -gt 0.02){throw 'PIM no concilia'}
 if([Math]::Abs($rs.Range('F15').Value2-68725290.04) -gt 0.02){throw 'Saldo no concilia'}
 $rs.Activate();$excel.ActiveWindow.DisplayGridlines=$false;$rs.Range('A1').Select()
 $wb.Save()
 Write-Output 'SAVED AND VERIFIED'
} finally {if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
