$ErrorActionPreference='Stop'
$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
try {
 $wb=$excel.Workbooks.Open((Join-Path (Get-Location) 'outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx'),0,$false)
 foreach($sn in @('Proyectos','PMI','Control','Certificacion PDF','Ejecucion SIAF')) {
  $s=$wb.Worksheets.Item($sn);$t=$s.ListObjects.Item(1)
  $t.DataBodyRange.Interior.Pattern=-4142
  $t.TableStyle='TableStyleMedium2';$t.ShowTableStyleRowStripes=$true
  $t.HeaderRowRange.Font.Bold=$true;$t.HeaderRowRange.Font.Color=16777215
  $t.HeaderRowRange.Interior.Color=5847830;$t.HeaderRowRange.RowHeight=44;$t.HeaderRowRange.WrapText=$true
  $t.DataBodyRange.NumberFormat='#,##0.00'
  $t.QueryTable.AdjustColumnWidth=$false;$t.QueryTable.PreserveFormatting=$true
 }
 $s=$wb.Worksheets.Item('Proyectos');$s.Columns.Item('A').ColumnWidth=10;$s.Columns.Item('B:C').ColumnWidth=11;$s.Columns.Item('D').ColumnWidth=70;$s.Columns.Item('E').ColumnWidth=32;$s.Columns.Item('F:U').ColumnWidth=20;$s.Columns.Item('V:AG').ColumnWidth=29
 $s.Range('A6:A2000').NumberFormat='0';$s.Range('B6:E2000').NumberFormat='@';$s.Range('V6:AG2000').NumberFormat='@';$s.Range('AC6:AC2000').NumberFormat='dd/mm/yyyy';$s.Range('P6:P2000').NumberFormat='0.0%'
 $s.Range('D6:E455').WrapText=$true;$s.Range('W6:AB455').WrapText=$true;$s.Range('A6:AG455').Rows.AutoFit()|Out-Null
 $s=$wb.Worksheets.Item('PMI');$s.Columns.Item('A').ColumnWidth=12;$s.Columns.Item('B').ColumnWidth=75;$s.Columns.Item('C').ColumnWidth=28;$s.Columns.Item('D:E').ColumnWidth=13;$s.Columns.Item('F:L').ColumnWidth=22;$s.Columns.Item('M').ColumnWidth=25
 $s.Range('A6:E375').NumberFormat='@';$s.Range('M6:M375').NumberFormat='@';$s.Range('B6:C375').WrapText=$true;$s.Range('A6:M375').Rows.AutoFit()|Out-Null
 $s=$wb.Worksheets.Item('Certificacion PDF');$s.Columns.Item('A:B').ColumnWidth=12;$s.Columns.Item('C').ColumnWidth=75;$s.Columns.Item('D:J').ColumnWidth=22;$s.Range('A6:C455').NumberFormat='@';$s.Range('I6:J455').NumberFormat='@';$s.Range('C6:C455').WrapText=$true;$s.Range('A6:J455').Rows.AutoFit()|Out-Null
 $s=$wb.Worksheets.Item('Ejecucion SIAF');$s.Columns.Item('A:D').ColumnWidth=12;$s.Columns.Item('E:I').ColumnWidth=22;$s.Range('A6:D495').NumberFormat='@'
 $s=$wb.Worksheets.Item('Control');$s.Columns.Item('A:B').ColumnWidth=35;$s.Columns.Item('C:E').ColumnWidth=23;$s.Columns.Item('F').ColumnWidth=64;$s.Range('A6:B61').NumberFormat='@';$s.Range('F6:F61').NumberFormat='@';$s.Range('F6:F61').WrapText=$true;$s.Range('A6:F61').Rows.AutoFit()|Out-Null
 $wb.Worksheets.Item('Guia').Range('A22').Value2='Vacíos en PMI'
 $wb.Worksheets.Item('Guia').Range('B22').Value2='Siete inversiones de la cartera recibida no informan monto 2026. Se conserva el vacío de la fuente; no representa cero.'
 $wb.Worksheets.Item('Guia').Range('A23').Value2='Primera apertura'
 $wb.Worksheets.Item('Guia').Range('B23').Value2='Si Excel pide habilitar las conexiones del archivo, habilítelas para actualizar. Las consultas solo leen los archivos locales indicados en Configuracion.'
 $wb.Worksheets.Item('Guia').Range('A22:B23').WrapText=$true;$wb.Worksheets.Item('Guia').Range('A22:B23').Font.Size=11;$wb.Worksheets.Item('Guia').Range('A22:B23').RowHeight=46
 $excel.CalculateFull();$wb.Worksheets.Item('Resumen').Activate();$wb.Save();Write-Output 'FINAL FORMATTING SAVED'
}finally{if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
