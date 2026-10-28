$ErrorActionPreference='Stop'
$root='E:\New folder\Economia\8vo\gore\Cash data request'
$spec=Get-Content -LiteralPath (Join-Path $root 'work/alex_spec.json') -Raw | ConvertFrom-Json
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$false
$excel.DisplayAlerts=$false
$excel.AutomationSecurity=3
$excel.AskToUpdateLinks=$false
Write-Output 'Excel iniciado'
try {
 $wb=$excel.Workbooks.Open((Join-Path $root 'outputs/alex_central/Alex_Sede_Central_2026.xlsx'),0,$false)
 Write-Output 'Libro abierto'
 $excel.Calculation=-4135
 $gi=$wb.Worksheets.Item('Gerencia de Infraestructura')
 $gi.Range('C438:C441').ClearContents()
 $gi.Range('C458:C485').ClearContents()
 Write-Output 'Referencias limpiadas'
 foreach($p in $spec.primary.PSObject.Properties){
  $ws=$wb.Worksheets.Item($p.Name)
  foreach($leaf in $p.Value.leaves){if($leaf.mode){$r=$ws.Rows.Item($leaf.row);if($r.RowHeight -lt 42){$r.RowHeight=42}}}
 }
 foreach($p in $spec.aliases.PSObject.Properties){
  Write-Output $p.Name
  $ws=$wb.Worksheets.Item($p.Name)
  foreach($leaf in $p.Value.leaves){
   $r=$ws.Rows.Item($leaf.row);if($r.RowHeight -lt 42){$r.RowHeight=42}
   foreach($c in @('J','K')){
    $cell=$ws.Range("$c$($leaf.row)");$cell.NumberFormat='General'
    $ref="'$($leaf.primary_sheet)'!$c$($leaf.primary_row)"
    $cell.Formula="=IF($ref=`"`",`"`",$ref)"
   }
  }
 }
 $wb.Worksheets.Item('ABS').Range('L5').WrapText=$true
 $wb.Worksheets.Item('ABS').Rows.Item(5).RowHeight=28
 $excel.CalculateFull()
 Write-Output 'Calculado'
 $wb.Save()
 Write-Output 'Celdas heredadas corregidas; formato ajustado y libro guardado.'
}finally{if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
