$ErrorActionPreference='Stop'
$path=Join-Path (Get-Location) 'outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx'
$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
try {
 $wb=$excel.Workbooks.Open($path,0,$false)
 $p=$wb.Worksheets.Item('Proyectos').ListObjects.Item('tProyectos')
 $s=$wb.Worksheets.Item('Seguimiento').ListObjects.Item('tSeguimiento')
 $n=$p.ListRows.Count
 $key=$s.ListColumns.Item('Clave').DataBodyRange.Cells.Item(1,1).Value2
 $new=$s.ListRows.Add();$new.Range.Cells.Item(1,1).Value2=$key
 $failed=$false
 try {$null=$p.QueryTable.Refresh($false)} catch {$failed=$true;Write-Output ('DUPLICATE REJECTED: '+$_.Exception.Message)}
 if(-not $failed){throw 'No se rechazó clave duplicada'}
 $new.Delete()
 if($p.ListRows.Count -ne $n){throw 'Se perdieron filas tras fallo'}
 $config=$wb.Worksheets.Item('Configuracion');$original=$config.Range('B6').Value2
 $config.Range('B6').Value2=Join-Path (Get-Location) 'ref docs/archivo_inexistente_prueba.pdf'
 $failed=$false
 try {$null=$wb.Worksheets.Item('Certificacion PDF').ListObjects.Item('tPDF').QueryTable.Refresh($false)} catch {$failed=$true;Write-Output ('MISSING FILE REJECTED: '+$_.Exception.Message)}
 if(-not $failed){throw 'No se rechazó archivo inexistente'}
 $config.Range('B6').Value2=$original
 # Discard this diagnostic session, including query failure state and temporary test inputs.
 Write-Output 'FAILURE TESTS PASSED; NO CHANGES SAVED'
}finally{if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
