$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$out=Join-Path $root 'outputs/alex_central'
$null=New-Item -ItemType Directory -Path $out -Force
$spec=Get-Content -LiteralPath 'work/alex_spec.json' -Raw | ConvertFrom-Json
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
function Sum-Leaf($ws,$row,$leaves){
 foreach($col in @('D','E','G')) {
  $terms=@($leaves | ForEach-Object {"$col$_"})
  $ws.Range("$col$row").Formula=if($terms.Count){'='+($terms -join '+')}else{'=0'}
 }
 $ws.Range("F$row").Formula="=D$row-E$row";$ws.Range("H$row").Formula="=D$row-G$row";$ws.Range("I$row").Formula="=IF(D$row=0,0,G$row/D$row)"
}
function Leaf-Data($ws,$leaf){
 $r=$leaf.row;$key=$leaf.key;$cui=$leaf.cui
 foreach($pair in @(@('D','PIM'),@('E','Certificacion'),@('G','Devengado'))){$c=$pair[0];$metric=$pair[1];$ws.Range("$c$r").Formula='=IF(COUNTIF(tSIAF[Clave],"'+$key+'")=0,"",SUMIF(tSIAF[Clave],"'+$key+'",tSIAF['+$metric+']))'}
 $ws.Range("F$r").Formula="=IF(D$r=`"`",`"`",D$r-E$r)";$ws.Range("H$r").Formula="=IF(D$r=`"`",`"`",D$r-G$r)";$ws.Range("I$r").Formula="=IF(D$r=`"`",`"`",IF(D$r=0,0,G$r/D$r))"
 if($leaf.mode){$ws.Range("J$r").Value2=[string]$leaf.mode}
 $ws.Range("K$r").Value2='Por completar'
 for($i=0;$i -lt 4;$i++) {
  $c=[string][char](76+$i);$year=2026+$i;$ix='INDEX(tCartera[PMI'+$year+'],MATCH("'+$cui+'",tCartera[CUI],0))'
  $ws.Range("$c$r").Formula='=IF(COUNTIF(tCartera[CUI],"'+$cui+'")=0,"Sin PMI",IF('+ $ix +'="","No informado",'+$ix+'))'
 }
}
try {
 $wb=$excel.Workbooks.Open((Join-Path $root 'ref docs/0001917176_1_SINVERSION (1).xlsm'),0,$false)
 $wb.SaveAs((Join-Path $out 'Alex_Sede_Central_2026.xlsx'),51)
 $excel.Calculation=-4135
 $links=$wb.LinkSources(1);if($links){foreach($link in $links){$wb.BreakLink($link,1)}}
 for($i=$wb.Names.Count;$i -ge 1;$i--){$n=$wb.Names.Item($i);if($n.RefersTo -match '#REF!|GET.CELL|OBTENER.CELDA'){$n.Delete()}}
 $keep=@('Gerencia de Infraestructura','Gerencias y Oficinas','GRDE','GRRNyDA','PETT','PECSA','DS','DEPI','ABS')
 for($i=$wb.Worksheets.Count;$i -ge 1;$i--){$ws=$wb.Worksheets.Item($i);if($keep -notcontains $ws.Name){$ws.Delete()}}
 $wb.Worksheets.Item('Gerencias y Oficinas').Rows.Item('315:431').Delete() | Out-Null
 # Native Excel preserves the original merged layout and styles while replacing obsolete links.
 foreach($name in @('SIAF','Cartera')) {
  $ws=$wb.Worksheets.Add([Type]::Missing,$wb.Worksheets.Item($wb.Worksheets.Count));$ws.Name=$name
  $ws.Range('A1').Value2='Archivo';$ws.Range('B1').Value2='Uso'
  $ws.Range('A2').Value2=Join-Path $root $(if($name -eq 'SIAF'){'ref docs/reporte20261006085201.xls'}else{'ref docs/Cartera_261006-090543.xlsx'})
  $ws.Range('B2').Value2='Reemplazar ruta al descargar un nuevo reporte'
  $route=$ws.ListObjects.Add(1,$ws.Range('A1:B2'),$null,1);$route.Name=if($name -eq 'SIAF'){'tRutaSIAF'}else{'tRutaPMI'}
  $ws.Range('A3').Value2='Fecha de corte declarada';$ws.Range('B3').Value2=46301;$ws.Range('B3').NumberFormat='dd/mm/yyyy'
 }
 $null=$wb.Queries.Add('SIAF_Central',[IO.File]::ReadAllText((Join-Path $root 'work/alex_siaf.m')))
 $null=$wb.Queries.Add('Cartera_Origen',[IO.File]::ReadAllText((Join-Path $root 'work/alex_pmi.m')))
 $null=$wb.Queries.Add('Cartera_Central','let P=Cartera_Origen, C=Table.Distinct(Table.SelectColumns(SIAF_Central,{"CUI"})), J=Table.NestedJoin(P,{"CUI"},C,{"CUI"},"S",JoinKind.Inner) in Table.RemoveColumns(J,{"S"})')
 foreach($cfg in @(@('SIAF','SIAF_Central','tSIAF'),@('Cartera','Cartera_Central','tCartera'))){
  $ws=$wb.Worksheets.Item($cfg[0]);$conn='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location='+$cfg[1]+';Extended Properties="";'
  $t=$ws.ListObjects.Add(0,$conn,$null,1,$ws.Range('A6'));$t.Name=$cfg[2];$t.QueryTable.CommandType=2;$t.QueryTable.CommandText='SELECT * FROM ['+$cfg[1]+']';$t.QueryTable.BackgroundQuery=$false
  $null=$t.QueryTable.Refresh($false);$t.TableStyle='TableStyleLight9';$t.DataBodyRange.NumberFormat='#,##0.00';$t.QueryTable.AdjustColumnWidth=$false;$t.QueryTable.PreserveFormatting=$true
  $ws.Columns.Item('A').ColumnWidth=65;$ws.Columns.Item('B').ColumnWidth=45;$ws.Columns.Item('C:F').ColumnWidth=22
  Write-Output ($cfg[0]+' '+$t.ListRows.Count)
 }
 foreach($sn in $keep | Where-Object {$_ -ne 'ABS'}) {
  $ws=$wb.Worksheets.Item($sn);$last=if($spec.primary.PSObject.Properties.Name -contains $sn){$spec.primary.$sn.last}else{$spec.aliases.$sn.last}
  $head=if($sn -eq 'Gerencia de Infraestructura'){3}else{4}
  for($c=10;$c -le 15;$c++){$ws.Range("I$head`:`I$last").Copy();$ws.Cells.Item($head,$c).PasteSpecial(-4122) | Out-Null}
  $ws.Range("D6:O$last").ClearContents();$ws.Range("J$head`:O$($head+1)").UnMerge()
  $titles=@('MODALIDAD DE EJECUCIÓN','ESTADO SITUACIONAL DEL PROYECTO','Monto Año 2026 (S/)','Monto Año 2027 (S/)','Monto Año 2028 (S/)','Monto Año 2029 (S/)')
  for($i=0;$i -lt 6;$i++){$c=10+$i;$rg=$ws.Range($ws.Cells.Item($head,$c),$ws.Cells.Item($head+1,$c));$rg.Merge();$rg.Cells.Item(1,1).Value2=$titles[$i];$rg.WrapText=$true}
  $ws.Columns.Item('J').ColumnWidth=32;$ws.Columns.Item('K').ColumnWidth=29;$ws.Columns.Item('L:O').ColumnWidth=20
  $ws.Range("J6:K$last").NumberFormat='@';$ws.Range("L6:O$last").NumberFormat='#,##0.00';$ws.Range("J6:K$last").WrapText=$true
  $ws.Range('A2:O2').UnMerge();$ws.Range('A2:O2').Merge();$ws.Range('A2').Formula='="COMPORTAMIENTO DE EJECUCIÓN DE PROYECTOS DE INVERSIÓN PÚBLICA 2026 - U.E. 001 SEDE CENTRAL AL "&TEXT(SIAF!B3,"dd/mm/yyyy")'
  $ws.PageSetup.PrintArea="`$A`$1:`$O`$$last";$ws.PageSetup.Orientation=2;$ws.PageSetup.Zoom=$false;$ws.PageSetup.FitToPagesWide=1;$ws.PageSetup.FitToPagesTall=$false
 }
 foreach($p in $spec.primary.PSObject.Properties) {
  $sn=$p.Name;$x=$p.Value;$ws=$wb.Worksheets.Item($sn)
  $show=@(1,2,3,4,5,$x.last)+@($x.leaves|ForEach-Object {$_.row})+@($x.totals.PSObject.Properties.Name|ForEach-Object {[int]$_})
  for($r=6;$r -le $x.last;$r++){if($show -notcontains $r){$ws.Rows.Item($r).Hidden=$true}}
  foreach($leaf in $x.leaves){Leaf-Data $ws $leaf}
  foreach($t in $x.totals.PSObject.Properties){Sum-Leaf $ws ([int]$t.Name) $t.Value}
 }
 $g=$wb.Worksheets.Item('Gerencias y Oficinas')
 $g.Rows.Item(6).Copy();$g.Rows.Item(316).PasteSpecial(-4122)|Out-Null;$g.Range('A316:C316').Merge();$g.Range('A316').Value2='METAS SIAF POR CLASIFICAR EN EL FORMATO'
 foreach($m in $spec.missing){
  $r=$m.row;$g.Rows.Item(8).Copy();$g.Rows.Item($r).PasteSpecial(-4122)|Out-Null;$g.Range("A$r`:O$r").UnMerge()
  $g.Range("A$r").Value2=[double]($r-316);$g.Range("B$r").Value2=[string]$m.meta;$g.Range("C$r").Value2=$m.cui+': '+$m.name
  Leaf-Data $g ([pscustomobject]@{row=$r;cui=$m.cui;key=$m.cui+'|'+$m.meta;mode='Por validar'})
  $g.Rows.Item($r).RowHeight=65;$g.Range("C$r").WrapText=$true
 }
 Sum-Leaf $g 316 @(317..322)
 $g.Rows.Item(314).Copy();$g.Rows.Item(323).PasteSpecial(-4122)|Out-Null;$g.Range('A323:C323').Merge();$g.Range('A323').Value2='TOTAL METAS POR CLASIFICAR';Sum-Leaf $g 323 @(317..322)
 $g.PageSetup.PrintArea='$A$1:$O$323'
 foreach($c in @('D','E','G')){$g.Range("${c}314").Formula="=SUM($($spec.primary.'Gerencias y Oficinas'.leaves|ForEach-Object {"$c$($_.row)"}|Join-String -Separator ','))+${c}316+'Gerencia de Infraestructura'!${c}905"}
 $gi=$wb.Worksheets.Item('Gerencia de Infraestructura');foreach($c in @('D','E','F','G','H','I')){$gi.Range("${c}5").Formula="='Gerencias y Oficinas'!${c}314"}
 $masterGroups=@{GRDE=6;GRRNyDA=192;PETT=217;PECSA=224;DS=231;DEPI=300}
 foreach($p in $spec.aliases.PSObject.Properties){
  $ws=$wb.Worksheets.Item($p.Name);$x=$p.Value;$show=@(1,2,3,4,5,6,$x.last)+@($x.leaves|ForEach-Object {$_.row})+@($x.totals.PSObject.Properties.Name|ForEach-Object {[int]$_})
  for($r=7;$r -lt $x.last;$r++){if($show -notcontains $r){$ws.Rows.Item($r).Hidden=$true}}
  $ws.Range("J6:K$($x.last)").NumberFormat='General'
  foreach($l in $x.leaves){foreach($c in @('D','E','F','G','H','I','J','K','L','M','N','O')){$ref="'$($l.primary_sheet)'!$c$($l.primary_row)";$ws.Range("$c$($l.row)").Formula="=IF($ref=`"`",`"`",$ref)"}}
  foreach($t in $x.totals.PSObject.Properties){Sum-Leaf $ws ([int]$t.Name) $t.Value}
  foreach($c in @('D','E','F','G','H','I')){$ws.Range("${c}6").Formula="='Gerencias y Oficinas'!$c$($masterGroups[$p.Name])";$ws.Range("$c$($x.last)").Formula="=${c}6"}
 }
 $abs=$wb.Worksheets.Item('ABS');$abs.Rows.Item(15).Insert()|Out-Null;$abs.Rows.Item(14).Copy();$abs.Rows.Item(15).PasteSpecial(-4122)|Out-Null
 $abs.Range('A15').Value2=8;$abs.Range('B15').Value2='METAS POR CLASIFICAR';$abs.Range('L5').Formula='="AL "&TEXT(SIAF!B3,"dd/mm/yyyy")'
 foreach($pair in @(@('E','D'),@('F','E'),@('G','F'),@('H','G'),@('I','H'))){$abs.Range($pair[0]+'15').Formula="='Gerencias y Oficinas'!$($pair[1])316";$abs.Range($pair[0]+'16').Formula='=SUM('+$pair[0]+'8:'+$pair[0]+'15)'}
 for($r=8;$r -le 16;$r++){foreach($pair in @(@('J','H'),@('K','G'),@('L','I'))){$abs.Range($pair[0]+$r).Formula="=IF(E$r=0,0,$($pair[1])$r/E$r)"}}
 $abs.Range('A18:L18').Merge();$abs.Range('A18').Formula='=IF(AND(ABS(E16-SUM(tSIAF[PIM]))<0.02,ABS(F16-SUM(tSIAF[Certificacion]))<0.02,ABS(H16-SUM(tSIAF[Devengado]))<0.02),"Control: los importes concilian con el reporte SIAF de Sede Central.","REVISAR: hay metas nuevas, ausentes o diferencias respecto al SIAF.")'
 $notes=@('Actualización: Datos > Actualizar todo. Cambiar rutas y fechas en las hojas SIAF y Cartera (ocultas).','Modalidad referencial según agrupación de Alex; confirmar en J. Estado editable en K de los dos reportes principales.','La programación PMI se repite si un CUI tiene varias metas. No sumar los años entre metas. Sin PMI no significa cero.','Se ocultan filas de la plantilla sin coincidencia CUI + meta. Las seis metas no coincidentes se muestran por clasificar.')
 for($i=0;$i -lt $notes.Count;$i++){$r=19+$i;$abs.Range("A$r`:L$r").Merge();$abs.Range("A$r").Value2=$notes[$i];$abs.Range("A$r").WrapText=$true;$abs.Rows.Item($r).RowHeight=30}
 $abs.PageSetup.PrintArea='$A$1:$L$22'
 $wb.Worksheets.Item('SIAF').Visible=0;$wb.Worksheets.Item('Cartera').Visible=0
 $excel.Calculation=-4105;$excel.CalculateFull()
 if([Math]::Abs($abs.Range('E16').Value2-905026984) -gt 0.02 -or [Math]::Abs($abs.Range('F16').Value2-843143127.18) -gt 0.02 -or [Math]::Abs($abs.Range('H16').Value2-433575772.65) -gt 0.02){throw "No concilia ABS: $($abs.Range('E16').Value2) $($abs.Range('F16').Value2) $($abs.Range('H16').Value2)"}
 $state=$gi.Range('K10').Value2;$gi.Range('K10').Value2='PRUEBA TEMPORAL';$wb.RefreshAll();$excel.CalculateUntilAsyncQueriesDone();$excel.CalculateFull();if($gi.Range('K10').Value2 -ne 'PRUEBA TEMPORAL'){throw 'Se perdió el estado editable'};$gi.Range('K10').Value2=$state
 $gi.Range('C438:C441').ClearContents();$gi.Range('C458:C485').ClearContents()
 foreach($p in $spec.primary.PSObject.Properties){$ws=$wb.Worksheets.Item($p.Name);foreach($leaf in $p.Value.leaves){if($leaf.mode -and $ws.Rows.Item($leaf.row).RowHeight -lt 42){$ws.Rows.Item($leaf.row).RowHeight=42}}}
 foreach($p in $spec.aliases.PSObject.Properties){$ws=$wb.Worksheets.Item($p.Name);foreach($leaf in $p.Value.leaves){if($ws.Rows.Item($leaf.row).RowHeight -lt 42){$ws.Rows.Item($leaf.row).RowHeight=42}}}
 $abs.Range('L5').WrapText=$true;$abs.Rows.Item(5).RowHeight=28
 $excel.CalculateFull();$gi.Activate();$gi.Range('A1').Select();$wb.Save()
 Write-Output ('SAVED: '+$wb.FullName+' | PIM '+$abs.Range('E16').Value2+' | CERT '+$abs.Range('F16').Value2+' | DEV '+$abs.Range('H16').Value2)
}finally{if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
