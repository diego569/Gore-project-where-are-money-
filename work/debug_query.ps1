$ErrorActionPreference='Stop'
$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
try {
 $wb=$excel.Workbooks.Open((Join-Path (Get-Location) 'outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx'),0,$true)
 $m=[IO.File]::ReadAllText((Join-Path (Get-Location) 'work/qFuentes.m'))
 $m=$m.Substring(0,$m.LastIndexOf("`nin"))+"`nin Table.Combine({Table.SelectRows(ChecksText,each Number.Abs([Diferencia])>0.02),#table({`"Control`",`"Resultado`"},{{`"ValidPDF`",Text.From(ValidPDF)},{`"ValidE`",Text.From(ValidE)},{`"EHeader`",Text.From(EValidHeader)},{`"ValidPMI`",Text.From(ValidPMI)},{`"MissingE`",Text.From(MissingE)},{`"PDFRows`",Text.From(Table.RowCount(PDF))},{`"PMIRows`",Text.From(Table.RowCount(PMI))}})})"
 $null=$wb.Queries.Add('Debug',$m)
 $ws=$wb.Worksheets.Add();$lo=$ws.ListObjects.Add(0,'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Debug;Extended Properties="";',$null,1,$ws.Range('A1'))
 $lo.QueryTable.CommandType=2;$lo.QueryTable.CommandText='SELECT * FROM [Debug]';$lo.QueryTable.BackgroundQuery=$false
 $null=$lo.QueryTable.Refresh($false)
 $rows=@();for($r=1;$r -le $lo.Range.Rows.Count;$r++) {$row=@();for($c=1;$c -le $lo.Range.Columns.Count;$c++){$row+=$ws.Cells.Item($r,$c).Text};$rows+=,$row};$rows|ConvertTo-Json -Depth 4
}finally{if($wb){$wb.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
