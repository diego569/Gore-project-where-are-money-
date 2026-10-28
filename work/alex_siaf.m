let
 Path=Text.From(Excel.CurrentWorkbook(){[Name="tRutaSIAF"]}[Content]{0}[Archivo]),
 Clean=(v)=>if v=null then "" else Text.Trim(Text.From(v)),
 Tokens=(v)=>List.Select(Text.SplitAny(v," #(lf)#(cr)#(tab)"),each _<>""),
 Digit=(s,n)=>Text.Length(s)=n and Text.Select(s,{"0".."9"})=s,
 Num=(v)=>if Clean(v)="" then 0 else if Value.Is(v,type number) then v else Number.FromText(Clean(v),"en-US"),
 W=Excel.Workbook(File.Contents(Path),null,true), Rows=Table.ToRows(W{[Name="resumen_marco"]}[Data]),Labels=List.Transform(Rows,each Clean(_{1})),
 ValidHeader=Text.Contains(Labels{0},"2026") and Text.Contains(Labels{5},"458 -") and Text.Contains(Labels{6},"UNIDAD EJECUTORA: 001 -"),
 R=Table.FromRecords(List.Transform(Rows,(r)=>let s=Clean(r{1}),t=Tokens(s),h=List.Count(t)>=4 and Digit(t{0},4) and Digit(t{1},4) and Digit(t{2},7) in [Texto=s,Raw=r,CUI=if h then t{2} else null,Meta=if h then Text.From(Number.FromText(t{0})) else null])),
 Fill=Table.FillDown(R,{"CUI","Meta"}),Totals=Table.SelectRows(Fill,each Text.StartsWith([Texto],"TOTAL META ")),
 Data=Table.FromRecords(List.Transform(Table.ToRecords(Totals),(r)=>[CUI=r[CUI],Meta=r[Meta],PIM=Num(r[Raw]{7}),Certificacion=Num(r[Raw]{9}),Devengado=Num(r[Raw]{17})])),
 Group=Table.Group(Data,{"CUI","Meta"},{{"PIM",each List.Sum([PIM]),type number},{"Certificacion",each List.Sum([Certificacion]),type number},{"Devengado",each List.Sum([Devengado]),type number}}),
 K=Table.AddColumn(Group,"Clave",each [CUI]&"|"&[Meta],type text),
 FI=List.PositionOf(Labels,"TOTAL UNIDAD EJECUTORA 001:"),F=Rows{FI+1},
 Valid=ValidHeader and Number.Abs(List.Sum(Group[PIM])-Num(F{8}))<0.02 and Number.Abs(List.Sum(Group[Certificacion])-Num(F{9}))<0.02 and Number.Abs(List.Sum(Group[Devengado])-Num(F{17}))<0.02
in if Valid then Table.ReorderColumns(K,{"Clave","CUI","Meta","PIM","Certificacion","Devengado"}) else error "El SIAF no corresponde a UE 001 / 2026 o no concilia con su total."
