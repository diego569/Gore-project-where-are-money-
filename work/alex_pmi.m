let
 Path=Text.From(Excel.CurrentWorkbook(){[Name="tRutaPMI"]}[Content]{0}[Archivo]),
 W=Excel.Workbook(File.Contents(Path),null,true),T=Table.PromoteHeaders(Table.Skip(W{[Item="Cartera",Kind="Sheet"]}[Data],5),[PromoteAllScalars=true]),
 P=Table.SelectColumns(T,{"Código Único","Nombre de inversión","Monto Año 2026 (S/)","Monto Año 2027 (S/)","Monto Año 2028 (S/)","Monto Año 2029 (S/)"}),
 R=Table.RenameColumns(P,List.Zip({Table.ColumnNames(P),{"CUI","Proyecto","PMI2026","PMI2027","PMI2028","PMI2029"}})),
 C=Table.TransformColumns(R,{{"CUI",each if _=null then "" else Text.From(_),type text}}),D=Table.SelectRows(C,each [CUI]<>""),
 Valid=Table.RowCount(D)=List.Count(List.Distinct(D[CUI]))
in if Valid then D else error "La cartera contiene CUI duplicados."
