let
    F=qFuentes,
    P=qProyectos,
    S=qSeguimiento,
    Row=(c,f,v,r,d,t)=>[Control=c,Fuente=f,Calculado=v,Referencia=r,Diferencia=d,Resultado=t],
    Extra=Table.FromRecords({
        Row("Proyectos / UE","PDF",Table.RowCount(F[PDF]),null,null,"Registros únicos"),
        Row("Repeticiones consolidadas","PDF",F[PDFFilas]-Table.RowCount(F[PDF]),null,null,"Filas por proyecto antes de consolidar menos claves únicas"),
        Row("Proyectos sin PMI","Cruce CUI",Table.RowCount(Table.SelectRows(P,each [Cruce PMI]="Sin coincidencia")),null,null,"Filtrar Sin coincidencia en Proyectos"),
        Row("Cartera sin ejecución PDF","Cruce CUI",Table.RowCount(Table.NestedJoin(F[PMI],{"CUI"},F[PDF],{"CUI"},"P",JoinKind.LeftAnti)),null,null,"Permanecen en PMI; no se infiere presupuesto"),
        Row("Sin fila de seguimiento","Seguimiento",Table.RowCount(Table.NestedJoin(P,{"Clave"},S,{"Clave"},"S",JoinKind.LeftAnti)),null,null,"Agregar nuevas claves en Seguimiento"),
        Row("Diferencia certificación UE 001","PDF menos XLS",List.Sum(Table.SelectRows(F[PDF],each [UE]="001")[Certificacion]),F[ERef][Certificacion],List.Sum(Table.SelectRows(F[PDF],each [UE]="001")[Certificacion])-F[ERef][Certificacion],"Fuentes con cortes distintos; no mezclar para calcular saldos"),
        Row("Corte certificación","PDF",null,null,null,F[PDF]{0}[Fecha PDF]),
        Row("Corte ejecución declarado","Configuracion",null,null,null,Text.From(F[Config]{[Parametro="Fecha ejecucion"]}[Valor])),
        Row("Corte PMI declarado","Configuracion",null,null,null,Text.From(F[Config]{[Parametro="Fecha PMI"]}[Valor])),
        Row("Actualización completada","Power Query",null,null,null,DateTime.ToText(DateTime.LocalNow(),"dd/MM/yyyy HH:mm:ss")),
        Row("Cobertura devengado","Excel SIAF",1,7,null,"Solo UE 001; no es total regional")
    })
in Table.Combine({F[Control],Extra})
