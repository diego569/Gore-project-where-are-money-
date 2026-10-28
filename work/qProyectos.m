let
    F=qFuentes,
    S0=qSeguimiento,
    Clean=(v)=>if v=null then "" else Text.Trim(Text.From(v)),
    S=Table.SelectRows(S0,each Clean([Clave])<>""),
    CheckS=Table.RowCount(S)=List.Count(List.Distinct(S[Clave])),
    A0=qAlex,
    A=Table.TransformColumns(A0,{{"UE",each Text.PadStart(Clean(_),3,"0"),type text},{"CUI",each Clean(_),type text}}),
    P=Table.AddColumn(F[PDF],"Clave",each "2026|"&[UE]&"|"&[CUI],type text),
    J1=Table.NestedJoin(P,{"UE","CUI"},F[EAgg],{"UE","CUI"},"E",JoinKind.LeftOuter),
    J2=Table.NestedJoin(J1,{"CUI"},F[PMI],{"CUI"},"M",JoinKind.LeftOuter),
    J3=Table.NestedJoin(J2,{"UE","CUI"},A,{"UE","CUI"},"A",JoinKind.LeftOuter),
    J4=Table.NestedJoin(J3,{"Clave"},S,{"Clave"},"S",JoinKind.LeftOuter),
    Param=(k)=>Text.From(F[Config]{[Parametro=k]}[Valor]),
    Rows=List.Transform(Table.ToRecords(J4),(r)=>
        let e=if Table.IsEmpty(r[E]) then null else r[E]{0},m=if Table.IsEmpty(r[M]) then null else r[M]{0},a=if Table.IsEmpty(r[A]) then null else r[A]{0},s=if Table.IsEmpty(r[S]) then null else r[S]{0},
            val=(key,def)=>if s=null or Clean(Record.Field(s,key))="" then def else Record.Field(s,key),
            mode=if a=null or Clean(a[Modalidad antecedente])="" then "Pendiente de validar" else a[Modalidad antecedente]&" (antecedente por validar)",
            ep=(k)=>if e=null then "No disponible" else Record.Field(e,k),mp=(k)=>if m=null then "No disponible" else Record.Field(m,k)
        in [Ejercicio=2026,UE=r[UE],CUI=r[CUI],Proyecto=r[Nombre],Gerencia=if a=null then "Sin clasificar" else a[Gerencia],#"PIA PDF"=r[PIA],#"Modificaciones PDF"=r[Modificaciones],#"PIM certificado"=r[PIM],Certificacion=r[Certificacion],#"Saldo sin certificar"=r[PIM]-r[Certificacion],#"PIM ejecucion"=ep("PIM"),#"Compromiso anual"=ep("Compromiso"),Devengado=ep("Devengado"),#"Saldo sin compromiso"=if e=null then "No disponible" else e[PIM]-e[Compromiso],#"Saldo sin devengar"=if e=null then "No disponible" else e[PIM]-e[Devengado],#"Avance devengado"=if e=null then "No disponible" else if e[PIM]=0 then "Sin PIM" else e[Devengado]/e[PIM],#"Diferencia PIM"=if e=null then "No comparable" else r[PIM]-e[PIM],#"PMI 2026"=mp("PMI 2026"),#"PMI 2027"=mp("PMI 2027"),#"PMI 2028"=mp("PMI 2028"),#"PMI 2029"=mp("PMI 2029"),#"Cruce PMI"=if m=null then "Sin coincidencia" else "Encontrado",Modalidad=val("Modalidad validada",mode),#"Estado situacional"=val("Estado situacional","Pendiente de actualizar"),Restriccion=val("Restriccion","Pendiente de evaluación"),#"Base legal"=val("Base legal",""),Observacion=val("Observacion",""),Responsable=val("Responsable",""),#"Fecha revision"=val("Fecha revision",""),#"Fecha PDF"=r[Fecha PDF],#"Fecha ejecucion"=if e=null then "No disponible" else Param("Fecha ejecucion"),#"Fecha PMI"=if m=null then "Sin coincidencia" else Param("Fecha PMI"),Clave=r[Clave]]),
    Result=Table.Sort(Table.FromRecords(Rows),{{"UE",Order.Ascending},{"Saldo sin certificar",Order.Descending}})
in if CheckS and F[MissingE]=0 then Result else error "Hay claves duplicadas en Seguimiento o proyectos de ejecución sin correspondencia en PDF. Revise antes de actualizar."
