import sys, json, re, collections, pathlib
sys.path.insert(0, str(pathlib.Path('.analysis-deps').resolve()))
import pdfplumber, openpyxl, xlrd

ROOT=pathlib.Path('ref docs')
num=lambda s: float(str(s).replace(',',''))
records=[]; totals={}; names={}; program=None
with pdfplumber.open(ROOT/'reporte20261006110417.pdf') as pdf:
    for pn,page in enumerate(pdf.pages,1):
        text=page.extract_text()
        ue=re.search(r'UNIDAD EJECUTORA: (\d{3}) - (.*?) \(\d+\)',text)
        if not ue: raise ValueError(f'UE no encontrada página {pn}')
        u,un=ue.groups(); names[u]=un
        lines=[]
        for word in sorted(page.extract_words(), key=lambda w:(w['top'],w['x0'])):
            line=next((l for l in lines[-3:] if abs(l[0]-word['top'])<1.8),None)
            if line is None: line=[word['top'],[]]; lines.append(line)
            line[1].append(word)
        for y,ws in lines:
            ws.sort(key=lambda w:w['x0']); label=' '.join(w['text'] for w in ws)
            if y<220: continue
            if re.match(r'^\d{4} ',label): program=label[:4]
            if re.match(r'^\d{7} ',label): project=label[:7]; pname=label[8:]
            if label.startswith('TOTAL PRODUCTO/PROYECTO') or label.startswith('TOTAL UNIDAD EJECUTORA') or label.startswith('TOTAL PLIEGO:'):
                vals=[num(w['text']) for w in ws if w['x0']>200 and re.fullmatch(r'-?[\d,]+(?:\.\d+)?',w['text'])]
                if len(vals)!=17: raise ValueError((pn,label,len(vals)))
                data=dict(zip(['PIA','Modificaciones','PIM']+[f'Mes{i:02}' for i in range(1,13)]+['Certificacion','Saldo'],vals))
                if label.startswith('TOTAL PRODUCTO/PROYECTO'):
                    code=re.search(r'PRODUCTO/PROYECTO (\d+)',label)[1]
                    records.append(dict(UE=u,Proyecto=code,Nombre=pname if code==project else '',Programa=program,Pagina=pn,**data))
                elif label.startswith('TOTAL UNIDAD'): totals[u]=data
                else: grand=data
for k in ['PIA','Modificaciones','PIM','Certificacion','Saldo']:
    assert abs(sum(r[k] for r in records)-grand[k])<.02,(k,sum(r[k] for r in records),grand[k])
groups=collections.defaultdict(list)
for r in records: groups[(r['UE'],r['Proyecto'])].append(r)
projects=[]
for (u,c),rs in groups.items():
    projects.append(dict(UE=u,Proyecto=c,Nombre=rs[0]['Nombre'],**{k:round(sum(r[k] for r in rs),2) for k in ['PIA','Modificaciones','PIM','Certificacion','Saldo']},Paginas=','.join(str(r['Pagina']) for r in rs)))

sw=xlrd.open_workbook(ROOT/'reporte20261006085201.xls'); s=sw.sheet_by_index(0)
print('XLS samples')
for i in range(s.nrows):
    if 'TOTAL' in str(s.cell_value(i,1)) and (i> s.nrows-20 or i<100): print(i+1,[(xlrd.formula.colname(j),v) for j,v in enumerate(s.row_values(i)) if v!=''])
print('PDF',len(records),'project rows',len(projects),'unique UE-project', 'duplicates',[(k,len(v)) for k,v in groups.items() if len(v)>1][:10])
print('PDF totals',grand)
exrows=[];rub=None;code=None;meta=None
for i in range(s.nrows):
    label=str(s.cell_value(i,1)).strip()
    if re.match(r'^\d{2} ',label): rub=label[:2]
    m=re.match(r'^(\d{4})\s+\d{4}\s+(\d{7})\s+',label)
    if m: meta,code=m.groups()
    if label.startswith('TOTAL META '):
        assert meta==label.split()[-1]
        exrows.append(dict(UE='001',Proyecto=code,Meta=meta,Rubro=rub,**{k:float(s.cell_value(i,j) or 0) for k,j in [('PIA',5),('PIM',7),('Certificacion',9),('Compromiso',13),('Devengado',17)]}))
footer=s.row_values(s.nrows-1)
for k,j in [('PIA',6),('PIM',8),('Certificacion',9),('Compromiso',13),('Devengado',17)]:
    assert abs(sum(r[k] for r in exrows)-footer[j])<.02,(k,sum(r[k] for r in exrows),footer[j])
eg=collections.defaultdict(list)
for r in exrows: eg[r['Proyecto']].append(r)
execution={k:{v:round(sum(r[v] for r in rs),2) for v in ['PIA','PIM','Certificacion','Compromiso','Devengado']} for k,rs in eg.items()}

pw=openpyxl.load_workbook(ROOT/'Cartera_261006-090543.xlsx',read_only=True,data_only=True)
pmi=[]
for row in pw['Cartera'].iter_rows(min_row=7,values_only=True):
    if row[4] is None: continue
    pmi.append(dict(Proyecto=str(row[4]),Nombre=row[7],Tipo=row[6],Prioridad=row[0],Prelacion=row[1],Costo=row[11],DevHistorico=row[12],PIM=row[13],**{f'PMI{2026+j}':row[14+j] for j in range(4)}))
assert len(pmi)==370 and len({r['Proyecto'] for r in pmi})==370
aw=openpyxl.load_workbook(ROOT/'0001917176_1_SINVERSION (1).xlsm',read_only=True,data_only=False)
amap=collections.defaultdict(list)
for sn in ['Gerencia de Infraestructura','Gerencias y Oficinas']:
    ue='001';ger='';mode=''
    for row in aw[sn].iter_rows():
        a=str(row[0].value or '').strip();c=str(row[2].value or '').strip()
        m=re.match(r'^(\d{3})[.\-]',a)
        if m: ue=m[1];ger=names.get(ue,'');mode=''
        if a.startswith(('GERENCIA','SUB GERENCIA','PROYECTO ESPECIAL','DIRECCIÓN','PROYECTOS SIN')):
            ger=a;mode=''
        if a.startswith('EJECUCIÓN POR'): mode=a
        m=re.match(r'^(\d{7})\D',c)
        if m: amap[(ue,m[1])].append((ger,mode,f'{sn}!C{row[2].row}'))
classifications=[]
for (u,c),rs in amap.items():
    gs=sorted({r[0] for r in rs if r[0]}); ms=sorted({r[1] for r in rs if r[1]})
    classifications.append(dict(UE=u,Proyecto=c,Gerencia=' / '.join(gs),Modalidad=' / '.join(ms),Fuente='; '.join(r[2] for r in rs),Revision='Revisar múltiples clasificaciones' if len(gs)>1 or len(ms)>1 else 'Antecedente Alex; pendiente de validación'))
print('execution',len(exrows),len(execution),'pmi',len(pmi),'classifications',len(classifications))
print('CUI sin PMI',sum(p['Proyecto'] not in {x['Proyecto'] for x in pmi} for p in projects))
out=dict(pdf_rows=records,projects=projects,ue_names=names,ue_totals=totals,grand=grand,execution=execution,execution_rows=exrows,pmi=pmi,classifications=classifications)
pathlib.Path('work/data.json').write_text(json.dumps(out,ensure_ascii=False),encoding='utf-8')
