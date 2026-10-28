import json,re,collections,pathlib,openpyxl,sys
sys.path.insert(0,'.analysis-deps')
import xlrd
d=json.loads(pathlib.Path('work/data.json').read_text(encoding='utf-8'))
w=openpyxl.load_workbook('ref docs/0001917176_1_SINVERSION (1).xlsm',data_only=False)
src=collections.defaultdict(list)
for r in d['execution_rows']:src[(r['Proyecto'],str(int(r['Meta'])))].append(r)
metanames={};s=xlrd.open_workbook('ref docs/reporte20261006085201.xls').sheet_by_index(0);code=meta=None
for i in range(s.nrows):
 t=str(s.cell_value(i,1));m=re.match(r'^(\d{4})\s+\d{4}\s+(\d{7})\s',t)
 if m:meta,code=m.groups()
 if t.startswith('Meta:'):
  n=re.sub(r'^Meta:\s*\d+\s*-\s*\d+\s*','',t).split(';')[0]
  metanames[(code,str(int(meta)))]=n
primary={};assigned=set();lookup={};spec={}
for sn,last in [('Gerencia de Infraestructura',905),('Gerencias y Oficinas',314)]:
 s=w[sn];leaves=[];heads=[];mode=''
 for row in s.iter_rows(max_row=last):
  a=str(row[0].value or '').strip();c=str(row[2].value or '');r=row[0].row;m=re.match(r'^(\d{7})\D',c);meta=row[1].value
  if a.startswith('EJECUCIÓN POR'):mode=a
  if a.startswith(('SUB GERENCIA','PROYECTOS SIN','GERENCIA','DIRECCIÓN','PROYECTO ESPECIAL')):mode=''
  if m:
   key=(m[1],str(int(meta))) if str(meta or '').isdigit() else None
   if key in src and key not in assigned:
    leaves.append(dict(row=r,cui=key[0],meta=key[1],mode=mode,key='|'.join(key)));assigned.add(key);lookup[key]=(sn,r)
  if a and not a.startswith('=') and row[3].value is not None and not m and not c:
   if a.startswith('TOTAL'):continue
   if sn=='Gerencia de Infraestructura':rank=0 if r==5 else 1 if r==6 else 2 if r in [7,175,210,307] else 3 if r in [8,128] else 4
   else:rank=1 if r in [6,192,217,224,231,300] else 2
   heads.append((r,rank))
 totals={}
 for r,rank in heads:
  end=next((rr for rr,rk in heads if rr>r and rk<=rank),last)
  totals[str(r)]=[l['row'] for l in leaves if r<l['row']<end]
 totals[str(last)]=[l['row'] for l in leaves]
 primary[sn]=dict(last=last,leaves=leaves,totals=totals)
missing=sorted(set(src)-assigned)
for i,k in enumerate(missing,317):lookup[k]=('Gerencias y Oficinas',i)
aliases={}
for sn in ['GRDE','GRRNyDA','PETT','PECSA','DS','DEPI']:
 s=w[sn];leaves=[]
 for row in s:
  r=row[0].row;c=str(row[2].value or '');m=re.match(r'^(\d{7})\D',c)
  if not m:continue
  # Prefer the original link to the primary report; its meta mapping is authoritative.
  link=re.search(r"'Gerencias y Oficinas'!D(\d+)",str(row[3].value or ''))
  match=next((l for l in primary['Gerencias y Oficinas']['leaves'] if link and l['row']==int(link[1])),None)
  meta=row[1].value;k=(m[1],str(int(meta))) if str(meta or '').isdigit() else None
  if match: k=(match['cui'],match['meta'])
  if k in lookup:leaves.append(dict(row=r,cui=k[0],meta=k[1],primary_sheet=lookup[k][0],primary_row=lookup[k][1]))
 # Alias totals keep the original grouping boundaries but sum only current detail.
 heads=[]
 for row in s:
  a=str(row[0].value or '').strip();c=str(row[2].value or '')
  if a and not a.startswith('=') and not c and row[3].value is not None and not a.startswith('TOTAL'):
   heads.append((row[0].row,1 if row[0].row==6 else 2))
 totals={}
 for r,rank in heads:
  end=next((rr for rr,rk in heads if rr>r and rk<=rank),s.max_row)
  totals[str(r)]=[l['row'] for l in leaves if r<l['row']<end]
 aliases[sn]=dict(last=s.max_row,leaves=leaves,totals=totals)
result=dict(primary=primary,aliases=aliases,missing=[dict(cui=k[0],meta=k[1],name=metanames[k],row=317+i) for i,k in enumerate(missing)],totals={k:sum(r[k] for r in d['execution_rows']) for k in ['PIM','Certificacion','Devengado']})
pathlib.Path('work/alex_spec.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
print('matched',len(assigned),'missing',len(missing),'totals',result['totals'])
print('aliases', {k:len(v['leaves']) for k,v in aliases.items()})
