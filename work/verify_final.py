import pathlib,json,openpyxl,zipfile,collections
path=pathlib.Path('outputs/gore_puno_v1/GORE_Puno_Inversiones_2026.xlsx')
d=json.loads(pathlib.Path('work/data.json').read_text(encoding='utf-8'))
w=openpyxl.load_workbook(path,data_only=True,read_only=True)
def rows(sheet):
 s=w[sheet];headers=[c.value for c in next(s.iter_rows(min_row=5,max_row=5))]
 return [dict(zip(headers,r)) for r in s.iter_rows(min_row=6,values_only=True) if r[0] is not None]
pdf=rows('Certificacion PDF');pro=rows('Proyectos');pmi=rows('PMI');ex=rows('Ejecucion SIAF')
assert len(pdf)==450 and len(pro)==450 and len(pmi)==370 and len(ex)==490
pm={(str(r['UE']).zfill(3),str(r['CUI'])):r for r in pdf}
for r in d['projects']:
 n=pm[(r['UE'],r['Proyecto'])]
 for k in ['PIA','Modificaciones','PIM','Certificacion','Saldo']: assert abs(r[k]-n[k])<.02,(k,r,n)
for r in pro:
 ue=str(r['UE']).zfill(3);c=str(r['CUI'])
 if ue!='001': assert all(r[k]=='No disponible' for k in ['Devengado','Compromiso anual','PIM ejecucion'])
 else:
  e=d['execution'][c]
  for k,j in [('Devengado','Devengado'),('Compromiso anual','Compromiso'),('PIM ejecucion','PIM')]: assert abs(r[k]-e[j])<.02
  assert abs(r['Saldo sin devengar']-(e['PIM']-e['Devengado']))<.02
 assert r['Restriccion']=='Pendiente de evaluación'
 assert r['Observacion'] in ('',None),r
assert len(set(r['Clave'] for r in pro))==450
assert sum(r['Cruce PMI']=='Sin coincidencia' for r in pro)==142
for r in pmi:
 source=next(x for x in d['pmi'] if x['Proyecto']==str(r['CUI']))
 for y in range(2026,2030):
  v=source[f'PMI{y}']; actual=r[f'PMI {y}']
  if v in ('',None): assert actual in ('',None)
  else: assert isinstance(actual,(int,float)) and abs(actual-v)<.02
errors=[]
for s in w:
 for row in s:
  for c in row:
   if c.data_type=='e':errors.append((s.title,c.coordinate,c.value))
assert not errors,errors
for c,v in [('D15',987989928),('E15',919264637.96),('F15',68725290.04),('H7',433575772.65)]:assert abs(w['Resumen'][c].value-v)<.02
with zipfile.ZipFile(path) as z:
 assert 'xl/connections.xml' in z.namelist()
 assert len([n for n in z.namelist() if n.startswith('xl/queryTables/queryTable') and n.endswith('.xml')])==5
 assert not any(n.startswith('xl/externalLinks/') for n in z.namelist())
 assert not any(n.endswith('vbaProject.bin') for n in z.namelist())
print(json.dumps({'projects':len(pro),'pmi':len(pmi),'execution_detail':len(ex),'missing_pmi':142,'native_query_tables':5,'formula_errors':0,'all_rows_reconciled':True},indent=2))
