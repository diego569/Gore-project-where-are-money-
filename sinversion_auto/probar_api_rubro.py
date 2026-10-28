"""Prueba de la API de Datos Abiertos del MEF: proyectos del Gobierno Regional de Puno (pliego 458) por unidad ejecutora, CUI y rubro.

Hallazgos que motivan este script (07/10/2026):
  - datastore_search_sql responde 404 (la ruta no existe); se usa datastore_search con "filters".
  - El certificado SSL del MEF estaba vencido: se intenta con verificación y, solo si falla por eso, se repite sin verificar
    (son datos públicos de solo lectura; no se envía ninguna credencial).
  - Los registros son mensuales: PIA/PIM vienen en MES_EJE = 0; certificado, compromiso y devengado son flujos por mes (se suman).
"""
import concurrent.futures as cf, collections, datetime, json, pathlib, sys, time
import requests, urllib3, xlsxwriter

URL = 'https://api.datosabiertos.mef.gob.pe/DatosAbiertos/v1/datastore_search'
RID = '615644aa-ef73-4358-b4e0-0c20931632f3'
FILTRO = {'ANO_EJE': '2026', 'PLIEGO': '458', 'TIPO_ACT_PROY': '2'}      # 2 = proyectos; 3 = actividades (gasto corriente)
PAG = 5000
estado_ssl = {'verificar': True}


def get(params):
    for intento in range(4):
        try:
            return requests.get(URL, params=params, timeout=400, verify=estado_ssl['verificar']).json()
        except requests.exceptions.SSLError:
            if not estado_ssl['verificar']:
                raise
            print('AVISO: certificado SSL del MEF inválido/vencido; se continúa sin verificarlo (datos públicos, solo lectura).')
            estado_ssl['verificar'] = False; urllib3.disable_warnings()
        except Exception as e:
            print('reintento', str(e)[:80]); time.sleep(5)
    raise SystemExit('No se pudo consultar la API del MEF.')


def descargar(filtro=FILTRO):
    p = dict(resource_id=RID, filters=json.dumps(filtro))
    total = int(get({**p, 'limit': 1})['result']['include_total'])
    offs = list(range(0, total, PAG))
    with cf.ThreadPoolExecutor(6) as ex:
        pags = list(ex.map(lambda o: get({**p, 'limit': PAG, 'offset': o})['records'], offs))
    filas = [r for pg in pags for r in pg]
    assert len(filas) == total, (len(filas), total)
    return filas


def agrupar(filas):
    f = lambda x: float(x or 0)
    g = collections.OrderedDict()
    for r in filas:
        k = (r['EJECUTORA'], r['PRODUCTO_PROYECTO'], r['RUBRO'])
        d = g.setdefault(k, dict(ue=k[0], ue_nombre=r['EJECUTORA_NOMBRE'], cui=k[1], proyecto=r['PRODUCTO_PROYECTO_NOMBRE'], rubro=k[2], rubro_nombre=r['RUBRO_NOMBRE'],
                                 pia=0.0, pim=0.0, cert=0.0, comp_mes=0.0, dev=0.0, giro=0.0))
        d['pia'] += f(r['MONTO_PIA']); d['pim'] += f(r['MONTO_PIM']); d['cert'] += f(r['MONTO_CERTIFICADO'])
        d['comp_mes'] += f(r['MONTO_COMPROMETIDO']); d['dev'] += f(r['MONTO_DEVENGADO']); d['giro'] += f(r['MONTO_GIRADO'])
    return sorted(g.values(), key=lambda d: (d['ue'], d['cui'], d['rubro']))


def main():
    t = time.time()
    filas = descargar()
    datos = agrupar(filas)
    meses = sorted({int(r['MES_EJE']) for r in filas})
    out = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else pathlib.Path(__file__).resolve().parent.parent / 'outputs' / f'proyectos_por_rubro_API_{datetime.date.today():%Y-%m-%d}.xlsx'
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(out))
    b = dict(font_name='Arial', font_size=9)
    hd = wb.add_format({**b, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#002060', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
    num = wb.add_format({**b, 'num_format': '#,##0;(#,##0);"-"'}); txt = wb.add_format(b); bold = wb.add_format({**b, 'bold': True})
    ws = wb.add_worksheet('Por UE-CUI-rubro')
    cab = ['UE', 'Unidad ejecutora', 'CUI', 'Proyecto', 'Rubro', 'Nombre del rubro', 'PIA', 'PIM', 'Certificación', 'Compromiso (flujo mensual acumulado)', 'Devengado', 'Girado']
    for j, h in enumerate(cab): ws.write(0, j, h, hd)
    ws.set_row(0, 40)
    for i, d in enumerate(datos, 1):
        for j, k in enumerate(['ue', 'ue_nombre', 'cui', 'proyecto', 'rubro', 'rubro_nombre']): ws.write_string(i, j, d[k], txt)
        for j, k in enumerate(['pia', 'pim', 'cert', 'comp_mes', 'dev', 'giro']): ws.write_number(i, 6 + j, d[k], num)
    ws.freeze_panes(1, 4); ws.autofilter(0, 0, len(datos), len(cab) - 1)
    for j, w in enumerate([5, 36, 9, 60, 7, 40, 14, 14, 14, 18, 14, 14]): ws.set_column(j, j, w)
    r2 = wb.add_worksheet('Resumen por rubro')
    por = collections.OrderedDict()
    for d in datos:
        x = por.setdefault((d['rubro'], d['rubro_nombre']), collections.Counter())
        for k in ('pia', 'pim', 'cert', 'comp_mes', 'dev', 'giro'): x[k] += d[k]
    for j, h in enumerate(['Rubro', 'Nombre', 'PIA', 'PIM', 'Certificación', 'Compromiso mensual acum.', 'Devengado', 'Girado']): r2.write(0, j, h, hd)
    for i, ((c, n), x) in enumerate(sorted(por.items()), 1):
        r2.write_string(i, 0, c, txt); r2.write_string(i, 1, n, txt)
        for j, k in enumerate(('pia', 'pim', 'cert', 'comp_mes', 'dev', 'giro')): r2.write_number(i, 2 + j, x[k], num)
    n = len(por) + 1
    r2.write(n, 1, 'TOTAL', bold)
    for j in range(2, 8): r2.write_formula(n, j, f'=SUM({chr(65 + j)}2:{chr(65 + j)}{n})', num, sum(x[('pia', 'pim', 'cert', 'comp_mes', 'dev', 'giro')[j - 2]] for x in por.values()))
    r2.set_column(0, 0, 7); r2.set_column(1, 1, 55); r2.set_column(2, 7, 18)
    r2.write(n + 2, 0, f'Fuente: API Datos Abiertos MEF (recurso {RID}), año 2026, pliego 458, solo proyectos. Meses con datos: {meses[0]}-{meses[-1]} (mes 0 = PIA/PIM). Descargado {datetime.datetime.now():%d/%m/%Y %H:%M}.', txt)
    wb.close()
    print(f'{len(filas)} registros -> {len(datos)} filas UE-CUI-rubro en {time.time() - t:.0f} s')
    print('PIM total', f"{sum(d['pim'] for d in datos):,.0f}", '| archivo:', out)


if __name__ == '__main__':
    main()
