"""Prueba aparte: ¿puede la API del MEF reemplazar el reporte SIAF (.xls)?
Compara, a nivel UE + código + meta, la API contra el XLS del SIAF (proyectos del pliego y gasto corriente de la UE 001).
Uso: python probar_api_reemplazo_siaf.py [ruta_PLIEGO.xls] [salida.xlsx]
"""
import collections, datetime, pathlib, sys
import xlsxwriter
from sinv import api_mef
from sinv.siaf import leer_siaf, texto_modalidad, MOD_RULES

RAIZ = pathlib.Path(__file__).resolve().parent
SIAF = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ.parent / 'ref docs' / 'nuevo' / 'PLIEGO--.xls'
OUT = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else RAIZ.parent / 'outputs' / f'Prueba_API_vs_SIAF_{datetime.date.today():%Y-%m-%d}.xlsx'
CAMPOS = [('pim', 'pim', 'PIM'), ('cert', 'cert', 'Certificación'), ('comp', 'comp', 'Compromiso anual'), ('dev', 'dev', 'Devengado')]
TOL = 1.0


def etiqueta_mod(esp):
    g, sg, sgd, e, ed, nombre = esp
    for k, v in MOD_RULES:
        if k in (nombre or ''):
            return v
    if g == '6' and sg == '3':
        return 'Equipamiento'
    if (g, sg, sgd, e) in (('6', '8', '1', '4'), ('6', '7', '1', '6')):
        return 'Personal/bienes/servicios'
    return 'Otro (' + (nombre or '')[:30] + ')'


def mod_api(m):
    pos = collections.Counter()
    for esp, v in m['espec'].items():
        if v > 0:
            pos[etiqueta_mod(esp)] += v
    return ' / '.join(k for k, _ in pos.most_common())


def comparar(api, siaf, solo_ue=None):
    sia = {(m['ue'], m['cui'], m['meta']): m for m in siaf if solo_ue is None or m['ue'] == solo_ue}
    res = dict(n_api=len(api), n_siaf=len(sia), solo_api=sorted(set(api) - set(sia)), solo_siaf=sorted(set(sia) - set(api)), campos={}, difs=[])
    comunes = [k for k in sia if k in api]
    for a, s, nombre in CAMPOS:
        dif = [(k, api[k][a], sia[k][s]) for k in comunes if abs(api[k][a] - sia[k][s]) > TOL]
        res['campos'][nombre] = dict(metas_dif=len(dif), de=len(comunes), api=sum(api[k][a] for k in comunes), siaf=sum(sia[k][s] for k in comunes))
        res['difs'] += [(k, nombre, x, y) for k, x, y in dif]
    return res, sia, comunes


def main():
    t0 = datetime.datetime.now()
    cli = api_mef.Cliente()
    siaf = leer_siaf(SIAF)
    # ---------------- proyectos del pliego
    filas_p = cli.descargar(api_mef.PROYECTOS)
    api_p = api_mef.agregar_metas(filas_p)
    sia_p = [m for m in siaf['metas'] if m['tipo'] == 'proyecto']
    res_p, sia_idx, comunes = comparar(api_p, sia_p)
    mod_dif = []
    for k in comunes:
        a = set(mod_api(api_p[k]).split(' / ')) - {''}; s = set(texto_modalidad(sia_idx[k]).split(' / '))
        if a and a != s and not (api_p[k]['pim'] <= 0):
            mod_dif.append((k, mod_api(api_p[k]), texto_modalidad(sia_idx[k])))
    fn_dif = [k for k in comunes if sia_idx[k]['fn'] != api_p[k]['fn']]
    # ---------------- gasto corriente UE 001
    filas_g = cli.descargar({'ANO_EJE': '2026', 'PLIEGO': '458', 'TIPO_ACT_PROY': '3', 'EJECUTORA': '001'})
    api_g = api_mef.agregar_metas(filas_g)
    sia_g = [m for m in siaf['metas'] if m['tipo'] == 'actividad' and m['ue'] == '001']
    res_g, _, _ = comparar(api_g, sia_g, '001')
    seg = (datetime.datetime.now() - t0).seconds

    wb = xlsxwriter.Workbook(str(OUT)); b = dict(font_name='Arial', font_size=9)
    hd = wb.add_format({**b, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#002060', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
    num = wb.add_format({**b, 'num_format': '#,##0.00;(#,##0.00);"-"'}); txt = wb.add_format(b); bold = wb.add_format({**b, 'bold': True}); tit = wb.add_format({**b, 'bold': True, 'font_size': 12})
    ws = wb.add_worksheet('Resumen')
    ws.write(0, 0, 'Prueba: ¿puede la API del MEF reemplazar al reporte SIAF (.xls)?', tit)
    ws.write(1, 0, f'API: recurso {api_mef.RID}, año 2026, pliego 458. SIAF: {SIAF.name}. Tolerancia por meta: S/ {TOL:.0f}. Descarga total en {seg} s.', txt)
    fila = [3]

    def bloque(titulo, res, extra):
        ws.write(fila[0], 0, titulo, bold); fila[0] += 1
        for j, h in enumerate(['Concepto', 'Metas en API', 'Metas en SIAF', 'Metas con diferencia', 'Suma API', 'Suma SIAF', 'Diferencia']):
            ws.write(fila[0], j, h, hd)
        fila[0] += 1
        for nombre, d in res['campos'].items():
            ws.write(fila[0], 0, nombre, txt); ws.write_number(fila[0], 1, res['n_api'], txt); ws.write_number(fila[0], 2, res['n_siaf'], txt)
            ws.write(fila[0], 3, f"{d['metas_dif']} de {d['de']}", txt); ws.write_number(fila[0], 4, d['api'], num); ws.write_number(fila[0], 5, d['siaf'], num)
            ws.write_number(fila[0], 6, d['api'] - d['siaf'], num); fila[0] += 1
        ws.write(fila[0], 0, f"Metas solo en API: {len(res['solo_api'])} · solo en SIAF: {len(res['solo_siaf'])}", txt); fila[0] += 1
        for e in extra:
            ws.write(fila[0], 0, e, txt); fila[0] += 1
        fila[0] += 1
    bloque('PROYECTOS (todas las UE del pliego)', res_p, [f'Modalidad distinta: {len(mod_dif)} de {len(comunes)} metas · Función (código) distinta: {len(fn_dif)}',
                                                          'La API trae además: nombre del proyecto, nombre de la función, nombre de la UE, rubro por meta y específica de gasto.'])
    bloque('GASTO CORRIENTE (solo UE 001, actividades)', res_g, [])
    ws.set_column(0, 0, 60); ws.set_column(1, 6, 18)
    d = wb.add_worksheet('Diferencias proyectos')
    for j, h in enumerate(['UE', 'CUI', 'Meta', 'Concepto', 'API', 'SIAF', 'API − SIAF']):
        d.write(0, j, h, hd)
    for i, (k, nombre, x, y) in enumerate(sorted(res_p['difs'], key=lambda t: -abs(t[2] - t[3])), 1):
        d.write(i, 0, k[0], txt); d.write(i, 1, k[1], txt); d.write(i, 2, k[2], txt); d.write(i, 3, nombre, txt)
        d.write_number(i, 4, x, num); d.write_number(i, 5, y, num); d.write_number(i, 6, x - y, num)
    d.set_column(0, 3, 12); d.set_column(4, 6, 18)
    g = wb.add_worksheet('Diferencias gasto corriente')
    for j, h in enumerate(['UE', 'Código', 'Meta', 'Concepto', 'API', 'SIAF', 'API − SIAF']):
        g.write(0, j, h, hd)
    for i, (k, nombre, x, y) in enumerate(sorted(res_g['difs'], key=lambda t: -abs(t[2] - t[3])), 1):
        g.write(i, 0, k[0], txt); g.write(i, 1, k[1], txt); g.write(i, 2, k[2], txt); g.write(i, 3, nombre, txt)
        g.write_number(i, 4, x, num); g.write_number(i, 5, y, num); g.write_number(i, 6, x - y, num)
    g.set_column(0, 3, 12); g.set_column(4, 6, 18)
    m = wb.add_worksheet('Modalidad distinta')
    for j, h in enumerate(['UE', 'CUI', 'Meta', 'Modalidad según API', 'Modalidad según SIAF']):
        m.write(0, j, h, hd)
    for i, (k, a, s) in enumerate(mod_dif, 1):
        for j, v in enumerate((k[0], k[1], k[2], a, s)):
            m.write(i, j, v, txt)
    m.set_column(0, 2, 10); m.set_column(3, 4, 70)
    wb.close()
    print('Proyectos:', {n: f"{d['metas_dif']}/{d['de']}" for n, d in res_p['campos'].items()}, 'solo API', len(res_p['solo_api']), 'solo SIAF', len(res_p['solo_siaf']), 'mod dif', len(mod_dif), 'fn dif', len(fn_dif))
    print('Gasto corriente UE001:', {n: f"{d['metas_dif']}/{d['de']}" for n, d in res_g['campos'].items()}, 'solo API', len(res_g['solo_api']), 'solo SIAF', len(res_g['solo_siaf']))
    for n, d in res_g['campos'].items():
        print(f'  {n}: API {d["api"]:,.0f} vs SIAF {d["siaf"]:,.0f}')
    print('segundos', seg, '->', OUT)


if __name__ == '__main__':
    main()
