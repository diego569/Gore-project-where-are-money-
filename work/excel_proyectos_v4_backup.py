"""Construye el Excel de PROYECTOS (todas las UE con proyectos) en el formato de Alex + columnas nuevas.

La hoja Base es la fuente; las hojas de anexo leen de ella por fórmula (con valores ya calculados
para que se vean sin recalcular). Para gasto corriente se reutilizan siaf.py, maestro.py y reglas.py.
"""
import collections, datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

from .reglas import datos_cui, evaluar, SEMAFOROS
from .siaf import texto_modalidad
from .maestro import ubicar, HOJA_INFRA, HOJA_GER, INFRA
from .formato import orden_hojas, aplicar as aplicar_formato

SIN = 'SIN CLASIFICAR (pendiente de asignar gerencia)'
NOCART = 'No registrado en cartera'
RUBROS = [('00', 'RO'), ('09', 'RDR'), ('15', 'FONCOR'), ('18', 'CANON'), ('19', 'ROOC')]
BASE_KEYS = ['ue', 'hoja', 'ger', 'grupo', 'fun', 'cui', 'meta', 'nombre', 'tipo', 'sem', 'disponible', 'pot', 'motivo', 'pim', 'cert', 'comp', 'dev', 's_cert',
             'cert_nc', 's_dev', 'pdev', 'costo', 'acum', 'avance', 'prog', 'f12f', 'm26', 'm27', 'm28', 'm29', 'estado_inv', 'cierre', 'avfis', 'mod', 'bl',
             'obs', 'nota', 'ejec', 'antony']
BASE_COLS = ['UE', 'Hoja', 'Gerencia / Oficina', 'Grupo (Alex)', 'Función', 'CUI', 'Meta', 'Proyecto', 'Tipo de inversión', 'Semáforo', 'Disponible legal (S/)',
             'Potencial rebajando certificación (S/)', 'Motivo / condiciones', 'PIM 2026', 'Certificación', 'Compromiso anual', 'Devengado', 'Saldo sin certificar (A-B)', 'Certificado no comprometido',
             'Saldo sin devengar (A-C)', '% Devengado', 'Costo actualizado', 'Devengado acumulado', '% Avance financiero acumulado', 'Programación 12-B 2026 (total)',
             'Fecha 12-B', 'PMI 2026 (S/)', 'PMI 2027 (S/)', 'PMI 2028 (S/)', 'PMI 2029 (S/)', 'Estado en Invierte', 'Cierre registrado', 'Avance físico (%)',
             'Modalidad de ejecución', 'Base legal', 'Observación', 'Nota de clasificación', 'Ejecutora (Silucia)', 'Lista Antony']
L = {n: xl_col_to_name(i) for i, n in enumerate(BASE_KEYS)}
B0 = 5   # primera fila de datos de Base (1-indexada)


def _fun_label(fn, funcs):
    return f"{fn}. {funcs.get(fn, '(sin nombre)')}" if fn else '(sin función)'


def armar_filas(siaf, cartera, maestro, cfg, silucia=None, mapa=None, ssi=None, f12b=None):
    rows = []; usados = set(); pendientes = []
    ssi = ssi or {}; f12b = f12b or {}
    # códigos sin ninguna cifra (la API los devuelve en cero): no son proyectos con presupuesto
    metas = [m for m in siaf['metas'] if m['tipo'] == 'proyecto' and (m['pim'] > 0 or m.get('pia', 0) > 0 or m['cert'] or m['comp'] or m['dev'])]
    por_cui = collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for m in metas:
        a = por_cui[m['cui']]; a[0] += m['pim']; a[1] += m['cert']; a[2] += m['dev']; a[3] += m['comp']
    dcui = {}
    metas.sort(key=lambda m: (m['ue'], -m['pim'], m['cui'], m['meta']))
    car = cartera['cartera']; funcs = maestro['funcs']
    for m in metas:
        ue = m['ue']; c = car.get(m['cui'])
        pl, nota = (None, '')
        if ue == '001':
            pl, nota = ubicar(maestro, ue, m['cui'], m['meta'], usados)
        nombre_ref = (c and c['nombre']) or (pl and pl['nombre']) or f"({m['nombre']} - nombre no disponible en fuentes)"
        nombre = f"{m['cui']}: {str(nombre_ref).strip()}"
        fl = _fun_label(m['fn'], funcs)
        sug = (silucia or {}).get((m['cui'], m['meta']))
        ejec = sug['ejecutora'] if sug else ''
        ger_sug = (mapa or {}).get(ejec) if (ejec and ue == '001') else None
        if ue == '001' and not (pl and pl['gerencia']) and ger_sug:
            hoja = HOJA_INFRA if ger_sug == INFRA else HOJA_GER
            grupo = '(grupo por asignar)' if hoja == HOJA_INFRA else ''
            path = (ger_sug,) + ((grupo, fl) if hoja == HOJA_INFRA else (fl,))
            orden = 10 ** 7 - m['pim']
            nota = (nota + ' · ' if nota else '') + f'Gerencia sugerida por Silucia (ejecutora "{ejec}"): confirmar en maestro.xlsx'
            if pl is None:
                pendientes.append(['001', m['cui'], m['meta'], hoja, ger_sug, '' if hoja != HOJA_INFRA else '', fl, 10 ** 6,
                                   'Silucia (sugerido) ' + str(datetime.date.today()), nombre[:300]])
        elif ue == '001' and pl and pl['gerencia']:
            hoja = pl['hoja'] or (HOJA_INFRA if pl['gerencia'] == INFRA else HOJA_GER)
            if hoja == HOJA_INFRA:
                grupo = pl['grupo'] or '(grupo por asignar)'
                path = (pl['gerencia'],) + tuple(grupo.split(' > ')) + (pl['funcion'] or fl,)
            else:
                grupo = ''
                path = (pl['gerencia'],) + ((pl['funcion'],) if pl['funcion'] else ())
            orden = pl['orden']
        elif ue == '001':
            hoja = 'Sin clasificar'; grupo = ''; path = (SIN, fl); orden = 10 ** 7 - m['pim']
            if pl is None:
                pendientes.append(['001', m['cui'], m['meta'], '', '', '', fl, 10 ** 6, 'Nuevo ' + str(datetime.date.today()), nombre[:300]])
        else:
            hoja = 'Otras UE'; grupo = ''; path = (f"{ue}. {siaf['ues'][ue]}", fl); orden = 10 ** 7 - m['pim']
        comp = maestro['comp'].get((ue, m['cui'], m['meta'])) or maestro['comp'].get((ue, m['cui'], '')) or {}
        if c:
            montos = [c['m2026'] or 0, c['m2027'] or 0, c['m2028'] or 0, c['m2029'] or 0]
        else:
            montos = [NOCART] * 4
        if m['cui'] not in dcui:
            pc = por_cui[m['cui']]
            dcui[m['cui']] = datos_cui(m['cui'], pc[0], pc[1], pc[2], pc[3], ssi.get(m['cui']), f12b.get(m['cui']), c, cfg)
        d = dcui[m['cui']]
        sem, dispo, motivo, pot = evaluar(m['cui'], m['pim'], m['cert'], m['comp'], d, cfg)
        rows.append(dict(ue=ue, hoja=hoja, ger=path[0], grupo=grupo, fun=path[-1] if len(path) > 1 else '', path=path, orden=orden, cui=m['cui'],
                         meta=m['meta'], nombre=nombre, pim=m['pim'], cert=m['cert'], comp=m['comp'], dev=m['dev'], sem=sem, d=d, tipo=d['tipo'], motivo=motivo, pot=pot,
                         disponible=dispo, mod=texto_modalidad(m), montos=montos, bl=comp.get('base_legal', ''), obs=comp.get('obs', ''),
                         nota=nota, car=c, fn=m['fn'], ejecutora=ejec))
    rows.sort(key=lambda r: (r['ue'] != '001', r['ue'], r['orden']))
    for i, r in enumerate(rows):
        r['n'] = B0 + i
    return rows, pendientes


def armar_rubros(rows, rubros, api=None):
    """Rubro por (UE, CUI). Fuente 1: API del MEF (separa por UE). Fuente 2 (respaldo): Consulta Amigable, a nivel pliego."""
    rubros = rubros or {}
    por = collections.OrderedDict(); ues_cui = collections.defaultdict(set)
    for r in rows:
        d = por.setdefault((r['ue'], r['cui']), dict(cui=r['cui'], ue=r['ue'], nombre=r['nombre'], pim=0.0))
        d['pim'] += r['pim']; ues_cui[r['cui']].add(r['ue'])
    std_codes = [c for c, _ in RUBROS]
    extras = [c for c in rubros if c not in std_codes]
    out = []
    for (ue, cui), d in por.items():
        compartido = False
        if api is not None:
            a = api['pim'].get((ue, cui), {})
            v = {c: a.get(c, 0.0) for c in std_codes}
            otros = sum(x for k, x in a.items() if k not in std_codes) or None
            fuente = 'API MEF'
            known = sum(v.values()) + (otros or 0.0)
            sin = d['pim'] - known
            est = 'Completo' if abs(sin) <= 1 else 'La API no suma el PIM del SIAF (diferencia ' + format(sin, ',.0f') + ')'
        else:
            fuente = 'Consulta Amigable'
            compartido = len(ues_cui[cui]) > 1
            v = {}
            for code in std_codes + extras:
                l = rubros.get(code)
                if compartido or l is None:
                    v[code] = None
                else:
                    pr = l['proyectos'].get(cui)
                    v[code] = pr['pim'] if pr else (0.0 if l['completo'] else None)
            otros = None if (not extras or compartido) else sum(x for x in (v[c] for c in extras) if x is not None)
            known = sum(x for x in (v[c] for c in std_codes) if x is not None) + (otros or 0.0)
            sin = d['pim'] - known
            faltan = [c + (' (incompleto)' if c in rubros else ' (sin archivo)') for c in std_codes if v[c] is None]
            if compartido:
                est = 'CUI en varias UE: la Consulta Amigable solo da el rubro a nivel pliego'
            elif faltan:
                est = 'Faltan rubros: ' + ', '.join(faltan)
            elif abs(sin) > 1:
                est = 'Los rubros no suman el PIM (diferencia ' + format(sin, ',.0f') + ')'
            else:
                est = 'Completo'
        out.append(dict(cui=cui, ue=ue, nombre=d['nombre'], pim=d['pim'], v=v, std=[v[c] for c in std_codes], otros=otros, sin=sin, estado=est,
                        compartido=compartido, fuente=fuente))
    return out


class Nodo:
    def __init__(self, nombre, nivel):
        self.nombre = nombre; self.nivel = nivel; self.hijos = collections.OrderedDict(); self.hojas = []; self.fila = None


def arbol(rs, clave=lambda r: r['path']):
    raiz = Nodo('', 0)
    for r in rs:
        n = raiz
        for lv, nm in enumerate(clave(r), 1):
            n = n.hijos.setdefault(nm, Nodo(nm, lv))
        n.hojas.append(r)
    return raiz


def construir(path, siaf, cartera, rows, cfg, meta_info, rubros=None, api=None):
    wb = xlsxwriter.Workbook(str(path), {'nan_inf_to_errors': True})
    wb.set_calc_mode('auto')
    base = dict(font_name='Arial', font_size=9)
    F = lambda **k: wb.add_format({**base, **k})
    f_txt = F(); f_wrap = F(text_wrap=True, valign='vcenter'); f_num = F(num_format='#,##0;(#,##0);"-"', align='right', valign='vcenter')
    f_in = F(num_format='#,##0;(#,##0);"-"', font_color='#0000FF', align='right')
    f_pct = F(num_format='0.0%', align='center', valign='vcenter'); f_ctr = F(align='center', valign='vcenter')
    f_bold = F(bold=True); f_title = F(bold=True, font_size=12); f_note = F(italic=True, font_color='#595959')
    f_hd = F(bold=True, font_color='#FFFFFF', bg_color='#002060', text_wrap=True, align='center', valign='vcenter', border=1)
    f_hd2 = F(bold=True, italic=True, bg_color='#B4C6E7', text_wrap=True, align='center', valign='vcenter', border=1)
    sub = {}
    for nm, col in (('yel', '#FFFF00'), ('sal', '#F4B084'), ('gry', '#D9D9D9'), ('navy', '#002060')):
        fc = '#FFFFFF' if nm == 'navy' else '#000000'
        sub[nm] = (F(bold=True, bg_color=col, font_color=fc, border=1), F(bold=True, bg_color=col, font_color=fc, border=1, num_format='#,##0;(#,##0);"-"'),
                   F(bold=True, bg_color=col, font_color=fc, border=1, num_format='0.0%', align='center'))
    f_leaf = F(border=1, text_wrap=True, valign='vcenter'); f_leafn = F(border=1, num_format='#,##0;(#,##0);"-"', align='right', valign='vcenter')
    f_leafp = F(border=1, num_format='0.0%', align='center', valign='vcenter'); f_leafc = F(border=1, align='center', valign='vcenter')
    colores = {'Disponible': '#C6EFCE', 'Con condiciones': '#FFEB9C', 'Rebajando certificación': '#DDEBF7', 'Restringido': '#F8CBAD', 'No mover': '#FFC7CE', 'Sin saldo útil': '#EDEDED', 'Bolsa genérica': '#D9E1F2'}

    def cond_disp(ws, rango):
        for k, col in colores.items():
            ws.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': f'"{k}"', 'format': wb.add_format({'bg_color': col})})

    fecha_txt = meta_info['fecha_corte'].strftime('%d/%m/%Y')
    ues_orden = sorted({r['ue'] for r in rows}, key=lambda u: (u != '001', u))
    nrows = len(rows); BL = B0 + nrows - 1

    # ------------------------------------------------------------ Config (se crea primero para que exista la referencia)
    anexos = []
    if any(r['hoja'] == HOJA_INFRA for r in rows): anexos.append(HOJA_INFRA)
    if any(r['hoja'] == HOJA_GER for r in rows): anexos.append(HOJA_GER)
    if any(r['hoja'] == 'Sin clasificar' for r in rows): anexos.append('Sin clasificar')
    if any(r['hoja'] == 'Otras UE' for r in rows): anexos.append('Otras UE')
    antony = meta_info.get('antony')
    AN = {it['cui']: it for it in (antony or {}).get('items', [])}
    nombres_h = ['Resumen', 'Base', 'Dónde hay plata'] + (['Lista Antony'] if antony else []) + ['Cambios'] + anexos + \
                ['Saldo al expediente', 'Rubro por CUI', 'Calidad de datos', 'Control', 'Reglas', 'Config']
    nombres_h = orden_hojas(nombres_h, meta_info.get('formato'))
    WS = {nm: wb.add_worksheet(nm) for nm in nombres_h}
    wsR = WS['Resumen']; wsB = WS['Base']; wsDH = WS['Dónde hay plata']; wsCH = WS['Cambios']; wsAN = WS.get('Lista Antony')
    wsA = {nm: WS[nm] for nm in anexos}
    wsS = WS['Saldo al expediente']; wsRB = WS['Rubro por CUI']; wsQ = WS['Calidad de datos']; wsC = WS['Control']; wsX = WS['Reglas']; wsG = WS['Config']
    wsG.write(0, 0, 'Configuración (editable)', f_title)
    wsG.write(2, 0, 'Parámetro', f_hd); wsG.write(2, 1, 'Valor', f_hd); wsG.write(2, 2, 'Significado', f_hd)
    f_inp = F(num_format='#,##0.##', bg_color='#FFF2CC', font_color='#0000FF', border=1)
    for k, (nm, v, txt) in enumerate([('Monto mínimo a listar (S/)', cfg['monto_minimo'], 'Saldo libre menor a este monto: semáforo "Sin saldo útil" (no entra en Dónde hay plata).'),
                                      ('Umbral devengado / PIM', cfg['umbral_ejecutado'], 'Devengado 2026 / PIM mayor o igual → "No mover".'),
                                      ('Umbral avance financiero acumulado', cfg['umbral_avance_acumulado'], 'Devengado acumulado / costo actualizado mayor → "No mover" (Art. 11 Ley 32513).')]):
        wsG.write(4 + k, 0, nm, f_txt); wsG.write_number(4 + k, 1, v, f_inp); wsG.write(4 + k, 2, txt, f_txt)
    wsG.write(3, 0, 'Estos valores son informativos: se cambian en config.json (monto_minimo, umbral_ejecutado, umbral_avance_acumulado) y se aplican en la siguiente corrida.', f_note)
    wsG.write(8, 0, 'Datos de esta corrida', f_bold)
    info = [('Fecha de corte', fecha_txt), ('Archivo SIAF', meta_info['siaf']), ('Fecha del archivo SIAF', meta_info['siaf_fecha']),
            ('Archivo Cartera', meta_info['cartera']), ('Fecha de la Cartera', cartera['fecha'].strftime('%d/%m/%Y %H:%M') if cartera['fecha'] else 'No disponible'),
            ('Generado', datetime.datetime.now().strftime('%d/%m/%Y %H:%M'))] + [(f'Fuente legal {i + 1}', t) for i, t in enumerate(meta_info.get('legal', []))]
    for i, (k, v) in enumerate(info):
        wsG.write(9 + i, 0, k, f_txt); wsG.write(9 + i, 1, v, f_txt)
    r0 = 9 + len(info) + 1
    wsG.write(r0, 0, 'Fuente del rubro de financiamiento', f_bold)
    if api is not None:
        wsG.write(r0 + 1, 0, 'API Datos Abiertos MEF', f_txt)
        wsG.write(r0 + 1, 1, 'USADA' + (' · certificado SSL no verificado' if api.get('sin_verificar') else ''), f_txt)
        wsG.write(r0 + 1, 2, f"{api['registros']:,} registros · meses {api['meses'][0]}-{api['meses'][-1]} · descargada {api['fecha']:%d/%m/%Y %H:%M} · rubro por UE y CUI", f_txt)
        r0 += 1
    else:
        wsG.write(r0 + 1, 0, 'API Datos Abiertos MEF', f_txt); wsG.write(r0 + 1, 1, 'NO USADA', f_txt)
        wsG.write(r0 + 1, 2, meta_info.get('api_detalle', ''), f_txt)
        wsG.write(r0 + 3, 0, 'Respaldo: reportes de rubro de la Consulta Amigable', f_bold)
        for j, h in enumerate(['Rubro', 'Estado', 'Detalle']):
            wsG.write(r0 + 4, j, h, f_hd)
        rb_rows = [(c, nm) for c, nm in RUBROS] + [(c, None) for c in (rubros or {}) if c not in dict(RUBROS)]
        for k, (c, nm) in enumerate(rb_rows):
            l = (rubros or {}).get(c)
            if l is None:
                est = 'NO CARGADO'; det = 'No hay reporte de este rubro ni copia de días anteriores.'
            else:
                est = ('Completo' if l['completo'] else 'INCOMPLETO') + ' · ' + str(l.get('origen', ''))
                det = f"{len(l['proyectos'])} proyectos · PIM del archivo {l['suma_pim']:,.0f} de {l['total']['pim']:,.0f}"
                if l.get('fecha'):
                    det += f" · consulta del {l['fecha']:%d/%m/%Y}"
                if not l['completo']:
                    det += ' · falta otra página del reporte (la Consulta Amigable parte los resultados de a 400)'
            wsG.write(r0 + 5 + k, 0, c + ' ' + (l['nombre'] if l else (nm or '')), f_txt); wsG.write(r0 + 5 + k, 1, est, f_txt); wsG.write(r0 + 5 + k, 2, det, f_txt)
    wsG.set_column(0, 0, 58); wsG.set_column(1, 1, 52); wsG.set_column(2, 2, 110)

    # ------------------------------------------------------------ Base
    wsB.write(0, 0, f'Base por UE-CUI-meta · proyectos · al {fecha_txt}', f_title)
    wsB.write(1, 0, 'Una fila por UE-CUI-meta. Filtre con las flechas de cada columna; Ctrl+B busca por CUI o nombre. Azul = dato de entrada; negro = fórmula o resultado de las reglas (hoja Reglas).', f_note)
    wsB.set_row(B0 - 2, 44)
    for j, h in enumerate(BASE_COLS):
        wsB.write(B0 - 2, j, h, f_hd)
    C = {k: q for q, k in enumerate(BASE_KEYS)}
    for r in rows:
        n = r['n']; i = n - 1; d = r['d']
        pdev = r['dev'] / r['pim'] if r['pim'] > 0 else 0
        for k in ('ue', 'hoja', 'ger', 'grupo', 'fun', 'cui', 'meta', 'nombre', 'tipo', 'sem'):
            wsB.write_string(i, C[k], r[k], f_wrap if k == 'nombre' else f_txt)
        wsB.write_number(i, C['disponible'], r['disponible'], f_num); wsB.write_number(i, C['pot'], r['pot'], f_num); wsB.write_string(i, C['motivo'], r['motivo'], f_txt)
        for k in ('pim', 'cert', 'comp', 'dev'):
            wsB.write_number(i, C[k], r[k], f_in)
        c = lambda nm: f'{L[nm]}{n}'
        wsB.write_formula(i, C['s_cert'], f"={c('pim')}-{c('cert')}", f_num, r['pim'] - r['cert'])
        wsB.write_formula(i, C['cert_nc'], f"={c('cert')}-{c('comp')}", f_num, r['cert'] - r['comp'])
        wsB.write_formula(i, C['s_dev'], f"={c('pim')}-{c('dev')}", f_num, r['pim'] - r['dev'])
        wsB.write_formula(i, C['pdev'], f"=IF({c('pim')}>0,{c('dev')}/{c('pim')},0)", f_pct, pdev)
        for k, v in (('costo', d['costo'] or None), ('acum', d['acum']), ('prog', d['prog'])):
            if v is None: wsB.write_string(i, C[k], 's/d', f_ctr)
            else: wsB.write_number(i, C[k], v, f_in)
        if d['avance'] is None: wsB.write_string(i, C['avance'], 's/d', f_ctr)
        else: wsB.write_number(i, C['avance'], d['avance'], f_pct)
        wsB.write_string(i, C['f12f'], str(d['f12_fecha'])[:10], f_ctr)
        for q, k in enumerate(('m26', 'm27', 'm28', 'm29')):
            v = r['montos'][q]
            if isinstance(v, str): wsB.write_string(i, C[k], v, f_txt)
            else: wsB.write_number(i, C[k], v, f_in)
        wsB.write_string(i, C['estado_inv'], d['estado_inv'], f_txt); wsB.write_string(i, C['cierre'], d['cierre'], f_txt)
        try: wsB.write_number(i, C['avfis'], float(d['avance_fisico']), f_num)
        except (TypeError, ValueError): wsB.write_string(i, C['avfis'], 's/d', f_ctr)
        for k in ('mod', 'bl', 'obs', 'nota', 'ejec'):
            wsB.write_string(i, C[k], r['ejecutora'] if k == 'ejec' else r[k], f_txt)
        wsB.write_string(i, C['antony'], '★ Antony' if r['cui'] in AN else '', f_txt)
    for j, w in enumerate([5, 22, 34, 28, 24, 9, 7, 60, 18, 15, 15, 15, 48, 14, 14, 14, 14, 14, 14, 14, 10, 15, 15, 12, 15, 11, 14, 14, 14, 14, 18, 18, 10, 30, 22, 22, 24, 28, 12]):
        wsB.set_column(j, j, w)
    wsB.freeze_panes(B0 - 1, 8); wsB.autofilter(B0 - 2, 0, BL - 1, len(BASE_COLS) - 1)
    cond_disp(wsB, f'{L["sem"]}{B0}:{L["sem"]}{BL}')

    # ------------------------------------------------------------ Dónde hay plata
    rb_list = armar_rubros(rows, rubros, api)
    rb_ue = {(x['ue'], x['cui']): x for x in rb_list}
    RUB_COD = [c for c, _ in RUBROS]; RUB_NOM = dict(RUBROS)
    pim_uc = collections.Counter()
    for r in rows:
        pim_uc[(r['ue'], r['cui'])] += r['pim']

    def rubros_fila(r):
        """PIM por rubro de esta fila (el del UE-CUI repartido entre sus metas) y el rubro principal."""
        x = rb_ue.get((r['ue'], r['cui'])); tot = pim_uc[(r['ue'], r['cui'])]; f = (r['pim'] / tot) if tot > 0 else 0.0
        vals = {c: (None if (x is None or x['v'][c] is None) else x['v'][c] * f) for c in RUB_COD}
        otros = None if (x is None or x['otros'] is None) else x['otros'] * f
        con = {c: v for c, v in vals.items() if v is not None}
        if otros:
            con['otros'] = otros
        if not con or max(con.values()) <= 0:
            pr = 's/d'
        else:
            k = max(con, key=con.get); pr = 'Otros' if k == 'otros' else f'{RUB_NOM[k]} ({k})'
        return vals, otros, pr

    ORD = {'Disponible': 0, 'Con condiciones': 1, 'Rebajando certificación': 2}
    util = sorted([r for r in rows if r['sem'] in ORD], key=lambda r: (ORD[r['sem']], -(r['disponible'] if r['sem'] != 'Rebajando certificación' else r['pot'])))
    for r in rows:
        r['rub'] = rubros_fila(r)
    wsDH.write(0, 0, f'¿Dónde hay plata? · proyectos con saldo libre · al {fecha_txt}', F(bold=True, font_size=14))
    wsDH.write(1, 0, 'Cada fila es un proyecto (UE-CUI-meta) con al menos S/ ' + format(cfg['monto_minimo'], ',.0f') + '. VERDE = saldo libre sin condiciones. AMARILLO = saldo libre con las condiciones del último campo. AZUL (rebajando certificación) = no hay saldo libre, pero hay dinero certificado sin comprometer ni programar en el 12-B: solo sirve si la UEI rebaja la certificación. ★ = proyecto de la lista de Antony. Los rubros son el PIM por fuente de financiamiento.', f_note)
    sv = sum(r['disponible'] for r in util if r['sem'] == 'Disponible'); sa = sum(r['disponible'] for r in util if r['sem'] == 'Con condiciones')
    f_k = F(bold=True, font_size=12, num_format='"S/ "#,##0', align='left'); f_kl = F(bold=True, font_color='#595959')
    wsDH.write(2, 0, 'Verde (disponible)', f_kl); wsDH.write_number(2, 4, sv, f_k); wsDH.write(2, 6, f'{sum(1 for r in util if r["sem"] == "Disponible")} proyectos', f_txt)
    wsDH.write(3, 0, 'Amarillo (con condiciones)', f_kl); wsDH.write_number(3, 4, sa, f_k); wsDH.write(3, 6, f'{sum(1 for r in util if r["sem"] == "Con condiciones")} proyectos', f_txt)
    wsDH.write(4, 0, 'Azul (rebajando certificación)', f_kl); wsDH.write_number(4, 4, sum(r['pot'] for r in util), f_k)
    wsDH.write(4, 6, f'{sum(1 for r in util if r["sem"] == "Rebajando certificación")} proyectos azules; la columna Potencial también existe en verdes y amarillos', f_txt)
    HDH = ['N°', 'Semáforo', 'Lista Antony', 'UE', 'Gerencia / Oficina', 'CUI', 'Proyecto', 'Tipo de inversión', 'Función', 'Disponible legal (S/)',
           'Potencial rebajando certificación (S/)', 'PIM 2026', 'Certificación', 'Rubro principal'] + [f'{nm} ({c})' for c, nm in RUBROS] + \
          ['Otros rubros', '% Avance acumulado', 'PMI 2026 (S/)', 'PMI 2027 (S/)', 'Motivo / condiciones']
    H0 = 7
    AN_FIRST = 9                                  # fila de Excel (1-indexada) del primer proyecto de la hoja "Lista Antony"
    an_pos = {c: k for k, c in enumerate(AN)}
    dh_fila = {}; base_fila = {}
    for r in rows:
        base_fila.setdefault(r['cui'], r['n'])
    wsDH.set_row(H0 - 1, 36)
    for j, h in enumerate(HDH):
        wsDH.write(H0 - 1, j, h, f_hd)
    f_wr = F(text_wrap=True, valign='top', border=1); f_ce = F(align='center', valign='top', border=1); f_nm = F(num_format='#,##0;(#,##0);"-"', valign='top', border=1)
    f_pc = F(num_format='0%', align='center', valign='top', border=1)
    OR = '#FFD966'
    g_ce = F(align='center', valign='top', border=1, bg_color=OR, bold=True); g_wr = F(text_wrap=True, valign='top', border=1, bg_color=OR, bold=True)
    g_lk = F(align='center', valign='top', border=1, bg_color=OR, bold=True, font_color='#7F6000', underline=1)
    for k, r in enumerate(util):
        i = H0 + k; d = r['d']; esA = r['cui'] in AN
        fc, fw = (g_ce, g_wr) if esA else (f_ce, f_wr)
        dh_fila.setdefault(r['cui'], i + 1)
        wsDH.write_number(i, 0, k + 1, fc); wsDH.write_string(i, 1, r['sem'], f_ce)
        if esA:
            wsDH.write_url(i, 2, f"internal:'Lista Antony'!A{AN_FIRST + an_pos[r['cui']]}", g_lk, string='★ Antony')
        else:
            wsDH.write_string(i, 2, '', f_ce)
        wsDH.write_string(i, 3, r['ue'], f_ce); wsDH.write_string(i, 4, r['ger'], f_wr); wsDH.write_string(i, 5, r['cui'], fc)
        wsDH.write_string(i, 6, r['nombre'], fw); wsDH.write_string(i, 7, r['tipo'], f_ce); wsDH.write_string(i, 8, r['fun'], f_wr)
        wsDH.write_number(i, 9, r['disponible'], f_nm); wsDH.write_number(i, 10, r['pot'], f_nm); wsDH.write_number(i, 11, r['pim'], f_nm); wsDH.write_number(i, 12, r['cert'], f_nm)
        vals, otros, pr = r['rub']
        wsDH.write_string(i, 13, pr, f_ce)
        for q, (c, _nm) in enumerate(RUBROS):
            if vals[c] is None: wsDH.write_string(i, 14 + q, 's/d', f_ce)
            else: wsDH.write_number(i, 14 + q, vals[c], f_nm)
        if otros is None: wsDH.write_string(i, 14 + len(RUBROS), '', f_ce)
        else: wsDH.write_number(i, 14 + len(RUBROS), otros, f_nm)
        o = 15 + len(RUBROS)
        if d['avance'] is None: wsDH.write_string(i, o, 's/d', f_ce)
        else: wsDH.write_number(i, o, d['avance'], f_pc)
        wsDH.write_number(i, o + 1, d['pmi26'], f_nm); wsDH.write_number(i, o + 2, d['pmi27'], f_nm); wsDH.write_string(i, o + 3, r['motivo'], f_wr)
    ult = H0 + len(util)
    for j, w in enumerate([5, 15, 11, 6, 34, 10, 62, 16, 26, 16, 16, 14, 14, 15] + [14] * (len(RUBROS) + 1) + [11, 13, 13, 60]):
        wsDH.set_column(j, j, w)
    if util:
        wsDH.autofilter(H0 - 1, 0, ult - 1, len(HDH) - 1)
        wsDH.conditional_format(H0, 9, ult - 1, 9, {'type': 'data_bar', 'bar_color': '#63BE7B', 'bar_solid': True})
        cond_disp(wsDH, f'B{H0 + 1}:B{ult}')
    wsDH.freeze_panes(H0, 7); wsDH.set_zoom(90)

    # ------------------------------------------------------------ Lista Antony
    HAN = ['N°', 'CUI', 'Proyecto (según la lista de Antony)', 'Rubro (Antony)', 'Anulación (Antony)', '¿Lo incluye el programa?', 'Semáforo del programa', 'Disponible legal (programa)',
           'Rubro principal (programa)', '¿Coincide el rubro?', '¿Coincide el monto?', 'Justificación del programa', 'Ver en Base', 'Ver en Dónde hay plata']
    if wsAN is not None:
        f_lk = F(align='center', valign='top', border=1, font_color='#0563C1', underline=1)
        by_cui = collections.defaultdict(list)
        for r in rows:
            by_cui[r['cui']].append(r)
        res = []
        for it in antony['items']:
            rs = by_cui.get(it['cui'], [])
            if not rs:
                res.append(dict(it=it, rs=rs, inc='No está en la Base', sem='', disp=0.0, pot=0.0, rub='', rcod='', just='El CUI no aparece en la Base de esta corrida (sin PIM 2026 o fuera del pliego). Revisar.'))
                continue
            r0 = max(rs, key=lambda x: x['pim']); sem = r0['sem']; disp = sum(x['disponible'] for x in rs); pot = sum(x['pot'] for x in rs)
            vals, otros, pr = r0['rub']
            if sem == 'Disponible': inc = 'Sí'; just = 'Coincide: ' + r0['motivo'] + '.'
            elif sem == 'Con condiciones': inc = 'Sí, con condiciones'; just = 'Coincide, con condiciones: ' + r0['motivo'] + '.'
            elif sem == 'Rebajando certificación': inc = 'Solo rebajando certificación'; just = r0['motivo'] + '.'
            elif sem == 'No mover': inc = 'No'; just = 'No lo proponemos: ' + r0['motivo'] + '. Es el límite del Art. 11 de la Ley 32513; la lista no lo considera.'
            elif sem == 'Restringido': inc = 'No (restringido)'; just = 'Restringido: ' + r0['motivo'] + '.'
            elif sem == 'Sin saldo útil': inc = 'No'; just = r0['motivo'] + '. ' + ('El 12-B programa casi todo el PIM y está en el PMI 2026.' if r0['d']['pmi26'] > 0 else '')
            else: inc = 'No'; just = r0['motivo'] + '.'
            res.append(dict(it=it, rs=rs, inc=inc, sem=sem, disp=disp, pot=pot, rub=pr, rcod=(pr.split('(')[-1].rstrip(')') if '(' in pr else ''), just=just.strip()))
        n_si = sum(1 for x in res if x['inc'] in ('Sí', 'Sí, con condiciones'))
        wsAN.write(0, 0, 'Lista de Antony vs. programa', F(bold=True, font_size=14))
        wsAN.write(1, 0, f"Fuente: {antony.get('fuente', '')} (importada {antony.get('importado', '')}). Cada fila dice si el programa también propone el proyecto y por qué. Use los enlaces azules para ir a la Base o a Dónde hay plata.", f_note)
        wsAN.write(2, 0, 'Proyectos en la lista', f_kl); wsAN.write_number(2, 4, len(res), F(bold=True, font_size=12))
        wsAN.write(3, 0, 'El programa también los propone', f_kl); wsAN.write_number(3, 4, n_si, F(bold=True, font_size=12))
        wsAN.write(4, 0, 'Monto de Antony / disponible del programa', f_kl)
        wsAN.write_number(4, 4, sum(x['it']['anulacion'] or 0 for x in res), F(bold=True, font_size=12, num_format='"S/ "#,##0')); wsAN.write_number(4, 7, sum(x['disp'] for x in res if x['inc'] in ('Sí', 'Sí, con condiciones')), F(bold=True, font_size=12, num_format='"S/ "#,##0'))
        wsAN.write(5, 0, 'No coinciden (por qué: columna Justificación)', f_kl); wsAN.write_number(5, 4, len(res) - n_si, F(bold=True, font_size=12))
        wsAN.set_row(7, 32)
        for j, h in enumerate(HAN):
            wsAN.write(7, j, h, f_hd)
        verde = F(align='center', valign='top', border=1, bg_color='#C6EFCE', bold=True); rojo = F(align='center', valign='top', border=1, bg_color='#FFC7CE', bold=True)
        amar = F(align='center', valign='top', border=1, bg_color='#FFEB9C', bold=True)
        for k, x in enumerate(res):
            i = 8 + k; it = x['it']
            wsAN.write_number(i, 0, k + 1, f_ce); wsAN.write_string(i, 1, it['cui'], f_ce); wsAN.write_string(i, 2, it['nombre'], f_wr); wsAN.write_string(i, 3, it['rubro'], f_wr)
            if it['anulacion'] is None: wsAN.write_string(i, 4, 's/d', f_ce)
            else: wsAN.write_number(i, 4, it['anulacion'], f_nm)
            wsAN.write_string(i, 5, x['inc'], verde if x['inc'] == 'Sí' else (amar if x['inc'].startswith(('Sí,', 'Solo')) else rojo))
            wsAN.write_string(i, 6, x['sem'], f_ce); wsAN.write_number(i, 7, x['disp'], f_nm); wsAN.write_string(i, 8, x['rub'], f_ce)
            ok_r = '' if not x['rs'] else ('Sí' if (it.get('rubro_cod') and it['rubro_cod'] == x['rcod']) else ('s/d' if x['rub'] == 's/d' else 'No'))
            wsAN.write_string(i, 9, ok_r, f_ce)
            if x['inc'] in ('Sí', 'Sí, con condiciones') and it['anulacion'] is not None:
                ok_m = 'Sí' if abs(it['anulacion'] - x['disp']) <= 1 else f"Difiere: saldo libre S/ {x['disp']:,.0f}"
            else:
                ok_m = ''
            wsAN.write_string(i, 10, ok_m, f_wr); wsAN.write_string(i, 11, x['just'], f_wr)
            if x['rs']:
                wsAN.write_url(i, 12, f"internal:'Base'!F{base_fila[it['cui']]}", f_lk, string='Ir a Base')
            else:
                wsAN.write_string(i, 12, '', f_ce)
            if it['cui'] in dh_fila:
                wsAN.write_url(i, 13, f"internal:'Dónde hay plata'!F{dh_fila[it['cui']]}", f_lk, string='Ir a Dónde hay plata')
            else:
                wsAN.write_string(i, 13, 'No figura (ver justificación)', f_ce)
        fin = 8 + len(res)
        wsAN.autofilter(7, 0, fin - 1, len(HAN) - 1)
        # proyectos que el programa propone y Antony no
        otros_p = [r for r in util if r['cui'] not in AN]
        q0 = fin + 2
        wsAN.write(q0, 0, f'Proyectos que el programa propone y que NO están en la lista de Antony ({len(otros_p)})', F(bold=True, font_size=12))
        wsAN.write(q0 + 1, 0, 'No significa que estén mal: son casos con PMI 2026, 12-B que debe actualizar la UEI, o certificación por rebajar. El motivo está en cada fila.', f_note)
        H2 = ['N°', 'CUI', 'Proyecto', 'Rubro principal', 'Disponible legal', 'Semáforo', 'Potencial rebajando cert.', 'Gerencia / Oficina', '', '', '', 'Motivo / condiciones', 'Ver en Base', 'Ver en Dónde hay plata']
        for j, h in enumerate(H2):
            if h: wsAN.write(q0 + 2, j, h, f_hd)
        for k, r in enumerate(otros_p):
            i = q0 + 3 + k
            wsAN.write_number(i, 0, k + 1, f_ce); wsAN.write_string(i, 1, r['cui'], f_ce); wsAN.write_string(i, 2, r['nombre'], f_wr); wsAN.write_string(i, 3, r['rub'][2], f_ce)
            wsAN.write_number(i, 4, r['disponible'], f_nm); wsAN.write_string(i, 5, r['sem'], f_ce); wsAN.write_number(i, 6, r['pot'], f_nm); wsAN.write_string(i, 7, r['ger'], f_wr)
            wsAN.write_string(i, 11, r['motivo'], f_wr)
            wsAN.write_url(i, 12, f"internal:'Base'!F{r['n']}", f_lk, string='Ir a Base'); wsAN.write_url(i, 13, f"internal:'Dónde hay plata'!F{dh_fila[r['cui']]}", f_lk, string='Ir a Dónde hay plata')
        cond_disp(wsAN, f'G9:G{fin}'); cond_disp(wsAN, f'F{q0 + 4}:F{q0 + 3 + len(otros_p)}')
        for j, w in enumerate([5, 10, 60, 20, 16, 24, 20, 16, 22, 14, 26, 72, 12, 24]):
            wsAN.set_column(j, j, w)
        wsAN.freeze_panes(8, 3); wsAN.set_zoom(90)

    # ------------------------------------------------------------ Anexos
    HDR = ['N°', 'META', 'PROYECTO', 'PIM 2026\n(A)', 'CERTIFICACIÓN\n(B)', 'SALDO\n(A-B)', 'DEVENGADO\n( C )', 'SALDO\n(A-C)', 'DEVEN\n%',
           'MODALIDAD\nDE\nEJECUCIÓN', 'TIPO DE\nINVERSIÓN', 'Monto Año 2026 (S/)', 'Monto Año 2027 (S/)', 'Monto Año 2028 (S/)',
           'Monto Año 2029 (S/)', 'BASE LEGAL', 'SEMÁFORO', '% AVANCE\nFINANCIERO\nACUMULADO', 'DISPONIBLE\nLEGAL (S/)', 'MOTIVO / CONDICIONES']
    W = [6, 8, 62, 15, 15, 14, 15, 15, 9, 28, 16, 15, 15, 15, 15, 18, 15, 13, 16, 52]
    tot_hoja = {}

    def escribir_anexo(ws, titulo, rs, nro):
        ws.write(0, 2, f'ANEXO N° {nro:02d}', f_title)
        ws.write(1, 1, f'COMPORTAMIENTO DE EJECUCIÓN DE PROYECTOS DE INVERSION PUBLICA DEL AÑO FISCAL 2026 · {titulo} · AL {fecha_txt}', f_bold)
        ws.set_row(3, 44)
        for j, h in enumerate(HDR):
            ws.write(3, j, h, f_hd if j <= 8 else f_hd2)
            ws.set_column(j, j, W[j])
        raiz = arbol(rs)
        cur = [5]; cnt = [0]; first = 5 + 1   # N° usa SUBTOTAL sobre columna B desde la primera fila de proyectos

        def vals(r):
            return dict(pim=r['pim'], cert=r['cert'], dev=r['dev'], dispo=r['disponible'])

        def emit(nd):
            """Devuelve (fila_excel_1idx, sumas) del nodo."""
            if nd.nivel > 0:
                nd.fila = cur[0]; cur[0] += 1
            filas_hoja = []; suma = collections.Counter()
            for r in nd.hojas:
                row = cur[0]; cur[0] += 1; cnt[0] += 1; filas_hoja.append(row); n = r['n']; i = row - 1
                v = vals(r)
                ws.write_formula(i, 0, f'=SUBTOTAL(103,$B${first}:B{row})', f_leafc, cnt[0])
                ws.write_number(i, 1, int(r['meta']), f_leafc)
                ws.write_string(i, 2, r['nombre'], f_leaf)   # texto fijo (no fórmula) para que Ctrl+B encuentre el CUI
                ws.write_formula(i, 3, f"=Base!{L['pim']}{n}", f_leafn, r['pim'])
                ws.write_formula(i, 4, f"=Base!{L['cert']}{n}", f_leafn, r['cert'])
                ws.write_formula(i, 5, f'=D{row}-E{row}', f_leafn, r['pim'] - r['cert'])
                ws.write_formula(i, 6, f"=Base!{L['dev']}{n}", f_leafn, r['dev'])
                ws.write_formula(i, 7, f'=D{row}-G{row}', f_leafn, r['pim'] - r['dev'])
                ws.write_formula(i, 8, f'=IF(D{row}=0,0,G{row}/D{row})', f_leafp, r['dev'] / r['pim'] if r['pim'] else 0)
                ws.write_formula(i, 9, f"=Base!{L['mod']}{n}", f_leaf, r['mod'])
                ws.write_formula(i, 10, f"=Base!{L['tipo']}{n}", f_leafc, r['tipo'])
                for k in range(4):
                    ws.write_formula(i, 11 + k, f"=Base!{L['m26'] if k == 0 else L['m27'] if k == 1 else L['m28'] if k == 2 else L['m29']}{n}",
                                     f_leafn, r['montos'][k])
                ws.write_formula(i, 15, f'=IF(Base!{L["bl"]}{n}="","",Base!{L["bl"]}{n})', f_leaf, r['bl'])
                ws.write_formula(i, 16, f"=Base!{L['sem']}{n}", f_leafc, r['sem'])
                ws.write_formula(i, 17, f"=Base!{L['avance']}{n}", f_leafp, r['d']['avance'] if r['d']['avance'] is not None else 's/d')
                ws.write_formula(i, 18, f"=Base!{L['disponible']}{n}", f_leafn, v['dispo'])
                ws.write_formula(i, 19, f"=Base!{L['motivo']}{n}", f_leaf, r['motivo'])
                suma.update({'D': r['pim'], 'E': r['cert'], 'G': r['dev'], 'S': v['dispo']})
            hijos = []
            for h in nd.hijos.values():
                fr, sm, _h = emit(h); hijos.append(fr); suma.update(sm)
            if nd.nivel == 0:
                return None, suma, hijos
            fila = nd.fila; i = fila - 1
            clave = 'yel' if nd.nivel == 1 else ('sal' if nd.nombre[:2].isdigit() and nd.nombre[2:3] in '.:' else 'gry')
            fx, fn_, fp = sub[clave]
            ws.write(i, 0, nd.nombre, fx)
            for j in (1, 2, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19):
                ws.write_blank(i, j, None, fx)
            for col, key in (('D', 'D'), ('E', 'E'), ('G', 'G'), ('S', 'S')):
                terms = []
                if filas_hoja: terms.append(f'SUM({col}{filas_hoja[0]}:{col}{filas_hoja[-1]})')
                terms += [f'{col}{k}' for k in hijos]
                ws.write_formula(i, ord(col) - 65, '=' + '+'.join(terms), fn_, suma[key])
            ws.write_formula(i, 5, f'=D{fila}-E{fila}', fn_, suma['D'] - suma['E'])
            ws.write_formula(i, 7, f'=D{fila}-G{fila}', fn_, suma['D'] - suma['G'])
            ws.write_formula(i, 8, f'=IF(D{fila}=0,0,G{fila}/D{fila})', fp, suma['G'] / suma['D'] if suma['D'] else 0)
            return fila, suma, None

        _, total, tops = emit(raiz)
        tr = cur[0] + 1; i = tr - 1
        fx, fn_, fp = sub['navy']
        ws.write(i, 0, f'TOTAL {titulo.upper()}', fx)
        for j in (1, 2, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19):
            ws.write_blank(i, j, None, fx)
        for col in 'DEGS':
            ws.write_formula(i, ord(col) - 65, '=' + '+'.join(f'{col}{k}' for k in tops), fn_, total[col])
        ws.write_formula(i, 5, f'=D{tr}-E{tr}', fn_, total['D'] - total['E'])
        ws.write_formula(i, 7, f'=D{tr}-G{tr}', fn_, total['D'] - total['G'])
        ws.write_formula(i, 8, f'=IF(D{tr}=0,0,G{tr}/D{tr})', fp, total['G'] / total['D'] if total['D'] else 0)
        ws.freeze_panes(4, 3); ws.set_zoom(90)
        cond_disp(ws, f'Q5:Q{tr}')
        tot_hoja[ws.get_name()] = (tr, total)

    for k, nm in enumerate(anexos, 1):
        rs = [r for r in rows if r['hoja'] == nm]
        titulo = {'Sin clasificar': 'PROYECTOS SIN GERENCIA ASIGNADA', 'Otras UE': 'OTRAS UNIDADES EJECUTORAS'}.get(nm, nm)
        escribir_anexo(wsA[nm], titulo, rs, k)

    # ------------------------------------------------------------ Resumen
    wsR.write(0, 0, f'Proyectos de inversión · GORE Puno · mapeo de disponibilidad · al {fecha_txt}', F(bold=True, font_size=14))
    wsR.write(1, 0, 'Montos en soles. Todo se calcula desde la hoja Base. "Disponible legal" = saldo libre según las reglas de la hoja Reglas (no es PIM − compromiso).', f_note)
    if meta_info.get('banner'):
        wsR.write(2, 0, meta_info['banner'], F(bold=True, font_color='#C00000', bg_color='#FFF2CC'))
    rng = lambda nm: f"Base!${L[nm]}${B0}:${L[nm]}${BL}"
    f_k = F(bold=True, font_size=13, num_format='"S/ "#,##0', align='left'); f_kl = F(bold=True, font_color='#595959')
    kp = [('Disponible (verde)', 'Disponible'), ('Con condiciones (amarillo)', 'Con condiciones'), ('Restringido (PROCOMPITE / Anexo I)', 'Restringido')]
    pot_tot = sum(r['pot'] for r in rows if r['sem'] in ('Disponible', 'Con condiciones', 'Rebajando certificación'))
    for k, (lab, sem) in enumerate(kp):
        v = sum(r['disponible'] for r in rows if r['sem'] == sem)
        wsR.write(3 + k, 0, lab, f_kl)
        wsR.write_formula(3 + k, 1, f'=SUMIFS({rng("disponible")},{rng("sem")},"{sem}")', f_k, v)
    wsR.write(6, 0, 'Potencial rebajando certificación', f_kl)
    wsR.write_formula(6, 1, '=' + '+'.join(f'SUMIFS({rng("pot")},{rng("sem")},"{q}")' for q in ('Disponible', 'Con condiciones', 'Rebajando certificación')), f_k, pot_tot)
    wsR.write(7, 0, 'Detalle ordenado por monto: hoja "Dónde hay plata". Lista completa y filtrable: hoja "Base".', f_note)
    SEMS = ['Disponible', 'Con condiciones', 'Restringido', 'No mover']
    heads = ['Unidad', 'Proyectos (UE-CUI-meta)', 'PIM 2026', 'Certificación', 'Devengado', 'Saldo sin certificar', 'Disponible legal · Disponible',
             'Disponible legal · Con condiciones', 'Disponible legal · Restringido', 'PIM · No mover', 'PIM · Sin saldo útil / rebajando cert. / bolsas / sin PIM', 'Potencial rebajando certificación']

    def tabla(r0, titulo, claves, colcrit, filtro_ue=None):
        wsR.write(r0, 0, titulo, f_bold)
        for j, h in enumerate(heads):
            wsR.write(r0 + 1, j, h, f_hd)
        wsR.set_row(r0 + 1, 44)
        tot = collections.Counter()
        for i, (clave, etiqueta, rs) in enumerate(claves):
            n = r0 + 3 + i; ix = n - 1
            wsR.write_string(ix, 0, etiqueta, f_leaf)
            crit = f'{rng(colcrit)},$A{n}' if colcrit != 'ue' else f'{rng("ue")},"{clave}"'
            if filtro_ue: crit = f'{rng("ue")},"{filtro_ue}",{rng("ger")},"{clave}"'
            val = lambda k: sum(r[k] for r in rs)
            wsR.write_formula(ix, 1, f'=COUNTIFS({crit})', f_leafn, len(rs)); tot[1] += len(rs)
            for j, k in ((2, 'pim'), (3, 'cert'), (4, 'dev')):
                wsR.write_formula(ix, j, f'=SUMIFS({rng(k)},{crit})', f_leafn, val(k)); tot[j] += val(k)
            wsR.write_formula(ix, 5, f'=C{n}-D{n}', f_leafn, val('pim') - val('cert')); tot[5] += val('pim') - val('cert')
            for q, sem in enumerate(SEMS[:3]):
                v = sum(r['disponible'] for r in rs if r['sem'] == sem); tot[6 + q] += v
                wsR.write_formula(ix, 6 + q, f'=SUMIFS({rng("disponible")},{crit},{rng("sem")},"{sem}")', f_leafn, v)
            v = sum(r['pim'] for r in rs if r['sem'] == 'No mover'); tot[9] += v
            wsR.write_formula(ix, 9, f'=SUMIFS({rng("pim")},{crit},{rng("sem")},"No mover")', f_leafn, v)
            v = sum(r['pim'] for r in rs if r['sem'] in ('Sin saldo útil', 'Rebajando certificación', 'Bolsa genérica', 'Sin PIM')); tot[10] += v
            wsR.write_formula(ix, 10, f'=C{n}-SUMIFS({rng("pim")},{crit},{rng("sem")},"Disponible")-SUMIFS({rng("pim")},{crit},{rng("sem")},"Con condiciones")'
                                      f'-SUMIFS({rng("pim")},{crit},{rng("sem")},"Restringido")-J{n}', f_leafn, v)
            v = sum(r['pot'] for r in rs if r['sem'] in ('Disponible', 'Con condiciones', 'Rebajando certificación')); tot[11] += v
            wsR.write_formula(ix, 11, '=' + '+'.join(f'SUMIFS({rng("pot")},{crit},{rng("sem")},"{q}")' for q in ('Disponible', 'Con condiciones', 'Rebajando certificación')), f_leafn, v)
        n = r0 + 3 + len(claves); ix = n - 1
        fx, fn_, fp = sub['gry']
        wsR.write(ix, 0, 'TOTAL', fx)
        for j in range(1, len(heads)):
            wsR.write_formula(ix, j, f'=SUM({xl_col_to_name(j)}{r0 + 3}:{xl_col_to_name(j)}{n - 1})', fn_, tot[j])
        return n + 2

    por_ue = [(u, f"{u}. {siaf['ues'][u]}", [r for r in rows if r['ue'] == u]) for u in ues_orden]
    nxt = tabla(9, 'Por unidad ejecutora', [(u, e, rs) for u, e, rs in por_ue], 'ue')
    gers = list(collections.OrderedDict.fromkeys(r['ger'] for r in rows if r['ue'] == '001'))
    por_ger = [(g, g, [r for r in rows if r['ue'] == '001' and r['ger'] == g]) for g in gers]
    tabla(nxt, 'UE 001 Sede Central · por gerencia / oficina', por_ger, 'ger')
    wsR.set_column(0, 0, 58); wsR.set_column(1, 1, 22); wsR.set_column(2, 11, 17); wsR.freeze_panes(0, 1)

    # ------------------------------------------------------------ Rubro por CUI
    rb_list = armar_rubros(rows, rubros, api)
    RB_LAST = 4 + len(rb_list)
    rb_idx = {}
    for k, rb in enumerate(rb_list):
        rb['fila'] = 5 + k; rb_idx[(rb['ue'], rb['cui'])] = rb
    fuente_txt = 'API Datos Abiertos MEF (por UE)' if api is not None else 'Consulta Amigable (nivel pliego; respaldo)'
    wsRB.write(0, 0, f'Rubro de financiamiento por UE y CUI · fuente: {fuente_txt} · al {fecha_txt}', f_title)
    wsRB.write(1, 0, 'PIM por rubro. "s/d" = sin dato (falta el reporte del rubro, está incompleto, o el CUI está en varias UE y la fuente no lo separa).', f_note)
    HRB = ['CUI', 'Proyecto', 'UE', 'PIM SIAF\n(esta UE)', 'RO (00)', 'RDR (09)', 'FONCOR (15)', 'CANON (18)', 'ROOC (19)', 'Otros rubros',
           'Suma rubros', 'PIM sin rubro\nasignado', 'Cobertura', 'Clave']
    wsRB.set_row(3, 36)
    for j, h in enumerate(HRB):
        wsRB.write(3, j, h, f_hd)
    for rb in rb_list:
        i = rb['fila'] - 1; n = rb['fila']
        wsRB.write_string(i, 0, rb['cui'], f_txt); wsRB.write_string(i, 1, rb['nombre'], f_txt); wsRB.write_string(i, 2, rb['ue'], f_txt)
        wsRB.write_number(i, 3, rb['pim'], f_in)
        for k, (code, _nm) in enumerate(RUBROS):
            val = rb['v'][code]
            if val is None:
                wsRB.write_string(i, 4 + k, 's/d', f_ctr)
            else:
                wsRB.write_number(i, 4 + k, val, f_in)
        if rb['otros'] is None:
            wsRB.write_string(i, 9, '', f_txt)
        else:
            wsRB.write_number(i, 9, rb['otros'], f_in)
        known = sum(x for x in rb['std'] if x is not None) + (rb['otros'] or 0.0)
        wsRB.write_formula(i, 10, f'=SUM(E{n}:J{n})', f_num, known)
        wsRB.write_formula(i, 11, f'=D{n}-K{n}', f_num, rb['pim'] - known)
        wsRB.write_string(i, 12, rb['estado'], f_txt)
        wsRB.write_formula(i, 13, f'=C{n}&"|"&A{n}', f_txt, rb['ue'] + '|' + rb['cui'])
    wsRB.set_column(0, 0, 10); wsRB.set_column(1, 1, 70); wsRB.set_column(2, 2, 6); wsRB.set_column(3, 11, 15); wsRB.set_column(12, 12, 62); wsRB.set_column(13, 13, 14)
    wsRB.freeze_panes(4, 2); wsRB.autofilter(3, 0, max(RB_LAST - 1, 4), 13)

    # ------------------------------------------------------------ Saldo al expediente
    proy = collections.OrderedDict()
    for r in rows:
        proy.setdefault((r['ue'], r['cui']), []).append(r)
    sx = []
    for (ue, cui), rs in proy.items():
        rs0 = max(rs, key=lambda r: r['pim'])
        sx.append(dict(ue=ue, cui=cui, rs=rs, metas=', '.join(sorted({str(int(r['meta'])) for r in rs})), nombre=rs0['nombre'], car=rs0['car'],
                       fun=rs0['fun'] if rs0['fun'] else '(sin función)', pim=sum(r['pim'] for r in rs)))
    HX = ['N°', 'META(S)', 'NOMBRE DEL PROYECTO', 'COSTO\nACTUALIZADO\n(Cartera)', 'EJEC. ACUM.\nAL 31/12/2025\n(Cartera)', 'RECURSOS\nORDINARIOS\n(00)', 'RDR\n(09)', 'FONCOR\n(15)', 'CANON\n(18)',
          'ROOC\n(19)', 'PIM SIN\nRUBRO\nASIGNADO', 'PIM 2026', 'SALDO AL EXPEDIENTE /\nINVERSIÓN', 'OBSERVACIÓN']
    wsS.write(0, 0, f'COMPORTAMIENTO DE PROYECTOS DE INVERSION PUBLICA AL {fecha_txt} · por función · saldo al expediente', f_title)
    wsS.write(1, 0, 'Saldo = Costo actualizado − Ejecución acumulada al 2025 − PIM 2026 (lo que falta financiar después de 2026). Un CUI se cuenta una vez por UE; no sumar el costo entre UE.', f_note)
    if api is None:
        faltan_rb = [c + ' ' + nm for c, nm in RUBROS if c not in (rubros or {}) or not rubros[c]['completo']]
        if faltan_rb:
            wsS.write(2, 0, 'AVISO: la API del MEF no estuvo disponible y faltan reportes de rubro de la Consulta Amigable (s/d): ' + ', '.join(faltan_rb) + '. Ver Config.', F(bold=True, font_color='#C00000'))
    else:
        wsS.write(2, 0, f"Rubro: API Datos Abiertos MEF (por UE y CUI), descargada {api['fecha']:%d/%m/%Y %H:%M}.", f_note)
    wsS.set_row(3, 46)
    for j, h in enumerate(HX):
        wsS.write(3, j, h, f_hd)
    for j, w in enumerate([6, 12, 62, 16, 16, 15, 14, 15, 14, 12, 12, 15, 18, 36]):
        wsS.set_column(j, j, w)
    ra = {(r['ue'], r['cui']): r for r in rows}
    def lab_ue(u): return f"{u}. {siaf['ues'][u]}"
    ntree = Nodo('', 0)
    for p in sx:
        n = ntree.hijos.setdefault(lab_ue(p['ue']), Nodo(lab_ue(p['ue']), 1)).hijos.setdefault(p['fun'], Nodo(p['fun'], 2))
        n.hojas.append(p)
    cur = [5]; cntp = [0]
    def emitS(nd):
        suma = collections.Counter(); filas_h = []; hijos = []
        if nd.nivel > 0:
            nd.fila = cur[0]; cur[0] += 1
        for p in sorted(nd.hojas, key=lambda p: -p['pim']):
            row = cur[0]; cur[0] += 1; cntp[0] += 1; i = row - 1; filas_h.append(row)
            car = p['car']; crit = f'{rng("ue")},"{p["ue"]}",{rng("cui")},"{p["cui"]}"'
            wsS.write_number(i, 0, cntp[0], f_leafc); wsS.write_string(i, 1, p['metas'], f_leafc); wsS.write_string(i, 2, p['nombre'], f_leaf)
            if car and car['costo'] is not None:
                wsS.write_number(i, 3, car['costo'], F(border=1, num_format='#,##0;(#,##0);"-"', font_color='#0000FF', align='right'))
                ej = car['dev_acum'] or 0.0
                wsS.write_number(i, 4, ej, F(border=1, num_format='#,##0;(#,##0);"-"', font_color='#0000FF', align='right'))
                saldo = car['costo'] - ej - p['pim']
                wsS.write_formula(i, 12, f'=D{row}-E{row}-L{row}', f_leafn, saldo)
                obs = 'Saldo negativo: costo menor que ejecución + PIM (revisar)' if saldo < -1 else ''
                suma.update({'D': car['costo'], 'E': ej, 'M': saldo})
            else:
                wsS.write_string(i, 3, NOCART, f_leaf); wsS.write_string(i, 4, NOCART, f_leaf); wsS.write_string(i, 12, NOCART, f_leaf)
                obs = 'Sin datos en Cartera: no se puede calcular el saldo'
            rb = rb_idx[(p['ue'], p['cui'])]
            if rb['compartido']:
                for k in range(6):
                    wsS.write_string(i, 5 + k, 's/d' if k < 5 else '', f_leaf)
                obs = (obs + ' · ' if obs else '') + 'CUI en varias UE: rubro solo a nivel pliego (hoja Rubro por CUI)'
            else:
                for k, (code, _nm) in enumerate(RUBROS):
                    val = rb['v'][code]
                    cached = 's/d' if val is None else val
                    colx = xl_col_to_name(4 + k)
                    wsS.write_formula(i, 5 + k, "=INDEX('Rubro por CUI'!$" + colx + "$5:$" + colx + "$" + str(RB_LAST) + ",MATCH(\"" + p['ue'] + "|" + p['cui'] + "\",'Rubro por CUI'!$N$5:$N$" + str(RB_LAST) + ",0))",
                                      f_leafn, cached)
                    if val is not None:
                        suma[chr(70 + k)] += val
                sinv_ = p['pim'] - sum(x for x in rb['std'] if x is not None)
                wsS.write_formula(i, 10, f'=L{row}-SUM(F{row}:J{row})', f_leafn, sinv_); suma['K'] += sinv_
                if rb['estado'] != 'Completo':
                    obs = (obs + ' · ' if obs else '') + 'Rubro: ' + rb['estado']
            wsS.write_formula(i, 11, f'=SUMIFS({rng("pim")},{crit})', f_leafn, p['pim']); suma['L'] += p['pim']
            wsS.write_string(i, 13, obs, f_leaf)
        for h in nd.hijos.values():
            fr, sm = emitS(h); hijos.append(fr); suma.update(sm)
        if nd.nivel == 0:
            return hijos, suma
        fila = nd.fila; i = fila - 1
        fx, fn_, fp = sub['yel' if nd.nivel == 1 else 'sal']
        wsS.write(i, 0, nd.nombre, fx)
        for j in range(1, 14): wsS.write_blank(i, j, None, fx)
        for col in 'DEFGHIJKLM':
            terms = []
            if filas_h: terms.append(f'SUM({col}{filas_h[0]}:{col}{filas_h[-1]})')
            terms += [f'{col}{k}' for k in hijos]
            wsS.write_formula(i, ord(col) - 65, '=' + '+'.join(terms), fn_, suma[col])
        return fila, suma
    emitS(ntree)
    wsS.freeze_panes(4, 3); wsS.set_zoom(90)

    # ------------------------------------------------------------ Control
    wsC.write(0, 0, 'Control y conciliación', f_title)
    for j, h in enumerate(['Control', 'Calculado', 'Referencia', 'Diferencia', 'Resultado']):
        wsC.write(2, j, h, f_hd)
    f_c = F(num_format='#,##0.00;(#,##0.00);"-"', border=1); f_t = F(border=1)
    cr = [3]
    def chk(label, formula, calc, ref, tol=0.5, texto=False):
        i = cr[0]; n = i + 1
        wsC.write_string(i, 0, label, f_t)
        wsC.write_formula(i, 1, formula, f_c, calc)
        if isinstance(ref, str) and ref.startswith('='):
            wsC.write_formula(i, 2, ref, f_c, meta_info.get('_ref_cache', {}).get(ref, 0))
        else:
            wsC.write_number(i, 2, ref, f_c)
        wsC.write_formula(i, 3, f'=B{n}-C{n}', f_c, calc - ref if not isinstance(ref, str) else 0)
        ok = abs(calc - ref) <= tol if not isinstance(ref, str) else True
        wsC.write_formula(i, 4, f'=IF(ABS(D{n})<={tol},"Conciliado","REVISAR")', f_t, 'Conciliado' if ok else 'REVISAR')
        cr[0] += 1
    allm = siaf['metas']
    for u in (ues_orden if siaf.get('ue_tot') else []):
        pim_p = sum(r['pim'] for r in rows if r['ue'] == u)
        pim_a = sum(m['pim'] for m in allm if m['ue'] == u and m['tipo'] != 'proyecto')
        tot = siaf['ue_tot'].get(u, {}).get('pim', 0)
        chk(f'UE {u}: PIM proyectos (Base) + PIM actividades (SIAF) = TOTAL UE del SIAF', f'=SUMIFS({rng("pim")},{rng("ue")},"{u}")+{pim_a}', pim_p + pim_a, tot)
    tot_all = sum(r['pim'] for r in rows)
    chk('PIM de todos los proyectos (Base)', f'=SUM({rng("pim")})', tot_all, tot_all)
    for nm in anexos:
        tr, t = tot_hoja[nm]
        chk(f'Hoja "{nm}": PIM = Base', f"='{nm}'!D{tr}", t['D'], sum(r['pim'] for r in rows if r['hoja'] == nm))
        chk(f'Hoja "{nm}": Devengado = Base', f"='{nm}'!G{tr}", t['G'], sum(r['dev'] for r in rows if r['hoja'] == nm))
    if api is None:
        for c, l in sorted((rubros or {}).items()):
            chk('Rubro ' + c + ': suma de proyectos del archivo = total del rubro en el reporte', '=' + repr(l['suma_pim']), l['suma_pim'], l['total']['pim'])
    else:
        pim_api = sum(sum(v.values()) for v in api['pim'].values())
        chk('API del MEF: PIM total de proyectos = PIM del SIAF', '=' + repr(pim_api), pim_api, sum(r['pim'] for r in rows), tol=1000)
    chk('Rubro por UE-CUI: PIM SIAF de la hoja = PIM de los proyectos (Base)', f"=SUM('Rubro por CUI'!D5:D{RB_LAST})", sum(x['pim'] for x in rb_list), sum(r['pim'] for r in rows))
    cr[0] += 1
    wsC.write(cr[0], 0, 'Alertas (informativo)', f_bold); cr[0] += 1
    sin_cl = [r for r in rows if r['hoja'] == 'Sin clasificar']
    sin_c = [r for r in rows if r['montos'][0] == NOCART]
    alert = [('Proyectos sin gerencia asignada (completar en maestro.xlsx, hoja Clasificacion)', len(sin_cl)),
             ('Filas sin datos en Cartera (montos 2026-2029 = "No registrado en cartera")', len(sin_c)),
             ('Filas con nota de clasificación', sum(1 for r in rows if r['nota'])),
             ('Filas con PIM > 0 sin ficha SSI ni 12-B (semáforo con datos incompletos)', sum(1 for r in rows if r['pim'] > 0 and not r['d']['tiene_ssi'] and not r['d']['tiene_f12'])),
             ('CUI con cobertura de rubro incompleta (ver hoja Rubro por CUI)', sum(1 for x in rb_list if x['estado'] != 'Completo')),
             ('UE-CUI cuyo rubro no se pudo repartir por UE (solo ocurre con la Consulta Amigable)', sum(1 for x in rb_list if x['compartido']))]
    for k, v in alert:
        wsC.write_string(cr[0], 0, k, f_t); wsC.write_number(cr[0], 1, v, f_t); cr[0] += 1
    if meta_info.get('avisos'):
        cr[0] += 1; wsC.write(cr[0], 0, 'Avisos de la corrida', f_bold); cr[0] += 1
        for a in meta_info['avisos']:
            wsC.write_string(cr[0], 0, a, F(font_color='#C00000')); cr[0] += 1
    if sin_cl:
        cr[0] += 1; wsC.write(cr[0], 0, 'Proyectos sin clasificar:', f_bold); cr[0] += 1
        for r in sin_cl:
            wsC.write_string(cr[0], 0, f"{r['cui']} meta {r['meta']} · PIM {r['pim']:,.0f} · {r['nombre'][:100]}", f_txt); cr[0] += 1
    wsC.set_column(0, 0, 100); wsC.set_column(1, 4, 20)

    # ------------------------------------------------------------ Cambios desde la corrida anterior
    cb = meta_info.get('cambios')
    HC = ['N°', 'Tipo de cambio', 'UE', 'CUI', 'Meta', 'Proyecto', 'Semáforo antes', 'Semáforo ahora', 'Disponible antes', 'Disponible ahora', 'Variación', 'PIM antes', 'PIM ahora']
    wsCH.write(0, 0, 'Qué cambió desde la corrida anterior', F(bold=True, font_size=14))
    if cb is None or cb.get('lista') is None:
        wsCH.write(1, 0, 'Esta es la primera corrida con historial: desde la próxima se listarán aquí los proyectos nuevos, los que cambiaron de semáforo y los que variaron su disponible o PIM.', f_note)
    else:
        lst = cb['lista']
        wsCH.write(1, 0, f"Comparado con la corrida del {cb['desde']:%d/%m/%Y %H:%M}. Se listan proyectos nuevos, que ya no aparecen, que cambiaron de semáforo o cuyo disponible/PIM varió en S/ 1 000 o más.", f_note)
        tot_n = {s_: sum(r['disponible'] for r in rows if r['sem'] == s_) for s_ in ('Disponible', 'Con condiciones')}
        tot_a = {s_: sum(c['d0'] for c in lst if c['sem0'] == s_) for s_ in ('Disponible', 'Con condiciones')}
        wsCH.write(2, 0, f'{len(lst)} cambios', F(bold=True))
        wsCH.set_row(3, 30)
        for j, h in enumerate(HC):
            wsCH.write(3, j, h, f_hd)
        for k, c in enumerate(lst):
            i = 4 + k
            wsCH.write_number(i, 0, k + 1, f_ctr); wsCH.write_string(i, 1, c['tipo'], f_txt); wsCH.write_string(i, 2, c['k'][0], f_ctr); wsCH.write_string(i, 3, c['k'][1], f_ctr)
            wsCH.write_string(i, 4, c['k'][2], f_ctr); wsCH.write_string(i, 5, c['nombre'], f_txt); wsCH.write_string(i, 6, c['sem0'], f_ctr); wsCH.write_string(i, 7, c['sem1'], f_ctr)
            for j, v in ((8, c['d0']), (9, c['d1']), (10, c['d1'] - c['d0']), (11, c['p0']), (12, c['p1'])):
                wsCH.write_number(i, j, v, f_num)
        if lst:
            wsCH.autofilter(3, 0, 3 + len(lst), len(HC) - 1)
            cond_disp(wsCH, f'G5:H{4 + len(lst)}')
            wsCH.conditional_format(4, 10, 3 + len(lst), 10, {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFFFFF', 'max_color': '#63BE7B', 'mid_type': 'num', 'mid_value': 0})
    for j, w in enumerate([5, 30, 6, 10, 7, 62, 18, 18, 15, 15, 14, 15, 15]):
        wsCH.set_column(j, j, w)
    wsCH.freeze_panes(4, 4)

    # ------------------------------------------------------------ Calidad de datos
    wsQ.write(0, 0, f'Calidad de datos · qué falta o hay que validar · al {fecha_txt}', F(bold=True, font_size=12))
    wsQ.write(1, 0, 'Los datos faltantes nunca se inventan: quedan como "s/d" y el semáforo lo indica en el motivo.', f_note)
    con_pim = [r for r in rows if r['pim'] > 0 and r['sem'] != 'Bolsa genérica']
    items = [
        ('Con PIM pero sin ficha SSI (tipo de inversión y avance acumulado no verificables)', [r for r in con_pim if not r['d']['tiene_ssi']], 'Reintentar mañana; si persiste, subir ConsultaAvanzada a "entrada".'),
        ('Con PIM pero sin Formato 12-B (saldo libre = PIM − certificado; sin programación)', [r for r in con_pim if not r['d']['tiene_f12']], 'Pedir a la OPMI/UEI que registre el 12-B.'),
        ('Con 12-B pero sin programación vigente (todos los meses en 0)', [r for r in con_pim if r['d']['tiene_f12'] and not r['d']['prog_ok']], 'La UEI debe actualizar el 12-B.'),
        ('Con PIM y no figuran en la Cartera PMI (montos 2026-2029 no disponibles)', [r for r in con_pim if not r['d']['en_cartera']], 'Normal para PROCOMPITE o inversiones no registradas en Invierte.'),
        ('Sin costo actualizado (no se calcula % de avance acumulado)', [r for r in con_pim if r['d']['avance'] is None], 'Revisar ficha de la inversión.'),
        ('Con cierre registrado o estado distinto de ACTIVO y aún con PIM', [r for r in con_pim if r['d']['cierre'].startswith('Sí') or r['d']['estado_inv'] in ('CERRADO', 'DESACTIVADO PERMANENTE')], 'Posible liberación de PIM: validar con la UEI.'),
        ('Sin gerencia asignada (UE 001)', [r for r in rows if r['hoja'] == 'Sin clasificar'], 'Completar en maestro.xlsx.'),
    ]
    for j, h in enumerate(['Indicador', 'Filas', 'PIM 2026 afectado', 'Qué hacer']):
        wsQ.write(3, j, h, f_hd)
    for k, (lab, rs, que) in enumerate(items):
        wsQ.write_string(4 + k, 0, lab, f_leaf); wsQ.write_number(4 + k, 1, len(rs), f_leafn); wsQ.write_number(4 + k, 2, sum(r['pim'] for r in rs), f_leafn); wsQ.write_string(4 + k, 3, que, f_leaf)
    q0 = 4 + len(items) + 2
    wsQ.write(q0 - 1, 0, 'Detalle (filtrable)', f_bold)
    for j, h in enumerate(['Indicador', 'UE', 'CUI', 'Proyecto', 'PIM 2026']):
        wsQ.write(q0, j, h, f_hd)
    q = q0 + 1
    for lab, rs, _que in items:
        for r in rs:
            wsQ.write_string(q, 0, lab[:60], f_txt); wsQ.write_string(q, 1, r['ue'], f_txt); wsQ.write_string(q, 2, r['cui'], f_txt)
            wsQ.write_string(q, 3, r['nombre'], f_txt); wsQ.write_number(q, 4, r['pim'], f_num); q += 1
    if q > q0 + 1:
        wsQ.autofilter(q0, 0, q - 1, 4)
    wsQ.set_column(0, 0, 80); wsQ.set_column(1, 1, 14); wsQ.set_column(2, 2, 18); wsQ.set_column(3, 3, 90); wsQ.set_column(4, 4, 16)

    # ------------------------------------------------------------ Reglas
    wsX.write(0, 0, 'Reglas aplicadas (se evalúan en este orden)', f_title)
    reglas = [
        'SEMÁFORO de cada fila UE-CUI-meta. Los parámetros viven en config.json (monto_minimo, umbral_ejecutado, umbral_avance_acumulado, anexo1_cuis).',
        '  1. Sin PIM: PIM = 0.',
        '  2. Bolsa genérica: códigos 2001621 (estudios de preinversión) y 2016766 (iniciativa a la competitividad). No son inversiones con CUI propio.',
        '  3. NO MOVER: devengado 2026 / PIM ≥ 90%, o avance financiero acumulado (devengado acumulado / costo actualizado) > 90% (Art. 11 Ley 32513).',
        '  4. RESTRINGIDO: tipo PROCOMPITE o CUI de la lista del Anexo I (recursos con destino específico). El Anexo I se carga en config.json → anexo1_cuis (pendiente de la OPMI).',
        '  5. SIN SALDO ÚTIL: saldo libre menor al monto mínimo (S/ 10 000). Si además hay S/ 10 000 o más certificados sin comprometer ni programar en el 12-B: REBAJANDO CERTIFICACIÓN (azul).',
        '  6. CON CONDICIONES (amarillo): hay saldo libre pero está programado en el PMI 2026, tiene monto en el PMI 2027, no tiene 12-B vigente o no tiene ficha SSI.',
        '  7. DISPONIBLE (verde): saldo libre y sin ninguna de las condiciones anteriores (no está en el PMI 2026 ni 2027).',
        '',
        'POTENCIAL rebajando certificación = menor entre (PIM − 12-B) y (PIM − compromiso), menos el saldo libre: lo certificado que no está comprometido ni programado. Solo se puede usar si la UEI rebaja la certificación.',
        'SALDO LIBRE (Disponible legal, por CUI y repartido entre sus filas): menor entre (PIM − programación 12-B 2026 actualizada) y (PIM − certificación). Si no hay 12-B vigente: PIM − certificación.',
        '  EXCEPCIÓN: si el proyecto NO está en el PMI 2026, el 12-B no limita el saldo (saldo libre = PIM − certificación) y se marca "Con condiciones": la UEI debe actualizar su 12-B (impedimento comprobable, num. 11.2).',
        '  La programación 12-B 2026 suma los 12 meses (enero a la fecha = devengado real; resto = programado). Lo no programado ni certificado es lo que puede cederse.',
        '  Restricciones del Art. 11 que NO se pueden verificar con datos públicos: deudas contractuales por devengar y fondos de transferencias condicionadas. Validar antes de decidir.',
        'PMI: 2026 = monto 2026 de la Cartera PMI; 2027 = monto 2027. Un proyecto fuera del PMI 2026 respalda el "impedimento comprobable" (num. 11.2); con monto en 2027 es más difícil ceder su presupuesto.',
        'TIPO DE INVERSIÓN (proyecto, IOARR, PROCOMPITE): informativo, salvo PROCOMPITE que es Restringido. No cambia el semáforo del resto.',
        'No se usa ningún % de compromiso: se trabaja con saldos (sin certificar, certificado no comprometido, sin devengar).',
        '',
        'FUENTES: PIM, certificado, compromiso y devengado = API de ejecución del MEF (o reporte SIAF); tipo, estado, cierre y costo = ficha SSI; programación y devengado acumulado = 12-B (Datos Abiertos);',
        '  PMI 2026-2029 = Cartera PMI; gerencia = maestro.xlsx (Alex) + Silucia. Ver hoja Config para la fecha de cada fuente.',
        'MODALIDAD DE EJECUCIÓN (derivada del SIAF, específicas con PIM > 0): Contrata · Administración directa · Expediente técnico · Estudio de preinversión · Laudos · Equipamiento · Personal/bienes/servicios.',
        'GERENCIA / GRUPO (UE 001): maestro.xlsx. Los proyectos nuevos aparecen en "Sin clasificar" y se agregan solos al maestro (completar la Gerencia).',
        'Otras UE: una sola hoja ("Otras UE") agrupada por unidad ejecutora y función.']
    for i, t in enumerate(reglas):
        wsX.write(2 + i, 0, t, f_bold if t.startswith('SEMÁFORO') or t.startswith('SALDO LIBRE') else f_txt)
    wsX.set_column(0, 0, 190)
    aplicar_formato(WS, meta_info.get('formato'), {'Dónde hay plata': HDH, 'Base': BASE_COLS, 'Cambios': HC, 'Rubro por CUI': HRB, 'Lista Antony': HAN})
    if not (meta_info.get('formato') and 'Resumen' in meta_info['formato'].get('ocultas', [])):
        wsR.activate()
    wb.close()
    return dict(anexos=anexos, tot_hoja={k: v[1] for k, v in tot_hoja.items()})
