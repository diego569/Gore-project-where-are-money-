"""Construye el Excel de PROYECTOS (todas las UE con proyectos) en el formato de Alex + columnas nuevas.

La hoja Base es la fuente; las hojas de anexo leen de ella por fórmula (con valores ya calculados
para que se vean sin recalcular). Para gasto corriente se reutilizan siaf.py, maestro.py y reglas.py.
"""
import collections, datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

from .reglas import clasificar, ESTADOS, DISPONIBILIDAD
from .siaf import texto_modalidad
from .maestro import ubicar, HOJA_INFRA, HOJA_GER, INFRA

SIN = 'SIN CLASIFICAR (pendiente de asignar gerencia)'
NOCART = 'No registrado en cartera'
RUBROS = [('00', 'RO'), ('09', 'RDR'), ('15', 'FONCOR'), ('18', 'CANON'), ('19', 'ROOC')]
BASE_COLS = ['UE', 'Hoja', 'Gerencia / Oficina', 'Grupo (Alex)', 'Función', 'CUI', 'Meta', 'Proyecto', 'PIM 2026', 'Certificación',
             'Compromiso anual', 'Devengado', 'Saldo sin certificar (A-B)', 'Certificado no comprometido', 'Saldo sin devengar (A-C)',
             '% Devengado', '% Comprometido', 'Estado situacional', 'Disponibilidad', 'Disponible (PIM − compromiso)', 'Modalidad de ejecución',
             'Monto Año 2026 (S/)', 'Monto Año 2027 (S/)', 'Monto Año 2028 (S/)', 'Monto Año 2029 (S/)', 'Base legal', 'Observación', 'Nota de clasificación', 'Ejecutora (Silucia)']
L = {n: xl_col_to_name(i) for i, n in enumerate(['ue', 'hoja', 'ger', 'grupo', 'fun', 'cui', 'meta', 'nombre', 'pim', 'cert', 'comp', 'dev', 's_cert',
                                                  'cert_nc', 's_dev', 'pdev', 'pcomp', 'estado', 'disp', 'disponible', 'mod', 'm26', 'm27', 'm28', 'm29', 'bl', 'obs', 'nota'])}
B0 = 5   # primera fila de datos de Base (1-indexada)


def _fun_label(fn, funcs):
    return f"{fn}. {funcs.get(fn, '(sin nombre)')}" if fn else '(sin función)'


def armar_filas(siaf, cartera, maestro, cfg, silucia=None, mapa=None):
    rows = []; usados = set(); pendientes = []
    metas = [m for m in siaf['metas'] if m['tipo'] == 'proyecto']
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
            hoja = f'UE {ue}'; grupo = ''; path = (f"{ue}. {siaf['ues'][ue]}", fl); orden = 10 ** 7 - m['pim']
        comp = maestro['comp'].get((ue, m['cui'], m['meta'])) or maestro['comp'].get((ue, m['cui'], '')) or {}
        if c:
            montos = [c['m2026'] or 0, c['m2027'] or 0, c['m2028'] or 0, c['m2029'] or 0]
        else:
            montos = [NOCART] * 4
        est, disp, dispo = clasificar(m['pim'], m['cert'], m['comp'], m['dev'], cfg)
        rows.append(dict(ue=ue, hoja=hoja, ger=path[0], grupo=grupo, fun=path[-1] if len(path) > 1 else '', path=path, orden=orden, cui=m['cui'],
                         meta=m['meta'], nombre=nombre, pim=m['pim'], cert=m['cert'], comp=m['comp'], dev=m['dev'], estado=est, disp=disp,
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
    colores = {'Disponible': '#C6EFCE', 'Posible': '#E2EFDA', 'Parcial': '#FFEB9C', 'Difícil': '#F8CBAD', 'No usar': '#D9D9D9'}

    def cond_disp(ws, rango):
        for k, col in colores.items():
            ws.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': f'"{k}"', 'format': wb.add_format({'bg_color': col})})

    fecha_txt = meta_info['fecha_corte'].strftime('%d/%m/%Y')
    ues_orden = sorted({r['ue'] for r in rows}, key=lambda u: (u != '001', u))
    nrows = len(rows); BL = B0 + nrows - 1

    # ------------------------------------------------------------ Config (se crea primero para que exista la referencia)
    wsR = wb.add_worksheet('Resumen')
    sheets_anexo = []
    # orden de hojas: Resumen, anexos, Saldo expediente, Base, Control, Config, Reglas
    anexos = []
    if any(r['hoja'] == HOJA_INFRA for r in rows): anexos.append(HOJA_INFRA)
    if any(r['hoja'] == HOJA_GER for r in rows): anexos.append(HOJA_GER)
    if any(r['hoja'] == 'Sin clasificar' for r in rows): anexos.append('Sin clasificar')
    anexos += [f'UE {u}' for u in ues_orden if u != '001']
    wsA = {nm: wb.add_worksheet(nm) for nm in anexos}
    wsS = wb.add_worksheet('Saldo al expediente')
    wsRB = wb.add_worksheet('Rubro por CUI')
    wsB = wb.add_worksheet('Base'); wsC = wb.add_worksheet('Control'); wsG = wb.add_worksheet('Config'); wsX = wb.add_worksheet('Reglas')

    wsG.write(0, 0, 'Configuración (editable)', f_title)
    wsG.write(2, 0, 'Parámetro', f_hd); wsG.write(2, 1, 'Valor', f_hd); wsG.write(2, 2, 'Significado', f_hd)
    f_inp = F(num_format='0%', bg_color='#FFF2CC', font_color='#0000FF', border=1)
    wsG.write(4, 0, 'Umbral "Ejecutado"', f_txt); wsG.write_number(4, 1, cfg['umbral_ejecutado'], f_inp)
    wsG.write(4, 2, 'Devengado / PIM mayor o igual a este valor → "Ejecutado" (no usar).', f_txt)
    wsG.write(5, 0, 'Umbral "Mayormente comprometido"', f_txt); wsG.write_number(5, 1, cfg['umbral_comprometido'], f_inp)
    wsG.write(5, 2, 'Compromiso anual / PIM mayor o igual a este valor → "Mayormente comprometido" (difícil).', f_txt)
    wsG.write(7, 0, 'Datos de esta corrida', f_bold)
    info = [('Fecha de corte', fecha_txt), ('Archivo SIAF', meta_info['siaf']), ('Fecha del archivo SIAF', meta_info['siaf_fecha']),
            ('Archivo Cartera', meta_info['cartera']), ('Fecha de la Cartera', cartera['fecha'].strftime('%d/%m/%Y %H:%M') if cartera['fecha'] else 'No disponible'),
            ('Generado', datetime.datetime.now().strftime('%d/%m/%Y %H:%M'))]
    for i, (k, v) in enumerate(info):
        wsG.write(8 + i, 0, k, f_txt); wsG.write(8 + i, 1, v, f_txt)
    r0 = 8 + len(info) + 1
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
    CE, CC = 'Config!$B$5', 'Config!$B$6'

    # ------------------------------------------------------------ Base
    wsB.write(0, 0, f'Base por UE-CUI-meta · proyectos · al {fecha_txt}', f_title)
    wsB.write(1, 0, 'Fuentes: SIAF (cifras, modalidad), Cartera PMI (montos 2026-2029), maestro.xlsx (gerencia, base legal). Azul = dato de entrada; negro = fórmula.', f_note)
    wsB.set_row(B0 - 2, 40)
    for j, h in enumerate(BASE_COLS):
        wsB.write(B0 - 2, j, h, f_hd)
    for r in rows:
        n = r['n']; i = n - 1
        est = r['estado']
        pdev = r['dev'] / r['pim'] if r['pim'] > 0 else 0; pcomp = r['comp'] / r['pim'] if r['pim'] > 0 else 0
        txt = [r['ue'], r['hoja'], r['ger'], r['grupo'], r['fun'], r['cui'], r['meta'], r['nombre']]
        for j, v in enumerate(txt):
            wsB.write_string(i, j, v, f_wrap if j == 7 else f_txt)
        for j, k in ((8, 'pim'), (9, 'cert'), (10, 'comp'), (11, 'dev')):
            wsB.write_number(i, j, r[k], f_in)
        c = lambda nm: f'{L[nm]}{n}'
        wsB.write_formula(i, 12, f"={c('pim')}-{c('cert')}", f_num, r['pim'] - r['cert'])
        wsB.write_formula(i, 13, f"={c('cert')}-{c('comp')}", f_num, r['cert'] - r['comp'])
        wsB.write_formula(i, 14, f"={c('pim')}-{c('dev')}", f_num, r['pim'] - r['dev'])
        wsB.write_formula(i, 15, f"=IF({c('pim')}>0,{c('dev')}/{c('pim')},0)", f_pct, pdev)
        wsB.write_formula(i, 16, f"=IF({c('pim')}>0,{c('comp')}/{c('pim')},0)", f_pct, pcomp)
        wsB.write_formula(i, 17, f'=IF({c("pim")}<=0,"Sin PIM",IF({c("dev")}/{c("pim")}>={CE},"Ejecutado",IF({c("comp")}/{c("pim")}>={CC},"Mayormente comprometido",'
                                 f'IF({c("comp")}>0.005,"Parcialmente comprometido",IF({c("cert")}>0.005,"Solo certificado","Sin certificar")))))', f_txt, est)
        e = c('estado')
        wsB.write_formula(i, 18, f'=IF({e}="Sin certificar","Disponible",IF({e}="Solo certificado","Posible",IF({e}="Parcialmente comprometido","Parcial",'
                                 f'IF({e}="Mayormente comprometido","Difícil",IF({e}="Ejecutado","No usar","Sin PIM")))))', f_txt, r['disp'])
        wsB.write_formula(i, 19, f'=IF(OR({c("pim")}<=0,{e}="Ejecutado"),0,MAX(0,{c("pim")}-{c("comp")}))', f_num, r['disponible'])
        wsB.write_string(i, 20, r['mod'], f_txt)
        for j, v in enumerate(r['montos']):
            if isinstance(v, str): wsB.write_string(i, 21 + j, v, f_txt)
            else: wsB.write_number(i, 21 + j, v, f_in)
        wsB.write_string(i, 25, r['bl'], f_txt); wsB.write_string(i, 26, r['obs'], f_txt); wsB.write_string(i, 27, r['nota'], f_txt); wsB.write_string(i, 28, r['ejecutora'], f_txt)
    wsB.set_column(0, 0, 5); wsB.set_column(1, 1, 24); wsB.set_column(2, 2, 36); wsB.set_column(3, 3, 30); wsB.set_column(4, 4, 26); wsB.set_column(5, 6, 9)
    wsB.set_column(7, 7, 60); wsB.set_column(8, 16, 14); wsB.set_column(17, 18, 22); wsB.set_column(19, 19, 16); wsB.set_column(20, 20, 30)
    wsB.set_column(21, 24, 14); wsB.set_column(25, 27, 24); wsB.set_column(28, 28, 30)
    wsB.freeze_panes(B0 - 1, 8); wsB.autofilter(B0 - 2, 0, BL - 1, len(BASE_COLS) - 1)
    cond_disp(wsB, f'S{B0}:S{BL}')

    # ------------------------------------------------------------ Anexos
    HDR = ['N°', 'META', 'PROYECTO', 'PIM 2026\n(A)', 'CERTIFICACIÓN\n(B)', 'SALDO\n(A-B)', 'DEVENGADO\n( C )', 'SALDO\n(A-C)', 'DEVEN\n%',
           'MODALIDAD\nDE\nEJECUCIÓN', 'ESTADO\nSITUACIONAL\nPROYECTO', 'Monto Año 2026 (S/)', 'Monto Año 2027 (S/)', 'Monto Año 2028 (S/)',
           'Monto Año 2029 (S/)', 'BASE LEGAL', 'DISPONIBILIDAD', 'CERTIFICADO NO\nCOMPROMETIDO', 'DISPONIBLE\n(PIM − COMPROM.)']
    W = [6, 8, 62, 15, 15, 14, 15, 15, 9, 28, 24, 15, 15, 15, 15, 18, 15, 16, 16]
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
            return dict(pim=r['pim'], cert=r['cert'], dev=r['dev'], nc=r['cert'] - r['comp'], dispo=r['disponible'])

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
                ws.write_formula(i, 10, f"=Base!{L['estado']}{n}", f_leaf, r['estado'])
                for k in range(4):
                    ws.write_formula(i, 11 + k, f"=Base!{L['m26'] if k == 0 else L['m27'] if k == 1 else L['m28'] if k == 2 else L['m29']}{n}",
                                     f_leafn, r['montos'][k])
                ws.write_formula(i, 15, f'=IF(Base!{L["bl"]}{n}="","",Base!{L["bl"]}{n})', f_leaf, r['bl'])
                ws.write_formula(i, 16, f"=Base!{L['disp']}{n}", f_leafc, r['disp'])
                ws.write_formula(i, 17, f"=Base!{L['cert_nc']}{n}", f_leafn, v['nc'])
                ws.write_formula(i, 18, f"=Base!{L['disponible']}{n}", f_leafn, v['dispo'])
                suma.update({'D': r['pim'], 'E': r['cert'], 'G': r['dev'], 'R': v['nc'], 'S': v['dispo']})
            hijos = []
            for h in nd.hijos.values():
                fr, sm, _h = emit(h); hijos.append(fr); suma.update(sm)
            if nd.nivel == 0:
                return None, suma, hijos
            fila = nd.fila; i = fila - 1
            clave = 'yel' if nd.nivel == 1 else ('sal' if nd.nombre[:2].isdigit() and nd.nombre[2:3] in '.:' else 'gry')
            fx, fn_, fp = sub[clave]
            ws.write(i, 0, nd.nombre, fx)
            for j in (1, 2, 9, 10, 11, 12, 13, 14, 15, 16):
                ws.write_blank(i, j, None, fx)
            for col, key in (('D', 'D'), ('E', 'E'), ('G', 'G'), ('R', 'R'), ('S', 'S')):
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
        for j in (1, 2, 9, 10, 11, 12, 13, 14, 15, 16):
            ws.write_blank(i, j, None, fx)
        for col in 'DEGRS':
            ws.write_formula(i, ord(col) - 65, '=' + '+'.join(f'{col}{k}' for k in tops), fn_, total[col])
        ws.write_formula(i, 5, f'=D{tr}-E{tr}', fn_, total['D'] - total['E'])
        ws.write_formula(i, 7, f'=D{tr}-G{tr}', fn_, total['D'] - total['G'])
        ws.write_formula(i, 8, f'=IF(D{tr}=0,0,G{tr}/D{tr})', fp, total['G'] / total['D'] if total['D'] else 0)
        ws.freeze_panes(4, 3); ws.set_zoom(90)
        cond_disp(ws, f'Q5:Q{tr}')
        tot_hoja[ws.get_name()] = (tr, total)

    for k, nm in enumerate(anexos, 1):
        rs = [r for r in rows if r['hoja'] == nm]
        titulo = {'Sin clasificar': 'PROYECTOS SIN GERENCIA ASIGNADA'}.get(nm, nm.upper() if nm.startswith('UE ') else nm)
        escribir_anexo(wsA[nm], titulo, rs, k)

    # ------------------------------------------------------------ Resumen
    wsR.write(0, 0, f'Proyectos de inversión · GORE Puno · mapeo de disponibilidad · al {fecha_txt}', f_title)
    wsR.write(1, 0, 'Montos en soles. Todo se calcula desde la hoja Base. Los umbrales se cambian en la hoja Config.', f_note)
    if meta_info.get('banner'):
        wsR.write(2, 0, meta_info['banner'], F(bold=True, font_color='#C00000', bg_color='#FFF2CC'))
    heads = ['Unidad', 'Proyectos (UE-CUI-meta)', 'PIM 2026', 'Certificación', 'Compromiso anual', 'Devengado', 'Saldo sin certificar', 'Disponible (PIM − comprom.)'] + \
            [f'PIM · {e}' for e in ESTADOS]
    rng = lambda nm: f"Base!${L[nm]}${B0}:${L[nm]}${BL}"

    def tabla(r0, titulo, claves, colcrit, filtro_ue=None):
        wsR.write(r0, 0, titulo, f_bold)
        for j, h in enumerate(heads):
            wsR.write(r0 + 1, j, h, f_hd)
        wsR.set_row(r0 + 1, 40)
        tot = collections.Counter()
        for i, (clave, etiqueta, rs) in enumerate(claves):
            n = r0 + 3 + i; ix = n - 1
            wsR.write_string(ix, 0, etiqueta, f_leaf)
            crit = f'{rng(colcrit)},$A{n}' if colcrit != 'ue' else f'{rng("ue")},"{clave}"'
            if filtro_ue: crit = f'{rng("ue")},"{filtro_ue}",{rng("ger")},"{clave}"'
            val = lambda k: sum(r[k] for r in rs)
            wsR.write_formula(ix, 1, f'=COUNTIFS({crit})', f_leafn, len(rs))
            for j, k in ((2, 'pim'), (3, 'cert'), (4, 'comp'), (5, 'dev')):
                wsR.write_formula(ix, j, f'=SUMIFS({rng(k)},{crit})', f_leafn, val(k)); tot[j] += val(k)
            wsR.write_formula(ix, 6, f'=C{n}-D{n}', f_leafn, val('pim') - val('cert')); tot[6] += val('pim') - val('cert')
            wsR.write_formula(ix, 7, f'=SUMIFS({rng("disponible")},{crit})', f_leafn, val('disponible')); tot[7] += val('disponible')
            tot[1] += len(rs)
            for k, e in enumerate(ESTADOS):
                v = sum(r['pim'] for r in rs if r['estado'] == e); tot[8 + k] += v
                wsR.write_formula(ix, 8 + k, f'=SUMIFS({rng("pim")},{crit},{rng("estado")},"{e}")', f_leafn, v)
        n = r0 + 3 + len(claves); ix = n - 1
        fx, fn_, fp = sub['gry']
        wsR.write(ix, 0, 'TOTAL', fx)
        for j in range(1, 8 + len(ESTADOS)):
            wsR.write_formula(ix, j, f'=SUM({xl_col_to_name(j)}{r0 + 3}:{xl_col_to_name(j)}{n - 1})', fn_, tot[j])
        return n + 2

    por_ue = [(u, f"{u}. {siaf['ues'][u]}", [r for r in rows if r['ue'] == u]) for u in ues_orden]
    nxt = tabla(3, 'Por unidad ejecutora', [(u, e, rs) for u, e, rs in por_ue], 'ue')
    gers = list(collections.OrderedDict.fromkeys(r['ger'] for r in rows if r['ue'] == '001'))
    por_ger = [(g, g, [r for r in rows if r['ue'] == '001' and r['ger'] == g]) for g in gers]
    tabla(nxt, 'UE 001 Sede Central · por gerencia / oficina', por_ger, 'ger')
    wsR.set_column(0, 0, 58); wsR.set_column(1, 1, 14); wsR.set_column(2, 13, 16); wsR.freeze_panes(0, 1)

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
             ('Filas con PIM > 0 y Estado "Sin PIM" (no debería ocurrir)', sum(1 for r in rows if r['pim'] > 0 and r['estado'] == 'Sin PIM')),
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

    # ------------------------------------------------------------ Reglas
    wsX.write(0, 0, 'Reglas aplicadas', f_title)
    reglas = [
        'ESTADO SITUACIONAL (se evalúa en este orden; umbrales en Config):',
        '  1. Sin PIM: PIM = 0.',
        '  2. Ejecutado: devengado ≥ umbral de ejecutado (90%) del PIM. Disponibilidad "No usar"; el disponible es 0.',
        '  3. Mayormente comprometido: compromiso anual ≥ umbral (50%) del PIM. Disponibilidad "Difícil".',
        '  4. Parcialmente comprometido: compromiso anual > 0 y menor al umbral. Disponibilidad "Parcial".',
        '  5. Solo certificado: certificación > 0 y compromiso = 0. Disponibilidad "Posible".',
        '  6. Sin certificar: certificación = 0. Disponibilidad "Disponible".',
        'DISPONIBLE (PIM − compromiso): 0 si está Ejecutado o sin PIM. Columna "Certificado no comprometido" = certificación − compromiso.',
        'CAVEAT: el compromiso anual difiere entre el reporte SIAF, la Consulta Amigable y la API del MEF (hasta S/ 9 M a nivel UE 001). Validar antes de decidir.',
        '',
        'MODALIDAD DE EJECUCIÓN (derivada del SIAF, específicas con PIM > 0, ordenadas por monto, separadas por " / "):',
        '  Contrata = "Costo de construcción por contrata" · Administración directa = "...por administración directa" · Expediente técnico · Estudio de preinversión ·',
        '  Laudos arbitrales · Equipamiento (2.6.3.x) · Personal/bienes/servicios (2.6.8.1.4 y 2.6.7.1.6; confirmar con Alex si equivale a administración directa).',
        '',
        'GERENCIA / GRUPO (UE 001): viene de maestro.xlsx (hoja Clasificacion, partiendo del Excel de Alex del 16/09). Los proyectos nuevos aparecen en "Sin clasificar" y se',
        'agregan solos al maestro con la Gerencia vacía: completar la Gerencia (y el Grupo si es Infraestructura) y en la siguiente corrida quedan ubicados.',
        'Otras UE: se agrupan por unidad ejecutora y función; no requieren maestro.',
        'MONTOS 2026-2029: Cartera PMI. "No registrado en cartera" = el CUI no está en el reporte de Cartera (no es cero).',
        'BASE LEGAL y OBSERVACIÓN: hoja Complementos de maestro.xlsx (UE, CUI, meta).']
    for i, t in enumerate(reglas):
        wsX.write(2 + i, 0, t, f_bold if t.endswith(':') else f_txt)
    wsX.set_column(0, 0, 160)
    wb.close()
    return dict(anexos=anexos, tot_hoja={k: v[1] for k, v in tot_hoja.items()})
