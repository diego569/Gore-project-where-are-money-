"""Excel de GASTO CORRIENTE (actividades; por ahora solo la UE 001), mismo estilo que el de proyectos.

La clasificación meta -> gerencia/oficina vive en maestro.xlsx (hoja ClasifGC). Si Silucia trae una ejecutora distinta de
"GASTO CORRIENTE" para la meta y esa ejecutora está en la hoja MapaEjecutoras, se usa como sugerencia.
"""
import collections, datetime
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from .reglas import clasificar, ESTADOS
from .excel_proyectos import Nodo, arbol

SIN_OF = 'SIN CLASIFICAR (completar en maestro.xlsx, hoja ClasifGC)'
RUBROS_GC = [('00', 'RO'), ('09', 'RDR'), ('13', 'DyT'), ('15', 'FONCOR'), ('18', 'CANON')]
GEN = [('2.1.', 'Personal y oblig. sociales'), ('2.3.', 'Bienes y servicios'), ('2.5.', 'Otros gastos')]
NOCART = 'No registrado'
BASE_H = ['UE', 'Gerencia / Oficina', 'Función', 'Producto', 'Actividad', 'Meta', 'Actividad / meta', 'PIM 2026', 'Certificación', 'Compromiso anual', 'Devengado',
          'Saldo sin certificar (A-B)', 'Certificado no comprometido', 'Saldo sin devengar (A-C)', '% Devengado', '% Comprometido', 'Estado situacional',
          'Disponibilidad', 'Disponible (PIM − compromiso)', 'PIM 2.1 Personal', 'PIM 2.3 Bienes y serv.', 'PIM 2.5 Otros gastos', 'PIM otras genéricas',
          'PIM · RO (00)', 'PIM · RDR (09)', 'PIM · DyT (13)', 'PIM · FONCOR (15)', 'PIM · CANON (18)', 'PIM · otros rubros', 'Ejecutora (Silucia)', 'Nota']
B0 = 5


def armar_filas_gc(siaf, maestro, cfg, ue='001', silucia=None, mapa=None, rubros_meta=None, nombres_fn=None):
    gc_cls, mapa_m = maestro.get('gc', {}), mapa if mapa is not None else maestro.get('mapa', {})
    funcs = dict(maestro['funcs']); funcs.update({k: v for k, v in (nombres_fn or {}).items() if k not in funcs})
    metas = sorted([m for m in siaf['metas'] if m['ue'] == ue and m['tipo'] == 'actividad'], key=lambda m: (-m['pim'], m['meta']))
    rows = []; pend = []
    for m in metas:
        fl = f"{m['fn']}. {funcs.get(m['fn'], '(sin nombre)')}" if m['fn'] else '(sin función)'
        sug = (silucia or {}).get((m['cui'], m['meta']))
        ejec = sug['ejecutora'] if sug else ''
        of = gc_cls.get((ue, m['meta']), '')
        nota = ''
        if not of and ejec and mapa_m.get(ejec):
            of = mapa_m[ejec]; nota = f'Oficina sugerida por Silucia (ejecutora "{ejec}"): confirmar'
        nombre = m['nombre'] + ((' · ' + m['meta_desc'].split(';')[0]) if m.get('meta_desc') and m['meta_desc'].split(';')[0].strip() != m['nombre'].strip() else '')
        if (ue, m['meta']) not in gc_cls:
            pend.append([ue, m['meta'], of, fl, nombre[:200], ('Silucia (sugerido) ' if of else 'Nuevo ') + str(datetime.date.today())])
        est, disp, dispo = clasificar(m['pim'], m['cert'], m['comp'], m['dev'], cfg)
        g = m['gen']; otras = sum(v for k, v in g.items() if k not in dict(GEN))
        rb = (rubros_meta or {}).get((m['cui'], m['meta']))
        if rb is None:
            rv = [None] * 6
        else:
            std = [rb.get(c, 0.0) for c, _ in RUBROS_GC]; rv = std + [sum(v for k, v in rb.items() if k not in dict(RUBROS_GC))]
        rows.append(dict(ue=ue, of=of or SIN_OF, fun=fl, prod=m['cui'], act=m['act'], meta=m['meta'], nombre=nombre, pim=m['pim'], cert=m['cert'], comp=m['comp'],
                         dev=m['dev'], estado=est, disp=disp, disponible=dispo, gen=[g.get(k, 0.0) for k, _ in GEN] + [otras], rub=rv, ejec=ejec, nota=nota))
    rows.sort(key=lambda r: (r['of'] == SIN_OF, r['of'], r['fun'], -r['pim']))
    for i, r in enumerate(rows):
        r['n'] = B0 + i
    return rows, pend


def construir_gc(path, siaf, rows, cfg, info, ue='001', rubro_fuente=''):
    wb = xlsxwriter.Workbook(str(path), {'nan_inf_to_errors': True}); wb.set_calc_mode('auto')
    b = dict(font_name='Arial', font_size=9)
    F = lambda **k: wb.add_format({**b, **k})
    f_txt = F(); f_wrap = F(text_wrap=True, valign='vcenter'); f_num = F(num_format='#,##0;(#,##0);"-"', align='right'); f_in = F(num_format='#,##0;(#,##0);"-"', align='right', font_color='#0000FF')
    f_pct = F(num_format='0.0%', align='center'); f_ctr = F(align='center', valign='vcenter'); f_bold = F(bold=True); f_title = F(bold=True, font_size=12); f_note = F(italic=True, font_color='#595959')
    f_hd = F(bold=True, font_color='#FFFFFF', bg_color='#002060', text_wrap=True, align='center', valign='vcenter', border=1)
    f_hd2 = F(bold=True, italic=True, bg_color='#B4C6E7', text_wrap=True, align='center', valign='vcenter', border=1)
    sub = {}
    for nm, col in (('yel', '#FFFF00'), ('sal', '#F4B084'), ('navy', '#002060')):
        fc = '#FFFFFF' if nm == 'navy' else '#000000'
        sub[nm] = (F(bold=True, bg_color=col, font_color=fc, border=1), F(bold=True, bg_color=col, font_color=fc, border=1, num_format='#,##0;(#,##0);"-"'),
                   F(bold=True, bg_color=col, font_color=fc, border=1, num_format='0.0%', align='center'))
    f_leaf = F(border=1, text_wrap=True, valign='vcenter'); f_leafn = F(border=1, num_format='#,##0;(#,##0);"-"', align='right', valign='vcenter')
    f_leafp = F(border=1, num_format='0.0%', align='center', valign='vcenter'); f_leafc = F(border=1, align='center', valign='vcenter')
    colores = {'Disponible': '#C6EFCE', 'Posible': '#E2EFDA', 'Parcial': '#FFEB9C', 'Difícil': '#F8CBAD', 'No usar': '#D9D9D9'}

    def cond(ws, rango):
        for k, col in colores.items():
            ws.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': f'"{k}"', 'format': wb.add_format({'bg_color': col})})
    fecha_txt = info['fecha_corte'].strftime('%d/%m/%Y'); BL = B0 + len(rows) - 1
    wsR = wb.add_worksheet('Resumen'); wsA = wb.add_worksheet('Gasto corriente'); wsB = wb.add_worksheet('Base GC')
    wsC = wb.add_worksheet('Control'); wsG = wb.add_worksheet('Config'); wsX = wb.add_worksheet('Reglas')

    # ---- Config
    wsG.write(0, 0, 'Configuración (editable)', f_title)
    for j, h in enumerate(['Parámetro', 'Valor', 'Significado']):
        wsG.write(2, j, h, f_hd)
    fin = F(num_format='0%', bg_color='#FFF2CC', font_color='#0000FF', border=1)
    wsG.write(4, 0, 'Umbral "Ejecutado"', f_txt); wsG.write_number(4, 1, cfg['umbral_ejecutado'], fin); wsG.write(4, 2, 'Devengado / PIM ≥ este valor → "Ejecutado" (no usar).', f_txt)
    wsG.write(7, 0, 'Datos de esta corrida', f_bold)
    for i, (k, v) in enumerate([('Fecha de corte', fecha_txt), ('Fuente de cifras', info['fuente_siaf']), ('Fuente de rubro', rubro_fuente or 'No disponible'),
                                ('Silucia', info.get('silucia', '')), ('Generado', datetime.datetime.now().strftime('%d/%m/%Y %H:%M'))]):
        wsG.write(8 + i, 0, k, f_txt); wsG.write(8 + i, 1, v, f_txt)
    wsG.set_column(0, 0, 38); wsG.set_column(1, 1, 60); wsG.set_column(2, 2, 90)
    CE = 'Config!$B$5'
    # ---- Base
    wsB.write(0, 0, f'Base gasto corriente UE {ue} · al {fecha_txt}', f_title)
    wsB.write(1, 0, 'Azul = dato de entrada; negro = fórmula. "s/d" en rubro = no se obtuvo de la API.', f_note); wsB.set_row(B0 - 2, 40)
    for j, h in enumerate(BASE_H):
        wsB.write(B0 - 2, j, h, f_hd)
    LB = {n: xl_col_to_name(i) for i, n in enumerate(['ue', 'of', 'fun', 'prod', 'act', 'meta', 'nombre', 'pim', 'cert', 'comp', 'dev', 's_cert', 'cert_nc', 's_dev', 'pdev',
                                                     'pcomp', 'estado', 'disp', 'disponible', 'g21', 'g23', 'g25', 'gotras', 'r00', 'r09', 'r13', 'r15', 'r18', 'rotros', 'ejec', 'nota'])}
    for r in rows:
        n = r['n']; i = n - 1; c = lambda k: f'{LB[k]}{n}'
        for j, k in enumerate(['ue', 'of', 'fun', 'prod', 'act', 'meta', 'nombre']):
            wsB.write_string(i, j, r[k], f_wrap if k == 'nombre' else f_txt)
        for j, k in ((7, 'pim'), (8, 'cert'), (9, 'comp'), (10, 'dev')):
            wsB.write_number(i, j, r[k], f_in)
        wsB.write_formula(i, 11, f"={c('pim')}-{c('cert')}", f_num, r['pim'] - r['cert'])
        wsB.write_formula(i, 12, f"={c('cert')}-{c('comp')}", f_num, r['cert'] - r['comp'])
        wsB.write_formula(i, 13, f"={c('pim')}-{c('dev')}", f_num, r['pim'] - r['dev'])
        wsB.write_formula(i, 14, f"=IF({c('pim')}>0,{c('dev')}/{c('pim')},0)", f_pct, r['dev'] / r['pim'] if r['pim'] > 0 else 0)
        wsB.write_formula(i, 15, f"=IF({c('pim')}>0,{c('comp')}/{c('pim')},0)", f_pct, r['comp'] / r['pim'] if r['pim'] > 0 else 0)
        wsB.write_formula(i, 16, f'=IF({c("pim")}<=0,"Sin PIM",IF({c("dev")}/{c("pim")}>={CE},"Ejecutado",'
                                 f'IF({c("comp")}>0.005,"Con compromiso",IF({c("cert")}>0.005,"Solo certificado","Sin certificar"))))', f_txt, r['estado'])
        e = c('estado')
        wsB.write_formula(i, 17, f'=IF({e}="Sin certificar","Disponible",IF({e}="Solo certificado","Posible",IF({e}="Con compromiso","Parcial",'
                                 f'IF({e}="Ejecutado","No usar","Sin PIM"))))', f_txt, r['disp'])
        wsB.write_formula(i, 18, f'=IF(OR({c("pim")}<=0,{e}="Ejecutado"),0,MAX(0,{c("pim")}-{c("comp")}))', f_num, r['disponible'])
        for j, v in enumerate(r['gen']):
            wsB.write_number(i, 19 + j, v, f_in)
        for j, v in enumerate(r['rub']):
            if v is None: wsB.write_string(i, 23 + j, 's/d', f_ctr)
            else: wsB.write_number(i, 23 + j, v, f_in)
        wsB.write_string(i, 29, r['ejec'], f_txt); wsB.write_string(i, 30, r['nota'], f_txt)
    wsB.set_column(0, 0, 5); wsB.set_column(1, 1, 40); wsB.set_column(2, 2, 30); wsB.set_column(3, 5, 10); wsB.set_column(6, 6, 70); wsB.set_column(7, 15, 14)
    wsB.set_column(16, 17, 22); wsB.set_column(18, 28, 14); wsB.set_column(29, 30, 34)
    wsB.freeze_panes(B0 - 1, 7); wsB.autofilter(B0 - 2, 0, max(BL - 1, B0 - 1), len(BASE_H) - 1); cond(wsB, f'R{B0}:R{max(BL, B0)}')
    rng = lambda k: f"'Base GC'!${LB[k]}${B0}:${LB[k]}${max(BL, B0)}"

    # ---- Anexo
    HD = ['N°', 'META', 'ACTIVIDAD / META', 'PIM 2026\n(A)', 'CERTIFICACIÓN\n(B)', 'SALDO\n(A-B)', 'COMPROMISO\nANUAL', 'DEVENGADO\n( C )', 'SALDO\n(A-C)', 'DEVEN\n%',
          'ESTADO\nSITUACIONAL', 'DISPONIBILIDAD', 'CERTIFICADO NO\nCOMPROMETIDO', 'DISPONIBLE\n(PIM − COMPROM.)', 'PIM 2.1\nPersonal', 'PIM 2.3\nBienes y servicios',
          'PIM 2.5\nOtros gastos', 'PIM otras\ngenéricas', 'RO\n(00)', 'RDR\n(09)', 'DyT\n(13)', 'FONCOR\n(15)', 'CANON\n(18)', 'Otros\nrubros', 'EJECUTORA\n(SILUCIA)']
    WD = [6, 8, 70, 15, 15, 14, 15, 15, 15, 9, 24, 15, 16, 16, 14, 14, 14, 14, 13, 13, 13, 13, 13, 13, 22]
    wsA.write(0, 2, 'ANEXO · GASTO CORRIENTE', f_title)
    wsA.write(1, 1, f'COMPORTAMIENTO DE EJECUCIÓN DE GASTO CORRIENTE (ACTIVIDADES) · UE {ue} · AÑO FISCAL 2026 · AL {fecha_txt}', f_bold)
    wsA.set_row(3, 44)
    for j, h in enumerate(HD):
        wsA.write(3, j, h, f_hd if j <= 9 else f_hd2); wsA.set_column(j, j, WD[j])
    raiz = arbol(rows, clave=lambda r: (r['of'], r['fun']))
    SUMCOLS = {'D': 'pim', 'E': 'cert', 'H': 'dev', 'G': 'comp', 'M': 'nc', 'N': 'disponible', 'O': 'g0', 'P': 'g1', 'Q': 'g2', 'R': 'g3',
               'S': 'r0', 'T': 'r1', 'U': 'r2', 'V': 'r3', 'W': 'r4', 'X': 'r5'}
    cur = [5]; cnt = [0]; first = 6

    def valores(r):
        v = dict(pim=r['pim'], cert=r['cert'], dev=r['dev'], comp=r['comp'], nc=r['cert'] - r['comp'], disponible=r['disponible'])
        for k in range(4): v[f'g{k}'] = r['gen'][k]
        for k in range(6): v[f'r{k}'] = r['rub'][k] or 0.0
        return v

    def emit(nd):
        suma = collections.Counter(); filas_h = []; hijos = []
        if nd.nivel > 0:
            nd.fila = cur[0]; cur[0] += 1
        for r in nd.hojas:
            row = cur[0]; cur[0] += 1; cnt[0] += 1; filas_h.append(row); n = r['n']; i = row - 1; v = valores(r)
            bg = lambda k: f"'Base GC'!{LB[k]}{n}"
            wsA.write_formula(i, 0, f'=SUBTOTAL(103,$B${first}:B{row})', f_leafc, cnt[0]); wsA.write_number(i, 1, int(r['meta']), f_leafc)
            wsA.write_string(i, 2, r['nombre'], f_leaf)   # texto fijo para que Ctrl+B funcione
            wsA.write_formula(i, 3, '=' + bg('pim'), f_leafn, r['pim']); wsA.write_formula(i, 4, '=' + bg('cert'), f_leafn, r['cert'])
            wsA.write_formula(i, 5, f'=D{row}-E{row}', f_leafn, r['pim'] - r['cert'])
            wsA.write_formula(i, 6, '=' + bg('comp'), f_leafn, r['comp']); wsA.write_formula(i, 7, '=' + bg('dev'), f_leafn, r['dev'])
            wsA.write_formula(i, 8, f'=D{row}-H{row}', f_leafn, r['pim'] - r['dev'])
            wsA.write_formula(i, 9, f'=IF(D{row}=0,0,H{row}/D{row})', f_leafp, r['dev'] / r['pim'] if r['pim'] else 0)
            wsA.write_formula(i, 10, '=' + bg('estado'), f_leaf, r['estado']); wsA.write_formula(i, 11, '=' + bg('disp'), f_leafc, r['disp'])
            wsA.write_formula(i, 12, '=' + bg('cert_nc'), f_leafn, v['nc']); wsA.write_formula(i, 13, '=' + bg('disponible'), f_leafn, v['disponible'])
            for k, key in enumerate(['g21', 'g23', 'g25', 'gotras']):
                wsA.write_formula(i, 14 + k, '=' + bg(key), f_leafn, r['gen'][k])
            for k, key in enumerate(['r00', 'r09', 'r13', 'r15', 'r18', 'rotros']):
                val = r['rub'][k]
                wsA.write_formula(i, 18 + k, '=' + bg(key), f_leafn if val is not None else f_leafc, 's/d' if val is None else val)
            wsA.write_formula(i, 24, '=' + bg('ejec'), f_leaf, r['ejec'])
            suma.update(v)
        for h in nd.hijos.values():
            fr, sm = emit(h); hijos.append(fr); suma.update(sm)
        if nd.nivel == 0:
            return hijos, suma
        fila = nd.fila; i = fila - 1
        fx, fn_, fp = sub['yel' if nd.nivel == 1 else 'sal']
        wsA.write(i, 0, nd.nombre, fx)
        for j in range(1, len(HD)): wsA.write_blank(i, j, None, fx)
        for col, key in SUMCOLS.items():
            terms = []
            if filas_h: terms.append(f'SUM({col}{filas_h[0]}:{col}{filas_h[-1]})')
            terms += [f'{col}{k}' for k in hijos]
            wsA.write_formula(i, ord(col) - 65, '=' + '+'.join(terms), fn_, suma[key])
        wsA.write_formula(i, 5, f'=D{fila}-E{fila}', fn_, suma['pim'] - suma['cert']); wsA.write_formula(i, 8, f'=D{fila}-H{fila}', fn_, suma['pim'] - suma['dev'])
        wsA.write_formula(i, 9, f'=IF(D{fila}=0,0,H{fila}/D{fila})', fp, suma['dev'] / suma['pim'] if suma['pim'] else 0)
        return fila, suma
    tops, total = emit(raiz)
    tr = cur[0] + 1; i = tr - 1; fx, fn_, fp = sub['navy']
    wsA.write(i, 0, f'TOTAL GASTO CORRIENTE UE {ue}', fx)
    for j in range(1, len(HD)): wsA.write_blank(i, j, None, fx)
    for col, key in SUMCOLS.items():
        wsA.write_formula(i, ord(col) - 65, '=' + '+'.join(f'{col}{k}' for k in tops) if tops else '=0', fn_, total[key])
    wsA.write_formula(i, 5, f'=D{tr}-E{tr}', fn_, total['pim'] - total['cert']); wsA.write_formula(i, 8, f'=D{tr}-H{tr}', fn_, total['pim'] - total['dev'])
    wsA.write_formula(i, 9, f'=IF(D{tr}=0,0,H{tr}/D{tr})', fp, total['dev'] / total['pim'] if total['pim'] else 0)
    wsA.freeze_panes(4, 3); wsA.set_zoom(90); cond(wsA, f'L5:L{tr}')

    # ---- Resumen
    wsR.write(0, 0, f'Gasto corriente · UE {ue} · mapeo de disponibilidad · al {fecha_txt}', f_title)
    wsR.write(1, 0, 'Montos en soles. Calculado desde la hoja Base GC. Umbrales en Config.', f_note)
    if info.get('banner'):
        wsR.write(2, 0, info['banner'], F(bold=True, font_color='#C00000', bg_color='#FFF2CC'))
    heads = ['Gerencia / Oficina', 'Metas', 'PIM 2026', 'Certificación', 'Compromiso anual', 'Devengado', 'Disponible (PIM − comprom.)', 'PIM 2.1 Personal', 'PIM 2.3 Bienes y serv.', 'PIM 2.5 Otros gastos'] + \
            [f'PIM · {e}' for e in ESTADOS]
    wsR.set_row(3, 40)
    for j, h in enumerate(heads):
        wsR.write(3, j, h, f_hd)
    ofs = list(collections.OrderedDict.fromkeys(r['of'] for r in rows))
    tot = collections.Counter()
    for k, of in enumerate(ofs):
        n = 5 + k; i = n - 1; rs = [r for r in rows if r['of'] == of]; crit = f'{rng("of")},$A{n}'
        wsR.write_string(i, 0, of, f_leaf); wsR.write_formula(i, 1, f'=COUNTIFS({crit})', f_leafn, len(rs)); tot[1] += len(rs)
        for j, key, kk in ((2, 'pim', 'pim'), (3, 'cert', 'cert'), (4, 'comp', 'comp'), (5, 'dev', 'dev'), (6, 'disponible', 'disponible')):
            val = sum(r[kk] for r in rs); tot[j] += val
            wsR.write_formula(i, j, f'=SUMIFS({rng(key)},{crit})', f_leafn, val)
        for q, key in enumerate(['g21', 'g23', 'g25']):
            val = sum(r['gen'][q] for r in rs); tot[7 + q] += val
            wsR.write_formula(i, 7 + q, f'=SUMIFS({rng(key)},{crit})', f_leafn, val)
        for q, e in enumerate(ESTADOS):
            val = sum(r['pim'] for r in rs if r['estado'] == e); tot[10 + q] += val
            wsR.write_formula(i, 10 + q, f'=SUMIFS({rng("pim")},{crit},{rng("estado")},"{e}")', f_leafn, val)
    n = 5 + len(ofs); i = n - 1; fx, fn_, fp = sub['sal']
    wsR.write(i, 0, 'TOTAL', fx)
    for j in range(1, 10 + len(ESTADOS)):
        wsR.write_formula(i, j, f'=SUM({xl_col_to_name(j)}5:{xl_col_to_name(j)}{n - 1})', fn_, tot[j])
    wsR.set_column(0, 0, 60); wsR.set_column(1, 15, 16)

    # ---- Control
    wsC.write(0, 0, 'Control y conciliación', f_title)
    for j, h in enumerate(['Control', 'Calculado', 'Referencia', 'Diferencia', 'Resultado']):
        wsC.write(2, j, h, f_hd)
    f_c = F(num_format='#,##0.00;(#,##0.00);"-"', border=1); f_t = F(border=1); cr = [3]

    def chk(label, formula, calc, ref, tol=0.5):
        i = cr[0]; n = i + 1
        wsC.write_string(i, 0, label, f_t); wsC.write_formula(i, 1, formula, f_c, calc); wsC.write_number(i, 2, ref, f_c)
        wsC.write_formula(i, 3, f'=B{n}-C{n}', f_c, calc - ref); wsC.write_formula(i, 4, f'=IF(ABS(D{n})<={tol},"Conciliado","REVISAR")', f_t, 'Conciliado' if abs(calc - ref) <= tol else 'REVISAR')
        cr[0] += 1
    pim_t = sum(r['pim'] for r in rows)
    chk('Hoja Gasto corriente: PIM = Base GC', f"='Gasto corriente'!D{tr}", total['pim'], pim_t)
    chk('Hoja Gasto corriente: Devengado = Base GC', f"='Gasto corriente'!H{tr}", total['dev'], sum(r['dev'] for r in rows))
    chk(f'PIM de la Base GC = PIM de actividades de la UE {ue} en la fuente de cifras', f"=SUM({rng('pim')})", pim_t,
        sum(m['pim'] for m in siaf['metas'] if m['ue'] == ue and m['tipo'] == 'actividad'))
    if siaf.get('ue_tot') and ue in siaf['ue_tot']:
        pp = sum(m['pim'] for m in siaf['metas'] if m['ue'] == ue and m['tipo'] == 'proyecto')
        chk(f'PIM UE {ue}: proyectos + gasto corriente = TOTAL UE del reporte SIAF', f'={pp}+SUM({rng("pim")})', pp + pim_t, siaf['ue_tot'][ue]['pim'])
    cr[0] += 1; wsC.write(cr[0], 0, 'Alertas (informativo)', f_bold); cr[0] += 1
    sin = [r for r in rows if r['of'] == SIN_OF]
    for k, v in [('Metas sin gerencia/oficina asignada (completar en maestro.xlsx, hoja ClasifGC)', len(sin)),
                 ('Metas sin rubro (la API de Datos Abiertos no respondió)', sum(1 for r in rows if r['rub'][0] is None)),
                 ('Metas con oficina sugerida por Silucia (confirmar)', sum(1 for r in rows if r['nota']))]:
        wsC.write_string(cr[0], 0, k, f_t); wsC.write_number(cr[0], 1, v, f_t); cr[0] += 1
    if info.get('avisos'):
        cr[0] += 1; wsC.write(cr[0], 0, 'Avisos de la corrida', f_bold); cr[0] += 1
        for a in info['avisos']:
            wsC.write_string(cr[0], 0, a, F(font_color='#C00000')); cr[0] += 1
    wsC.set_column(0, 0, 100); wsC.set_column(1, 4, 20)
    # ---- Reglas
    wsX.write(0, 0, 'Reglas aplicadas', f_title)
    for i, t in enumerate(['Estado por saldos (sin umbral de % comprometido): Sin PIM · Ejecutado (devengado ≥ 90% del PIM, no usar) · Con compromiso · Solo certificado · Sin certificar (disponible).',
                           'Ojo: en gasto corriente (planillas, servicios básicos, contratos) un compromiso alto es normal y lo "no comprometido" puede ser pago por venir; el estado es una primera aproximación.',
                           'GERENCIA / OFICINA: maestro.xlsx, hoja ClasifGC (UE + meta). Las metas nuevas se agregan solas con la oficina vacía (celda amarilla).',
                           'Silucia (api/metas): las metas de gasto corriente llegan todas con la ejecutora "GASTO CORRIENTE" (sin oficina). Si Silucia empieza a traer otra ejecutora y esta figura en la hoja MapaEjecutoras, se sugiere esa oficina.',
                           'GENÉRICAS: 2.1 personal y obligaciones sociales, 2.3 bienes y servicios, 2.5 otros gastos; "otras" incluye 2.2, 2.4, 2.6 y 2.7.',
                           'RUBRO: PIM por rubro de financiamiento por meta, desde la API de Datos Abiertos del MEF ("s/d" si no respondió).']):
        wsX.write(2 + i, 0, t, f_txt)
    wsX.set_column(0, 0, 170)
    wb.close()
    return dict(total=total, tr=tr)
