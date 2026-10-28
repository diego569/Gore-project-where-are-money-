"""Lector del reporte SIAF 'Resumen del marco presupuestal y la ejecución del gasto' (.xls).

Devuelve metas a nivel (UE, código producto/proyecto, meta). Sirve para proyectos (código que empieza con 2)
y para gasto corriente (código que empieza con 3): el filtrado se hace después, no aquí.
"""
import re, collections
import xlrd

HDR = re.compile(r'^(\d{4})\s+(\d{4})\s+(\d{7})\s+(\d{7})\s+(.*?)(?:\s+(\d{2})\s+(\d{3})\s+(\d{4}))?\s*$')
META_RE = re.compile(r'^Meta:\s*(\d+)\s*-\s*(\d+)\s+(.*)$')
UE_RE = re.compile(r'UNIDAD EJECUTORA: (\d{3}) - (.*?)\s*\((\d+)\)')
RUBRO_RE = re.compile(r'^(\d{2}) ([A-ZÁÉÍÓÚÑ].*)$')
COLS = dict(pia=5, pim=7, cert=9, comp=13, dev=17)
UE_TOT_COLS = dict(pia=6, pim=8, cert=9, comp=13, dev=17)

MOD_RULES = [('CONTRATA', 'Contrata'), ('ADMINISTRACION DIRECTA', 'Administración directa'),
             ('EXPEDIENTES TECNICOS', 'Expediente técnico'), ('PREINVERSION', 'Estudio de preinversión'),
             ('LAUDOS', 'Laudos arbitrales')]


def arreglar_texto(s):
    """El XLS suele venir con tildes dañadas (AÃ‘O). Se intenta recuperar."""
    if not isinstance(s, str) or ('Ã' not in s and 'â' not in s):
        return s
    for enc in ('cp1252', 'latin1'):
        try:
            return s.encode(enc).decode('utf-8')
        except Exception:
            continue
    return s


def etiqueta_modalidad(code, desc):
    for k, v in MOD_RULES:
        if k in desc:
            return v
    if code.startswith('2.6. 3'):
        return 'Equipamiento'
    if re.match(r'^2\.6\. (8  1\. 4|7  1\. 6)', code):
        return 'Personal/bienes/servicios'
    return 'Otro (' + desc[:30] + ')'


def es_reporte_siaf(path):
    try:
        bk = xlrd.open_workbook(path, on_demand=True)
        try:
            sh = bk.sheet_by_index(0)
            txt = ' '.join(arreglar_texto(str(sh.cell_value(r, 1))) for r in range(min(8, sh.nrows)))
            return 'RESUMEN DEL MARCO PRESUPUESTAL' in txt.upper()
        finally:
            bk.release_resources()
    except Exception:
        return False


def leer_siaf(path):
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    num = lambda v: float(v) if v != '' else 0.0
    metas = {}; ues = {}; ue_tot = {}; pliego = None
    ue = None; cur = None; rub = None; con_rubro = False
    for r in range(sh.nrows):
        t = arreglar_texto(str(sh.cell_value(r, 1)))
        m = UE_RE.match(t)
        if m:
            if m.group(1) != ue:      # el encabezado se repite en cada página: solo reinicia al cambiar de UE
                cur = None; rub = None
            ue = m.group(1); ues[ue] = m.group(2).strip()
            continue
        if t.startswith('TOTAL UNIDAD EJECUTORA'):
            ue_tot[ue] = {k: num(sh.cell_value(r + 1, c)) for k, c in UE_TOT_COLS.items()}
            continue
        if t.startswith('TOTAL PLIEGO'):
            pliego = {k: num(sh.cell_value(r + 1, c)) for k, c in UE_TOT_COLS.items()}
            continue
        h = HDR.match(t)
        if h:
            sec, _prod, cui, act, nom, fn, dvf, grp = h.groups()
            key = (ue, cui, sec)
            cur = metas.get(key)
            if cur is None:
                cur = metas[key] = dict(ue=ue, cui=cui, act=act, meta=sec, fn=fn or '', nombre=nom.strip(),
                                        tipo='proyecto' if cui[0] == '2' else 'actividad',
                                        pia=0.0, pim=0.0, cert=0.0, comp=0.0, dev=0.0,
                                        rubros=collections.defaultdict(lambda: dict(pim=0.0, cert=0.0, comp=0.0, dev=0.0, nombre='')),
                                        mods=collections.Counter(), mods_any=collections.Counter(), gen=collections.defaultdict(float), meta_desc='', finalidad='')
            elif fn and not cur['fn']:
                cur['fn'] = fn
            continue
        if t.startswith('Meta:') and cur is not None:
            mm = META_RE.match(t)
            if mm and not cur.get('meta_desc'):
                cur['finalidad'] = mm.group(2); cur['meta_desc'] = mm.group(3).strip()
            continue
        if sh.cell_value(r, 2) == '' and RUBRO_RE.match(t) and not t.startswith(('TOTAL', 'Meta')):
            rr = RUBRO_RE.match(t); rub = (rr.group(1), rr.group(2).strip()); con_rubro = True
            continue
        if cur is None:
            continue
        if re.match(r'^2\.\d\.', t):
            v = {k: num(sh.cell_value(r, c)) for k, c in COLS.items()}
            for k in v:
                cur[k] += v[k]
            if rub:
                d = cur['rubros'][rub[0]]; d['nombre'] = rub[1]
                for k in ('pim', 'cert', 'comp', 'dev'):
                    d[k] += v[k]
            lab = etiqueta_modalidad(t, arreglar_texto(str(sh.cell_value(r, 2))))
            cur['mods'][lab] += max(v['pim'], 0)
            cur['mods_any'][lab] += sum(abs(x) for x in v.values())
            cur['gen'][t[:5].strip()] += v['pim']
    for mm in metas.values():
        mm['rubros'] = {k: dict(v) for k, v in mm['rubros'].items()}
        mm['gen'] = dict(mm['gen'])
    return dict(metas=list(metas.values()), ues=ues, ue_tot=ue_tot, pliego=pliego, con_rubro=con_rubro)


def texto_modalidad(m):
    c = {k: v for k, v in m['mods'].items() if v > 0} or {k: v for k, v in m['mods_any'].items() if v > 0}
    return ' / '.join(k for k, _ in sorted(c.items(), key=lambda kv: -kv[1])) if c else 'Sin datos'
