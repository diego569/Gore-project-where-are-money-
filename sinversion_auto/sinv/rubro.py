"""Lector de los reportes de la Consulta Amigable (MEF) por Rubro de financiamiento.

Cada archivo (.xls que en realidad es HTML) corresponde a UN rubro y trae los proyectos de todo el pliego.
La Consulta Amigable parte los resultados en páginas de 400 filas: si el rubro tiene más, hay un segundo archivo.
Se detecta si falta alguna página comparando la suma de filas con el total del rubro que trae el propio reporte.
"""
import re
from html.parser import HTMLParser

RUBRO_RE = re.compile(r'^Rubro (\d{2}):\s*(.+)$')
CUI_RE = re.compile(r'^(\d{7}):\s*(.*)$', re.S)
PAG_RE = re.compile(r'Resultados\s+(\d+)\s*-\s*(\d+)\s+de\s+(\d+)')
FECHA_RE = re.compile(r'Fecha de la Consulta:\s*(\d{1,2})-(\w+)-(\d{4})')
MESES = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8, 'setiembre': 9, 'septiembre': 9,
         'octubre': 10, 'noviembre': 11, 'diciembre': 12}


class _Tablas(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tablas = []; self._fila = None; self._celda = None; self._prof = 0

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.tablas.append([])
        elif tag == 'tr' and self.tablas:
            self._fila = []
        elif tag in ('td', 'th') and self._fila is not None:
            self._celda = []

    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self._celda is not None and self._fila is not None:
            self._fila.append(re.sub(r'\s+', ' ', ''.join(self._celda).replace('\xa0', ' ')).strip()); self._celda = None
        elif tag == 'tr' and self._fila is not None and self.tablas:
            self.tablas[-1].append(self._fila); self._fila = None

    def handle_data(self, data):
        if self._celda is not None:
            self._celda.append(data)


def _leer_texto(path):
    raw = open(path, 'rb').read()
    for enc in ('utf-8', 'cp1252'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('latin1')


def _num(s):
    s = (s or '').replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def es_rubro(path):
    try:
        with open(path, 'rb') as f:
            ini = f.read(4000).decode('latin1')
        return '<' in ini[:200] and 'Consulta Amigable' in _leer_texto(path)[:6000] and bool(re.search(r'Rubro \d{2}:', _leer_texto(path)))
    except Exception:
        return False


def leer_rubro(path):
    txt = _leer_texto(path)
    p = _Tablas(); p.feed(txt)
    rub = None; total = None; proy = {}
    for tabla in p.tablas:
        for fila in tabla:
            if not fila:
                continue
            m = RUBRO_RE.match(fila[0])
            if m and len(fila) >= 4:
                rub = (m.group(1), m.group(2).strip())
                total = dict(pia=_num(fila[1]), pim=_num(fila[2]), cert=_num(fila[3]), comp=_num(fila[4]) if len(fila) > 4 else 0.0,
                             dev=_num(fila[6]) if len(fila) > 6 else 0.0)
                continue
            c = CUI_RE.match(fila[0])
            if c and len(fila) >= 8:
                proy[c.group(1)] = dict(nombre=c.group(2).strip(), pia=_num(fila[1]), pim=_num(fila[2]), cert=_num(fila[3]),
                                        comp=_num(fila[4]), dev=_num(fila[6]))
    if rub is None:
        raise ValueError('El archivo no indica el rubro de financiamiento (no se encontró la línea "Rubro NN: ...").')
    fecha = None
    f = FECHA_RE.search(txt)
    if f:
        import datetime
        mes = MESES.get(f.group(2).lower())
        if mes:
            fecha = datetime.date(int(f.group(3)), mes, int(f.group(1)))
    pg = PAG_RE.search(txt)
    suma = sum(v['pim'] for v in proy.values())
    completo = abs(suma - total['pim']) <= 1.0
    return dict(rubro=rub[0], nombre=rub[1], fecha=fecha, proyectos=proy, total=total, completo=completo, suma_pim=suma,
                pagina=(int(pg.group(2)), int(pg.group(3))) if pg else None, archivo=str(path))


def combinar(lecturas):
    """Une las páginas de un mismo rubro. Devuelve {codigo: lectura combinada}."""
    por = {}
    for l in lecturas:
        r = por.get(l['rubro'])
        if r is None:
            por[l['rubro']] = dict(l, archivos=[l['archivo']])
            continue
        r['proyectos'].update(l['proyectos']); r['archivos'].append(l['archivo'])
        r['suma_pim'] = sum(v['pim'] for v in r['proyectos'].values())
        r['completo'] = abs(r['suma_pim'] - r['total']['pim']) <= 1.0
        if l['fecha'] and (not r['fecha'] or l['fecha'] > r['fecha']):
            r['fecha'] = l['fecha']
    return por
