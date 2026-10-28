"""Lista de proyectos que Antony señaló como utilizables (Word o Excel con columnas CUI, NOMBRE, RUBRO, ANULACIÓN).

Se importa una vez (actualizar.exe --lista-antony "<archivo>") y queda en lista_antony.json; el Excel de proyectos la resalta en cada corrida.
"""
import datetime, json, pathlib, re, zipfile

ARCHIVO = 'lista_antony.json'
RUBRO_COD = [('FONCOR', '15'), ('CANON', '18'), ('DIRECTAMENTE', '09'), ('RDR', '09'), ('ORDINARIO', '00'), ('OPERACIONES', '19'), ('ROOC', '19')]


def _monto(t):
    m = re.sub(r'[^\d.,]', '', t or '')
    if not m:
        return None
    m = m.replace(',', '')
    try:
        return float(m)
    except ValueError:
        return None


def codigo_rubro(texto):
    t = (texto or '').upper()
    for k, c in RUBRO_COD:
        if k in t:
            return c
    return ''


def _de_docx(p):
    xml = zipfile.ZipFile(p).read('word/document.xml').decode('utf-8')
    items = []
    for tr in re.findall(r'<w:tr[ >].*?</w:tr>', xml, flags=re.S):
        celdas = []
        for tc in re.findall(r'<w:tc>.*?</w:tc>', tr, flags=re.S):
            t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tc)).strip()
            celdas.append(t)
        if celdas and re.fullmatch(r'\d{7}', celdas[0] or ''):
            items.append(dict(cui=celdas[0], nombre=celdas[1] if len(celdas) > 1 else '', rubro=celdas[2] if len(celdas) > 2 else '',
                              anulacion=_monto(celdas[3]) if len(celdas) > 3 else None))
    return items


def _de_xlsx(p):
    import openpyxl
    ws = openpyxl.load_workbook(p, data_only=True).active
    items = []
    for row in ws.iter_rows(values_only=True):
        if row and row[0] is not None and re.fullmatch(r'\d{7}', str(row[0]).strip()):
            g = lambda i: row[i] if len(row) > i else None
            items.append(dict(cui=str(row[0]).strip(), nombre=str(g(1) or ''), rubro=str(g(2) or ''), anulacion=g(3) if isinstance(g(3), (int, float)) else _monto(str(g(3) or ''))))
    return items


def importar(raiz, archivo):
    p = pathlib.Path(archivo)
    items = _de_docx(p) if p.suffix.lower() == '.docx' else _de_xlsx(p)
    if not items:
        raise SystemExit('No se encontraron filas con CUI de 7 dígitos en ' + p.name + ' (la primera columna debe ser el CUI).')
    for it in items:
        it['rubro_cod'] = codigo_rubro(it['rubro'])
    d = dict(fuente=p.name, importado=datetime.datetime.now().isoformat(timespec='minutes'), items=items)
    (pathlib.Path(raiz) / ARCHIVO).write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding='utf-8')
    return d


def cargar(raiz):
    p = pathlib.Path(raiz) / ARCHIVO
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return None
