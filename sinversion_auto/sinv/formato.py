"""Formato de la oficina: orden de hojas, hojas ocultas, anchos de columna y zoom que el usuario dejó en un Excel retocado a mano.

Se guarda una sola vez (actualizar.exe --guardar-formato "<Excel retocado>") en formato_oficina.json y se aplica en todas las corridas siguientes.
Los anchos de las hojas con tabla (Dónde hay plata, Base, Cambios, Rubro por CUI) se guardan por TÍTULO de columna, así sobreviven aunque se agreguen columnas nuevas.
"""
import json, pathlib, shutil, tempfile

ARCHIVO = 'formato_oficina.json'
PADDING = 0.71                      # xlsxwriter guarda ancho + 0,71; al leer un Excel hay que restarlo para reproducir el mismo ancho
FILA_TITULOS = {'Dónde hay plata': 7, 'Base': 4, 'Cambios': 4, 'Rubro por CUI': 4, 'Lista Antony': 8}


def extraer(xlsx):
    import openpyxl
    tmp = pathlib.Path(tempfile.mkdtemp()) / 'copia.xlsx'
    shutil.copy(str(xlsx), str(tmp))                     # el original puede estar abierto en Excel
    wb = openpyxl.load_workbook(str(tmp))
    fmt = {'orden': [ws.title for ws in wb.worksheets], 'ocultas': [ws.title for ws in wb.worksheets if ws.sheet_state != 'visible'], 'hojas': {}}
    for ws in wb.worksheets:
        ancho = {}; oculto = []
        for cd in ws.column_dimensions.values():
            if cd.min is None or cd.max is None:
                continue
            for i in range(cd.min, min(cd.max, cd.min + 200) + 1):
                if cd.width:
                    ancho[i - 1] = round(max(cd.width - PADDING, 1.0), 2)
                if cd.hidden:
                    oculto.append(i - 1)
        h = {'zoom': ws.sheet_view.zoomScale, 'anchos': {str(k): v for k, v in sorted(ancho.items())}, 'columnas_ocultas': sorted(oculto)}
        fila = FILA_TITULOS.get(ws.title)
        if fila:
            tit = [c.value for c in ws[fila]]
            h['por_titulo'] = {str(t): ancho[i] for i, t in enumerate(tit) if t is not None and i in ancho}
            h['titulos_ocultos'] = [str(tit[i]) for i in oculto if i < len(tit) and tit[i] is not None]
        fmt['hojas'][ws.title] = h
    return fmt


def guardar(raiz, xlsx):
    fmt = extraer(xlsx)
    (pathlib.Path(raiz) / ARCHIVO).write_text(json.dumps(fmt, ensure_ascii=False, indent=1), encoding='utf-8')
    return fmt


def cargar(raiz):
    p = pathlib.Path(raiz) / ARCHIVO
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return None


def orden_hojas(nombres, fmt):
    """Ordena las hojas que existen en esta corrida según el orden guardado. Las hojas nuevas (sin lugar guardado) van después de 'Dónde hay plata'."""
    if not fmt:
        return list(nombres)
    guardado = [n for n in fmt['orden'] if n in nombres]
    nuevas = [n for n in nombres if n not in guardado]
    out = list(guardado)
    pos = out.index('Dónde hay plata') + 1 if 'Dónde hay plata' in out else len(out)
    for n in nuevas:
        out.insert(pos, n); pos += 1
    return out


def aplicar(hojas, fmt, titulos):
    """hojas: {nombre: worksheet}; titulos: {nombre: [títulos de columna]} de las hojas con tabla. Se llama justo antes de cerrar el libro."""
    if not fmt:
        return
    for nm, ws in hojas.items():
        h = fmt['hojas'].get(nm)
        if not h:
            continue
        tit = titulos.get(nm)
        if tit and h.get('por_titulo') is not None:
            for i, t in enumerate(tit):
                w = h['por_titulo'].get(str(t))
                if w:
                    ws.set_column(i, i, w, None, {'hidden': True} if str(t) in h.get('titulos_ocultos', []) else {})
        else:
            for i, w in h.get('anchos', {}).items():
                ws.set_column(int(i), int(i), w, None, {'hidden': True} if int(i) in h.get('columnas_ocultas', []) else {})
        if h.get('zoom'):
            ws.set_zoom(int(h['zoom']))
        if nm in fmt.get('ocultas', []):
            ws.hide()
