"""Historial de corridas: guarda una foto de la Base en cada corrida y compara con la anterior ("qué cambió").

Las fotos viven en historial\\base_AAAAMMDD_HHMM.csv.gz (se conservan las últimas 60). Se comparan por (UE, CUI, meta).
"""
import csv, datetime, gzip, io, pathlib

CAMPOS = ['ue', 'cui', 'meta', 'nombre', 'sem', 'pim', 'cert', 'disponible', 'pot']
MIN_CAMBIO = 1000.0      # variaciones de disponible menores a este monto no se listan (ruido)


def _clave(r):
    return (str(r['ue']), str(r['cui']), str(r['meta']))


def guardar(raiz, rows, cuando=None, conservar=60):
    d = pathlib.Path(raiz) / 'historial'; d.mkdir(exist_ok=True)
    cuando = cuando or datetime.datetime.now()
    p = d / f'base_{cuando:%Y%m%d_%H%M}.csv.gz'
    buf = io.StringIO(); w = csv.writer(buf); w.writerow(CAMPOS)
    for r in rows:
        w.writerow([r['ue'], r['cui'], r['meta'], r['nombre'][:200], r['sem'], round(r['pim'], 2), round(r['cert'], 2), round(r['disponible'], 2), round(r['pot'], 2)])
    with gzip.open(p, 'wt', encoding='utf-8') as f:
        f.write(buf.getvalue())
    viejos = sorted(d.glob('base_*.csv.gz'))[:-conservar]
    for v in viejos:
        try:
            v.unlink()
        except OSError:
            pass
    return p


def cargar_anterior(raiz, antes_de=None):
    """Foto más reciente anterior a 'antes_de' (por defecto, la última guardada). Devuelve (dict, datetime) o (None, None)."""
    d = pathlib.Path(raiz) / 'historial'
    if not d.exists():
        return None, None
    antes_de = antes_de or datetime.datetime.now()
    for p in sorted(d.glob('base_*.csv.gz'), reverse=True):
        try:
            t = datetime.datetime.strptime(p.stem.replace('.csv', '')[5:], '%Y%m%d_%H%M')
        except ValueError:
            continue
        if t >= antes_de.replace(second=0, microsecond=0):
            continue
        try:
            with gzip.open(p, 'rt', encoding='utf-8') as f:
                filas = list(csv.DictReader(f))
        except Exception:
            continue
        return {(x['ue'], x['cui'], x['meta']): x for x in filas}, t
    return None, None


def comparar(rows, prev):
    """Lista de cambios entre la foto anterior y las filas actuales."""
    if prev is None:
        return None
    ahora = {_clave(r): r for r in rows}; out = []
    for k, r in ahora.items():
        a = prev.get(k)
        if a is None:
            out.append(dict(tipo='Nuevo', k=k, nombre=r['nombre'], sem0='', sem1=r['sem'], d0=0.0, d1=r['disponible'], p0=0.0, p1=r['pim'], pot1=r['pot']))
            continue
        d0 = float(a['disponible']); p0 = float(a['pim'])
        tipos = []
        if a['sem'] != r['sem']:
            tipos.append('Cambió el semáforo')
        if abs(r['disponible'] - d0) >= MIN_CAMBIO:
            tipos.append('Cambió el disponible')
        if abs(r['pim'] - p0) >= MIN_CAMBIO:
            tipos.append('Cambió el PIM')
        if tipos:
            out.append(dict(tipo=' + '.join(tipos), k=k, nombre=r['nombre'], sem0=a['sem'], sem1=r['sem'], d0=d0, d1=r['disponible'], p0=p0, p1=r['pim'], pot1=r['pot']))
    for k, a in prev.items():
        if k not in ahora:
            out.append(dict(tipo='Ya no aparece', k=k, nombre=a['nombre'], sem0=a['sem'], sem1='', d0=float(a['disponible']), d1=0.0, p0=float(a['pim']), p1=0.0, pot1=0.0))
    out.sort(key=lambda c: -abs(c['d1'] - c['d0']))
    return out
