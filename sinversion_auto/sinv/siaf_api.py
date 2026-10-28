"""Reemplazo del reporte SIAF (.xls) por la API de Datos Abiertos del MEF.

Devuelve la misma estructura que siaf.leer_siaf (metas por UE + código + meta), de modo que el resto de la herramienta no cambia.
Alcance: proyectos de todo el pliego y gasto corriente de las UE pedidas (por defecto solo la 001: bajar todo el gasto corriente del
pliego son ~130 mil registros). Comparación con el XLS el 07/10/2026: PIM exacto; certificación/compromiso/devengado difieren
en el orden de 0.1% (desfase horario).
"""
import collections
from . import api_mef
from .siaf import MOD_RULES


def _etiqueta(g, sg, sgd, e, nombre):
    for k, v in MOD_RULES:
        if k in (nombre or ''):
            return v
    if g == '6' and sg == '3':
        return 'Equipamiento'
    if (g, sg, sgd, e) in (('6', '8', '1', '4'), ('6', '7', '1', '6')):
        return 'Personal/bienes/servicios'
    return 'Otro (' + (nombre or '')[:30] + ')'


def metas_desde_filas(filas, tipo):
    g = {}
    for r in filas:
        k = (r['EJECUTORA'], r['PRODUCTO_PROYECTO'], r['SEC_FUNC'].zfill(4))
        d = g.get(k)
        if d is None:
            d = g[k] = dict(ue=k[0], cui=k[1], act=r['ACTIVIDAD_ACCION_OBRA'], meta=k[2], fn=r['FUNCION'], fn_nombre=r['FUNCION_NOMBRE'],
                            nombre=r['ACTIVIDAD_ACCION_OBRA_NOMBRE'] or r['PRODUCTO_PROYECTO_NOMBRE'], prod_nombre=r['PRODUCTO_PROYECTO_NOMBRE'], tipo=tipo,
                            pia=0.0, pim=0.0, cert=0.0, comp=0.0, dev=0.0, ue_nombre=r['EJECUTORA_NOMBRE'],
                            rubros=collections.defaultdict(lambda: dict(pim=0.0, nombre='')), mods=collections.Counter(), mods_any=collections.Counter(),
                            gen=collections.defaultdict(float), meta_desc=r.get('META_NOMBRE') or '', finalidad=r.get('FINALIDAD') or '')
        f = lambda x: float(x or 0)
        pia, pim, cert, comp, dev = f(r['MONTO_PIA']), f(r['MONTO_PIM']), f(r['MONTO_CERTIFICADO']), f(r['MONTO_COMPROMETIDO_ANUAL']), f(r['MONTO_DEVENGADO'])
        d['pia'] += pia; d['pim'] += pim; d['cert'] += cert; d['comp'] += comp; d['dev'] += dev
        rb = d['rubros'][r['RUBRO']]; rb['pim'] += pim; rb['nombre'] = r['RUBRO_NOMBRE']
        lab = _etiqueta(r['GENERICA'], r['SUBGENERICA'], r['SUBGENERICA_DET'], r['ESPECIFICA'], r['ESPECIFICA_DET_NOMBRE'])
        d['mods'][lab] += max(pim, 0); d['mods_any'][lab] += abs(pia) + abs(pim) + abs(cert) + abs(comp) + abs(dev)
        d['gen']['2.' + r['GENERICA'] + '.'] += pim
    for d in g.values():
        d['rubros'] = {k: dict(v) for k, v in d['rubros'].items()}; d['gen'] = dict(d['gen'])
    return list(g.values())


def siaf_desde_filas(f_proy, f_gc, sin_verificar=False, fuente='api'):
    metas = metas_desde_filas(f_proy, 'proyecto') + metas_desde_filas(f_gc or [], 'actividad')
    ues = {m['ue']: m['ue_nombre'] for m in metas}
    return dict(metas=metas, ues=ues, ue_tot={}, pliego=None, con_rubro=False, fuente=fuente, filas_proy=f_proy, filas_gc=f_gc or [],
                sin_verificar=sin_verificar, meses=sorted({int(r['MES_EJE']) for r in f_proy}))


def leer_siaf_api(cliente=None, ue_gasto_corriente=('001',), raiz=None):
    """Descarga proyectos (pliego) y gasto corriente (UE indicadas) y arma la estructura tipo SIAF. Si se da 'raiz', guarda la copia buena."""
    c = cliente or api_mef.Cliente()
    f_proy = c.descargar(api_mef.PROYECTOS)
    f_gc = []
    for ue in ue_gasto_corriente:
        f_gc += c.descargar({'ANO_EJE': '2026', 'PLIEGO': '458', 'TIPO_ACT_PROY': '3', 'EJECUTORA': ue})
    if raiz:
        try:
            api_mef.guardar_filas(raiz, f_proy, f_gc)
        except OSError:
            pass
    s = siaf_desde_filas(f_proy, f_gc, c.sin_verificar_usado)
    s['metodo'] = c.metodo; s['notas'] = c.notas
    return s
