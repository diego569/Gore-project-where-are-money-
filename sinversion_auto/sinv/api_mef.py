"""Cliente de la API de Datos Abiertos del MEF (CKAN, recurso de ejecución del gasto diario/mensual).

Notas verificadas el 07/10/2026:
  - Solo existe datastore_search (con "filters" exactos); datastore_search_sql responde 404.
  - Los registros son mensuales: PIA y PIM en MES_EJE = 0; certificado/compromiso/devengado/girado son flujos por mes (se suman).
  - SEC_FUNC viene sin ceros a la izquierda ('480'); el SIAF usa 4 dígitos ('0480').
  - El certificado SSL del MEF estuvo vencido; ver parámetro permitir_sin_verificar.
"""
import concurrent.futures as cf, collections, csv, datetime, gzip, io, json, pathlib, time
import requests, urllib3

URL = 'https://api.datosabiertos.mef.gob.pe/DatosAbiertos/v1/datastore_search'
DUMP_URL = 'https://api.datosabiertos.mef.gob.pe/DatosAbiertos/v1/dump'
RID = '615644aa-ef73-4358-b4e0-0c20931632f3'
PAG = 5000
PROYECTOS = {'ANO_EJE': '2026', 'PLIEGO': '458', 'TIPO_ACT_PROY': '2'}       # 2 = proyectos; 3 = actividades (gasto corriente)


class ErrorAPI(Exception):
    pass


class Cliente:
    def __init__(self, permitir_sin_verificar=True, timeout=(15, 120), reintentos=1):
        self.verificar = True; self.permitir = permitir_sin_verificar; self.sin_verificar_usado = False
        self.timeout = timeout; self.reintentos = reintentos
        self.metodo = ''; self.notas = []

    def get(self, params):
        ultimo = None
        for _ in range(self.reintentos + 1):
            try:
                r = requests.get(URL, params=params, timeout=self.timeout, verify=self.verificar)
                if r.status_code != 200:
                    raise ErrorAPI(f'HTTP {r.status_code}')
                return r.json()
            except requests.exceptions.SSLError as e:
                if self.verificar and self.permitir:
                    urllib3.disable_warnings(); self.verificar = False; self.sin_verificar_usado = True
                    continue
                raise ErrorAPI('Certificado SSL del MEF inválido o vencido (config api_sin_verificar_certificado = false).') from e
            except (requests.exceptions.RequestException, ValueError, ErrorAPI) as e:
                ultimo = e; time.sleep(3)
        raise ErrorAPI(f'No se pudo consultar la API del MEF: {str(ultimo)[:150]}')

    def descargar_dump(self, filtro, espera=10, max_espera=360):
        """Descarga masiva: el servidor arma un CSV filtrado (1.ª llamada: "Creando consulta"; luego devuelve el CSV)."""
        p = {'id': RID, 'filters': json.dumps(filtro)}
        t0 = time.time(); ultimo = ''
        while time.time() - t0 < max_espera:
            try:
                r = requests.get(DUMP_URL, params=p, timeout=(15, 300), verify=self.verificar)
            except requests.exceptions.SSLError as e:
                if self.verificar and self.permitir:
                    urllib3.disable_warnings(); self.verificar = False; self.sin_verificar_usado = True
                    continue
                raise ErrorAPI('Certificado SSL del MEF inválido o vencido.') from e
            except requests.exceptions.RequestException as e:
                raise ErrorAPI(f'dump: {str(e)[:120]}')
            if r.status_code != 200:
                raise ErrorAPI(f'dump: HTTP {r.status_code}')
            if 'csv' in r.headers.get('content-type', '').lower():
                texto = r.content.decode('utf-8-sig', errors='replace')
                filas = list(csv.DictReader(io.StringIO(texto)))
                if not filas:
                    raise ErrorAPI('dump: el CSV llegó solo con encabezados (la tabla del MEF está vacía).')
                return filas
            ultimo = r.text[:80].replace('\n', ' ')
            time.sleep(espera)
        raise ErrorAPI(f'dump: no estuvo listo en {max_espera} s ({ultimo}).')

    def descargar(self, filtro=PROYECTOS, hilos=6, usar_dump=True):
        """Primero la descarga masiva (dump); si falla, la API paginada (datastore_search)."""
        if usar_dump:
            try:
                filas = self.descargar_dump(filtro)
                self.metodo = 'dump (descarga masiva)'
                return filas
            except ErrorAPI as e:
                self.notas.append(str(e))
        try:
            filas = self.descargar_paginado(filtro, hilos)
        except ErrorAPI as e:
            raise ErrorAPI('; '.join(self.notas + [f'paginada: {e}']))
        self.metodo = 'API paginada' + (' (el dump falló)' if usar_dump else '')
        return filas

    def descargar_paginado(self, filtro=PROYECTOS, hilos=6):
        p = dict(resource_id=RID, filters=json.dumps(filtro))
        d = self.get({**p, 'limit': 1})
        total = int(d['result']['include_total'])
        if total == 0:
            raise ErrorAPI('La API no devolvió registros para el filtro pedido.')
        offs = list(range(0, total, PAG))
        with cf.ThreadPoolExecutor(hilos) as ex:
            pags = list(ex.map(lambda o: self.get({**p, 'limit': PAG, 'offset': o})['records'], offs))
        filas = [r for pg in pags for r in pg]
        if len(filas) != total:
            raise ErrorAPI(f'Descarga incompleta: {len(filas)} de {total} registros.')
        return filas


def _f(x):
    return float(x or 0)


def pim_por_ue_cui_rubro(filas):
    """{(ue, cui): {rubro: pim}} y nombres de rubro."""
    out = collections.defaultdict(lambda: collections.defaultdict(float)); nombres = {}
    for r in filas:
        out[(r['EJECUTORA'], r['PRODUCTO_PROYECTO'])][r['RUBRO']] += _f(r['MONTO_PIM'])
        nombres[r['RUBRO']] = r['RUBRO_NOMBRE']
    return {k: dict(v) for k, v in out.items()}, nombres


def agregar_metas(filas):
    """Agrega a nivel (UE, código, meta SEC_FUNC de 4 dígitos). Sirve para comparar/reemplazar el reporte SIAF."""
    g = {}
    for r in filas:
        k = (r['EJECUTORA'], r['PRODUCTO_PROYECTO'], r['SEC_FUNC'].zfill(4))
        d = g.setdefault(k, dict(ue=k[0], cui=k[1], meta=k[2], fn=r['FUNCION'], fn_nombre=r['FUNCION_NOMBRE'], nombre=r['PRODUCTO_PROYECTO_NOMBRE'],
                                 pia=0.0, pim=0.0, cert=0.0, comp=0.0, comp_mes=0.0, dev=0.0, giro=0.0, rubros=collections.defaultdict(float),
                                 espec=collections.defaultdict(float)))
        d['pia'] += _f(r['MONTO_PIA']); d['pim'] += _f(r['MONTO_PIM']); d['cert'] += _f(r['MONTO_CERTIFICADO'])
        d['comp'] += _f(r['MONTO_COMPROMETIDO_ANUAL']); d['comp_mes'] += _f(r['MONTO_COMPROMETIDO']); d['dev'] += _f(r['MONTO_DEVENGADO'])
        d['giro'] += _f(r['MONTO_GIRADO']); d['rubros'][r['RUBRO']] += _f(r['MONTO_PIM'])
        d['espec'][(r['GENERICA'], r['SUBGENERICA'], r['SUBGENERICA_DET'], r['ESPECIFICA'], r['ESPECIFICA_DET'], r['ESPECIFICA_DET_NOMBRE'])] += max(_f(r['MONTO_PIM']), 0)
    return g


def rubros_listos(cliente=None):
    """Descarga los proyectos del pliego y devuelve lo necesario para la hoja de rubros."""
    c = cliente or Cliente()
    filas = c.descargar(PROYECTOS)
    pim, nombres = pim_por_ue_cui_rubro(filas)
    return dict(pim=pim, nombres=nombres, registros=len(filas), fecha=datetime.datetime.now(), sin_verificar=c.sin_verificar_usado,
                meses=sorted({int(r['MES_EJE']) for r in filas}))


def pim_por_meta_rubro(filas):
    """{(codigo_producto, meta4): {rubro: pim}} y {codigo_funcion: nombre}. Sirve para gasto corriente (rubro por meta)."""
    out = collections.defaultdict(lambda: collections.defaultdict(float)); fn = {}
    for r in filas:
        out[(r['PRODUCTO_PROYECTO'], r['SEC_FUNC'].zfill(4))][r['RUBRO']] += _f(r['MONTO_PIM'])
        fn[r['FUNCION']] = r['FUNCION_NOMBRE']
    return {k: dict(v) for k, v in out.items()}, fn


# ---------------------------------------------------------------- última descarga buena (respaldo cuando la API cae)
CAMPOS_COPIA = ['ANO_EJE', 'MES_EJE', 'EJECUTORA', 'EJECUTORA_NOMBRE', 'SEC_FUNC', 'TIPO_ACT_PROY', 'PRODUCTO_PROYECTO', 'PRODUCTO_PROYECTO_NOMBRE',
                'ACTIVIDAD_ACCION_OBRA', 'ACTIVIDAD_ACCION_OBRA_NOMBRE', 'META_NOMBRE', 'FINALIDAD', 'FUNCION', 'FUNCION_NOMBRE', 'GENERICA', 'SUBGENERICA',
                'SUBGENERICA_DET', 'ESPECIFICA', 'ESPECIFICA_DET', 'ESPECIFICA_DET_NOMBRE', 'RUBRO', 'RUBRO_NOMBRE', 'MONTO_PIA', 'MONTO_PIM', 'MONTO_CERTIFICADO',
                'MONTO_COMPROMETIDO', 'MONTO_COMPROMETIDO_ANUAL', 'MONTO_DEVENGADO', 'MONTO_GIRADO']


def guardar_filas(raiz, proy=None, gc=None, fecha=None):
    """Guarda (comprimido) la última descarga buena de proyectos y/o gasto corriente. Solo se reemplaza lo que llegó."""
    p = pathlib.Path(raiz) / 'cache_api_filas.json.gz'
    actual = {}
    if p.exists():
        try:
            actual = json.loads(gzip.open(p, 'rt', encoding='utf-8').read())
        except Exception:
            actual = {}
    ahora = (fecha or datetime.datetime.now()).isoformat()
    for clave, filas in (('proy', proy), ('gc', gc)):
        if filas:
            actual[clave] = dict(fecha=ahora, filas=[{k: r.get(k) for k in CAMPOS_COPIA} for r in filas])
    with gzip.open(p, 'wt', encoding='utf-8') as f:
        f.write(json.dumps(actual, ensure_ascii=False))


def leer_filas(raiz):
    """Devuelve dict clave -> (filas, fecha) para 'proy' y 'gc' (las que existan)."""
    p = pathlib.Path(raiz) / 'cache_api_filas.json.gz'
    if not p.exists():
        return {}
    try:
        d = json.loads(gzip.open(p, 'rt', encoding='utf-8').read())
        return {k: (v['filas'], datetime.datetime.fromisoformat(v['fecha'])) for k, v in d.items()}
    except Exception:
        return {}
