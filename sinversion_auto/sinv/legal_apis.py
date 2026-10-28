"""Datos por CUI que la API de ejecución del MEF no trae: ficha del Banco de Inversiones (SSI) y Formato 12-B (Datos Abiertos).

SSI  : POST https://ofi5.mef.gob.pe/invierteWS/Ssi/traeDetInvSSI  (id=CUI, tipo=1)  -> tipo de inversión, estado, cierre, costo, avance acumulado.
12-B : datastore_search del dataset de seguimiento (programación mensual actualizada + devengado acumulado), un CUI por consulta.
Ambos se guardan en un caché local por CUI (la ficha cambia poco: 7 días; el 12-B: ~1 día) para no insistir al MEF.
Si una API cae se usa lo guardado (cualquier antigüedad) y se avisa; si no hay nada, el CUI queda "sin dato" (nunca se inventa).
"""
import concurrent.futures as cf, datetime, gzip, json, pathlib, time
import requests, urllib3

SSI_URL = 'https://ofi5.mef.gob.pe/invierteWS/Ssi/traeDetInvSSI'
F12_URL = 'https://api.datosabiertos.mef.gob.pe/DatosAbiertos/v1/datastore_search'
F12_RID = 'c275fa9f-5c61-4313-828d-0827277bdd97'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) SINVERSION/2.0'
SSI_CAMPOS = ['TIPO_FORMATO', 'ESTADO', 'SITUACION', 'MARCO', 'CIERRE_REGISTRADO', 'COSTO_ACTUALIZADO', 'DEV_ACUMULADO', 'AVAN_FISICO', 'MODAL_EJEC',
              'IND_REG_PMI', 'ENTIDAD', 'NIVEL', 'MTO_PMI_1', 'MTO_PMI_2', 'MTO_PMI_3', 'MTO_PMI_4', 'TIENE_F12B', 'FEC_FIN_EJ', 'ET_REGISTRADO', 'IND_PARALIZADA']
F12_CAMPOS = ['DEVENGADO_ACUMULADO', 'FECHA_ULT_ACT_F12B', 'ULT_PERIODO_REG_F12B', 'AVANCE_EJECUCION', 'AVANCE_FISICO', 'ESTADO'] + [f'MONTO_ACTUALIZADO_{i}' for i in range(1, 13)]


class Falla(Exception):
    pass


def _leer(p):
    try:
        return json.loads(gzip.open(p, 'rt', encoding='utf-8').read())
    except Exception:
        return {}


def _guardar(p, d):
    with gzip.open(p, 'wt', encoding='utf-8') as f:
        f.write(json.dumps(d, ensure_ascii=False))


def _num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


class _Base:
    nombre = ''; archivo = ''; ttl = datetime.timedelta(days=7)

    def __init__(self, raiz, verificar=False, hilos=4):
        self.p = pathlib.Path(raiz) / self.archivo; self.verificar = verificar; self.hilos = hilos
        self.cache = _leer(self.p); self.notas = []; self.fallidos = 0; self.nuevos = 0
        self.s = requests.Session(); self.s.headers.update({'User-Agent': UA})
        if not verificar:
            urllib3.disable_warnings()

    def _uno(self, cui):
        raise NotImplementedError

    def obtener(self, cuis, limite_seg=900):
        """Devuelve {cui: datos|None}. Consulta solo los CUI que faltan o están vencidos. Se rinde rápido si la API está caída."""
        ahora = datetime.datetime.now(); pend = []
        for c in cuis:
            e = self.cache.get(c)
            if not e or ahora - datetime.datetime.fromisoformat(e['t']) > self.ttl:
                pend.append(c)
        if pend and hasattr(self, '_lote'):
            try:
                got = self._lote(pend)
                for c in pend:
                    self.cache[c] = {'t': datetime.datetime.now().isoformat(), 'd': got.get(c, {})}
                self.nuevos += len(pend); pend = []
                _guardar(self.p, self.cache)
            except Exception as e:
                self.notas.append(f'{self.nombre}: la descarga masiva falló ({str(e)[:100]}); se consulta de a un CUI (lento).')
        if pend:
            t0 = time.time(); seguidos = 0; hechos = 0
            with cf.ThreadPoolExecutor(self.hilos) as ex:
                futs = {ex.submit(self._uno, c): c for c in pend}
                for f in cf.as_completed(futs):
                    c = futs[f]
                    try:
                        d = f.result(); seguidos = 0; hechos += 1
                        self.cache[c] = {'t': datetime.datetime.now().isoformat(), 'd': d}; self.nuevos += 1
                    except Exception:
                        self.fallidos += 1; seguidos += 1
                    if (seguidos >= 12 and hechos == 0) or time.time() - t0 > limite_seg:
                        for g in futs:
                            g.cancel()
                        self.notas.append(f'{self.nombre}: la API no respondió (se interrumpió tras {self.fallidos} fallas); se usa lo guardado.')
                        break
            if self.nuevos:
                try:
                    _guardar(self.p, self.cache)
                except OSError:
                    pass
        if self.fallidos and not self.notas:
            self.notas.append(f'{self.nombre}: {self.fallidos} CUI no respondieron; se usa lo guardado de ellos (o quedan sin dato).')
        return {c: (self.cache.get(c) or {}).get('d') for c in cuis}

    def antiguedad_max_dias(self, cuis):
        ts = [datetime.datetime.fromisoformat(self.cache[c]['t']) for c in cuis if c in self.cache]
        return (datetime.datetime.now() - min(ts)).days if ts else None


class SSI(_Base):
    nombre = 'Ficha SSI (Banco de Inversiones)'; archivo = 'cache_ssi.json.gz'; ttl = datetime.timedelta(days=7)

    def _uno(self, cui):
        h = {'X-Requested-With': 'XMLHttpRequest', 'Referer': 'https://ofi5.mef.gob.pe/ssi/'}
        r = self.s.post(SSI_URL, data={'id': cui, 'tipo': 1}, headers=h, timeout=(15, 40), verify=self.verificar)
        if r.status_code != 200:
            raise Falla(f'HTTP {r.status_code}')
        d = r.json()
        if not d:
            return {}          # el CUI no tiene ficha (respuesta válida y vacía)
        d = d[0]
        out = {k: d.get(k) for k in SSI_CAMPOS}
        for k in ('COSTO_ACTUALIZADO', 'DEV_ACUMULADO', 'MTO_PMI_1', 'MTO_PMI_2', 'MTO_PMI_3', 'MTO_PMI_4'):
            out[k] = _num(out[k])
        return out


class F12B(_Base):
    nombre = 'Formato 12-B (Datos Abiertos MEF)'; archivo = 'cache_f12b.json.gz'; ttl = datetime.timedelta(hours=20)

    def _lote(self, cuis, espera=8, max_espera=300):
        """Una sola descarga (dump con lista de CUI) en vez de una consulta por CUI."""
        import csv, io, json as _j
        t0 = time.time(); ultimo = ''
        while time.time() - t0 < max_espera:
            r = self.s.get(F12_URL.replace('datastore_search', 'dump'), params={'id': F12_RID, 'filters': _j.dumps({'CODIGO_UNICO': list(cuis)})},
                           timeout=(15, 300), verify=self.verificar)
            if r.status_code != 200:
                raise Falla(f'HTTP {r.status_code}')
            if 'csv' in r.headers.get('content-type', '').lower():
                filas = list(csv.DictReader(io.StringIO(r.content.decode('utf-8-sig', errors='replace'))))
                if not filas:
                    raise Falla('el CSV llegó vacío (la tabla del 12-B no está disponible)')
                out = {}
                for d in filas:
                    o = {k: d.get(k) for k in F12_CAMPOS}
                    o['PROG'] = [_num(d.get(f'MONTO_ACTUALIZADO_{i}')) for i in range(1, 13)]
                    o['DEVENGADO_ACUMULADO'] = _num(o['DEVENGADO_ACUMULADO'])
                    for k in [f'MONTO_ACTUALIZADO_{i}' for i in range(1, 13)]:
                        o.pop(k, None)
                    out[str(d['CODIGO_UNICO'])] = o
                return out
            ultimo = r.text[:60].replace(chr(10), ' ')
            time.sleep(espera)
        raise Falla(f'el dump no estuvo listo en {max_espera} s ({ultimo})')

    def _uno(self, cui):
        import json as _j
        r = self.s.get(F12_URL, params={'resource_id': F12_RID, 'filters': _j.dumps({'CODIGO_UNICO': cui}), 'limit': 5}, timeout=(15, 60), verify=self.verificar)
        if r.status_code != 200:
            raise Falla(f'HTTP {r.status_code}')
        recs = r.json().get('records')
        if recs is None:
            raise Falla('la tabla del 12-B no está disponible (vacía o recargándose)')
        if not recs:
            return {}
        d = recs[0]
        out = {k: d.get(k) for k in F12_CAMPOS}
        out['PROG'] = [_num(d.get(f'MONTO_ACTUALIZADO_{i}')) for i in range(1, 13)]
        out['DEVENGADO_ACUMULADO'] = _num(out['DEVENGADO_ACUMULADO'])
        for k in [f'MONTO_ACTUALIZADO_{i}' for i in range(1, 13)]:
            out.pop(k, None)
        return out


def cargar(raiz, cuis, cfg, avisos):
    """Devuelve (ssi, f12b, detalle). Cada uno es {cui: datos|None}; vacío si está desactivado o no hay nada."""
    import os
    cuis = sorted(set(cuis)); det = []
    res = {}
    for clave, cls, activo in (('ssi', SSI, cfg.get('usar_ssi', True)), ('f12b', F12B, cfg.get('usar_f12b', True))):
        if not activo or os.environ.get('SINV_LEGAL_OFF'):
            res[clave] = {}; det.append(f'{cls.nombre}: desactivado'); continue
        o = cls(raiz, verificar=not cfg.get('api_sin_verificar_certificado', True))
        res[clave] = o.obtener(cuis)
        avisos += o.notas
        con = sum(1 for v in res[clave].values() if v)
        edad = o.antiguedad_max_dias(cuis)
        det.append(f'{cls.nombre}: {con} de {len(cuis)} CUI con dato · consultados ahora {o.nuevos}' + (f' · el más viejo tiene {edad} día(s)' if edad else ''))
        if edad is not None and edad > 3 and clave == 'f12b':
            avisos.append(f'12-B: hay datos de hace {edad} días (la API no respondió); la programación puede estar desactualizada.')
    return res['ssi'], res['f12b'], det
