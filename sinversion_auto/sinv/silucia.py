"""Cliente de la API Silucia del GORE Puno (metas, con la ejecutora/oficina responsable).

Verificado el 07/10/2026:
  - Requiere "Authorization: Bearer <token>" y un User-Agent de navegador (con el User-Agent por defecto de python-requests
    el firewall responde 403 "Bot blocked").
  - Filtro de año: filter_by[anio]=2026; per_page=100 (557 metas en 6 páginas).
  - Las cifras de dinero de esta API (pim, certificación, etc.) NO son confiables (muchas metas traen pim = 1): solo se usa la
    clasificación (ejecutora, función) y la llave (prod_proy = CUI/código de producto, codmeta = meta SIAF de 4 dígitos).
  - Las metas de gasto corriente (cui_flag = 'C') llegan todas con la ejecutora 'GASTO CORRIENTE' (id 0), sin oficina.
  - El token de la prueba duraba 10 horas (iat/exp dentro del JWT), no meses.
"""
import base64, datetime, json, pathlib
import requests

BASE = 'https://sistemas.regionpuno.gob.pe/siluciav2-api/api/'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) SINVERSION/2.0'


class TokenFaltante(Exception):
    pass


class TokenVencido(Exception):
    pass


class ErrorSilucia(Exception):
    pass


def leer_token(raiz):
    """Lee token_silucia.txt (con o sin la palabra Bearer). Devuelve el JWT sin prefijo."""
    p = pathlib.Path(raiz) / 'token_silucia.txt'
    if not p.exists():
        raise TokenFaltante('No existe token_silucia.txt: pegue allí el token de Silucia (una sola línea).')
    t = ' '.join(p.read_text(encoding='utf-8', errors='ignore').split())
    if t.lower().startswith('bearer '):
        t = t[7:]
    t = t.strip()
    if t.count('.') != 2 or len(t) < 60 or not t.startswith('eyJ'):
        raise TokenFaltante('token_silucia.txt está vacío o no tiene un token válido (debe empezar con "eyJ").')
    return t


def vencimiento(token):
    """Fecha/hora de vencimiento según el propio JWT (sin verificar firma). None si no se puede leer."""
    try:
        pl = json.loads(base64.urlsafe_b64decode(token.split('.')[1] + '=' * (-len(token.split('.')[1]) % 4)))
        return datetime.datetime.fromtimestamp(pl['exp'])
    except Exception:
        return None


def metas(raiz, anio=2026, ahora=None, aviso_dias=3):
    """Devuelve (dict[(codigo_producto, meta4)] -> info, lista_de_avisos). Lanza TokenFaltante/TokenVencido/ErrorSilucia."""
    ahora = ahora or datetime.datetime.now()
    tok = leer_token(raiz); avisos = []
    exp = vencimiento(tok)
    if exp is not None:
        if exp <= ahora:
            raise TokenVencido(f'El token de Silucia venció el {exp:%d/%m/%Y %H:%M}. Renuévelo y péguelo en token_silucia.txt.')
        if exp - ahora <= datetime.timedelta(days=aviso_dias):
            avisos.append(f'El token de Silucia vence el {exp:%d/%m/%Y %H:%M}: renuévelo pronto (token_silucia.txt).')
    h = {'Authorization': 'Bearer ' + tok, 'Accept': 'application/json', 'User-Agent': UA}
    filas = []; pag = 1
    while True:
        try:
            r = requests.get(BASE + 'metas', headers=h, params={'per_page': 100, 'filter_by[anio]': anio, 'page': pag}, timeout=120)
        except requests.exceptions.RequestException as e:
            raise ErrorSilucia(f'No se pudo conectar con Silucia: {str(e)[:120]}')
        if r.status_code in (401, 403) and 'Bot' not in r.text:
            raise TokenVencido('Silucia rechazó el token (vencido o inválido). Renuévelo y péguelo en token_silucia.txt.')
        if r.status_code != 200:
            raise ErrorSilucia(f'Silucia respondió HTTP {r.status_code}: {r.text[:100]}')
        d = r.json(); filas += d['data']
        if pag >= d['last_page']:
            break
        pag += 1
    if len(filas) != d['total']:
        raise ErrorSilucia(f'Silucia: descarga incompleta ({len(filas)} de {d["total"]}).')
    out = {}
    for m in filas:
        uej = m.get('uejecutora') or {}
        out[(str(m.get('prod_proy') or ''), str(m.get('codmeta') or '').zfill(4))] = dict(
            ejecutora_id=m.get('iduejecutora'), ejecutora=(uej.get('desuejecutora') or '').strip(), funcion=(m.get('funcion') or {}).get('desfuncion', ''),
            tipo='gasto_corriente' if m.get('cui_flag') == 'C' else 'proyecto', desmeta=m.get('desmeta') or '')
    return out, avisos


# ---------------------------------------------------------------- copia local (Silucia cambia ~cada 3 meses: no hace falta token a diario)
def guardar_cache(raiz, data):
    p = pathlib.Path(raiz) / 'cache_silucia.json'
    p.write_text(json.dumps({'fecha': datetime.datetime.now().isoformat(), 'metas': {f'{k[0]}|{k[1]}': v for k, v in data.items()}}, ensure_ascii=False), encoding='utf-8')


def leer_cache(raiz):
    p = pathlib.Path(raiz) / 'cache_silucia.json'
    if not p.exists():
        return None, None
    try:
        d = json.loads(p.read_text(encoding='utf-8'))
        return {tuple(k.split('|')): v for k, v in d['metas'].items()}, datetime.datetime.fromisoformat(d['fecha'])
    except Exception:
        return None, None
