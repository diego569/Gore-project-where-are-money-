"""Actualiza el Excel de PROYECTOS a partir de los archivos descargados.

Obligatorios (2): reporte SIAF (.xls) y Cartera del Banco de Inversiones (.xlsx).
Opcionales (hasta 4-6): reportes de la Consulta Amigable por Rubro de financiamiento (.xls, uno por rubro; si un rubro
trae más de 400 proyectos, también la 2.ª página). Si no se suben, se usan los últimos conocidos y el Excel lo avisa.

Uso:
  actualizar.exe                  -> procesa lo que haya en la carpeta entrada
  actualizar.exe --vigilar        -> solo procesa si hay archivos nuevos y completos (para la tarea programada)
  opciones: --entrada RUTA  --salida RUTA  --fecha AAAA-MM-DD  --aviso (mensaje en pantalla)
"""
import argparse, collections, concurrent.futures as cf, datetime, json, os, pathlib, shutil, sys, time, traceback

if sys.stdout is None:   # ejecutable sin consola (tarea programada)
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

if getattr(sys, 'frozen', False):
    RAIZ = pathlib.Path(sys.executable).resolve().parent
else:
    RAIZ = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ))

from sinv.siaf import leer_siaf, es_reporte_siaf
from sinv.cartera import leer_cartera, es_cartera
from sinv.rubro import leer_rubro, es_rubro, combinar
from sinv import api_mef, silucia, excel_gasto, legal_apis, historial, formato
from sinv.siaf_api import leer_siaf_api, siaf_desde_filas
from sinv.cartera_api import leer_cartera_api
from sinv.maestro import cargar_maestro, agregar_pendientes, asegurar_estructura, cargar_extras, agregar_pendientes_gc
from sinv.reglas import cargar_config
from sinv import excel_proyectos as ep

CACHE = RAIZ / 'cache_rubro.json'
CACHE_API = RAIZ / 'cache_api_rubro.json'


def aviso(titulo, texto):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, texto, titulo, 0x40 | 0x1000)
    except Exception:
        pass


def clasificar_archivos(carpeta):
    """Devuelve (siaf, cartera, [rubros]) del más reciente de cada tipo. Ignora temporales ~$ y subcarpetas."""
    sia = car = None; rub = []
    arch = sorted([p for p in carpeta.iterdir() if p.is_file() and not p.name.startswith('~$')], key=lambda p: p.stat().st_mtime, reverse=True)
    for p in arch:
        suf = p.suffix.lower()
        if suf == '.xls':
            if sia is None and es_reporte_siaf(p):
                sia = p
            elif es_rubro(p):
                rub.append(p)
        elif suf == '.xlsx' and car is None and es_cartera(p):
            car = p
    return sia, car, rub


def previo(entrada):
    """Último par SIAF+Cartera ya procesado (carpeta procesados/AAAA-MM-DD)."""
    base = entrada / 'procesados'
    if not base.exists():
        return None
    for d in sorted([x for x in base.iterdir() if x.is_dir()], reverse=True):
        s, c, _ = clasificar_archivos(d)
        if s and c:
            return s, c, datetime.date.fromisoformat(d.name)
    return None


def estable(p, seg=20):
    return (time.time() - p.stat().st_mtime) >= seg


def _a_json(l):
    d = {k: v for k, v in l.items() if k not in ('archivo', 'archivos', 'origen')}
    d['fecha'] = l['fecha'].isoformat() if l.get('fecha') else None
    return d


def _de_json(d):
    d = dict(d); d['fecha'] = datetime.date.fromisoformat(d['fecha']) if d.get('fecha') else None
    return d


def cargar_rubros(paths, avisos):
    cache = {}
    if CACHE.exists():
        try:
            cache = {k: _de_json(v) for k, v in json.loads(CACHE.read_text(encoding='utf-8')).items()}
        except Exception:
            avisos.append('No se pudo leer cache_rubro.json; se ignora.')
    nuevas = []
    for p in paths:
        try:
            nuevas.append(leer_rubro(p))
        except Exception as e:
            avisos.append(f'No se pudo leer el reporte de rubro {p.name}: {e}')
    res = {}
    for code, l in combinar(nuevas).items():
        l['origen'] = 'cargado ahora'
        c = cache.get(code)
        if not l['completo'] and c and c.get('fecha') == l.get('fecha'):   # llegó otra página el mismo día
            l['proyectos'] = {**c['proyectos'], **l['proyectos']}
            l['suma_pim'] = sum(v['pim'] for v in l['proyectos'].values())
            l['completo'] = abs(l['suma_pim'] - l['total']['pim']) <= 1.0
        res[code] = l
    for code, c in cache.items():
        if code not in res:
            c['origen'] = f"copia del {c['fecha']:%d/%m/%Y}, no se subió reporte nuevo" if c.get('fecha') else 'copia anterior'
            res[code] = c
    if res:
        CACHE.write_text(json.dumps({k: _a_json(v) for k, v in res.items()}, ensure_ascii=False), encoding='utf-8')
    return res


def descargar_api(cfg, siaf, avisos):
    """Baja en paralelo los registros de proyectos y de gasto corriente (UE 001) de Datos Abiertos. Si la API está caída, se rinde rápido."""
    if not cfg.get('usar_api', True) or os.environ.get('SINV_API_OFF'):
        return None, None, 'Desactivada (config.json: usar_api = false).'
    cli = api_mef.Cliente(permitir_sin_verificar=cfg.get('api_sin_verificar_certificado', True))
    ue_gc = '001'; fp = siaf.get('filas_proy'); fg = siaf.get('filas_gc'); err = ''
    tareas = {}
    with cf.ThreadPoolExecutor(2) as ex:
        if fp is None:
            tareas['p'] = ex.submit(cli.descargar, api_mef.PROYECTOS)
        if fg is None:
            tareas['g'] = ex.submit(cli.descargar, {'ANO_EJE': '2026', 'PLIEGO': '458', 'TIPO_ACT_PROY': '3', 'EJECUTORA': ue_gc})
        for k, f in tareas.items():
            try:
                r = f.result()
                if k == 'p': fp = r
                else: fg = r
            except Exception as e:
                err = f'La API del MEF no respondió o falló: {str(e)[:200]}'
    if cli.sin_verificar_usado:
        avisos.append('API del MEF consultada SIN verificar el certificado SSL (vencido/inválido); el PIM se valida contra el SIAF.')
    if fp and fg is not None and not siaf.get('desde_copia') and tareas:
        try:
            api_mef.guardar_filas(RAIZ, fp if 'p' in tareas else None, fg if 'g' in tareas else None)
        except OSError:
            pass
    return fp, fg, err


def obtener_api(cfg, siaf_rows, avisos, filas_proy=None, error=''):
    """Rubro por UE-CUI desde la API de Datos Abiertos. Devuelve (api|None, detalle). Valida contra el PIM del SIAF."""
    if not cfg.get('usar_api', True) or os.environ.get('SINV_API_OFF'):
        return None, 'Desactivada (config.json: usar_api = false).'
    if filas_proy is None:
        return None, error or 'La API del MEF no respondió.'
    pim, nombres = api_mef.pim_por_ue_cui_rubro(filas_proy)
    api = dict(pim=pim, nombres=nombres, registros=len(filas_proy), fecha=datetime.datetime.now(), sin_verificar=False,
               meses=sorted({int(r['MES_EJE']) for r in filas_proy}))
    pim_siaf = collections.defaultdict(float)
    for r in siaf_rows:
        pim_siaf[(r['ue'], r['cui'])] += r['pim']
    malos = [k for k, v in pim_siaf.items() if abs(sum(api['pim'].get(k, {}).values()) - v) > 1]
    if len(malos) > max(3, 0.02 * len(pim_siaf)):
        return None, f'La API se descartó: su PIM no coincide con el SIAF en {len(malos)} de {len(pim_siaf)} proyectos-UE.'
    if malos:
        avisos.append(f'API del MEF: {len(malos)} proyecto(s)-UE con PIM distinto al SIAF (ver columna Cobertura de "Rubro por CUI").')
    try:
        CACHE_API.write_text(json.dumps({'fecha': api['fecha'].isoformat(), 'registros': api['registros'], 'meses': api['meses'], 'sin_verificar': api['sin_verificar'],
                                         'pim': {f'{k[0]}|{k[1]}': v for k, v in api['pim'].items()}, 'nombres': api['nombres']}), encoding='utf-8')
    except OSError:
        pass
    return api, ''


def api_desde_cache():
    if not CACHE_API.exists():
        return None
    try:
        d = json.loads(CACHE_API.read_text(encoding='utf-8'))
        return dict(pim={tuple(k.split('|')): v for k, v in d['pim'].items()}, nombres=d['nombres'], registros=d['registros'],
                    fecha=datetime.datetime.fromisoformat(d['fecha']), sin_verificar=d['sin_verificar'], meses=d['meses'])
    except Exception:
        return None


def obtener_silucia(cfg, avisos, acciones):
    """Metas de Silucia (ejecutora/oficina). Nunca detiene la corrida. Silucia cambia ~cada 3 meses: si el token venció se usa la
    última copia local; solo se pide renovar el token cuando la copia no existe o tiene más de 'silucia_cache_dias' días."""
    alerta = RAIZ / 'ALERTA_TOKEN.txt'
    if not cfg.get('usar_silucia', True) or os.environ.get('SINV_SILUCIA_OFF'):
        return {}, 'Desactivada (config.json: usar_silucia = false).'
    try:
        data, av = silucia.metas(RAIZ, anio=cfg.get('anio', 2026))
        avisos += av
        if av:
            acciones.append(av[0])
        elif alerta.exists():
            alerta.unlink()
        try:
            silucia.guardar_cache(RAIZ, data)
        except OSError:
            pass
        return data, f'{len(data)} metas leídas de Silucia (en vivo)'
    except (silucia.TokenFaltante, silucia.TokenVencido) as e:
        motivo = str(e)
    except Exception as e:
        motivo = f'Silucia no respondió ({str(e)[:120]})'
    cache, cuando = silucia.leer_cache(RAIZ)
    limite = cfg.get('silucia_cache_dias', 90)
    if cache is not None:
        dias = (datetime.datetime.now() - cuando).days
        if dias <= limite:
            return cache, f'{len(cache)} metas de la copia local del {cuando:%d/%m/%Y} (sin token vigente; Silucia cambia cada ~3 meses)'
        msg = f'Silucia: {motivo} La copia local tiene {dias} días (más de {limite}); puede estar desactualizada.'
    else:
        msg = f'Silucia: {motivo} No hay copia local.'
    acciones.append('ACCIÓN REQUERIDA: renovar el token de Silucia (token_silucia.txt) para actualizar la clasificación.')
    try:
        alerta.write_text(f'{datetime.datetime.now():%d/%m/%Y %H:%M}'+chr(10)+msg+chr(10)+f'Pegue el token nuevo (una sola línea) en token_silucia.txt, junto a actualizar.exe, o use RENOVAR_TOKEN_SILUCIA.bat.', encoding='utf-8')
    except OSError:
        pass
    avisos.append(msg + ' Se usa solo el maestro para clasificar.')
    return (cache or {}), 'No disponible (ver avisos)'


def correr(args):
    entrada = pathlib.Path(args.entrada); salida = pathlib.Path(args.salida)
    entrada.mkdir(exist_ok=True); salida.mkdir(exist_ok=True)
    cfg = cargar_config(RAIZ)
    modo_siaf = os.environ.get('SINV_FUENTE_SIAF') or cfg.get('fuente_siaf', 'archivo')        # archivo | api | auto
    api_cart = cfg.get('usar_api_cartera', True) and not os.environ.get('SINV_CARTERA_API_OFF')
    sia, car, rub = clasificar_archivos(entrada)
    if args.vigilar:                                   # modo "archivos": solo actúa cuando llegan archivos nuevos y completos
        if not (sia or car or rub):
            return 'nada'
        if not all(estable(p) for p in (sia, car, *rub) if p):
            return 'copiando'
    reutiliza = None
    if modo_siaf == 'archivo' and not sia or (not api_cart and not car):
        if rub:
            reutiliza = previo(entrada)
        if reutiliza is None:
            faltan = [n for n, v, req in (('reporte SIAF (.xls)', sia, modo_siaf == 'archivo'), ('Cartera (.xlsx)', car, not api_cart)) if req and not v]
            raise SystemExit('Faltan archivos en la carpeta "entrada": ' + ', '.join(faltan) + '.')
    mpath = RAIZ / 'maestro.xlsx'
    if not mpath.exists():
        raise SystemExit('No existe maestro.xlsx. Ejecute primero sembrar_maestro.py (una sola vez).')
    asegurar_estructura(mpath)
    avisos = []; acciones = []; usados = []; banner = ''
    # ---------------- SIAF
    siaf = None; fuente_siaf = ''; f_siaf = datetime.date.today()
    if reutiliza:
        sia, car_prev, fecha_previa = reutiliza; car = car if car else car_prev
        avisos.append(f'No se subieron SIAF y Cartera nuevos: se reutilizaron los del {fecha_previa:%d/%m/%Y}.')
    if modo_siaf in ('api', 'auto'):
        try:
            cli = api_mef.Cliente(permitir_sin_verificar=cfg.get('api_sin_verificar_certificado', True))
            siaf = leer_siaf_api(cli, ue_gasto_corriente=tuple(cfg.get('ue_gasto_corriente', ['001'])), raiz=RAIZ)
            fuente_siaf = 'API Datos Abiertos MEF · ' + (siaf.get('metodo') or 'API')
            for n_ in siaf.get('notas', []):
                avisos.append('API del MEF: ' + n_ + ' Se usó la API paginada.')
            if siaf['sin_verificar']:
                avisos.append('API del MEF consultada SIN verificar el certificado SSL; el PIM se validó contra el SIAF/Cartera.')
        except Exception as e:
            avisos.append(f'La API del MEF no devolvió las cifras ({str(e)[:150]}).' + (' Se usa el reporte SIAF (.xls).' if sia else ''))
            if not sia:
                copia = api_mef.leer_filas(RAIZ)
                dias_max = cfg.get('cache_cifras_dias', 3)
                if copia.get('proy') and (datetime.datetime.now() - copia['proy'][1]).days <= dias_max:
                    fp_, cuando = copia['proy']; fg_ = copia.get('gc', ([], None))[0]
                    siaf = siaf_desde_filas(fp_, fg_, fuente='copia'); siaf['desde_copia'] = True
                    fuente_siaf = f'COPIA de la API del MEF del {cuando:%d/%m/%Y %H:%M} (la API no respondió)'
                    banner = f'CIFRAS DESACTUALIZADAS: copia de la API del MEF del {cuando:%d/%m/%Y %H:%M}. Suba el reporte SIAF (.xls) o espere a que la API se recupere.'
                    acciones.append('ATENCIÓN: ' + banner)
                    avisos.append(banner)
                else:
                    raise SystemExit('No se pudieron obtener las cifras: la API del MEF falló, no hay reporte SIAF (.xls) en "entrada" y no hay una copia reciente '
                                     f'(máximo {dias_max} días). Suba el reporte SIAF (.xls) a la carpeta "entrada" y vuelva a ejecutar.')
    if siaf is None:
        if not sia:
            raise SystemExit('Falta el reporte SIAF (.xls) en la carpeta "entrada".')
        siaf = leer_siaf(sia); fuente_siaf = f'Reporte SIAF {sia.name}'; f_siaf = datetime.date.fromtimestamp(sia.stat().st_mtime)
        if not reutiliza: usados.append(sia)
        for u, t in siaf['ue_tot'].items():
            s = sum(m['pim'] for m in siaf['metas'] if m['ue'] == u)
            if abs(s - t['pim']) > 1:
                avisos.append(f'UE {u}: la suma de metas del SIAF ({s:,.0f}) no coincide con el total de la UE ({t["pim"]:,.0f}). Revisar formato del reporte.')
    # ---------------- Cartera
    cart = None; fuente_cart = ''
    if api_cart:
        try:
            cart = leer_cartera_api(cfg.get('api_sin_verificar_certificado', True)); fuente_cart = 'API Banco de Inversiones (MEF)'
            if cart['sin_verificar']:
                avisos.append('Cartera del MEF consultada SIN verificar el certificado SSL (no se pudo validar en esta PC).')
        except Exception as e:
            avisos.append(f'La Cartera por API falló ({str(e)[:150]}).' + (' Se usa el archivo de Cartera (.xlsx).' if car else ''))
    if cart is None:
        if not car:
            raise SystemExit('No se pudo obtener la Cartera: la API del MEF falló y no hay Cartera (.xlsx) en "entrada".')
        cart = leer_cartera(car); fuente_cart = f'Archivo {car.name}'
        if not reutiliza: usados.append(car)
    maestro = cargar_maestro(mpath); maestro['gc'], maestro['mapa'] = cargar_extras(mpath)
    # ---------------- fecha de corte
    if reutiliza:
        fecha = fecha_previa
    elif args.fecha:
        fecha = datetime.date.fromisoformat(args.fecha)
    elif fuente_cart.startswith('Archivo') and cart['fecha']:
        fecha = cart['fecha'].date()
    else:
        fecha = datetime.date.today()
    if fuente_siaf.startswith('Reporte') and fuente_cart.startswith('Archivo') and cart['fecha'] and abs((cart['fecha'].date() - f_siaf).days) > 0:
        avisos.append(f'El archivo SIAF es del {f_siaf:%d/%m/%Y} y la Cartera del {cart["fecha"]:%d/%m/%Y}: no son del mismo día.')
    # ---------------- Silucia (clasificación)
    silu, silu_det = obtener_silucia(cfg, avisos, acciones)
    # ---------------- proyectos
    # ---------------- datos por CUI que la API de ejecución no trae: ficha SSI (tipo, cierre, costo) y 12-B (programación, devengado acumulado)
    cuis_legal = sorted({m['cui'] for m in siaf['metas'] if m['tipo'] == 'proyecto' and (m['pim'] > 0 or m.get('pia', 0) > 0)} - set(cfg.get('bolsas_genericas', [])))
    try:
        ssi_d, f12_d, legal_det = legal_apis.cargar(RAIZ, cuis_legal, cfg, avisos)
    except Exception as e:
        ssi_d, f12_d, legal_det = {}, {}, [f'Datos legales no disponibles ({str(e)[:100]}).']
        avisos.append(f'No se pudieron obtener SSI/12-B ({str(e)[:100]}): el semáforo queda con datos incompletos.')
    f_a1 = RAIZ / 'anexo1_cuis.txt'          # lista del Anexo I (un CUI por línea; la entrega la OPMI)
    if f_a1.exists():
        import re as _re
        cfg['anexo1_cuis'] = sorted(set(cfg.get('anexo1_cuis', [])) | set(_re.findall(r'\d{7}', f_a1.read_text(encoding='utf-8', errors='ignore'))))
        legal_det.append(f"Anexo I: {len(cfg['anexo1_cuis'])} CUI (anexo1_cuis.txt)")
    elif not cfg.get('anexo1_cuis'):
        legal_det.append('Anexo I: lista no cargada (pendiente de la OPMI; ver LEEME)')
    rows, pend = ep.armar_filas(siaf, cart, maestro, cfg, silucia=silu, mapa=maestro['mapa'], ssi=ssi_d, f12b=f12_d)
    n_pend = agregar_pendientes(mpath, pend)
    # ---------------- rubro (proyectos)
    filas_p, filas_gc, err_api = descargar_api(cfg, siaf, avisos)
    api, api_detalle = obtener_api(cfg, rows, avisos, filas_proy=filas_p, error=err_api)
    rubros = cargar_rubros(rub, avisos)
    usados += rub
    if api is None:
        avisos.append('Rubro: ' + api_detalle)
        api_c = api_desde_cache()
        if api_c and (datetime.datetime.now() - api_c['fecha']).days <= 7:
            api = api_c; avisos.append(f"Rubro: se usa la última copia de la API del MEF ({api_c['fecha']:%d/%m/%Y %H:%M}); el PIM por rubro cambia poco de un día a otro.")
        else:
            avisos.append('Rubro: se usa la Consulta Amigable (respaldo).')
    for c in sorted({c for c, _ in ep.RUBROS} - set(rubros) if api is None else ()):
        avisos.append(f'Rubro {c}: no hay reporte (ni copia anterior). Sus columnas quedan "s/d".')
    for c, l in sorted(rubros.items() if api is None else ()):
        if not l['completo']:
            avisos.append(f'Rubro {c}: reporte INCOMPLETO ({len(l["proyectos"])} proyectos; PIM {l["suma_pim"]:,.0f} de {l["total"]["pim"]:,.0f}). Falta la otra página.')
        elif 'copia' in l.get('origen', ''):
            avisos.append(f'Rubro {c}: {l["origen"]}.')
        if l.get('fecha') and l['fecha'] != fecha:
            avisos.append(f'Rubro {c}: la consulta es del {l["fecha"]:%d/%m/%Y} y el corte del Excel es {fecha:%d/%m/%Y}.')
    # ---------------- gasto corriente
    ue_gc = '001'
    if filas_gc is None and err_api:
        avisos.append('Rubro del gasto corriente: ' + err_api + ' Queda "s/d".')
    rubros_meta, nombres_fn = (api_mef.pim_por_meta_rubro(filas_gc) if filas_gc else (None, {}))
    rows_gc, pend_gc = excel_gasto.armar_filas_gc(siaf, maestro, cfg, ue=ue_gc, silucia=silu, mapa=maestro['mapa'], rubros_meta=rubros_meta, nombres_fn=nombres_fn)
    n_pend_gc = agregar_pendientes_gc(mpath, pend_gc)
    # ---------------- escribir Excel
    carpeta = salida / f'{fecha:%Y-%m-%d}'; carpeta.mkdir(exist_ok=True)

    def destino_de(nombre):
        d = carpeta / f'{nombre}_{fecha:%Y-%m-%d}.xlsx'
        try:
            if d.exists():
                d.unlink()
        except PermissionError:
            d = carpeta / f'{nombre}_{fecha:%Y-%m-%d}_{datetime.datetime.now():%H%M}.xlsx'
            avisos.append(f'El Excel {nombre} estaba abierto: se guardó con otro nombre (con la hora).')
        return d
    d_proy = destino_de('SINVERSION_PROYECTOS'); d_gc = destino_de('SINVERSION_GASTO_CORRIENTE')
    info = dict(fecha_corte=fecha, siaf=fuente_siaf, siaf_fecha=f_siaf.strftime('%d/%m/%Y'), cartera=fuente_cart, avisos=avisos, api_detalle=api_detalle,
                fuente_siaf=fuente_siaf, silucia=silu_det, banner=banner, legal=legal_det)
    ahora_ts = datetime.datetime.now()
    prev, prev_t = historial.cargar_anterior(RAIZ, ahora_ts)
    info['cambios'] = dict(lista=historial.comparar(rows, prev), desde=prev_t)
    info['formato'] = formato.cargar(RAIZ)
    if not info['formato']:
        avisos.append('No hay formato_oficina.json: el Excel sale con el formato por defecto (ver LEEME: GUARDAR_FORMATO).')
    ep.construir(d_proy, siaf, cart, rows, cfg, info, rubros, api)
    try:
        historial.guardar(RAIZ, rows, ahora_ts)
    except OSError:
        avisos.append('No se pudo guardar el historial de la corrida (carpeta "historial").')
    excel_gasto.construir_gc(d_gc, siaf, rows_gc, cfg, info, ue=ue_gc,
                             rubro_fuente='API Datos Abiertos MEF (por meta)' if rubros_meta else 'No disponible')
    arch = entrada / 'procesados' / f'{fecha:%Y-%m-%d}'
    for p in usados:
        try:
            arch.mkdir(parents=True, exist_ok=True)
            if (arch / p.name).exists():
                p.unlink()
            else:
                shutil.move(str(p), str(arch / p.name))
        except OSError:
            avisos.append(f'No se pudo mover {p.name} a "procesados" (¿está abierto en Excel?). Muévalo a mano para que no se vuelva a procesar.')
    sin_proy = sum(1 for r in rows if r['hoja'] == 'Sin clasificar'); sin_gc = sum(1 for r in rows_gc if r['of'].startswith('SIN CLASIFICAR'))
    resumen = acciones + [f'Listo: {d_proy}', f'       {d_gc}', f'Proyectos (UE-CUI-meta): {len(rows)} · sin clasificar: {sin_proy}',
                          f'Gasto corriente UE {ue_gc} (metas): {len(rows_gc)} · sin oficina: {sin_gc}',
                          f'Fuentes: cifras = {fuente_siaf}; cartera = {fuente_cart}; Silucia = {silu_det}',
                          f'Agregados al maestro para clasificar: {n_pend} proyectos, {n_pend_gc} metas de gasto corriente'] + [f'AVISO: {a}' for a in avisos]
    (salida / 'ULTIMO_RESULTADO.txt').write_text('\n'.join([f'{datetime.datetime.now():%d/%m/%Y %H:%M}'] + resumen), encoding='utf-8')
    return resumen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--entrada', default=str(RAIZ / 'entrada')); ap.add_argument('--salida', default=str(RAIZ / 'salida'))
    ap.add_argument('--guardar-formato', metavar='EXCEL', help='guarda el orden, hojas ocultas y anchos de un Excel retocado a mano')
    ap.add_argument('--fecha'); ap.add_argument('--vigilar', action='store_true', help='solo si llegaron archivos nuevos (tarea cada 10 min)')
    ap.add_argument('--diario', action='store_true', help='corrida completa por horario (tarea diaria con APIs)'); ap.add_argument('--aviso', action='store_true')
    args = ap.parse_args()
    if args.guardar_formato:
        try:
            if args.guardar_formato:
                f = formato.guardar(RAIZ, args.guardar_formato)
                m = f"Formato guardado: {len(f['orden'])} hojas ({len(f['ocultas'])} ocultas), anchos y zoom. Se aplicará en las próximas actualizaciones."
                print(m); aviso('SINVERSION', m)
        except SystemExit as e:
            print('ERROR:', e); aviso('SINVERSION: error', str(e)); sys.exit(1)
        except Exception as e:
            print('ERROR:', e); aviso('SINVERSION: error', str(e)); sys.exit(1)
        return
    try:
        r = correr(args)
        if isinstance(r, list):
            print('\n'.join(r))
            if args.aviso:
                aviso('SINVERSION actualizado', '\n'.join([x for x in r if x.startswith(('ACCIÓN', 'Listo', 'Proyectos', 'Gasto'))][:5]))
    except SystemExit as e:
        msg = str(e)
        pathlib.Path(args.salida).mkdir(exist_ok=True)
        (pathlib.Path(args.salida) / 'ERROR.txt').write_text(f'{datetime.datetime.now():%d/%m/%Y %H:%M}\n{msg}', encoding='utf-8')
        print('ERROR:', msg)
        if args.aviso:
            aviso('SINVERSION: no se pudo actualizar', msg)
        sys.exit(1)
    except Exception:
        tb = traceback.format_exc()
        pathlib.Path(args.salida).mkdir(exist_ok=True)
        (pathlib.Path(args.salida) / 'ERROR.txt').write_text(f'{datetime.datetime.now():%d/%m/%Y %H:%M}\nOcurrió un error inesperado. Envíe este archivo al soporte.\n\n{tb}', encoding='utf-8')
        (RAIZ / 'logs').mkdir(exist_ok=True)
        (RAIZ / 'logs' / f'error_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt').write_text(tb, encoding='utf-8')
        print(tb)
        if args.aviso:
            aviso('SINVERSION: error', 'Ocurrió un error inesperado. Revise salida\\ERROR.txt')
        sys.exit(1)


if __name__ == '__main__':
    main()
