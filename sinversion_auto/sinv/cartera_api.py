"""Cartera PMI directo desde el Banco de Inversiones (MEF), sin descargar el Excel.

POST https://ofi5.mef.gob.pe/invierte/Pmi/traeListaCarteraSector  (ddlSector=96 Gobiernos Regionales, ddlGobiernoRegional=458 Puno)
Verificado el 07/10/2026: devuelve JSON con los mismos 370 registros que el Excel (costo, devengado acumulado, programación 2026-2029);
solo difirió el costo de un CUI (la API es más reciente). El certificado SSL de ofi5.mef.gob.pe no se pudo verificar desde esta PC.
"""
import datetime
import requests, urllib3

URL = 'https://ofi5.mef.gob.pe/invierte/Pmi/traeListaCarteraSector'
REF = 'https://ofi5.mef.gob.pe/invierte/pmi/consultapmi'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) SINVERSION/2.0'
FORM = {'ddlSector': '96', 'ddlGobiernoRegional': '458', 'ddlDepartamento': '', 'ddlMunProvincial': '', 'ddlMunDistrital': '', 'txtCodigoUnico': ''}


class ErrorCartera(Exception):
    pass


def _f(v):
    try:
        return None if v in (None, '') else float(v)
    except (TypeError, ValueError):
        return None


def leer_cartera_api(permitir_sin_verificar=True, timeout=240):
    verificar = True; sin_verif = False; ultimo = None
    for _ in range(3):
        try:
            r = requests.post(URL, data=FORM, headers={'User-Agent': UA, 'X-Requested-With': 'XMLHttpRequest', 'Referer': REF}, timeout=timeout, verify=verificar)
            if r.status_code != 200:
                raise ErrorCartera(f'HTTP {r.status_code}')
            filas = r.json()
            break
        except requests.exceptions.SSLError as e:
            if verificar and permitir_sin_verificar:
                urllib3.disable_warnings(); verificar = False; sin_verif = True; continue
            raise ErrorCartera('Certificado SSL del MEF inválido (config api_sin_verificar_certificado = false).') from e
        except (requests.exceptions.RequestException, ValueError, ErrorCartera) as e:
            ultimo = e
    else:
        raise ErrorCartera(f'No se pudo consultar la Cartera del MEF: {str(ultimo)[:150]}')
    if not isinstance(filas, list) or not filas:
        raise ErrorCartera('La Cartera del MEF llegó vacía.')
    out = {}
    for x in filas:
        cui = str(x.get('CODIGO_UNICO') or '').strip()
        if not cui:
            continue
        out[cui] = dict(cui=cui, nombre=x.get('NOMBRE_INVERSION'), tipo=x.get('NOMBRE_TIPOINVERSION'), funcion=x.get('NOMBRE_FUNCION'),
                        costo=_f(x.get('COSTO')), dev_acum=_f(x.get('DEVENGADO_ACUMULADO')), pim=_f(x.get('PIM')),
                        m2026=_f(x.get('PROGRAMACION_INVERSION_ANIO0')), m2027=_f(x.get('PROGRAMACION_INVERSION_ANIO1')),
                        m2028=_f(x.get('PROGRAMACION_INVERSION_ANIO2')), m2029=_f(x.get('PROGRAMACION_INVERSION_ANIO3')),
                        fin_ejecucion=x.get('FIN_EJECUCION_STR'), uei=x.get('DES_UNIDAD_UEI'))
    return dict(cartera=out, fecha=datetime.datetime.now(), sin_verificar=sin_verif, registros=len(filas))
