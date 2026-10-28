"""VERSIÓN PUNTUAL (no forma parte del programa): genera el libro de proyectos de hoy con la hoja "Lista Antony" y los proyectos de Antony resaltados.

Uso:  python generar_version_antony.py
Lee lista_antony.json (esta misma carpeta), corre el programa normal con las APIs, y deja el resultado en ..\\..\\outputs\\SINVERSION_PROYECTOS_<fecha>_CON_LISTA_ANTONY.xlsx.
No toca la carpeta salida, el historial ni el maestro del programa diario.
"""
import argparse, datetime, json, pathlib, shutil, sys, tempfile

AQUI = pathlib.Path(__file__).resolve().parent
RAIZ = AQUI.parent.parent / 'sinversion_auto'
sys.path.insert(0, str(RAIZ))
import actualizar                                   # noqa: E402

ANT = json.loads((AQUI / 'lista_antony.json').read_text(encoding='utf-8'))
construir_original = actualizar.ep.construir


def construir_con_antony(path, siaf, cartera, rows, cfg, info, rubros=None, api=None):
    return construir_original(path, siaf, cartera, rows, cfg, dict(info, antony=ANT), rubros, api)


actualizar.ep.construir = construir_con_antony
actualizar.historial.guardar = lambda *a, **k: None          # esta corrida no entra al historial diario

tmp = pathlib.Path(tempfile.mkdtemp())
args = argparse.Namespace(entrada=str(RAIZ / 'entrada'), salida=str(tmp), fecha=None, vigilar=False, diario=False, aviso=False, guardar_formato=None)
actualizar.correr(args)
for p in tmp.glob('*/SINVERSION_PROYECTOS_*.xlsx'):
    destino = AQUI.parent.parent / 'outputs' / (p.stem + '_CON_LISTA_ANTONY.xlsx')
    try:
        shutil.copy(str(p), str(destino))
    except PermissionError:                       # el libro anterior está abierto en Excel
        destino = destino.with_name(destino.stem + f'_{datetime.datetime.now():%H%M}.xlsx'); shutil.copy(str(p), str(destino))
    print('Listo:', destino)
