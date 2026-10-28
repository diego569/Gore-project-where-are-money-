"""Se ejecuta UNA vez: crea maestro.xlsx a partir del Excel de Alex (SINVERSION 16/09)."""
import sys, pathlib
from sinv.maestro import sembrar_desde_alex, guardar_maestro
raiz = pathlib.Path(__file__).resolve().parent
alex = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else raiz.parent / 'ref docs' / '0001917176_1_SINVERSION (1).xlsm'
extra = raiz.parent / 'ref docs' / 'nuevo' / '0001846478_1_LINVERSION.xlsm'
clas, funcs = sembrar_desde_alex(alex, extra if extra.exists() else None)
destino = raiz / 'maestro.xlsx'
if destino.exists():
    sys.exit('maestro.xlsx ya existe; no se sobrescribe (bórralo si quieres sembrarlo de nuevo).')
guardar_maestro(destino, clas, [], funcs); from sinv.maestro import asegurar_estructura; asegurar_estructura(destino)
print('maestro creado:', len(clas), 'filas de clasificación;', len(funcs), 'funciones')
