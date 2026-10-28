"""Reglas de clasificación de disponibilidad. Los umbrales viven en config.json y en la hoja Configuración del Excel."""
import json, pathlib

DEFAULT = {'umbral_ejecutado': 0.90, 'usar_api': True, 'api_sin_verificar_certificado': True,
           'fuente_siaf': 'archivo', 'usar_api_cartera': True, 'usar_silucia': True, 'anio': 2026, 'silucia_cache_dias': 90, 'cache_cifras_dias': 3, 'ue_gasto_corriente': ['001']}

ESTADOS = ['Sin certificar', 'Solo certificado', 'Con compromiso', 'Ejecutado', 'Sin PIM']
DISPONIBILIDAD = {'Sin certificar': 'Disponible', 'Solo certificado': 'Posible', 'Con compromiso': 'Parcial',
                  'Ejecutado': 'No usar', 'Sin PIM': 'Sin PIM'}


def cargar_config(raiz):
    p = pathlib.Path(raiz) / 'config.json'
    cfg = dict(DEFAULT)
    if p.exists():
        cfg.update(json.loads(p.read_text(encoding='utf-8')))
    return cfg


def clasificar(pim, cert, comp, dev, cfg):
    """Misma lógica que la fórmula de la hoja Base. Devuelve (estado, disponibilidad, disponible)."""
    if pim <= 0:
        e = 'Sin PIM'
    elif dev / pim >= cfg['umbral_ejecutado']:
        e = 'Ejecutado'
    elif comp > 0.005:
        e = 'Con compromiso'
    elif cert > 0.005:
        e = 'Solo certificado'
    else:
        e = 'Sin certificar'
    disp = 0.0 if (pim <= 0 or e == 'Ejecutado') else max(0.0, pim - comp)
    return e, DISPONIBILIDAD[e], disp


# ---------------------------------------------------------------------------------------------------------------------
# Reglas v3 (por saldos y reglas legales). Sin umbral de "% comprometido".
# ---------------------------------------------------------------------------------------------------------------------
DEFAULT.update({'monto_minimo': 10000, 'umbral_avance_acumulado': 0.90, 'usar_ssi': True, 'usar_f12b': True, 'anexo1_cuis': [],
                'bolsas_genericas': ['2001621', '2016766']})
SEMAFOROS = ['Disponible', 'Con condiciones', 'Rebajando certificación', 'Restringido', 'No mover', 'Sin saldo útil', 'Bolsa genérica', 'Sin PIM']
TIPOS = {'PROYECTO DE INVERSION': 'Proyecto de inversión', 'IOARR': 'IOARR', 'INVERSIONES IOARR': 'IOARR', 'PROYECTO PROCOMPITE': 'PROCOMPITE',
         'PROYECTO GENERICO': 'Genérico'}


def tipo_inversion(ssi, car):
    t = (ssi or {}).get('TIPO_FORMATO') or (car and car.get('tipo')) or ''
    return TIPOS.get(str(t).upper().strip(), 'Proyecto de inversión' if t else 'Sin dato')


def datos_cui(cui, pim_cui, cert_cui, dev_cui, comp_cui, ssi, f12, car, cfg):
    """Cálculos que dependen del CUI completo (todas sus UE/metas): avance acumulado, programación 12-B, saldo libre y PMI."""
    ssi = ssi or None; f12 = f12 or None
    tipo = tipo_inversion(ssi, car)
    costo = (ssi or {}).get('COSTO_ACTUALIZADO') or (car and car.get('costo')) or 0.0
    if f12:
        acum = f12['DEVENGADO_ACUMULADO']; fuente_acum = '12-B'
    elif ssi:
        acum = ssi['DEV_ACUMULADO']; fuente_acum = 'SSI'
    elif car and car.get('dev_acum') is not None:
        acum = (car['dev_acum'] or 0) + dev_cui; fuente_acum = 'Cartera + SIAF 2026'
    else:
        acum = None; fuente_acum = ''
    if acum is not None:
        acum = max(acum, dev_cui)
    avance = (acum / costo) if (acum is not None and costo > 0) else None
    prog_ok = bool(f12) and sum(f12['PROG']) > 0
    prog = sum(f12['PROG']) if prog_ok else None
    sin_cert = max(0.0, pim_cui - cert_cui)
    libre = max(0.0, min(pim_cui - prog, sin_cert)) if prog_ok else sin_cert
    tope = max(0.0, min(pim_cui - prog, pim_cui - comp_cui)) if prog_ok else max(0.0, pim_cui - comp_cui)   # si además se rebaja la certificación no comprometida
    pmi26_ = (car or {}).get('m2026') or 0.0
    ind_pmi = (ssi or {}).get('IND_REG_PMI') or ''
    en_pmi = pmi26_ > 0 or ind_pmi == 'SI'      # la ficha SSI sabe si está en el PMI de cualquier entidad; la Cartera solo trae el del GR
    # Fuera del PMI 2026 el 12-B no frena: el proyecto no se ejecutará (impedimento comprobable, num. 11.2) y la UEI debe actualizar su 12-B
    libre_estricto = libre
    if prog_ok and not en_pmi:
        libre = sin_cert
    pot = max(0.0, tope - libre)
    pmi26 = (car or {}).get('m2026') or 0.0; pmi27 = (car or {}).get('m2027') or 0.0
    return dict(pot=pot, libre_estricto=libre_estricto, en_pmi=en_pmi, ind_pmi=ind_pmi, entidad=(ssi or {}).get('ENTIDAD') or '', nivel=(ssi or {}).get('NIVEL') or '', comp_cui=comp_cui, tipo=tipo, costo=costo, acum=acum, fuente_acum=fuente_acum, avance=avance, prog_ok=prog_ok, prog=prog, libre=libre, sin_cert=sin_cert,
                pmi26=pmi26, pmi27=pmi27, en_cartera=car is not None, tiene_ssi=bool(ssi), tiene_f12=bool(f12), pim_cui=pim_cui, dev_cui=dev_cui,
                estado_inv=(ssi or {}).get('ESTADO') or '', cierre=(ssi or {}).get('CIERRE_REGISTRADO') or '', avance_fisico=(ssi or {}).get('AVAN_FISICO'),
                f12_fecha=(f12 or {}).get('FECHA_ULT_ACT_F12B') or '')


def evaluar(cui, pim, cert, comp, d, cfg):
    """Semáforo y disponible legal de una fila (UE-CUI-meta). d = datos_cui(...). Devuelve (semaforo, disponible, motivo, potencial)."""
    if pim <= 0:
        return 'Sin PIM', 0.0, 'PIM igual a 0', 0.0
    if cui in set(cfg.get('bolsas_genericas', [])):
        return 'Bolsa genérica', max(0.0, pim - cert), 'Código genérico (no es una inversión con CUI propio)', 0.0
    reparto = (pim / d['pim_cui']) if d['pim_cui'] > 0 else 1.0
    disp = min(max(0.0, pim - cert), d['libre'] * reparto)
    pot = min(max(0.0, pim - comp) - disp, d['pot'] * reparto)
    pot = max(0.0, pot)
    if pim > 0 and d['dev_cui'] / d['pim_cui'] >= cfg['umbral_ejecutado']:
        return 'No mover', 0.0, f"Devengado 2026 ≥ {cfg['umbral_ejecutado']:.0%} del PIM", 0.0
    if d['avance'] is not None and d['avance'] > cfg['umbral_avance_acumulado']:
        return 'No mover', 0.0, f"Avance financiero acumulado {d['avance']:.0%} (> {cfg['umbral_avance_acumulado']:.0%})", 0.0
    if d['tipo'] == 'PROCOMPITE' or cui in set(cfg.get('anexo1_cuis', [])):
        return 'Restringido', disp, 'PROCOMPITE' if d['tipo'] == 'PROCOMPITE' else 'Lista Anexo I (recursos con destino específico)', pot
    if disp < cfg['monto_minimo']:
        if pot >= cfg['monto_minimo']:
            return 'Rebajando certificación', disp, f"Sin saldo libre; hay S/ {pot:,.0f} certificados sin comprometer ni programar en el 12-B (requiere rebajar la certificación)", pot
        return 'Sin saldo útil', disp, f"Saldo libre menor a S/ {cfg['monto_minimo']:,.0f}", pot
    cond = []
    if d['prog_ok'] and not d['en_pmi'] and d['libre'] - d['libre_estricto'] >= 1:
        cond.append(f"El 12-B programa S/ {d['prog']:,.0f} de un PIM de S/ {d['pim_cui']:,.0f}: la UEI debe actualizarlo (el proyecto no está en ningún PMI)")
    if d['pmi26'] > 0:
        cond.append('Está programado en el PMI 2026')
    elif d['ind_pmi'] == 'SI':
        cond.append('Figura en un PMI (' + (d['entidad'] or 'otra entidad') + ') sin monto 2026 en la Cartera del GR: validar con la OPMI')
    if d['pmi27'] > 0:
        cond.append(f"Tiene monto en el PMI 2027 (S/ {d['pmi27']:,.0f})")
    if not d['tiene_f12']:
        cond.append('Sin Formato 12-B: validar con la OPMI')
    elif not d['prog_ok']:
        cond.append('12-B sin programación vigente: validar con la OPMI')
    if not d['tiene_ssi'] and d['avance'] is None:
        cond.append('Sin ficha SSI: avance acumulado no verificable')
    if cond:
        return 'Con condiciones', disp, ' · '.join(cond), pot
    return 'Disponible', disp, 'No está en el PMI 2026 ni 2027; saldo libre según 12-B', pot
