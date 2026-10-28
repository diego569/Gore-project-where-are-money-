"""Primera propuesta de clasificación meta -> gerencia/oficina para el GASTO CORRIENTE de la UE 001 (para validar con un responsable del GORE).

No hay una fuente que relacione estas metas con oficinas (Silucia las trae todas como "GASTO CORRIENTE"), así que la propuesta se basa SOLO en:
  1. el nombre de la actividad y de la meta (a veces nombra al área: "DESARROLLO SOCIAL REGIONAL", "OREDIS"...),
  2. la clasificación que Alex ya hizo de los proyectos de la misma función (ej.: función 23 PROTECCION SOCIAL -> Desarrollo Social),
  3. el sentido estándar de las actividades del clasificador presupuestal (control, defensa judicial, planeamiento...).
Cada fila lleva su evidencia y un nivel de confianza. Las denominaciones de unidades que no figuran en el Excel de Alex ni en Silucia
se marcan "por confirmar": el validador debe corregirlas con el nombre oficial.
Uso: python propuesta_clasificacion_gc.py [PLIEGO.xls] [salida.xlsx]
"""
import datetime, pathlib, sys, collections
import xlsxwriter
from sinv.siaf import leer_siaf

RAIZ = pathlib.Path(__file__).resolve().parent
SIAF = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ.parent / 'ref docs' / 'nuevo' / 'PLIEGO--.xls'
OUT = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else RAIZ.parent / 'outputs' / 'Propuesta_clasificacion_gasto_corriente_UE001.xlsx'

GRI = 'GERENCIA DE INFRAESTRUCTURA'; GRDS = 'GERENCIA REGIONAL DE DESARROLLO SOCIAL'; GRDE = 'GERENCIA REGIONAL DE DESARROLLO ECONOMICO'
GRRN = 'GERENCIA REGIONAL DE RECURSOS NATURALES Y GESTION AMBIENTAL'
# unidades que ya existen en el Excel de Alex / Silucia (denominación firme)
EXISTENTES = {GRI, GRDS, GRDE, GRRN, 'ORSLP - Supervisión y liquidación de proyectos (ejecutora en Silucia)', 'OREDIS - Oficina Regional de Atención a la Persona con Discapacidad'}

# (clave en nombre de actividad/meta, función opcional, unidad, confianza, evidencia)
REGLAS = [
    (['OREDIS'], None, 'OREDIS - Oficina Regional de Atención a la Persona con Discapacidad', 'Alta', 'La descripción de la meta nombra a OREDIS.'),
    (['DESARROLLO SOCIAL REGIONAL'], None, GRDS, 'Alta', 'La actividad se llama "Desarrollo social regional".'),
    (['DESARROLLO ECONOMICO REGIONAL'], None, GRDE, 'Alta', 'La actividad se llama "Desarrollo económico regional".'),
    (['GESTION DE LA INFRAESTRUCTURA'], None, GRI, 'Alta', 'La actividad se llama "Gestión de la infraestructura".'),
    (['ASESORAMIENTO Y GESTION DEL MEDIO AMBIENTE', 'AREAS NATURALES'], None, GRRN, 'Alta', 'Actividad de medio ambiente; coincide con la gerencia de Recursos Naturales y Gestión Ambiental.'),
    (['SUPERVISION Y LIQUIDACION'], None, 'ORSLP - Supervisión y liquidación de proyectos (ejecutora en Silucia)', 'Alta', 'La actividad es la supervisión y liquidación de proyectos; Silucia tiene la ejecutora "ORSLP - LIQUIDACION".'),
    (['ACCIONES DE CONTROL Y AUDITORIA'], None, 'Órgano de Control Institucional', 'Alta', 'Actividad estándar de control gubernamental.'),
    (['DEFENSA JUDICIAL'], None, 'Procuraduría Pública Regional', 'Alta', 'Actividad estándar de defensa judicial del Estado.'),
    (['ASESORAMIENTO TECNICO Y JURIDICO'], None, 'Oficina Regional de Asesoría Jurídica', 'Media', 'Asesoramiento de naturaleza jurídica.'),
    (['PLANEAMIENTO Y PRESUPUESTO', 'DEMARCACION Y ORDENAMIENTO TERRITORIAL'], None, 'Gerencia Regional de Planeamiento, Presupuesto y Acondicionamiento Territorial', 'Media', 'Planeamiento/presupuesto y demarcación territorial.'),
    (['GESTION DE RECURSOS HUMANOS'], None, 'Oficina de Recursos Humanos', 'Media', 'Actividad estándar de recursos humanos.'),
    (['IMAGEN INSTITUCIONAL'], None, 'Oficina de Imagen Institucional', 'Media', 'La actividad se llama "Imagen institucional".'),
    (['CONDUCCION Y ORIENTACION SUPERIOR', 'DIRECCION Y GESTION REGIONAL'], None, 'Alta Dirección (Gobernación / Gerencia General)', 'Media', 'Conducción de la política institucional y dirección regional.'),
    (['SEGUIMIENTO DE FISCALIZACION DE LOS CONSEJOS REGIONALES'], None, 'Consejo Regional', 'Baja', 'La meta habla de fiscalización de consejos regionales: parece función del Consejo Regional.'),
    (['PRESERVACION Y DIFUSION DE ACERVOS'], None, 'Archivo Regional', 'Media', 'Conservación y catalogación de archivos y documentos.'),
    (['TRANSFERENCIA DE RECURSOS', 'TRANSFERENCIA FINANCIERA'], None, 'POR DEFINIR (transferencias a otras entidades)', 'Baja', 'Son transferencias financieras (suman ~S/ 28 M); no es evidente qué oficina las gestiona.'),
    (['OBLIGACIONES PREVISIONALES', 'PAGO DE PENSIONES'], None, 'POR DEFINIR (pensiones y cesantes)', 'Baja', 'Pago de pensiones; probablemente Recursos Humanos/Administración, sin evidencia.'),
    (['ADMINISTRACION DEUDA INTERNA', 'SERVICIO DE LA DEUDA'], None, 'POR DEFINIR (servicio de la deuda)', 'Baja', 'Servicio de la deuda; probablemente Administración/Tesorería, sin evidencia.'),
    (['EQUIPO MECANICO'], None, GRI, 'Media', 'Mantenimiento de equipo mecánico (maquinaria): Silucia tiene la ejecutora "GRI - MAQUINARIA"; confirmar si la maquinaria depende de Infraestructura.'),
    (['GESTION ADMINISTRATIVA', 'PAGO DE PLANILLAS', 'REPRO-AFP'], ('03',), 'Oficina Regional de Administración', 'Media', 'Gestión administrativa transversal de la sede (planillas, mantenimiento, servicios).'),
    (['GESTION ADMINISTRATIVA', 'GARANTIZAR LA APLICACION DE NORMAS LABORALES'], ('07',), 'Dirección Regional de Trabajo y Promoción del Empleo', 'Baja', 'Gestión administrativa dentro de la función Trabajo.'),
    (['GESTION ADMINISTRATIVA'], None, 'POR DEFINIR (gestión administrativa de un sector)', 'Baja', 'Gestión administrativa dentro de una función sectorial; debe pertenecer a la dirección sectorial correspondiente.'),
    (['GESTION DEL PROGRAMA', 'MATERIAS LABORALES', 'NORMATIVA LABORAL', 'FUERZA LABORAL', 'INTERMEDIACION LABORAL', 'TRABAJO FORZOSO'], ('07',), 'Dirección Regional de Trabajo y Promoción del Empleo', 'Media', 'Productos de la función Trabajo (consultas, formalización, intermediación laboral).'),
    (['RIESGO DE DESASTRES', 'EMERGENCIAS Y DESASTRES', 'KITS PARA LA ASISTENCIA', 'SIMULACROS', 'BRIGADAS', 'ZONAS CRITICAS Y FAJAS MARGINALES'], None, GRRN, 'Media',
     'Alex ubica en esta gerencia los proyectos de función 05 (orden público y seguridad: defensa ribereña) y 10 (agropecuaria).'),
    (['PROTECCION Y ATENCION', 'CUIDADO Y PROTECCION', 'MEJORAR LA CALIDAD DE VIDA', 'PERSONAS CON DISCAPACIDAD'], ('23',), GRDS, 'Media', 'Alex ubica en Desarrollo Social los proyectos de la función 23 (protección social); centros de acogida y discapacidad.'),
    (['MANTENIMIENTO DE INFRAESTRUCTURA'], None, GRI, 'Media', 'Mantenimiento de infraestructura; coincide con la Gerencia de Infraestructura.'),
    (['POTABILIZACION', 'PRESTADORES', 'AREAS TECNICAS MUNICIPALES', 'MONITOREO DE LA PRESTACION'], ('18',), 'Dirección Regional de Vivienda, Construcción y Saneamiento', 'Baja', 'Productos de saneamiento (agua y desinfección); no hay evidencia de qué unidad los ejecuta.'),
    (['HIERRO Y ACIDO FOLICO', 'CALIDAD DEL AGUA PARA EL CONSUMO', 'PROGRAMA ARTICULADO'], ('20',), 'Dirección Regional de Salud / Desarrollo Social', 'Baja', 'Productos de salud dentro de la sede central; no hay evidencia de la unidad.'),
    (['FORTALECIMIENTO DE LOS GOBIERNOS REGIONALES'], None, 'Gerencia Regional de Planeamiento, Presupuesto y Acondicionamiento Territorial', 'Baja', 'Fortalecimiento de funciones sectoriales transferidas; la actividad está bajo "Planeamiento y presupuesto".'),
]


def proponer(m, fn_nombre):
    txt = (m['nombre'] + ' | ' + m.get('meta_desc', '')).upper()
    for claves, fns, unidad, conf, ev in REGLAS:
        if fns and m['fn'] not in fns:
            continue
        if any(c in txt for c in claves):
            if unidad.startswith('POR DEFINIR (gestión administrativa'):
                ev += f" Función: {fn_nombre.get(m['fn'], m['fn'])}."
            return unidad, conf, ev
    return 'POR DEFINIR', 'Baja', 'Ninguna regla aplica: sin evidencia en el nombre de la actividad.'


def main():
    d = leer_siaf(SIAF)
    gc = sorted([m for m in d['metas'] if m['ue'] == '001' and m['tipo'] == 'actividad'], key=lambda m: (-m['pim'], m['meta']))
    FN = {'03': 'PLANEAMIENTO, GESTION Y RESERVA DE CONTINGENCIA', '05': 'ORDEN PUBLICO Y SEGURIDAD', '07': 'TRABAJO', '09': 'TURISMO', '10': 'AGROPECUARIA',
          '13': 'MINERIA', '15': 'TRANSPORTE', '17': 'AMBIENTE', '18': 'SANEAMIENTO', '20': 'SALUD', '23': 'PROTECCION SOCIAL', '24': 'PREVISION SOCIAL', '25': 'DEUDA PUBLICA'}
    filas = []
    for m in gc:
        u, c, ev = proponer(m, FN)
        filas.append(dict(m=m, unidad=u, conf=c, ev=ev, firme=u in EXISTENTES))
    wb = xlsxwriter.Workbook(str(OUT)); b = dict(font_name='Arial', font_size=9)
    F = lambda **k: wb.add_format({**b, **k})
    hd = F(bold=True, font_color='#FFFFFF', bg_color='#002060', text_wrap=True, align='center', valign='vcenter', border=1)
    txt = F(text_wrap=True, valign='top', border=1); num = F(num_format='#,##0', valign='top', border=1); title = F(bold=True, font_size=12); nota = F(italic=True, font_color='#595959', text_wrap=True)
    inp = F(bg_color='#FFF2CC', text_wrap=True, valign='top', border=1)
    colc = {'Alta': '#C6EFCE', 'Media': '#FFEB9C', 'Baja': '#F8CBAD'}

    ws0 = wb.add_worksheet('Cómo validar')
    ws0.set_column(0, 0, 130)
    for i, t in enumerate([
        ('PROPUESTA DE CLASIFICACIÓN DE LAS METAS DE GASTO CORRIENTE (UE 001) POR GERENCIA / OFICINA', title),
        ('Estado: BORRADOR para validar. Nada de esto viene de una fuente oficial: no existe hoy un archivo que relacione estas metas con oficinas (Silucia las trae todas como "GASTO CORRIENTE").', nota),
        ('', None),
        ('Cómo se hizo: se leyó el nombre de cada actividad y meta del SIAF y se aplicaron reglas (hoja "Reglas"). Cada fila muestra la evidencia y una confianza:', txt),
        ('   ALTA  = el nombre de la actividad nombra al área, o es una actividad estándar de un órgano concreto (control, defensa judicial).', txt),
        ('   MEDIA = coincide con lo que Alex ya clasificó para la misma función, o es una actividad típica de esa unidad.', txt),
        ('   BAJA  = suposición o "POR DEFINIR": hace falta que el GORE diga a qué oficina pertenece.', txt),
        ('', None),
        ('Qué debe hacer la persona que valida (hoja "Metas"):', F(bold=True)),
        ('   1. En la columna amarilla "Validación" elegir: Conforme / Corregir / No sé.', txt),
        ('   2. Si corrige, escribir en "Unidad correcta" el NOMBRE OFICIAL de la gerencia u oficina.', txt),
        ('   3. Las unidades marcadas "denominación por confirmar" no figuran en el Excel de Alex ni en Silucia: escribir el nombre oficial.', txt),
        ('   4. Prestar atención a las filas de mayor PIM (están ordenadas de mayor a menor) y a las "POR DEFINIR".', txt),
        ('', None),
        ('Después: el archivo validado se carga a maestro.xlsx (hoja ClasifGC) con  python importar_validacion_gc.py <archivo>.  Las metas nuevas del año siguiente se agregan solas.', txt)]):
        ws0.write(i, 0, t[0], t[1] or F())
    ws = wb.add_worksheet('Metas')
    H = ['N°', 'Meta', 'Función', 'Producto', 'Actividad (SIAF)', 'Descripción de la meta', 'PIM 2026', 'Unidad propuesta', 'Denominación por confirmar', 'Confianza', 'Evidencia', 'Validación', 'Unidad correcta', 'Comentarios']
    ws.set_row(0, 34)
    for j, h in enumerate(H): ws.write(0, j, h, hd)
    for j, w in enumerate([5, 7, 22, 10, 40, 48, 13, 46, 12, 10, 55, 14, 40, 30]): ws.set_column(j, j, w)
    for i, f in enumerate(filas, 1):
        m = f['m']
        ws.write_number(i, 0, i, txt); ws.write_string(i, 1, m['meta'], txt); ws.write_string(i, 2, f"{m['fn']}. {FN.get(m['fn'], '')}", txt); ws.write_string(i, 3, m['cui'], txt)
        ws.write_string(i, 4, f"{m['act']}: {m['nombre']}", txt); ws.write_string(i, 5, m['meta_desc'].split(';')[0], txt); ws.write_number(i, 6, m['pim'], num)
        ws.write_string(i, 7, f['unidad'], txt); ws.write_string(i, 8, 'No' if f['firme'] else 'Sí', txt)
        ws.write_string(i, 9, f['conf'], F(bg_color=colc[f['conf']], border=1, align='center', valign='top')); ws.write_string(i, 10, f['ev'], txt)
        ws.write_blank(i, 11, None, inp); ws.write_blank(i, 12, None, inp); ws.write_blank(i, 13, None, inp)
    ws.data_validation(1, 11, len(filas), 11, {'validate': 'list', 'source': ['Conforme', 'Corregir', 'No sé']})
    ws.freeze_panes(1, 2); ws.autofilter(0, 0, len(filas), len(H) - 1)
    wr = wb.add_worksheet('Resumen por unidad')
    por = collections.OrderedDict()
    for f in filas:
        x = por.setdefault(f['unidad'], dict(n=0, pim=0.0, conf=collections.Counter(), firme=f['firme']))
        x['n'] += 1; x['pim'] += f['m']['pim']; x['conf'][f['conf']] += 1
    for j, h in enumerate(['Unidad propuesta', 'Metas', 'PIM 2026', 'Alta', 'Media', 'Baja', 'Denominación por confirmar']): wr.write(0, j, h, hd)
    for i, (u, x) in enumerate(sorted(por.items(), key=lambda kv: -kv[1]['pim']), 1):
        wr.write_string(i, 0, u, txt); wr.write_number(i, 1, x['n'], txt); wr.write_number(i, 2, x['pim'], num)
        for j, k in enumerate(['Alta', 'Media', 'Baja']): wr.write_number(i, 3 + j, x['conf'][k], txt)
        wr.write_string(i, 6, 'No' if x['firme'] else 'Sí', txt)
    n = len(por) + 1
    wr.write(n, 0, 'TOTAL', F(bold=True)); wr.write_formula(n, 1, f'=SUM(B2:B{n})', F(bold=True), sum(x['n'] for x in por.values())); wr.write_formula(n, 2, f'=SUM(C2:C{n})', F(bold=True, num_format='#,##0'), sum(x['pim'] for x in por.values()))
    wr.set_column(0, 0, 70); wr.set_column(1, 6, 14)
    conf_pim = collections.Counter()
    for f in filas: conf_pim[f['conf']] += f['m']['pim']
    wr.write(n + 2, 0, 'PIM según confianza de la propuesta:', F(bold=True))
    for k, c in enumerate(['Alta', 'Media', 'Baja']):
        wr.write(n + 3 + k, 0, c, txt); wr.write_number(n + 3 + k, 2, conf_pim[c], num)
    rg = wb.add_worksheet('Reglas')
    for j, h in enumerate(['Palabras clave (en actividad o meta)', 'Función', 'Unidad propuesta', 'Confianza', 'Evidencia']): rg.write(0, j, h, hd)
    for i, (cl, fns, u, c, ev) in enumerate(REGLAS, 1):
        rg.write_string(i, 0, ' | '.join(cl), txt); rg.write_string(i, 1, ', '.join(fns) if fns else 'todas', txt); rg.write_string(i, 2, u, txt); rg.write_string(i, 3, c, txt); rg.write_string(i, 4, ev, txt)
    rg.set_column(0, 0, 60); rg.set_column(1, 1, 10); rg.set_column(2, 2, 60); rg.set_column(3, 3, 10); rg.set_column(4, 4, 90)
    wb.close()
    print(len(filas), 'metas ->', OUT)
    print({c: (sum(1 for f in filas if f['conf'] == c), round(conf_pim[c])) for c in ('Alta', 'Media', 'Baja')})
    print('POR DEFINIR:', [(f['m']['meta'], f['m']['nombre'][:40]) for f in filas if f['unidad'].startswith('POR DEFINIR')][:12])


if __name__ == '__main__':
    main()
