# AGENTS.md — SINVERSION (GORE Puno): "¿dónde hay plata que podemos usar?"

Guía para cualquier agente de IA (o persona) que tome este proyecto. Léela completa antes de tocar código.
Idioma de trabajo con el usuario: **español**, respuestas concisas, sin inventar datos.

## 1. Qué es y para qué sirve

Programa en Python (empaquetado como `actualizar.exe`) para el **Gobierno Regional de Puno (pliego 458)** que, con datos del MEF (SIAF / Banco de Inversiones),
genera **dos Excel** con los que la oficina decide **de qué proyectos se puede ceder presupuesto (anular/habilitar)** y de cuáles no:

1. `SINVERSION_PROYECTOS_AAAA-MM-DD.xlsx` — todas las UE con proyectos de inversión (clasificados por gerencia/oficina en el formato de Alex para la UE 001).
2. `SINVERSION_GASTO_CORRIENTE_AAAA-MM-DD.xlsx` — gasto corriente solo de la UE 001 (Sede Central), por gerencia/oficina y función.

Se actualiza solo (tarea programada de Windows, lunes a viernes 08:30 y 14:30), usando **APIs primero** y pidiendo archivos manuales solo si las APIs fallan.
Etapas siguientes previstas: Power BI y versión web/Docker (ver sección 12).

Vocabulario: **habilitador / cedente / anulador** = proyecto que cede presupuesto; **habilitado** = el que lo recibe. **PIM** = presupuesto modificado, **PIA** = inicial,
**CUI** = código único de inversión (7 dígitos), **meta** = secuencia funcional (4 dígitos), **UE** = unidad ejecutora, **UEI/UF/OPMI** = unidades del sistema Invierte.pe,
**12-B** = formato de seguimiento con la programación mensual de devengado, **PMI** = programa multianual de inversiones, **rubro** = fuente de financiamiento (00 RO, 09 RDR, 15 FONCOR, 18 CANON, 19 ROOC).

## 2. Mapa del repositorio

Raíz git: `E:\New folder\Economia\8vo\gore\Cash data request\` (hay carpetas ajenas: `vuelos_precios`, `Tutorial.mp4`, `*.rar` = respaldos viejos).

```
AGENTS.md / CLAUDE.md                 (este archivo; CLAUDE.md solo lo importa)
sinversion_auto/                      EL PROGRAMA (todo lo importante está aquí)
  actualizar.py                       orquestador (correr(), main()); se compila a actualizar.exe
  actualizar.exe                      ejecutable PyInstaller one-file --noconsole (HAY QUE RECOMPILAR tras cambiar código; ver 5)
  config.json                         parámetros (ver 7)
  maestro.xlsx                        clasificación de Alex: hojas Clasificacion, Complementos, Funciones, Listas, ClasifGC, MapaEjecutoras
  formato_oficina.json                orden de hojas, hojas ocultas, anchos y zoom que dejó la oficina (ver 9)
  token_silucia.txt                   SOLO un placeholder; nunca guardar tokens reales
  cache_*.json(.gz)                   cachés (ver 8); historial/ = fotos de cada corrida (60 últimas)
  entrada/  salida/AAAA-MM-DD/        archivos de respaldo manuales / Excel generados; salida/ULTIMO_RESULTADO.txt y ERROR.txt
  *.bat                               ACTUALIZAR_AHORA, ACTUALIZAR_SOLO_APIS, INSTALAR/DESINSTALAR_TAREA_*, RENOVAR_TOKEN_SILUCIA, GUARDAR_FORMATO
  LEEME.txt                           manual para usuarios finales (mantenerlo al día cuando cambien reglas o hojas)
  validar_listas.py                   compara el libro generado con el Excel del practicante (genera outputs/Validacion_vs_practicante_*.xlsx)
  sembrar_maestro.py, importar_validacion_gc.py, propuesta_clasificacion_gc.py, probar_api_*.py   utilitarios de una sola vez
  sinv/
    api_mef.py        cliente API Datos Abiertos MEF (ejecución del gasto): dump masivo → paginada; caché de última descarga buena
    siaf.py           lector del reporte SIAF .xls (respaldo); siaf_api.py arma lo mismo desde la API
    cartera.py / cartera_api.py   Cartera PMI (xlsx / API ofi5.mef.gob.pe)
    legal_apis.py     ficha SSI (tipo, estado, cierre, costo, PMI, entidad) y Formato 12-B por CUI, con caché
    silucia.py        API Silucia (ejecutora por meta) con caché de 90 días
    rubro.py          respaldo: reportes HTML-xls de Consulta Amigable por rubro
    maestro.py        lectura/escritura de maestro.xlsx y ubicación de cada meta en gerencia/grupo
    reglas.py         MOTOR DE REGLAS (semáforo, saldos, 90%, PMI, 12-B) — aquí viven las decisiones de negocio
    excel_proyectos.py  arma el libro de proyectos (xlsxwriter) — hojas, columnas, fórmulas
    excel_gasto.py    arma el libro de gasto corriente
    formato.py        aplica/guarda el formato de la oficina; historial.py = foto y comparación ("Cambios")
outputs/                              entregables y validaciones puntuales (Excel); no son parte del programa
ref docs/                             insumos de referencia (SIAF .xls, Cartera, Excel de Alex, ConsultaAvanzada, YOVELI GORE febrero 2026…)
work/                                 scripts viejos, respaldos de código (excel_proyectos_v*_backup.py) y lo puntual de Antony (work/antony_hoy/)
```

## 3. Cómo iniciar y operar

Windows 11. **El `python` del PATH es 2.7 (BioTime)**: usar siempre
`C:\Users\servidor\AppData\Local\Programs\Python\Python312\python.exe` (en Git Bash: `/c/Users/servidor/AppData/Local/Programs/Python/Python312/python.exe`).
Dependencias: `requests urllib3 openpyxl xlsxwriter xlrd pyinstaller` (`pip install ...`). LibreOffice (`C:\Program Files\LibreOffice\program\soffice.exe`) se usa para verificar fórmulas.

```bash
cd sinversion_auto
python actualizar.py                       # corrida normal (APIs; usa entrada\ si hace falta)
python actualizar.py --fecha 2026-10-09    # fuerza la fecha de corte; --salida RUTA / --entrada RUTA
python actualizar.py --guardar-formato "Excel_retocado.xlsx"   # guarda orden/ocultas/anchos de la oficina
./actualizar.exe --diario --aviso          # lo que ejecuta la tarea programada (muestra mensaje al terminar)
```
Variables de entorno de prueba: `SINV_FUENTE_SIAF=api|archivo|auto`, `SINV_API_OFF`, `SINV_SILUCIA_OFF`, `SINV_CARTERA_API_OFF`, `SINV_LEGAL_OFF` (desactiva SSI y 12-B).
Una corrida con APIs tarda 1–5 min (más si hay que rebajar fichas SSI: ~2 min). **Ejecuta comandos largos en segundo plano.**

Resultado: `salida\AAAA-MM-DD\` + `salida\ULTIMO_RESULTADO.txt` (o `ERROR.txt` en español; trazas en `logs\`).
**Cuidado:** si el Excel de destino ya existe, el programa lo **borra y lo reescribe** (si está abierto en Excel guarda otro nombre con la hora). Nunca edites a mano el libro de `salida\`
esperando conservarlo: lo que el usuario retoca se conserva vía `formato_oficina.json` (sección 9).

## 4. Flujo de `actualizar.correr()`

1. Lee `config.json`; decide fuente de cifras (`fuente_siaf`: `archivo` | `api` | `auto`).
2. **Cifras SIAF**: API Datos Abiertos (dump → paginada) → si falla, `.xls` en `entrada\` → si no hay, última descarga buena (`cache_api_filas.json.gz`, máx. `cache_cifras_dias`=3, franja roja
   "CIFRAS DESACTUALIZADAS") → si no hay nada, se detiene con error claro.
3. **Cartera PMI**: API (`traeListaCarteraSector`, sector 96, GR 458) → si falla, xlsx en `entrada\`.
4. **Silucia** (ejecutora por meta, para sugerir gerencia): en vivo si hay token vigente, si no la copia local (≤ 90 días); si no, solo el maestro.
5. **Datos legales por CUI** (`legal_apis.cargar`): ficha SSI (TTL 7 días) y 12-B (TTL 20 h, **una descarga masiva** `dump` con lista de CUI; por CUI individual tarda 7–17 s y es solo respaldo).
6. `excel_proyectos.armar_filas` → una fila por **UE-CUI-meta**, con gerencia (maestro/Silucia), `datos_cui()` + `evaluar()` (semáforo). Proyectos nuevos se agregan solos a `maestro.xlsx` (hoja Clasificacion) sin gerencia.
7. Rubro por UE-CUI: API (dump) → copia → Consulta Amigable (respaldo, solo a nivel pliego).
8. Gasto corriente UE 001: `excel_gasto.armar_filas_gc` (rubro por meta, clasificación de la hoja ClasifGC).
9. Historial: compara con la foto anterior (`historial.py`) → hoja "Cambios"; carga `formato_oficina.json`; escribe los dos Excel; guarda la foto; mueve los insumos usados a `entrada\procesados\fecha`.

## 5. Después de cambiar código (checklist obligatorio)

1. Respaldar el archivo que vas a tocar (`cp sinv/excel_proyectos.py ../work/excel_proyectos_vN_backup.py`). **Nunca hagas parches por rango de líneas/marcadores sin respaldo**: ya se borraron secciones enteras por error.
2. Correr `python actualizar.py` y mirar `ULTIMO_RESULTADO.txt` / `ERROR.txt`.
3. **Recalcular con LibreOffice** y comparar valores en caché vs recalculados (debe dar 0 diferencias): convertir el xlsx con
   `soffice --headless --convert-to xlsx --outdir out archivo.xlsx` y comparar con openpyxl `data_only=True` celda a celda. La hoja `Control` debe decir "Conciliado" en todo.
4. **Recompilar el exe** (si no, la tarea diaria sigue con el código viejo):
   `python -m PyInstaller --onefile --noconsole --name actualizar --distpath <dir>/dist --workpath <dir>/work --specpath <dir>/spec actualizar.py` y copiar `dist\actualizar.exe` a `sinversion_auto\`.
5. Si cambiaste reglas, hojas o parámetros: actualizar `LEEME.txt` y el texto de la hoja `Reglas` (está escrito a mano en `excel_proyectos.py`, no se genera de las reglas).
6. Probar el `.exe` (`./actualizar.exe --fecha AAAA-MM-DD`).

## 6. Fuentes de datos y particularidades verificadas

| Dato | Fuente | Notas |
|---|---|---|
| PIM, PIA, certificado, comprometido anual, devengado, rubro | API Datos Abiertos MEF `api.datosabiertos.mef.gob.pe/DatosAbiertos/v1` (recurso `615644aa-ef73-4358-b4e0-0c20931632f3`) | Registros **mensuales**: PIA/PIM solo en `MES_EJE=0`; lo demás son flujos que se suman. `SEC_FUNC` sin ceros (usar `zfill(4)`). `dump?id=&filters=` es asíncrono ("Creando consulta" → CSV; una tabla vacía devuelve CSV de solo encabezados). Solo `datastore_search` con filtros exactos (no hay SQL). El certificado SSL a veces está vencido (se reintenta con `verify=False` y se avisa). La API se cae/queda vacía por ratos (no es bloqueo). |
| Tipo de inversión, estado, cierre, costo actualizado, indicador PMI, entidad formuladora | Ficha SSI: `POST https://ofi5.mef.gob.pe/invierteWS/Ssi/traeDetInvSSI` (`id`=CUI, `tipo`=1) | Necesita User-Agent de navegador, `X-Requested-With` y `Referer`. `DEV_ACUM_ANT`/`DEV_ANO_VIGENTE` vienen en 0: no usarlos. Cachea solo campos elegidos (`SSI_CAMPOS`). |
| Programación 2026 por mes y devengado acumulado | Dataset 12-B Datos Abiertos, recurso `c275fa9f-5c61-4313-828d-0827277bdd97` (`CODIGO_UNICO`) | `MONTO_ACTUALIZADO_1..12`: ene→fecha = devengado real, resto = programado. `dump` acepta lista de CUI. ~80 CUI no tienen registro. |
| Cartera PMI (montos 2026-2029, costo, devengado acumulado al cierre del año anterior) | `POST https://ofi5.mef.gob.pe/invierte/Pmi/traeListaCarteraSector` (ddlSector=96, ddlGobiernoRegional=458) | Solo trae el PMI **del GR**. Cert SSL falla en esta PC (`verify=False`). |
| Gerencia/oficina | `maestro.xlsx` (de Alex, 16/09) + Silucia `sistemas.regionpuno.gob.pe/siluciav2-api/api/metas` | Bearer JWT de ~10 h; requiere User-Agent de navegador; las metas de gasto corriente llegan todas como "GASTO CORRIENTE". |
| Respaldo manual | `.xls` del SIAF, Cartera `.xlsx`, Consulta Amigable por rubro | **No** se raspa Consulta Amigable ni se evaden CAPTCHAs/bot-protection. |

**Seguridad:** nunca guardes tokens, contraseñas ni claves en archivos del proyecto ni en commits. `token_silucia.txt` debe contener solo el placeholder (el lector exige que un token empiece con `eyJ`).
Si el usuario pega un token "para probar", úsalo en memoria y no lo persistas.

## 7. Reglas de negocio vigentes (`sinv/reglas.py`)

Base legal consultada: **Art. 11 de la Ley 32513 (Presupuesto 2026)** — para anular/ceder en inversiones: avance financiero acumulado no mayor a 90 %, sin deudas contractuales
devengables en el año, no fondos de transferencias condicionadas; el num. 11.2 exige *impedimento comprobable*. El **usuario confirmó que la regla del 90 % es correcta**.
Lo que NO se puede verificar con datos públicos (validar con la OPMI): deudas contractuales por devengar, transferencias condicionadas, lista del **Anexo I** (recursos con destino específico).

**Parámetros (`config.json`)**: `monto_minimo` 10000 · `umbral_ejecutado` 0.90 · `umbral_avance_acumulado` 0.90 · `anexo1_cuis` [] (o archivo `anexo1_cuis.txt` junto al exe, un CUI por línea) ·
`usar_ssi`/`usar_f12b` · `usar_api`, `fuente_siaf`, `usar_api_cartera`, `usar_silucia`, `anio` 2026, `ue_gasto_corriente` ["001"], `cache_cifras_dias` 3, `silucia_cache_dias` 90, `api_sin_verificar_certificado`.

**Cálculo por CUI (`datos_cui`)** — se evalúa sobre todo el CUI (todas sus UE/metas) y se reparte entre sus filas proporcionalmente al PIM:
- `avance` = devengado acumulado / costo actualizado. Devengado acumulado: 12-B → si no, SSI → si no, Cartera + SIAF 2026; nunca menor al devengado 2026. Costo: SSI → Cartera.
- `libre` (saldo libre) = `min(PIM − Σ12-B, PIM − certificado)` si hay 12-B vigente; si no, `PIM − certificado`.
- Si el proyecto **no está en ningún PMI** (sin monto 2026 en la Cartera **y** `IND_REG_PMI` de la ficha SSI ≠ "SI"), el 12-B no frena: `libre = PIM − certificado` y se agrega la condición "la UEI debe actualizar el 12-B".
- `pot` (potencial rebajando certificación) = `min(PIM − Σ12-B, PIM − comprometido) − libre` (lo certificado que no está comprometido ni programado).

**Semáforo (`evaluar`, en este orden)**: Sin PIM → Bolsa genérica (códigos `2001621`, `2016766`) → **No mover** (devengado 2026/PIM ≥ `umbral_ejecutado` **o** avance acumulado > `umbral_avance_acumulado`)
→ **Restringido** (tipo PROCOMPITE o CUI en `anexo1_cuis`) → **Sin saldo útil** (saldo libre < `monto_minimo`; si el potencial ≥ monto mínimo: **Rebajando certificación**) →
**Con condiciones** (en PMI 2026, monto 2027 en PMI, sin 12-B o 12-B sin programación, 12-B que debe actualizar la UEI, figura en PMI de otra entidad, sin ficha SSI) → **Disponible**.
Tipo de inversión (proyecto / IOARR / PROCOMPITE / genérico) es **informativo**, salvo PROCOMPITE = Restringido.
**Gasto corriente** usa `clasificar()` (estados por saldos: Sin PIM · Ejecutado (dev ≥ umbral_ejecutado) · Con compromiso · Solo certificado · Sin certificar); **no existe** el antiguo 50 % de compromiso (el usuario lo eliminó).

**Hallazgos de dominio (no están en el código, importan para decidir):**
- El criterio de Antony/practicante reproduce exacto: certificación 0 + `IND_REG_PMI = NO` + no PROCOMPITE + no genérico (sus 16 proyectos). Antony ignora proyectos que están en PMI (incluidos PMI de otras entidades).
- Un proyecto puede figurar en el PMI de una **municipalidad/ministerio** (ficha SSI `NIVEL` GL/GN) aunque la UEI del GR lo ejecute; la Cartera del GR no lo trae. Es una incógnita si esos fondos (transferencias) pueden anularse: **pendiente de la OPMI**.
- Evidencia del Excel de febrero (`ref docs/YOVELI GORE- de META FEBRERO 2026.xlsx`, hoja HABILITADORES = "posibles habilitadores para GRDS", `DATA` = PIM de febrero): 94 proyectos con PMI 2026 cedieron ≥ 50 % de su PIM
  (S/ 171 M). O sea, **estar en el PMI no impide ceder**; hoy el programa lo trata como bloqueo solo cuando además el 12-B programa todo el PIM. Decisión **pendiente** (preguntar a la OPMI si el 12-B vigente limita la anulación).
- El `proyecto genérico` 2658446 ("Proyectos para el desarrollo local") se parece a una bolsa; hoy sale "Con condiciones". Pendiente de confirmar con la OPMI.
- Listas de comparación: practicante `C:\Users\servidor\Downloads\Proyectos Gore-Puno Completo-OFICIAL.xlsx` (hoja RESUMEN: A = NO-PMI sin movimiento, G = NO-PMI con movimiento, O = PMI sin cert/comp, R = PMI solo certificaron).

## 8. Cachés y respaldos (todos en `sinversion_auto/`)

`cache_api_filas.json.gz` (última descarga buena de la API, proy + gasto corriente) · `cache_api_rubro.json` (rubro, 7 días) · `cache_rubro.json` (Consulta Amigable) ·
`cache_silucia.json` (90 días) · `cache_ssi.json.gz` (7 días, por CUI) · `cache_f12b.json.gz` (20 h, por CUI) · `historial/base_AAAAMMDD_HHMM.csv.gz` (60 últimas).
Si cambias qué campos se guardan de SSI/12-B, **borra el caché correspondiente** (si no, las entradas viejas no tendrán los campos nuevos hasta vencer el TTL).
Principio: si una API falla se usa lo guardado y **se avisa** (hojas Config/Control/Resumen y `ULTIMO_RESULTADO.txt`); nunca se inventa un dato (queda `s/d`).

## 9. Estructura del Excel de proyectos (`excel_proyectos.construir`)

Hojas (orden por defecto; la oficina las reordena/oculta con `formato_oficina.json`): Resumen · Base · Dónde hay plata · Cambios · Gerencia de Infraestructura · Gerencias y Oficinas · Sin clasificar ·
Otras UE (una sola hoja para las demás UE) · Saldo al expediente · Rubro por CUI · Calidad de datos · Control · Reglas · Config.

- **Base**: tabla filtrable, una fila por UE-CUI-meta. Columnas en `BASE_KEYS`/`BASE_COLS` (mismo orden; `L[...]` da la letra de columna que usan las fórmulas de otras hojas). Los resultados de las reglas se escriben como **valores** (no fórmulas); solo algunos derivados son fórmulas.
  **Nombres de proyecto = texto fijo** (no fórmulas): Ctrl+B de Excel busca en fórmulas y no encontraba el CUI (causó pérdida de confianza). Mantener texto fijo en toda hoja donde se busque por CUI.
- **Dónde hay plata**: verde/amarillo/azul con saldo ≥ `monto_minimo`, ordenado; columnas por **título** (`HDH`), incluye rubros (PIM por fuente repartido entre metas) y "Rubro principal".
- **Anexos de Alex** (Infraestructura, Gerencias y Oficinas, Sin clasificar, Otras UE): árbol con subtotales; las celdas leen de Base por fórmula (con valor en caché ya calculado: `write_formula(..., valor)`).
- Todas las fórmulas llevan **valor en caché** para verse sin recalcular; por eso el paso 3 de la sección 5 es obligatorio.
- **Formato de la oficina**: la oficina retoca un Excel generado (orden, ocultas, anchos, zoom) y se guarda con `--guardar-formato` / `GUARDAR_FORMATO.bat`. `formato.aplicar` lo reproduce en cada corrida: hojas con tabla por título de columna,
  el resto por posición. Los anchos de xlsxwriter llevan +0,71 de relleno (ya compensado). No se copian paneles inmovilizados.
- **Gancho para versiones puntuales**: `construir()` acepta `meta_info['antony']` (lista de CUI resaltados + hoja "Lista Antony"). **Está inactivo en el programa diario por decisión del usuario**
  (lo de Antony es solo puntual). Se activa únicamente con `work/antony_hoy/generar_version_antony.py`, que no toca `salida/`, el historial ni el maestro. No lo conectes al flujo diario.

## 10. Receta: cambios frecuentes

**Quitar (desactivar) el filtro del 90 %** — es la regla "No mover" de `evaluar()` en `sinv/reglas.py` (dos condiciones: devengado 2026/PIM y avance acumulado). Opciones, de menor a mayor impacto:
1. *Sin tocar código*: en `config.json` poner `"umbral_avance_acumulado": 2` y/o `"umbral_ejecutado": 2` (valores > 1 nunca se alcanzan). Ojo: `umbral_ejecutado` también gobierna el estado "Ejecutado" del **gasto corriente** (`clasificar()` y la fórmula de `excel_gasto.py`).
2. *Limpio (recomendado si es definitivo)*: borrar las dos condiciones `No mover` en `reglas.evaluar` (líneas con `umbral_ejecutado` y `umbral_avance_acumulado`) y, si ya no se usa, el semáforo `No mover` en `SEMAFOROS`, `colores` (`excel_proyectos.construir`) y las columnas "PIM · No mover" del Resumen (`SEMS`/`heads`/`tabla`).
3. En ambos casos actualizar los textos fijos que dicen 90 %: hoja `Reglas` y hoja `Config` (`excel_proyectos.py`), nota de `excel_gasto.py`, `LEEME.txt`, y `validar_listas.py`. Recompilar el exe (sección 5).
   (Antes de hacerlo, recordar que el usuario confirmó que la regla es correcta; pedir confirmación explícita si no queda claro que quiere eliminarla.)

**Cambiar el monto mínimo** → `config.json: monto_minimo` (se aplica en la siguiente corrida; el Excel solo lo muestra).
**Tratar el PMI solo como condición (no bloqueo)** → en `reglas.datos_cui`, la excepción `if prog_ok and not en_pmi: libre = sin_cert` pasaría a aplicarse siempre (o quitar `en_pmi` de esa condición); la condición "Está programado en el PMI 2026" ya se muestra como amarillo en `evaluar`.
**Tratar los proyectos genéricos como bolsas** → añadir su CUI a `bolsas_genericas` (config/DEFAULT) o, mejor, marcar `d['tipo'] == 'Genérico'` en `evaluar` antes del bloque "No mover".
**Cargar el Anexo I** → crear `anexo1_cuis.txt` junto al exe (un CUI por línea); `actualizar.py` lo mezcla con `anexo1_cuis` de la config; pasan a "Restringido".
**Agregar una columna a Base** → añadir la clave en `BASE_KEYS` y el título en `BASE_COLS` (mismo orden), escribirla en el bucle de Base (`C['clave']`), ajustar el ancho; si debe verse en Dónde hay plata, añadirla a `HDH`, a `ANCHO_DH` y al diccionario `v` de la fila.
**Agregar un semáforo** → `SEMAFOROS` y `evaluar` en `reglas.py`; color en `colores`; columnas del Resumen; hoja Reglas; LEEME.
**Agregar un dato de SSI/12-B** → `SSI_CAMPOS`/`F12_CAMPOS` en `legal_apis.py`, usarlo en `datos_cui`, **borrar el caché** (sección 8), probar.
**Cambiar horarios de la tarea** → `INSTALAR_TAREA_DIARIA.bat` (schtasks: lunes a viernes 08:30 y 14:30). La tarea debe correr en una PC encendida; el exe solo se probó en esta PC.
**Nueva clasificación de gasto corriente** → completar la hoja `ClasifGC` del maestro (propuesta pendiente de validar en `outputs/Propuesta_clasificacion_gasto_corriente_UE001.xlsx`; cargar con `python importar_validacion_gc.py <archivo>`).
**Comparar con otra lista** → `python validar_listas.py "<practicante.xlsx>"` o un script puntual en `work/` (como `work/antony_hoy/`), sin acoplarlo al programa.

## 11. Trampas conocidas

- Git Bash/heredocs: las secuencias `\b`, `\d`, `\n` dentro de cadenas Python escritas desde el shell se corrompen (una vez `\b` se volvió retroceso y rompió una regex). Para parches grandes **escribe un archivo `.py`** y ejecútalo; verifica con `grep`.
- Archivos de Excel abiertos bloquean la escritura (`PermissionError`); el programa guarda con otro nombre y avisa.
- `xlrd` y `openpyxl` mantienen bloqueado el archivo si no se cierran (problema de mover a `procesados`).
- Los reportes SIAF/Consulta Amigable traen texto con **mojibake** (`arreglar_texto` en `siaf.py`); el encabezado de UE se repite por página (no reiniciar el estado del parser).
- Un mismo CUI puede estar en varias UE/metas: las cifras legales (12-B, avance) son por CUI y se reparten por PIM; no sumes el costo entre UE.
- Códigos con todo en 0 que devuelve la API (6 el 09/10) se excluyen; las bolsas `2001621` y `2016766` no son inversiones.
- `MODALIDAD` se deriva de los clasificadores del SIAF (ver hoja Reglas); "Personal/bienes/servicios" como administración directa está **por confirmar** con Alex.
- No uses `rm -rf` con variables (el entorno lo bloquea); borra archivos concretos.

## 12. Estado y pendientes (al 09/10/2026)

Hecho: pipeline con APIs y respaldos; reglas por saldos; SSI + 12-B; hojas Dónde hay plata, Cambios, Calidad de datos, Lista (puntual) de Antony; formato de oficina; validación contra el practicante; exe recompilado.
Pendiente / abierto:
1. **OPMI**: lista del Anexo I; si el 12-B vigente limita la anulación; si el PMI de otra entidad (municipal/sectorial) impide ceder; si el código genérico se puede redistribuir; fondos de transferencias condicionadas.
2. Decidir el tratamiento de PMI vs 12-B (sección 7) — el usuario dijo "déjalo como está" hasta tener respuesta.
3. Validación de la propuesta de clasificación del gasto corriente (80 metas aún "sin oficina").
4. Instalar la tarea programada en la PC definitiva y probar el exe allí; asignar los 2 proyectos sin gerencia en `maestro.xlsx`.
5. **Iteración 4**: modelo y tablero de Power BI (base de proyectos + historial diario) y versión web/Docker (hay servidor propio, p. ej. dhyrium.online).
6. Opcional: User-Agent descriptivo con correo de contacto para la API del MEF (el usuario no ha dado el correo).

## 13. Cómo trabaja el usuario (respétalo)

- Quiere **planear antes de construir** en cambios grandes y confirma por iteraciones; pregunta antes de decisiones de reglas que cambian resultados.
- Pide **no inventar datos**: lo que falta se muestra como `s/d` y se explica.
- Lo puntual (listas de Antony, validaciones) va **fuera del programa** (`work/`, `outputs/`); el programa diario debe quedar simple y estable.
- Los usuarios finales de la oficina no son técnicos: todo mensaje al usuario final en español claro; el Excel debe permitir buscar y filtrar con facilidad.
- Prefiere explicaciones concisas y con cifras comprobables (CUI concretos, montos, fuente de cada dato).
