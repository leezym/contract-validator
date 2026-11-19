# contract-validator

Valida archivos `.json`, `.xml` y `.pdf` contra el contrato de `data/specs/contrato_ejemplo.xlsx`.

Guía de operación paso a paso: [`docs/INSTRUCTIVO.md`](docs/INSTRUCTIVO.md).

**Uso normal (todo en uno):** se entregan los PDF, XML y JSON y la automatización valida la estructura contra el
`.xlsx`, consulta la BD Oracle (solo lectura) y compara el contenido, con un veredicto por documento
(`Estructura: CUMPLE/NO CUMPLE`, `Información vs BD: CORRECTA/INCORRECTA/NO VERIFICADA`, `RESULTADO`):

```powershell
$env:PYTHONPATH="src"   # servidor en config/oracle.local.json; usuario y contraseña en .env (CV_ORACLE_USER / CV_ORACLE_PASSWORD)
python -m contract_validator verificar entrada/ --out reports/verificacion.xlsx
```

**Principio:** solo se valida lo que definen los specs entregados (el libro Excel y los ejemplos de `data/specs/`).
No se agregan reglas por criterio propio; si algo no está en los specs, no se valida y se reporta como hallazgo de contrato.

## Flujo

```bash
pip install -r requirements.txt

# 1) libro original -> contrato plano (el original no se modifica)
PYTHONPATH=src python -m contract_validator build --source data/specs/contrato_ejemplo.xlsx --out data/specs/contrato_plano.xlsx

# 2) validar (tipo de documento por contenido; si no, por nombre de archivo; o --doc-type)
PYTHONPATH=src python -m contract_validator data/specs/pbs-notas-debito.json data/specs/53_16_01_001_131126_020656_*.xml --out reports/resultado.xlsx

# 3) auditar el mapeo JSON <-> XML_TQ del contrato (sin archivos reales)
PYTHONPATH=src python -m contract_validator audit-mapping [--doc-type FAC] [--out reports/auditoria_mapeo.xlsx]

# 4) auditar que el contrato plano sea fiel al libro original (lector independiente)
PYTHONPATH=src python -m contract_validator audit-plano [--fail-on-error] [--out reports/auditoria_plano.xlsx]

# tests
PYTHONPATH=src pytest -q
```

En PowerShell: `$env:PYTHONPATH="src"` antes de ejecutar.

## Hoja `contrato` (generada)

`tipo_doc | formato | ruta | tipo | valores | requerido | cambio | ruta_anterior | ruta_contraparte | descripcion | hoja_origen | fila_origen | nota`

- `tipo_doc`: CCF (03), FAC (01), FAC_EXP (11), NCR (05), NDB (06), REM (04), INVALIDACION, RETORNO.
- **Fuente de los tags (confirmada por el cliente): hojas `XML` y `JSON` del libro.** En la hoja `XML` cada tipo tiene su
  columna (`CCF - 03`, `F -01`, `FE - 11`, `NC - 05`, `ND - 06`, `NOTA REMISION - 04`, `EVENTO RETORNO`) con la ruta JSON
  v4; su tag XML es el `XML_TQ` de la misma fila (RETORNO: columna `XML-ER`; sus Tipo/Requerido salen de la fila cuyo
  `TAG_MIN_HACIENDA` coincide). La hoja `JSON` es la base v3 por tipo y solo se usa para verificar (diferencias → `lint`).
  `EVENTO INVALIDACION` no tiene tags XML en la hoja `XML`: ese tipo se sigue leyendo de su hoja `Invalidacion`.
  Las hojas por tipo solo aportan la `Descripcion` cuando la hoja `XML` la deja vacía (etiqueta para el PDF).
- `formato`: JSON = ruta de la columna del tipo; XML = `XML_TQ` (`DatosEntrada/...`, `CabeceraEntrada/...`).
- `cambio`: del texto `Adicionar` / `Eliminar` / `Renombrar A -> B` de la columna del tipo; las filas XML quedan `SIN_CAMBIO`
  (la hoja `XML` no trae `Accion`).
- `requerido`: SI / NO / SI_NULL desde la columna `Requerido` de la hoja `XML` (o `Obligatorio` en Invalidacion).
  `Si/null/ por tipoDTE` y `Si (null)` = requerido y admite null, **solo para los tipos cuya columna de la hoja `XML` marca el campo**.
- `valores`: catálogo cuando la columna Tipo del libro es una lista (`00, 01`).
- `ruta_contraparte`: tag equivalente en el otro formato; `nota = posicional` cuando el libro indica posición (`OtrosTotales(4)`, `[2]`).
- Rutas con `/`; el segmento `items` y `[]` se omiten. En XML la ruta unificada es `DatosEntrada/InformacioGeneral/...` (sin «n»): el libro ya está escrito así (2026-11-18); si aparece una variante, `build` la unifica y deja un aviso en `lint`.
- La hoja `Validar` se ignora. La hoja `lint` lista lo que el libro deja ambiguo o se omitió (con la etiqueta `XML [TIPO]`,
  porque una fila de la hoja `XML` alimenta a varios tipos). Los ejemplos PBS `data/specs/pbs-*.json` siguen las hojas
  por tipo antiguas: sus diferencias con la hoja `XML` están documentadas en `tests/test_validator.py` (`SHEET_VS_GOLDEN`;
  decisión 2026-11-17: manda la hoja `XML`).

## Qué se valida

**Estructura y contenido contra el contrato** (comparador):

| Código | Significado |
|---|---|
| `NEW_MISSING`, `REQUIRED_MISSING`, `RENAMED_MISSING` | campo requerido ausente (no se exige el contenido de un contenedor `null`, vacío u opcional ausente). En XML también: cada **posición** que indica el libro (`valorOtrosTotales(4)` → `...[4]`) y cada tag de ítem en **todos** los ítems |
| `XML_OPTIONAL_MISSING` (aviso) | tag XML del libro que no llegó y cuyo `Requerido` es `No`, `N/A` o vacío: no es error, pero queda a la vista |
| `NULL_NOT_ALLOWED` | requerido `Si` con valor nulo (`Si (null)` y `Si/null/ por tipoDTE` lo admiten) |
| `TYPE_MISMATCH` | el valor no tiene el tipo de la columna `Tipo` del libro |
| `REMOVED_STILL_PRESENT`, `OLD_NAME_PRESENT` | campo eliminado o nombre anterior aún presente |
| `UNEXPECTED_FIELD` | campo presente y no definido en el contrato (`CabeceraEntrada` está definida en la hoja `XML`; `Solo XML` marca tags sin campo JSON) |
| `VALUE_NOT_IN_ENUM` (aviso) | valor fuera de un catálogo. **Hoy inactivo:** la regla `catalogos_como_string` genera esos campos como `string`, sin lista |
| `FORMAT_INVALID`, `CONSISTENCY_ERROR` | número de control: `DTE-##-<bloque>-<15 dígitos>` y `##` = tipo de documento (única regla de contenido que define el Excel; la longitud del bloque central no se valida porque el Excel y los ejemplos difieren) |
| `DOCTYPE_UNKNOWN`, `DOCTYPE_NAME_MISMATCH` (aviso), `ENCODING` (aviso), `PARSE_ERROR` | tipo no determinado, nombre de archivo ≠ contenido, XML no UTF-8 leído como cp1252, archivo ilegible |

**No se validan:** valores fijos o formatos de las notas del libro, formatos de fecha/hora, NIT/DUI, teléfonos, correos, sumas y totales, ni equivalencias semánticas:
los specs no los definen. XML: se retira la envoltura SOAP, `null` y vacío = sin valor.

## Cruce XML <-> JSON (`crosscheck.py`)

Compara el XML y el JSON **del mismo documento** con el mapeo del propio Excel (`ruta_contraparte`). Igualdad **exacta** de texto; solo `null`/vacío = sin valor.

```bash
PYTHONPATH=src python -m contract_validator doc.xml doc.json            # emparejamiento por UUID
PYTHONPATH=src python -m contract_validator doc.xml doc.json --pair doc.xml doc.json
# --no-cross lo desactiva
```

| Código | Significado |
|---|---|
| `XREF_MISMATCH` | valor distinto |
| `XREF_FORMAT_DIFF` (aviso) | mismo valor con distinto formato (`05` vs `5`, `DD/MM/YYYY` vs ISO, mayúsculas); lo decide el equipo. Los números con el mismo valor decimal (`285.6` vs `285.60`) son iguales por la regla `decimales_equivalentes` y no generan hallazgo |
| `XREF_MISSING` | el valor existe en un lado y falta o es nulo en el otro |
| `XREF_COUNT_MISMATCH` | distinta cantidad de valores presentes, sin contar `null` (`Items` vs `cuerpoDocumento`) |
| `XREF_DOCTYPE` | tipos de documento distintos |

Solo se omiten los mapeos que el propio libro marca como posicionales. Una diferencia puede ser del documento **o del mapeo del Excel**.
Aún no hay pares reales (la solución no está implementada): el cruce está probado con un par sintético construido desde el mapeo.

## PDF vs XML (`pdf_text.py`, `pdf_check.py`)

Valida la información impresa en los PDF contra el XML del mismo documento. Hay dos PDF por documento, pero **solo se validan los que dicen MINH** en el nombre (`CCF_SAL_MINH_...`, el oficial de Hacienda con Sello de Recepción).
El interno de TQ (`CCF_SAL_<fecha>_...`) se ignora por completo (también el vacío `NCR_SAL_2026...pdf`). Requiere `pdftotext` (Xpdf/Poppler) en el PATH o `PDFTOTEXT=<ruta>`; no hay dependencias de Python nuevas.

```bash
PYTHONPATH=src python -m contract_validator data/specs/*.pdf data/specs/*.xml --out reports/resultado.xlsx   # empareja por Código de Generación
PYTHONPATH=src python -m contract_validator doc.pdf doc.xml --pair-pdf doc.xml doc.pdf                         # explícito
# --no-pdf ignora los PDF;  --pdf-config otra.json  usa otras normalizaciones
```

No existe un mapeo PDF↔XML en el Excel, así que la estrategia son 3 capas apoyadas en los specs:

1. **Identidad:** Código de Generación (= `CabeceraEntrada/uuid`), Número de Control (= `secuenciaDocumento`) y título del documento (= tipo del XML).
2. **Valores con ancla:** etiquetas del PDF que coinciden **exacto** con una `Descripcion` del Excel (p.ej. `Modelo de Facturación`, `Tipo de Transmisión`; las de ítem y `Notas` se excluyen por ambiguas);
   pares etiqueta/valor de `Notas` del propio XML; filas de ítems del PDF (código, cantidad, precio, descuento) contra `Items[i]` y cantidad de filas.
3. **Contención:** cada valor del XML debe aparecer en el PDF y cada importe del PDF debe existir en el XML.

| Código | Nivel | Significado |
|---|---|---|
| `PDF_EMPTY`, `PDF_TOOL_MISSING` | ERROR | PDF MINH sin contenido / falta `pdftotext` |
| `PDF_IDENTITY`, `PDF_DOCTYPE` | ERROR | uuid o número de control distintos; título del PDF ≠ tipo del XML |
| `PDF_MISMATCH` | ERROR | valor anclado distinto (etiqueta del Excel, catálogo, fila de ítem) |
| `PDF_VALUE_MISSING` | ERROR | valor de `Notas` o código de ítem que el XML declara y no aparece en el PDF |
| `PDF_COUNT_MISMATCH` | ERROR | distinta cantidad de filas de ítems (PDF MINH) |
| `PDF_NO_PAIR` | WARNING (ERROR con `--pair-pdf`) | no hay XML con ese Código de Generación |
| `PDF_NOT_PRINTED` | WARNING | valor del XML que no aparece en el PDF; si el PDF no lo imprime, agréguelo a `no_impreso` |
| `PDF_VALUE_NOT_IN_XML` | WARNING | importe del PDF que no existe entre los valores del XML |
| `PDF_IGNORED` | WARNING | se pasó explícitamente (`--pair-pdf`) un PDF interno: se ignora |

**Presentación (`data/specs/pdf_presentacion.json`, PROPUESTA INICIAL pendiente de aprobación):** lista explícita y editable de lo único que se acepta como igual
(fechas `DD/MM/YYYY`↔ISO, separador de miles y ceros decimales, guiones/espacios en NIT/NRC/teléfono), los catálogos tomados del Excel (`versionFormato`: 1=Previo; `tipoOperacion`: 1=Normal),
y los tags `no_impreso`. Todo lo demás es `PDF_MISMATCH`. Los acentos y mayúsculas son exactos salvo que se active en el archivo.

Estado de las pruebas: hay dos pares reales. `CCF_SAL_MINH_*` ↔ `53_16_01_001_131126_020656_*.xml` (en `data/specs`, entregado por el usuario): 0 errores y 3 avisos;
`NCR_SAL_MINH_*` ↔ `tests/fixtures/tq-notas-credito.xml` (reconstruido del XML NCR anterior): 0 errores. Los PDF MINH de EXP y REM no tienen XML (salen como `PDF_NO_PAIR`).
No hay PDF de los demás conceptos (FAC, ND, retorno, invalidación): no se exige; un XML sin PDF no produce ningún hallazgo de PDF.
El XML `53_16_01_105_009867_000002_*.xml` es un **evento de retorno** (no trae `tipoFactura`; el tipo se toma de la operación SOAP `EventoRetornoEntrada`).
Los PDF MINH de CCF, EXP y REM no tienen XML en `data/specs` (salen como `PDF_NO_PAIR`).

## `audit-mapping`

Reporta hechos del contrato, sin juicios: `NO_XML_SOURCE` (campo JSON requerido sin tag XML; ERROR si no admite null, aviso si lo admite),
`MULTI_SOURCE` (un tag XML alimenta varios campos JSON), `XML_NO_DEST` (tag XML requerido sin destino JSON),
`DIFF_ACROSS_TYPES` (el mismo campo apunta a tags distintos según el tipo) y `POSITIONAL` (informativo).
`--fail-on-error` devuelve exit 1 si hay ERROR. El `.xlsx` trae `Resumen` y `Hallazgos` con hoja y fila de origen del libro.

## `audit-plano`

Las reglas salen del libro; los documentos reales no son fuente de reglas. `data/specs/ajustes_contrato.json` solo tiene vigentes `decimales_equivalentes` (`285.6` = `285.60` en el cruce XML↔JSON) y `catalogos_como_string` (campos con lista de valores → `string`, aceptando también números sin comillas en el JSON; indicación del usuario); en lo demás el plano sigue fielmente `contrato_ejemplo.xlsx`, incluidos sus cruces (`apendice.campo` ↔ `Notas/etiqueta`).

Comprueba, con código propio e independiente de `build_contract.py`, que `contrato_plano.xlsx` refleje `contrato_ejemplo.xlsx`:
cobertura de filas (`ROW_LOST`), trazabilidad (`ROW_UNTRACEABLE`), campos (`FIELD_DIFF`: ruta, cambio, ruta anterior, catálogo, contraparte XML, posicional, requerido),
reciprocidad JSON↔XML (`COUNTERPART_ASYMMETRIC`), tipos completos (`TYPE_INCOMPLETE`) y que regenerar el plano dé lo mismo (`PLANO_STALE`).
Clasifica además `lint` (`LINT_DUPLICATE`, `LINT_NOT_PATH`, `LINT_NO_REQ`...) como decisiones pendientes del libro. No valida tipos de dato.

## Valores esperados desde la BD (Fase B: `queries/`, `expected.py`)

Referencia de contenido = tablas de negocio Oracle del esquema de la compañía del documento (`data/specs/companias.json`:
53 → `VENTASSAL`, 41 → `VENTASSVC`) y `VENTASTQ`, no el JSON que guarda el generador.
Clave: `DOC.PREIMPRESO` = Código de Generación. Detalle y hallazgos en `docs/fase_b_descubrimiento.md`.

- `data/specs/mapeo_bd.xlsx`: ruta del contrato × tipo → categoría (`DIRECTA`, `TRANSFORMACION`, `CATALOGO`, `SIN_FUENTE`,
  `NO_APLICA`), fuente, fórmula y evidencia. Aprobado el 2026-11-16; `SIN_FUENTE` se reportará como `DB_MISSING`.
  El 2026-11-17 pasaron a tener fuente `ambiente`, `horEmi`, `tipoMoneda`, `totalLetras`, `uniMedida`, `psv`,
  `noGravado`, `descuNoSuj`, `descuExenta`, `descuGravada`, `porcentajeDescuento`, `saldoFavor` y `totalNoGravado`
  (ver `docs/INSTRUCTIVO.md` §6). `normalizaciones_bd.json`: `tolerancia_montos` (0.01) y `tolerancia_hora_segundos`
  (120, para `horEmi`; aprobado 2026-11-17).
- `data/specs/normalizaciones_bd.json`: normalizaciones aprobadas BD ↔ documentos (tolerancia de montos 0.01).
- `queries/valores_esperados.sql`: un solo `SELECT` (Oracle, solo lectura) parametrizado con `:codigo_generacion`, `:desde`,
  `:hasta`, `:max_docs`. Devuelve `CODIGO_GENERACION | TIPO_DOC | RUTA | VALOR`; arreglos con sufijo `[n]` (1-based; ítems =
  `DET_VENTAS` con `CANTDESP <> 0` en el orden del generador: FEX por `ORDEN_FACTURAS`; demás por `PADRE`, `ORDEN_FACTURAS`) y apéndice con `[apendiceNN]`.
- `load_expected(source)` lee el resultado exportado (`.tsv`/`.csv`/`.xlsx`, p. ej. desde DBeaver) o un iterable de filas y
  devuelve `{codigo_generacion: {ruta: valor}}`, solo texto.

```powershell
$env:PYTHONPATH="src"; python -c "from contract_validator.expected import load_expected; print(len(load_expected('reports/esperados.tsv')))"
```

Alcance: CCF, FAC, FAC_EXP, NCR, NDB, REM (INVALIDACION y RETORNO sin fuente ubicada).

### Consulta automática a la BD (`db_fetch.py`) y veredicto (`verdict.py`)

La automatización ejecuta el query ella misma con `oracledb` (modo thin). Conexión en `config/oracle.local.json`
(sin credenciales; plantilla en `config/oracle.example.json`); usuario y contraseña en el archivo `.env` del proyecto
(`CV_ORACLE_USER`, `CV_ORACLE_PASSWORD`; ignorado por git; plantilla `.env.example`; las variables del sistema tienen
prioridad). Salvaguardas:
única sentencia `SELECT`/`WITH`, `SET TRANSACTION READ ONLY`, rollback y cierre, consulta solo por los UUID recibidos.

```powershell
python -m contract_validator verificar entrada/                     # todo en uno + veredicto (reports/verificacion_<fecha>.xlsx)
python -m contract_validator fetch-expected --cia 53 --codigo-generacion <UUID>... --out reports/esperados.tsv   # solo extraer
python -m contract_validator fetch-expected --desde 2026-11-01 --hasta 2026-11-15 --max-docs 200
python -m contract_validator validate entrada/ --expected reports/esperados.tsv   # comparar contra un archivo
# --expected-db (en validate) consulta en línea; --save-expected guarda lo consultado; --db-config otra conexión
# --mapeo-bd otro.xlsx  --bd-config otras.json  cambian el mapeo aprobado y las normalizaciones de BD
```

Veredicto por documento (consola y hoja `Veredicto`): Estructura vs `.xlsx` (CUMPLE/NO CUMPLE), Información vs BD
(CORRECTA/INCORRECTA/NO VERIFICADA), Coherencia XML/JSON, Validación del PDF (CUMPLE/NO CUMPLE/NO VERIFICADO/NO RECIBIDO/IGNORADO/NO VALIDADO) y RESULTADO (CUMPLE/NO CUMPLE/INCOMPLETO). Exit code de
`verificar`: 0 todo cumple, 1 algún incumplimiento, 2 algo sin verificar (p. ej. sin conexión a la BD, que no aborta).

### Comparación contra la BD (`db_check.py`)

Empareja cada JSON, XML y PDF MINH con sus valores esperados por Código de Generación. Solo compara rutas que el contrato
define para el tipo y que están aprobadas en `mapeo_bd.xlsx` (DIRECTA, TRANSFORMACION, CATALOGO). Igualdad exacta de
texto; números sin ceros a la izquierda se comparan como números con la tolerancia de `normalizaciones_bd.json`.

- JSON: por ruta, con índice de arreglo (`cuerpoDocumento.cantidad[2]`) y apéndice por `campo`.
- XML: vía `ruta_contraparte`; los mapeos posicionales del libro se omiten (como en el cruce). El apéndice se ubica por la
  contraparte de `apendice.campo`; si el contrato no permite ubicarlo se avisa como `DB_MISSING` (no comparado).
- PDF MINH: Código de Generación y Número de Control, etiquetas ancladas en la `Descripcion` del Excel (igual que
  `pdf_check`, con las normalizaciones de `pdf_presentacion.json`) y bloque de cada ítem (cantidad, precio, descuento).

Códigos: `DB_MISMATCH_JSON`, `DB_MISMATCH_XML`, `DB_MISMATCH_PDF` (ERROR: difiere, falta en el documento o distinta
cantidad de ítems); `DB_MISSING` (ERROR si el documento no está en el resultado del query; WARNING, uno por documento,
con los campos sin fuente en BD o sin valor en BD, que no se pueden verificar).

## Estructura

```
src/contract_validator/
  build_contract.py  # libro original (hojas 'XML' y 'JSON') -> hoja plana 'contrato' + 'lint' + 'ajustes'
  spec_loader.py     # lee la hoja plana (o el formato simple antiguo)
  flatten.py         # JSON/XML -> rutas; lectura tolerante de encoding, quita envoltura SOAP
  comparator.py      # reglas por tipo_doc/formato
  rules.py           # número de control (único contenido definido por el Excel)
  crosscheck.py      # cruce XML <-> JSON
  pdf_text.py        # texto de PDF (pdftotext -raw), identidad, MINH por nombre
  pdf_check.py       # PDF MINH vs XML
  audit_mapping.py   # auditoría del mapeo
  expected.py        # load_expected: valores esperados desde el resultado del query
  db_check.py        # JSON/XML/PDF contra los valores esperados de BD
  db_fetch.py        # ejecuta el query en Oracle (solo lectura)
  verdict.py         # veredicto por documento (estructura / información BD / coherencia)
  explain.py         # explicación en lenguaje natural de cada hallazgo (columna 'explicacion' de Hallazgos)
  report.py  cli.py
queries/             # valores_esperados.sql (Oracle, solo lectura)
config/              # oracle.example.json (plantilla); oracle.local.json (servidor, local, no versionado)
.env                 # usuario y contraseña de Oracle (local, no versionado; plantilla .env.example)
docs/                # traspasos y descubrimiento de la Fase B
data/specs/          # libro y ejemplos entregados (contrato_plano.xlsx se genera; ajustes_contrato.json lo complementa)
entrada/             # documentos reales a verificar (PDF, XML, JSON)
tests/               # pytest
```
