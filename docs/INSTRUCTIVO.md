# Instructivo: cómo funciona la automatización de validación de DTE

Este instructivo explica, de punta a punta, qué hace el validador, qué necesita, cómo se ejecuta y cómo leer el
resultado. Está pensado para quien lo opera (QA, soporte) y para quien lo mantiene.

---

## 1. Qué hace

Se le entregan los archivos de un documento tributario electrónico (DTE) de El Salvador —**PDF, XML y JSON**— y con un
solo comando (`verificar`) responde, por documento:

1. **¿La estructura cumple el contrato `.xlsx`?** → `CUMPLE` / `NO CUMPLE`
2. **¿La información es correcta según la base de datos?** → `CORRECTA` / `INCORRECTA` / `NO VERIFICADA`

Para llegar a eso, revisa **cuatro cosas** (la consulta a la base la hace la propia automatización, en solo lectura):

| # | Qué compara | Contra qué | Módulo |
|---|---|---|---|
| 1 | **Estructura** de cada `.json` y `.xml` (campos nuevos, eliminados, renombrados, requeridos, nulos, catálogos) | El contrato del libro Excel | `comparator.py` |
| 2 | **XML ↔ JSON** del mismo documento, campo a campo | El mapeo `ruta_contraparte` del contrato | `crosscheck.py` |
| 3 | **PDF oficial (MINH) ↔ XML** del mismo documento | Etiquetas del Excel, ítems y valores del XML | `pdf_check.py` |
| 4 | **Contenido** de JSON, XML y PDF | Los **valores reales de la base de datos** (Oracle) | `db_check.py` |

Principio que rige todo: **solo se valida lo que definen los specs** (el libro Excel, el mapeo aprobado y las
normalizaciones aprobadas). Nada se "interpreta": si algo no está definido, no se valida y se informa.

Tipos de documento cubiertos: CCF (03), FAC (01), FAC_EXP (11), NCR (05), NDB (06), REM (04).
INVALIDACION y RETORNO se validan en estructura, pero **no** contra la base de datos (su fuente no está ubicada).

```
 contrato_ejemplo.xlsx ──build──► contrato_plano.xlsx ──────────────┐
                                                                    ▼
  entrada: .pdf / .xml / .json ──────────────────────────────►  verificar ──► VEREDICTO por documento
        │ (Códigos de Generación)                                   ▲          + reporte .xlsx (Veredicto,
        ▼                                                           │            Resumen, Hallazgos)
  Oracle (solo lectura) ◄── queries/valores_esperados.sql ──────────┘
        mapeo_bd.xlsx + normalizaciones_bd.json (qué y cómo se compara)
```

---

## 2. Requisitos

- Python 3.12 y las dependencias: `pip install -r requirements.txt` (openpyxl, pytest, oracledb).
- `pdftotext` (Xpdf o Poppler) en el PATH, en la instalación de Git for Windows o en la variable `PDFTOTEXT=<ruta>`, para leer los PDF.
- Para la validación contra BD: usuario de **solo lectura** en Oracle, red/VPN hacia el servidor, conexión en
  `config/oracle.local.json` y usuario/contraseña en el archivo `.env` del proyecto (ver Paso 2).
- Todos los comandos se ejecutan desde la raíz del repositorio. En PowerShell, antes de cualquier comando:

```powershell
$env:PYTHONPATH="src"
```

---

## 3. Archivos que intervienen

| Archivo | Para qué sirve | ¿Se edita a mano? |
|---|---|---|
| `data/specs/contrato_ejemplo.xlsx` | Libro original del contrato | **No** (es la fuente; nunca se modifica) |
| `data/specs/contrato_plano.xlsx` | Contrato "aplanado" que usa el validador (hojas `contrato`, `lint` y `ajustes`) | **No**: se regenera con `build` (una edición a mano se pierde y `audit-plano` la marca `PLANO_STALE`) |
| `data/specs/ajustes_contrato.json` | Excepciones al libro **indicadas por el usuario** (campos que se agregan o cambian respecto del libro), con motivo, evidencia y fecha de aprobación | Sí, con aprobación |
| `data/specs/pdf_presentacion.json` | Normalizaciones aceptadas PDF ↔ XML (fechas, miles, guiones) y tags que el PDF no imprime | Sí, con aprobación |
| `data/specs/mapeo_bd.xlsx` | Qué columna/tabla de la BD alimenta cada campo del contrato, por tipo de documento | Sí, con aprobación |
| `data/specs/normalizaciones_bd.json` | Normalizaciones aceptadas BD ↔ documentos (hoy: tolerancia de 0.01 en montos y de 120 s en la hora de emisión) | Sí, con aprobación |
| `queries/valores_esperados.sql` | Query Oracle (solo `SELECT`) que la automatización ejecuta para traer los valores esperados por documento | Solo si cambia el mapeo |
| `config/oracle.local.json` | Servidor Oracle (host, puerto, SID/servicio), **sin credenciales**. No se versiona | Sí, una vez por equipo |
| `.env` | Usuario y contraseña de Oracle (`CV_ORACLE_USER`, `CV_ORACLE_PASSWORD`). No se versiona (plantilla `.env.example`) | Sí, una vez por equipo |
| `config/oracle.example.json` | Plantilla del anterior | No |
| `docs/fase_b_descubrimiento.md` | Esquema de la BD, reglas y hallazgos | Documentación |

---

## 4. Paso a paso

### Paso 1 — Generar el contrato plano (solo cuando cambia el libro)

```powershell
python -m contract_validator build --source data/specs/contrato_ejemplo.xlsx --out data/specs/contrato_plano.xlsx
python -m contract_validator audit-plano --fail-on-error
```

`build` convierte el libro en una tabla plana (una fila por campo, tipo de documento y formato) y le aplica los ajustes
aprobados de `data/specs/ajustes_contrato.json` (hoja `ajustes` del plano: valor del libro vs valor ajustado).
`audit-plano` lo verifica con un lector independiente: debe terminar con **0 errores** antes de seguir; cada ajuste
aparece como `AJUSTE_APLICADO` (INFO). `--sin-ajustes` genera solo desde el libro.

**Fuente de las reglas: el libro.** Las reglas salen de `contrato_ejemplo.xlsx`. Los JSON, XML y PDF reales son lo que
se valida y pueden tener errores: si difieren del libro se reporta el hallazgo, no se cambia la regla. Si el usuario
indica una excepción al libro, **no editar `contrato_plano.xlsx`** (se pierde al regenerar): agregar una entrada en
`data/specs/ajustes_contrato.json` y volver a ejecutar `build` + `audit-plano`:

```json
{
  "id": "X01",
  "accion": "agregar",                       // o "modificar" (cambia campos de una fila existente)
  "tipo_doc": "CCF", "formato": "XML",
  "ruta": "DatosEntrada/InformacioGeneral/nuevoCampoTQ",
  "campos": {"tipo": "string", "requerido": "NO", "cambio": "SIN_CAMBIO", "descripcion": "..."},
  "motivo": "Excepción indicada por el usuario: ...",
  "evidencia": "Indicación del usuario o fila del libro",
  "aprobado": "AAAA-MM-DD"
}
```

`tipo_doc` acepta un tipo (`"CCF"`), una lista (`["CCF", "FAC"]`) o `"*"` (todos los tipos de ese formato).

**Fuente del plano.** Desde 2026-11-17 (confirmado por el cliente) los tags reales salen de las hojas **`XML`** y
**`JSON`** del libro: en la hoja `XML` cada tipo de documento tiene su columna con la ruta JSON (con
`Adicionar`/`Eliminar`/`Renombrar`) y el tag XML es el `XML_TQ` de la misma fila (RETORNO usa la columna `XML-ER`); la
hoja `JSON` es la base v3 y solo se usa para verificar. INVALIDACION se sigue leyendo de su hoja propia porque la hoja
`XML` no trae sus tags XML. Las marcas `Eliminar` de las hojas por tipo también se aplican.

**Excepciones vigentes** (indicación del usuario, 2026-11-18):

- `decimales_equivalentes`: en el cruce XML↔JSON, los números con el mismo valor y distinto formato decimal (`285.6` y
  `285.60`, `2260` y `2260.00`) son iguales y no generan error ni aviso. Los códigos con ceros a la izquierda (`05` y `5`)
  y las fechas siguen como `XREF_FORMAT_DIFF`.
- `catalogos_como_string`: los campos que en el libro tienen una lista de valores se generan como `string`, sin lista,
  en el JSON y en su tag XML. Por eso no sale `VALUE_NOT_IN_ENUM`, y en esos campos se acepta tanto `"1"` como `1` en el JSON.

No hay ajustes por fila. En lo demás el plano sigue fielmente `contrato_ejemplo.xlsx`:

- `Si/null/ por tipoDTE` y `Si (null)`: el tag es requerido y admite valor nulo.
- La columna `Tipo` se valida (`TYPE_MISMATCH`).
- Los campos que el libro no define se reportan como `UNEXPECTED_FIELD`.
- `CabeceraEntrada` está en la hoja `XML` (`uuid` en la fila 8 y los demás tags en las filas 216–225): requeridos en los
  6 tipos de factura, con el dato esperado en `Descripcion`. `Solo XML` en la columna del tipo indica un tag sin campo
  JSON. Se valida presencia, valor no vacío y tipo; los valores fijos descritos no se comparan.
- `DatosEntrada/Referencia` (`tipoDocumentoReferencia`, `codigoReferencia`, `numeroDocumentoReferencia`,
  `fechaEmisionReferencia`) y sus campos `documentoRelacionado/*` solo aplican a NC y ND entre los tipos de factura; el
  evento de Retorno conserva los suyos (columna `XML-ER`).
- `Solo null: <ruta>` en la columna del tipo indica que el campo JSON debe llegar con valor `null`; si trae datos da
  `VALUE_NOT_NULL`. Aplica a `documentoRelacionado` en CCF, Factura, Factura de Exportación y Nota de Remisión.
- Si un campo dice `Eliminar` en una columna `TAG_MIN_HACIENDA` (hoja `XML` u hoja del tipo), **no va en el JSON**: si llega
  da `REMOVED_STILL_PRESENT`. Su tag XML puede venir, pero no es obligatorio.
- Con la misma marca, las filas 226–229 definen `InformacionEmisor/nitCompania`, `codigoEstablecimientoTQ`, `codigoSRI`
  (admite null) y `puntoEmisionTQ` en `DatosEntrada/InformacioGeneral`.
- `codigoEmpresa`, `firmaElectronica` y `selloRecibido` están en el libro (hoja `XML` filas 213–215 y hoja `JSON` filas
  140–142): van en la raíz del JSON de los 6 tipos de factura, son requeridos, de tipo `string` y sin null. No aplican
  al XML ni a los eventos de invalidación y retorno.
- Los cruces son los de la hoja `XML`, también `apendice/campo` ↔ `Notas/etiqueta`.

**Ejemplo (salida real):**

```text
PS> python -m contract_validator build --source data/specs/contrato_ejemplo.xlsx --out data/specs/contrato_plano.xlsx
Contrato plano: data/specs/contrato_plano.xlsx
  filas: 1370  lint: {'ERROR': 0, 'WARNING': 57, 'INFO': 111}  ajustes: ninguno
  CCF          JSON 126
  CCF          XML  79
  FAC          JSON 126
  ...
  RETORNO      XML  54

PS> python -m contract_validator audit-plano --fail-on-error
Hallazgos: 176  (ERROR=0)
  WARNING LINT_DUPLICATE             41   CCF=9, FAC=6, FAC_EXP=6, ...
  INFO    AJUSTE_APLICADO             7   CCF=7
  INFO    LINT_NO_REQ               104   CCF=7, FAC=5, ...
Reporte: reports/auditoria_plano.xlsx
```

Lo que importa es `ERROR=0`. Los `WARNING`/`INFO` son decisiones pendientes del libro (hoja `lint`), no bloquean.

### Paso 2 — Configurar la conexión a la base de datos (una sola vez por equipo)

La automatización consulta Oracle por sí misma (driver `oracledb`, modo thin: no requiere cliente Oracle ni DBeaver).

1. Instalar dependencias (incluye el driver): `pip install -r requirements.txt`.
2. Servidor en `config/oracle.local.json` (no se versiona). Ya viene con el servidor de pruebas `svmtest`:

   ```json
   { "host": "PRU950.TECNOQUIMICAS.COM", "port": 1530, "sid": "svmtest" }
   ```

   Para otro ambiente, copiar `config/oracle.example.json` (acepta `sid`, `service` o `dsn`).
3. **Usuario y contraseña en el archivo `.env` de la raíz del proyecto** (ya existe vacío; plantilla en `.env.example`).
   Abrirlo con un editor de texto, completar y guardar:

   ```dotenv
   CV_ORACLE_USER=usuario_solo_lectura
   CV_ORACLE_PASSWORD=la_contraseña
   ```

   - `.env` está en `.gitignore`: no se versiona. No copiarlo a otras carpetas ni pegar su contenido en chats o tickets.
   - Lo leen tanto el comando ejecutado desde una terminal como el agente de Claude Code; no hace falta reiniciar nada.
   - Si en el sistema existe una variable con el mismo nombre, esa tiene prioridad sobre `.env`.
   - En `.env` también pueden ir `CV_ORACLE_HOST`, `CV_ORACLE_PORT`, `CV_ORACLE_SID`, `CV_ORACLE_SERVICE` o
     `CV_ORACLE_DSN` para apuntar a otro servidor (precedencia: variable del sistema > `.env` > `config/oracle.local.json`).
   - Otro archivo de credenciales (p. ej. por ambiente): `$env:CV_ENV_FILE = ".env.qa"`.
4. Estar conectado a la red/VPN que llega al servidor.

Comprobar que la configuración está completa (no muestra valores):

```powershell
python -c "from contract_validator.db_fetch import load_db_config as L; c=L(); print('usuario:', bool(c.get('user')), '| contraseña:', bool(c.get('password')))"
# usuario: True | contraseña: True
```

Comprobar la conexión con un documento conocido:

```powershell
python -m contract_validator fetch-expected --codigo-generacion 13062C60-FF5A-BECC-E063-9B78A8C07752 --out reports/prueba_conexion.tsv
# 1 documento(s), 86 valor(es) esperado(s) -> reports/prueba_conexion.tsv
```

Salvaguardas incorporadas: el SQL se rechaza si no es una única consulta `SELECT`/`WITH`; la sesión abre con
`SET TRANSACTION READ ONLY` (Oracle impide cualquier escritura); al terminar se hace rollback y se cierra la conexión;
se consulta solo por los Códigos de Generación de los archivos recibidos (o con tope de documentos en un rango).

### Paso 3 — Verificar los documentos (todo automático)

Se entregan los archivos de entrada (PDF, XML y JSON, sueltos o en carpetas) y **un solo comando** hace todo:

```powershell
python -m contract_validator verificar <archivos y/o carpetas>
```

Internamente:

1. **Estructura:** cada JSON y XML contra el contrato `.xlsx` (campos nuevos, eliminados, renombrados, requeridos, nulos, catálogos).
2. **Coherencia entre archivos:** XML ↔ JSON del mismo documento. El PDF oficial (MINH) se valida aparte, contra el XML y contra la BD.
3. **Consulta a la BD:** toma el Código de Generación de cada archivo y ejecuta `queries/valores_esperados.sql` en
   Oracle (solo lectura) para exactamente esos documentos.
4. **Información vs BD:** compara JSON, XML y PDF contra los valores consultados (`mapeo_bd.xlsx`, `normalizaciones_bd.json`).
5. **Veredicto por documento** en consola y en la hoja **Veredicto** (primera hoja) del Excel
   `reports/verificacion_<fecha_hora>.xlsx` (o el que se indique con `--out`).

| Línea del veredicto | Responde | Valores |
|---|---|---|
| Estructura vs contrato (.xlsx) | ¿Los JSON/XML cumplen el contrato? | `CUMPLE` / `NO CUMPLE` / `SIN JSON/XML` (solo llegó PDF) |
| Información vs base de datos | ¿Los datos de JSON y XML son los de la BD? | `CORRECTA` / `INCORRECTA` / `NO VERIFICADA` (con el motivo) |
| Coherencia XML / JSON | ¿El XML y el JSON dicen lo mismo? | `OK` / `CON DIFERENCIAS` / `NO APLICA` |
| Validación del PDF | ¿El PDF MINH coincide con el XML y con la BD? | `CUMPLE` / `NO CUMPLE` / `NO VERIFICADO` / `NO RECIBIDO` / `IGNORADO` / `NO VALIDADO` |
| **RESULTADO** | Conclusión del documento | **`CUMPLE`** / **`NO CUMPLE`** / **`INCOMPLETO`** (nada incumple pero falta verificar algo) |

Código de salida: `0` todos los documentos CUMPLE · `1` algún documento NO CUMPLE · `2` sin incumplimientos pero
algo quedó NO VERIFICADO (por ejemplo, sin conexión a la BD).

Si la base no está disponible (sin VPN, sin usuario o sin contraseña), **no se detiene**: valida la estructura y la
coherencia igual y marca la información como `NO VERIFICADA` indicando qué falta.

**Ejemplos de ejecución:**

```powershell
# a) Un documento: su PDF oficial, su XML y su JSON
python -m contract_validator verificar entrada/CCF_SAL_MINH_531601001131126020656.pdf entrada/53_16_01_001_131126_020656_*.xml entrada/CCF_13062C60.json

# b) Todo lo que dejó QA en una carpeta (recursivo), con nombre de reporte propio
python -m contract_validator verificar entrada/ --out reports/verificacion_lote_noviembre.xlsx

# c) Guardar además lo que se consultó en la BD (para auditoría; contiene datos reales)
python -m contract_validator verificar entrada/ --save-expected reports/esperados_lote_noviembre.tsv

# d) Otro ambiente de base de datos
python -m contract_validator verificar entrada/ --db-config config/oracle.qa.local.json

# e) Forzar el tipo cuando el archivo no lo indica
python -m contract_validator verificar entrada/documento.json --doc-type NCR

# f) En un pipeline (CI): falla si algún documento NO CUMPLE
python -m contract_validator verificar entrada/ --out reports/ci.xlsx
if ($LASTEXITCODE -eq 1) { Write-Error "Hay documentos que NO CUMPLEN"; exit 1 }
```

**Ejemplo de salida (real, recortada; tres documentos):**

```text
...detalle por archivo (estructura, cruce, PDF, BD)...

====================================================================================================
VEREDICTO POR DOCUMENTO
====================================================================================================
Documento 13062C60-FF5A-BECC-E063-9B78A8C07752  (CCF)  archivos: PDF, XML
  Estructura vs contrato (.xlsx) : NO CUMPLE      REQUIRED_MISSING x1, UNEXPECTED_FIELD x1
  Información vs base de datos   : INCORRECTA     1 diferencia(s) en 71 valor(es) comparado(s): DB_MISMATCH_XML x1
  Coherencia XML / JSON          : OK             sin diferencias
  Validación del PDF             : NO RECIBIDO    no se recibió PDF MINH de este documento
  RESULTADO                      : NO CUMPLE
----------------------------------------------------------------------------------------------------
Documento 150CDB5F-98E0-57AB-E063-9B78A8C0B3CA  (REM)  archivos: JSON, PDF
  Estructura vs contrato (.xlsx) : NO CUMPLE      UNEXPECTED_FIELD x14, NULL_NOT_ALLOWED x2, REQUIRED_MISSING x1
  Información vs base de datos   : INCORRECTA     1 diferencia(s) en 173 valor(es) comparado(s): DB_MISMATCH_JSON x1
  Coherencia XML / JSON          : OK             sin diferencias
  Validación del PDF             : NO RECIBIDO    no se recibió PDF MINH de este documento
  RESULTADO                      : NO CUMPLE
----------------------------------------------------------------------------------------------------
Documento E9B3C4D4-F5A7-7B8C-9D0E-2F2A3B4C5D6E  (FAC)  archivos: JSON
  Estructura vs contrato (.xlsx) : NO CUMPLE      NULL_NOT_ALLOWED x8
  Información vs base de datos   : NO VERIFICADA  El Código de Generación E9B3C4D4-... no está en el resultado del query de BD
  Coherencia XML / JSON          : NO APLICA      no hay otro archivo del mismo documento para cruzar
  Validación del PDF             : NO RECIBIDO    no se recibió PDF MINH de este documento
  RESULTADO                      : NO CUMPLE
----------------------------------------------------------------------------------------------------
TOTAL: 3 documento(s) — CUMPLE: 0, NO CUMPLE: 3, INCOMPLETO: 0
Reporte: reports/verificacion_20261116_163010.xlsx
```

Y sin conexión a la base (por ejemplo, sin contraseña definida):

```text
  Información vs base de datos   : NO VERIFICADA  no se pudo consultar la BD: Falta la contraseña: CV_ORACLE_PASSWORD en el archivo .env del proyecto
```

**Qué revisar después:** en el Excel, la hoja **Veredicto** da la conclusión por documento; la hoja **Hallazgos**
tiene el detalle (filtrar por `severidad = ERROR` y por `formato`: `CCF/XML`, `XML<->JSON`, `PDF(MINH)<->XML`,
`BD<->JSON`…). Estructura y coherencia → equipo de integración; `BD<->…` → datos/generador; PDF → plantilla de impresión.
En `BD<->…`, `BD='…' documento='…'` es valor esperado vs encontrado y `falta en el documento` significa que la BD tiene
valor y el archivo no. El aviso `DB_MISSING` lista campos que **no se pueden verificar** (sin fuente en la BD).

### Ejemplo completo de una corrida (de principio a fin)

Caso: QA deja en `entrada/` los PDF, XML y JSON de los documentos emitidos en noviembre.

```powershell
# 0) Preparar la sesión (una vez por terminal)
cd C:\ruta\contract-validator
$env:PYTHONPATH="src"
#    (usuario y contraseña de Oracle ya completados en .env — Paso 2)

# 1) (solo si cambió el libro) regenerar y auditar el contrato
python -m contract_validator build
python -m contract_validator audit-plano --fail-on-error

# 2) Verificar todo: estructura + coherencia + consulta a la BD + comparación + veredicto
python -m contract_validator verificar entrada/ --out reports/verificacion_noviembre.xlsx

# 3) Conclusión
switch ($LASTEXITCODE) { 0 {"Todos los documentos CUMPLEN"} 1 {"Hay documentos que NO CUMPLEN"} 2 {"Hay información NO VERIFICADA (revise conexión a BD)"} }
```

Con el skill `validate-contract` en Claude Code basta con pedir, por ejemplo: *"verifica los archivos de la carpeta
entrada/"*; el agente ejecuta `verificar` y responde con el veredicto de cada documento.

---

---

## 5. Cómo leer el resultado

Primero el **detalle**: cada archivo o par sale como `[OK]` o `[FAIL]` con su tipo, formato y hallazgos. Al final, el
**VEREDICTO POR DOCUMENTO** (ver Paso 3). El `.xlsx` trae tres hojas: **Veredicto** (una fila por documento, con la
validación del PDF y el resultado), **Resumen** (una fila por archivo o cruce) y **Hallazgos** (una fila por hallazgo, con
validación, código, campo, **explicacion** y mensaje). Código de salida de `verificar`: **0** todo CUMPLE, **1** algún NO CUMPLE,
**2** algo NO VERIFICADO.

En **Hallazgos**, la columna **explicacion** dice el hallazgo en lenguaje natural, con el nombre del campo tal como
está en el libro (columna `Descripcion`), el archivo afectado y los valores; la columna **mensaje** conserva el texto
técnico. Ejemplos reales:

| codigo | explicacion | mensaje (técnico) |
|---|---|---|
| `REQUIRED_MISSING` | Falta «Monto Total de la Operación» (resumen.montoTotalOperacion) en el JSON. El contrato lo exige para el Comprobante de Crédito Fiscal. | Campo requerido no encontrado (cambio: SIN_CAMBIO, XML fila 156) |
| `XREF_FORMAT_DIFF` | «Tipo de documento» (receptor.tipoDocumento) tiene el mismo valor en el JSON y en el XML, pero escrito distinto (05 vs 5). Es solo de formato. | receptor.tipoDocumento <-> …: mismo valor con distinto formato JSON='05' XML='5' |
| `DB_MISMATCH_JSON` | Según la base de datos, el apéndice «apendice01» (apendice.campo) debería estar en el JSON, pero el JSON no lo trae. | apendice.campo[apendice01]: BD='apendice01' falta en el documento |

La salida `--format json` incluye el mismo texto en `explicacion` de cada hallazgo.

**Severidad:** `ERROR` = incumple lo definido. `WARNING` = algo a revisar que no se puede afirmar como error.

### Códigos más frecuentes

| Código | Nivel | Qué significa | Qué hacer |
|---|---|---|---|
| `NEW_MISSING`, `REQUIRED_MISSING`, `RENAMED_MISSING` | ERROR | Falta un campo requerido por el contrato. En XML se revisa además cada posición indicada en el libro (`valorOtrosTotales[4]`) y que los tags de ítem vengan en todos los ítems | Revisar el generador del documento |
| `APENDICE_NOT_RECEIVED` | WARNING | La BD arma un apéndice que no llegó en el XML | Confirmar que su ausencia es correcta |
| `VALUE_NOT_NULL` | ERROR | Un campo que el libro define como siempre `null` trae información | Revisar el generador del documento |
| `PDF_NOT_RECEIVED`, `PDF_IGNORED`, `PDF_NOT_VALIDATED` | WARNING | No se validó un PDF MINH: no llegó, solo llegó uno interno o se usó `--no-pdf` | Adjuntar el PDF MINH del documento |
| `XML_OPTIONAL_MISSING` | WARNING | Un tag XML del libro no llegó, pero su `Requerido` es `No`, `N/A` o vacío | Confirmar que su ausencia es correcta para ese documento |
| `UNEXPECTED_FIELD` | ERROR | Campo que el contrato no define | Si es un renombre no declarado, proponerlo en el libro |
| `NULL_NOT_ALLOWED` | ERROR | Campo requerido con valor nulo | Revisar el generador |
| `XREF_MISMATCH` / `XREF_MISSING` | ERROR | XML y JSON difieren / un lado no trae el valor | Puede ser del documento **o del mapeo del libro** |
| `XREF_FORMAT_DIFF` | WARNING | Mismo valor con distinto formato | Decide el equipo |
| `PDF_MISMATCH`, `PDF_IDENTITY` | ERROR | El PDF oficial no coincide con el XML | Revisar la impresión |
| `PDF_NOT_PRINTED` | WARNING | Valor del XML que no aparece en el PDF | Si no se imprime por diseño, agregarlo a `no_impreso` |
| `DB_MISMATCH_JSON` / `_XML` / `_PDF` | ERROR | El documento difiere de la BD, le falta un valor que la BD tiene, o tiene otra cantidad de ítems | Revisar el generador o los datos |
| `DB_MISSING` | ERROR | El Código de Generación del archivo no existe en la BD consultada (o en el archivo de valores) | Confirmar el ambiente (pruebas/producción) y que el documento esté registrado |
| `DB_MISSING` | WARNING | Campos del documento sin fuente en BD (no verificables) | Informativo; se lista una vez por documento |
| `TYPE_MISMATCH` | ERROR | El valor no tiene el tipo de la columna `Tipo` del libro | Revisar el generador del documento |

### Situaciones conocidas (no son fallas del validador)

- **Documentos emitidos con una versión anterior del esquema** (v1/v3): mostrarán campos del contrato como faltantes
  (`ivaRete` se llamaba `ivaRete1`, no traen `distrito` ni `observaciones`) y `identificacion.version` distinta.
- **Datos que cambiaron después de emitir:** el correo del receptor o la descripción del artículo pueden diferir en
  documentos antiguos porque el maestro se editó después.
- **Texto con acentos dañados** (`FARMACÃ‰UTICOS`) en el JSON certificado guardado en la BD: es un defecto real de
  esos archivos (doble codificación), no de la comparación.
- **Apéndice (`<Notas>` del XML / `apendice[]` del JSON):** el query genera `campo`, `etiqueta` y `valor` de
  `apendice01`…`04` con la misma lógica del generador (`VENTASAPL.IN_VEN_FACTURA_ELECTRONIC3_PKG.XML_NOTAS_SV_P`;
  exportación `XML_NOTAS_EXP_SV_P`): etiquetas fijas y valores de `DOC`/`DOC_EXP`. Cada apéndice es la unión
  de varios datos, así que se compara **completo**, ubicado por su nombre `apendiceNN` (en el XML, en `<campo>` o `<etiqueta>`). `campo`, `etiqueta` y `valor` son
  obligatorios en cada apéndice (hoja `XML`, filas 198–202):
  una diferencia da un hallazgo sobre `apendice[apendiceNN]`. Si está en la BD y no llega en el JSON da
  `DB_MISMATCH_JSON` (y `XREF_MISSING` si está en el XML), aunque `resumen.observaciones` tenga los mismos datos; si no
  llega en el XML da el aviso `APENDICE_NOT_RECEIVED`.
- **Régimen de exportación en el PDF:** el código es `tipoRegimen` (abreviado, p. ej. `EX-1`) + `.` + `regimen`. Con ese código se
  busca en `pdf_presentacion.json` el texto impreso, que puede ocupar más de una línea del PDF (catálogo `REGIMEN` de `FE_PARAMETROS`).
- **Apéndice en el PDF:** los datos de `apendice01` (`Regional`, `Bodega`, `Concepto`, `Pedido`, `Plazo`) deben verse en el PDF con su
  etiqueta; se comparan con el XML y con la BD. La lista está en `apendices_impresos` de `pdf_presentacion.json`.
- **PDF de exportación:** `recintoFiscal` y `regimen` se imprimen como texto y la BD trae el código; se traducen con
  los catálogos de `pdf_presentacion.json` (aprobados 2026-11-17). `regimen` trae los 9 códigos del parámetro `REGIMEN`;
  si un texto impreso difiere, se ajusta en ese catálogo tal como lo imprime el PDF.

---

## 6. Qué no se valida contra la BD

Campos sin columna en la base (se reportan como `DB_MISSING` aviso): `emisor.code`, `apendice05+` (mensajes de
`PBS_MENSAJES_FACTURA_P`), descripción larga de ítems en FAC (textos de licitación) y los campos nuevos de NCR/NDB sin
muestra. La lista completa, por tipo, está en `mapeo_bd.xlsx` (categoría `SIN_FUENTE`).

Desde 2026-11-17 la consulta también calcula estos 13 campos, que antes quedaban como `DB_MISSING` (la lógica sale del
código del generador y se confirmó con los 7 JSON certificados):

| Campo JSON | Valor esperado (consulta) | Origen |
|---|---|---|
| `identificacion.ambiente` | `01` si la BD es `SVMPROD`, si no `00` | `UT_VEN_GENERAL_PKG.RETORNA_AMBIENTE_MINH_F` |
| `identificacion.horEmi` | `DOC_EXTENDIDO.FEC_AUTORIZA − 1 h` (HH:MM:SS), con tolerancia de 120 s | El generador usa `SYSDATE − 1 h` segundos antes de autorizar |
| `identificacion.tipoMoneda` | `VENTASSAL.CONTROL.MONEDA` (`TIPO = '00'`) | Tag `<moneda>` del generador |
| `resumen.descuNoSuj`, `descuExenta` | `0` (CCF, FAC, REM, NCR, NDB) | `XML_CREA_DESCUENTOCARGOS_SV_P` |
| `resumen.descuGravada` | `0` | ídem |
| `resumen.porcentajeDescuento` | `0` (no en NCR/NDB) | ídem |
| `resumen.totalNoGravado` | `0` (no en REM) | `XML_INFOGENERAL_SV_P` (`totalSinImpuestos`) |
| `resumen.saldoFavor` | `0` (CCF, FAC, FAC_EXP) | Lo completa el proveedor (no viene en el XML TQ) |
| `resumen.totalLetras` | `VALOR_EN_LETRAS_F(total) CON cc/100` + moneda del grupo de compañías | `OBTENER_TOTALENLETRAS_SV_F`; NC/ND: retención 1 % sin redondear |
| `cuerpoDocumento.uniMedida` | `59` | `XML_ITEMS_*_SV_P` |
| `cuerpoDocumento.noGravado` | `0` (no en REM) | `XML_CREA_ITEMS_SV_P` (`baseImponible[4]`) |
| `cuerpoDocumento.psv` | `0` (CCF, FAC) | Lo completa el proveedor |

En Factura de Exportación también se obtienen: `emisor.tipoItemExpor` (`FE_PARAMETROS`, `TIPO_ITEM`), `receptor.tipoDocumento`
(`37`, fijo en el generador) y `resumen.observaciones` (`DOC_EXP.OBSERVACIONES` más el texto de aduana cuando hay fletes o
seguros con Incoterm CIF, CIP o DAP; `.` si queda vacío). `resumen.totalNoOnerosas` no tiene fuente en la BD.

`ambiente` refleja la BD consultada: un documento emitido en producción y copiado a pruebas sale como diferencia real.
Para `resumen.totalLetras` la consulta llama a la misma función del generador
(`VENTASAPL.LN_VEN_ENTIDADES_PKG.VALOR_EN_LETRAS_F`), que no escribe en la BD (la sesión es de solo lectura).

---

## 7. Mantenimiento

| Si cambia… | Hacer |
|---|---|
| El libro Excel | `build` → `audit-plano` (0 errores) → revisar la hoja `lint` → `pytest -q` |
| El usuario indica una excepción a lo que dice el libro | Agregar el ajuste en `data/specs/ajustes_contrato.json` → `build` → `audit-plano` (0 errores, ajuste como `AJUSTE_APLICADO`) |
| Una regla de la BD (columna, fórmula, catálogo) | Editar `mapeo_bd.xlsx` (categoría, fuente, transformación, **aprobado**) y ajustar `queries/valores_esperados.sql` en la misma ruta; volver a ejecutar el query y la validación sobre documentos conocidos |
| Una normalización aceptada | Editar `normalizaciones_bd.json` (BD) o `pdf_presentacion.json` (PDF). Solo con aprobación: el validador no acepta equivalencias que no estén escritas ahí |
| Un campo pasa a tener fuente | Cambiar su fila de `SIN_FUENTE` a `DIRECTA`/`TRANSFORMACION`/`CATALOGO` con `aprobado = SI` y agregarlo al query |

Categorías de `mapeo_bd.xlsx`:

- `DIRECTA`: columna tal cual.
- `TRANSFORMACION`: fórmula (concatenar, quitar guion, redondear…).
- `CATALOGO`: búsqueda en tablas de parámetros.
- `SIN_FUENTE`: no existe en la BD → `DB_MISSING`.
- `NO_APLICA`: eliminado en el contrato.

Solo las filas con `aprobado = SI` se comparan.

Pruebas automáticas (deben quedar en verde tras cualquier cambio; usan datos sintéticos):

```powershell
python -m pytest -q
```

---

## 8. Reglas de seguridad

- Contra la BD: **solo lectura** (la automatización lo fuerza: solo `SELECT`, sesión `READ ONLY`, rollback); usar un
  usuario de base de datos de solo lectura.
- Usuario y contraseña de Oracle van únicamente en `.env` (ignorado por git) o en variables del sistema;
  `config/oracle.local.json` rechaza la contraseña si se escribe ahí. No compartir ni copiar `.env`.
- No leer ni exportar las filas de credenciales de `VENTASTQ.FE_PARAMETROS`.
- Credenciales fuera del repositorio. Los XML de ejemplo traen `password` y `token` en `CabeceraEntrada`: no
  difundirlos y rotarlos si fueron reales.
- Datos reales (NIT, DUI, nombres, montos) no van a fixtures, tickets ni logs compartidos.

---

## 9. Problemas comunes

| Síntoma | Causa probable | Solución |
|---|---|---|
| `PDF_TOOL_MISSING` | No se encuentra `pdftotext` | Instalar Xpdf/Poppler o definir `PDFTOTEXT` |
| `NO VERIFICADA … Falta la contraseña` | `CV_ORACLE_PASSWORD` vacío en `.env` (o el comando se ejecutó fuera de la raíz del proyecto) | Completar `.env` (Paso 2) y ejecutar desde la raíz |
| `NO VERIFICADA … Falta el usuario` | `CV_ORACLE_USER` vacío en `.env` | Completarlo |
| `NO VERIFICADA … No se pudo conectar a Oracle` | Sin VPN/red, host/puerto/SID errados o credenciales inválidas | Revisar VPN y `config/oracle.local.json`; probar con `fetch-expected` |
| Todos los documentos con `DB_MISSING` ERROR | Los documentos son de otro ambiente (p. ej. producción vs pruebas) o, con `--expected`, el archivo no los incluye | Apuntar a la BD correcta (`--db-config`) o regenerar el archivo |
| `faltan columnas en la fuente de valores esperados` | El export no tiene `CODIGO_GENERACION`, `RUTA`, `VALOR` en la primera fila | Exportar con encabezados |
| `valor duplicado para … ` | El export mezcla dos ejecuciones con valores distintos | Exportar una sola ejecución |
| Muchos `DB_MISMATCH` en documentos viejos | Versión anterior del esquema o datos maestros editados | Ver §5 "Situaciones conocidas" |
| `ORA-01489` o resultados truncados al explorar | Concatenaciones largas en DBeaver | Usar el query provisto tal cual |

---

## Anexo A — Validaciones sueltas (sin BD)

Para revisar solo estructura, cruce y PDF, sin consultar la base:

```powershell
python -m contract_validator <archivos o carpetas> --out reports/resultado.xlsx
```

- Acepta archivos sueltos, comodines o carpetas (busca `.json`, `.xml` y `.pdf`).
- El tipo de documento se detecta por el contenido; si no, por el nombre del archivo; o se fuerza con `--doc-type CCF`.
- XML y JSON del mismo documento se emparejan solos por **Código de Generación**. Para forzar un par:
  `--pair doc.xml doc.json`; para un PDF: `--pair-pdf doc.xml doc.pdf`.
- Solo se validan los PDF **MINH** (oficiales de Hacienda: `*_SAL_MINH_*`); los internos (`*_SAL_<fecha>_*`) se ignoran.
- Opciones: `--no-cross` (sin cruce XML↔JSON), `--no-pdf` (sin PDF), `--format json`.

```powershell
python -m contract_validator entrada/ --no-cross --no-pdf --format json > reports/estructura.json   # solo estructura
python -m contract_validator a.xml b.json c.pdf --pair a.xml b.json --pair-pdf a.xml c.pdf        # pares manuales
```

Ejemplo de salida (real, recortada):

```text
[OK] entrada\53_16_01_001_131126_020656_...xml  [CCF/XML] (campos del contrato: 79, errores: 0, avisos: 0)
[FAIL] entrada\13062C60-FF5A-BECC-E063-9B78A8C07752.json  [CCF/JSON] (campos del contrato: 127, errores: 1, avisos: 0)
   ERROR   REQUIRED_MISSING       DatosEntrada.Items.TotalImpuestoItem.tipoImpuesto: Campo requerido no encontrado (cambio: SIN_CAMBIO, XML fila 131)
[OK] entrada\CCF_SAL_MINH_531601001131126020656.pdf <-> entrada\53_16_01_001_...xml  [CCF/PDF(MINH)<->XML] (errores: 0, avisos: 3)
```

Cuando el hallazgo es de un campo del libro, el mensaje cita su origen, p. ej. `(... XML fila 156)`: fila de la hoja
`XML` del libro donde está definido el campo.

## Anexo B — Valores esperados por archivo (modo manual o extracción por rango)

Cuando no se quiere consultar en línea durante la verificación (por ejemplo, para congelar los valores de un corte):

```powershell
# Extraer por documentos o por rango de fecha de autorización (tope de documentos)
python -m contract_validator fetch-expected --codigo-generacion <UUID> [<UUID> ...] --out reports/esperados.tsv
python -m contract_validator fetch-expected --desde 2026-11-01 --hasta 2026-11-15 --max-docs 200 --out reports/esperados.tsv

# Verificar usando ese archivo en lugar de consultar la BD
python -m contract_validator verificar entrada/ --expected reports/esperados.tsv
```

Si no hay acceso desde la máquina que ejecuta el validador, el mismo `queries/valores_esperados.sql` puede ejecutarse
en DBeaver (parámetros `:codigo_generacion`, `:desde`, `:hasta` en `YYYY-MM-DD`, `:max_docs`) y exportarse con
encabezados a TSV/CSV/XLSX (columnas `CODIGO_GENERACION | TIPO_DOC | RUTA | VALOR`); la grilla de DBeaver muestra 200
filas por defecto, así que conviene usar *Exportar datos*.

Contenido de ejemplo del archivo (primeras filas de un CCF):

```text
CODIGO_GENERACION                       TIPO_DOC  RUTA                               VALOR
13062C60-FF5A-BECC-E063-9B78A8C07752    CCF       cuerpoDocumento.codigo[1]          2026770
13062C60-FF5A-BECC-E063-9B78A8C07752    CCF       identificacion.numeroControl       DTE-03-M001P001-000000000002561
13062C60-FF5A-BECC-E063-9B78A8C07752    CCF       resumen.totalPagar                 1796.13
```

**Cómo arma el query cada valor (detalle en `mapeo_bd.xlsx`):** clave `DOC.PREIMPRESO` = Código de Generación; tipo por
`DOC.TIPO_DCTO_EMITIR` + `DOC.CPTO` (cpto 055 = REM); encabezado de `DOC`, `DOC_EXTENDIDO`, `CLIENTES`, `CIA`, `DOC_EXP`,
`VENTASTQ.FE_PARAMETROS` y `PAIS`; ítems de `DET_VENTAS` (`CANTDESP <> 0`, orden por `ARTICULO`) con descripción de `ART`
/ `ART_PAISES` + lotes de `LOTES_ART`; arreglos con sufijo `[n]` y apéndice por nombre (`apendice.valor[apendice02]`).
`VENTASTQ.FE_PARAMETROS` guarda credenciales: el query la filtra por nombre de parámetro y nunca las lee.
