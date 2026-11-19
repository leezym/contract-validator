---
name: validate-contract
description: Valida la completitud y correctitud de archivos .xml, .json y PDF de DTE contra un contrato en .xlsx (campos nuevos, eliminados, renombrados, obligatoriedad, cruce XML/JSON, PDF) y su contenido contra los valores reales de la base de datos Oracle, que consulta ella misma en solo lectura. Úsala cuando el usuario pida validar, comparar o auditar payloads/mensajes XML o JSON o PDF contra la especificación en Excel o contra la BD.
---

# validate-contract

Compara `.xml` / `.json` contra el contrato de `data/specs/contrato_ejemplo.xlsx`. Los tags reales salen de las hojas
**`XML`** (una columna por tipo: CCF, FAC, FAC_EXP, NCR, NDB, REM, RETORNO con `XML-ER`) y **`JSON`** (base v3, verificación);
INVALIDACION sale de su hoja propia. El libro se convierte primero a una hoja plana. `Eliminar` en una columna `TAG_MIN_HACIENDA`
(hoja `XML` u hoja del tipo): el campo no va en el JSON y su tag XML es opcional.

## REGLA OBLIGATORIA: siempre comparar contra la BD

Cada vez que el usuario pida validar, verificar, revisar o auditar archivos (XML, JSON o PDF, aunque solo mencione un PDF o la estructura),
**ejecutar siempre `verificar`**, que incluye la comparación contra la BD Oracle. **No usar** `validate` ni el comando sin subcomando
(no consultan la BD) salvo que el usuario pida expresamente omitirla. Nunca entregar un resultado "solo estructura/PDF" sin la BD:
si la BD no se pudo consultar, reportarlo como `NO VERIFICADA`, indicar el motivo (usuario/contraseña en `.env`, VPN, `config/oracle.local.json`)
y ofrecer reintentar. Si el usuario entrega una carpeta (p. ej. `entrada/`), pasarla completa a `verificar`.

## Flujo principal (lo que se usa normalmente)

El usuario entrega los PDF, XML y JSON de uno o varios documentos. La automatización hace TODO con un comando:

```
PYTHONPATH=src python -m contract_validator verificar <archivos y/o carpetas>... [--out reports/<nombre>.xlsx]
```

1. Valida la **estructura** de cada JSON/XML contra el contrato `.xlsx` (y la coherencia XML↔JSON y PDF↔XML).
2. Toma los Códigos de Generación de los archivos, **consulta la BD Oracle** en solo lectura (`queries/valores_esperados.sql`)
   y **compara** el contenido de JSON, XML y PDF contra esos valores.
3. Termina con el **VEREDICTO POR DOCUMENTO** (consola y hoja `Veredicto` del Excel, que sale primero):
   - `Estructura vs contrato (.xlsx)`: CUMPLE / NO CUMPLE
   - `Información vs base de datos`: CORRECTA / INCORRECTA / NO VERIFICADA (con el motivo)
   - `Coherencia XML / JSON`: OK / CON DIFERENCIAS / NO APLICA
   - `Validación del PDF`: CUMPLE / NO CUMPLE / NO VERIFICADO / NO RECIBIDO / IGNORADO / NO VALIDADO
   - `RESULTADO`: CUMPLE / NO CUMPLE / INCOMPLETO
   Exit code: `0` todo CUMPLE, `1` algún NO CUMPLE, `2` nada incumple pero algo quedó NO VERIFICADO.
4. Si no hay conexión a la BD (falta usuario, contraseña o VPN) **no se aborta**: la estructura se valida igual y la
   información queda `NO VERIFICADA` con el motivo; indicar al usuario qué falta (ver "Valores esperados desde la BD").
5. Responder al usuario con el veredicto de cada documento (las dos preguntas: ¿la estructura cumple el .xlsx? ¿la
   información es correcta según la BD?), los errores principales agrupados por código y la ruta del reporte. Por defecto
   el reporte va a `reports/verificacion_<fecha_hora>.xlsx`.

## Pasos de preparación y uso avanzado

1. Si `data/specs/contrato_plano.xlsx` no existe o el libro cambió:
   `PYTHONPATH=src python -m contract_validator build --source data/specs/contrato_ejemplo.xlsx --out data/specs/contrato_plano.xlsx`
   y revisar la hoja `lint` (nunca editar el libro original sin que lo pidan).
   `build` aplica además los ajustes aprobados de `data/specs/ajustes_contrato.json` (hoja `ajustes`). **Nunca editar
   `contrato_plano.xlsx` a mano**: las reglas salen del libro; un JSON, XML o PDF real nunca justifica cambiarlas (si difiere, es un hallazgo). Solo si el usuario indica una excepción, agregar el
   ajuste (id, accion `modificar`/`agregar`, tipo_doc, formato, ruta con `/`, campos, motivo, evidencia, aprobado) y
   regenerar. No agregar ajustes sin aprobación explícita. `tipo_doc` acepta un tipo, una lista o `*` (todos).
   Antes de regenerar, comparar el plano vigente con lo que generaría `build`: si el usuario lo editó a mano, convertir
   esas ediciones en ajustes (preguntando el motivo) para no perderlas. Si el plano está abierto en Excel/LibreOffice
   (`PermissionError` o archivo `.~lock.contrato_plano.xlsx#`), pedir que lo cierre.
   Reglas vigentes: `catalogos_como_string` (campos con lista de valores → `string`, sin `VALUE_NOT_IN_ENUM`) y `decimales_equivalentes` (`285.6` = `285.60` en el cruce XML↔JSON). En lo demás el plano sigue fielmente el libro, incluidos sus cruces.
1b. Tras generar el plano: `PYTHONPATH=src python -m contract_validator audit-plano --fail-on-error` (debe dar 0 ERROR; reporte en `reports/auditoria_plano.xlsx`).
2. Primera vez: `pip install -r requirements.txt`.
3. Solo si el usuario pide omitir la BD: `PYTHONPATH=src python -m contract_validator <archivo|carpeta|glob>... [--doc-type NDB] --format json --out reports/<nombre>.xlsx`
   (el tipo se detecta por `identificacion.tipoDte`, por el nombre del archivo o con `--doc-type`).
3b. `verificar` = `validate --expected-db` + reporte por defecto + veredicto. Con `validate` también puede usarse
   `--expected <export.tsv>` (valores de un archivo en vez de consultar) o `--expected-db`.
4. Resumir por archivo OK/FAIL y los hallazgos agrupados por código; indicar la ruta del reporte. Exit code `0` = OK,
   `1` = errores de validación, `2` = no se pudo consultar la BD.

## Formato plano (hoja `contrato`)

`tipo_doc, formato(JSON|XML), ruta, tipo, valores, requerido(SI|NO|SI_NULL), cambio, ruta_anterior, ruta_contraparte, descripcion, hoja_origen, fila_origen, nota`.
JSON se valida contra la columna v3→v4; XML contra `XML_TQ`. El formato simple antiguo
(`campo|tipo|requerido|cambio|campo_anterior`) sigue aceptándose.

## Reglas

| Cambio | Se espera en el archivo |
|---|---|
| NUEVO | presente (si requerido) y con el tipo correcto |
| ELIMINADO | **ausente** |
| RENOMBRADO | nombre nuevo presente y nombre anterior **ausente** |
| SIN_CAMBIO | presente si requerido, tipo correcto |
| (no listado) | `UNEXPECTED_FIELD` (ERROR; WARNING con `--allow-extra`) |

`SI` = presente y no nulo; `SI_NULL` = presente, puede ser `null`; `NO` = opcional. Si un contenedor está `null`/vacío u
opcional y ausente, no se exige su contenido. En XML: se retira la envoltura SOAP, los tags de `CabeceraEntrada` son los de la hoja `XML` (`Solo XML` = sin campo JSON; nunca mostrar `usuario`, `password` ni `token`),
el texto `null` es valor nulo y un archivo no-UTF-8 se lee como cp1252 con aviso `ENCODING`.

## Principio: solo lo que definen los specs

Se valida únicamente lo definido por `data/specs/contrato_ejemplo.xlsx` y los ejemplos entregados. No agregar reglas por criterio propio
(formatos, totales, equivalencias, heurísticas); si algo no está en los specs, reportarlo como hallazgo del contrato y preguntar.
Los tipos de dato se validan con la columna `Tipo` (`TYPE_MISMATCH`). Única regla de contenido propia del Excel: número de control `DTE-##-<bloque>-<15 dígitos>` con `##` = tipo de documento.

## Cruce XML <-> JSON

Si entre los archivos hay un XML y un JSON del mismo documento (mismo UUID) se cruzan solos; también `--pair XML JSON` (`--no-cross` lo desactiva).
Igualdad exacta de texto (`null`/vacío = sin valor); los números con el mismo valor decimal (`285.6`, `285.60`) son iguales; otras diferencias de formato (`05` vs `5`, fechas) salen como aviso `XREF_FORMAT_DIFF`.
Una diferencia puede ser del documento o del mapeo del Excel: reportar ambas posibilidades. Detalle en `README.md`.

## PDF vs XML

De los `.pdf` de `data/specs` solo se validan los oficiales `XXX_SAL_MINH_...` (el interno `XXX_SAL_<fecha>_...` se ignora) contra el XML con el mismo Código de Generación:
`python -m contract_validator <pdfs> <xml> [--pair-pdf XML PDF] [--no-pdf] [--pdf-config ...]`. Necesita `pdftotext` (o `PDFTOTEXT=<ruta>`).
Capas: identidad (uuid, número de control, título) → valores anclados en las `Descripcion` del Excel, `Notas` del XML e ítems → contención (valor del XML en el PDF y viceversa).
Las normalizaciones aceptadas viven en `data/specs/pdf_presentacion.json` (propuesta pendiente de aprobación): no agregar normalizaciones por cuenta propia; ante `PDF_NOT_PRINTED`
sugerir al usuario añadir el tag a `no_impreso` solo si el PDF realmente no lo imprime. Un PDF sin XML pareja es `PDF_NO_PAIR` (aviso), no un error de contenido. Detalle en `README.md`.

## Auditoría del mapeo

`PYTHONPATH=src python -m contract_validator audit-mapping [--doc-type X] [--out reports/auditoria_mapeo.xlsx]` reporta hechos del contrato sin
archivos reales: `NO_XML_SOURCE`, `MULTI_SOURCE`, `XML_NO_DEST`, `DIFF_ACROSS_TYPES`, `POSITIONAL`. Resumir por código y tipo; no modificar el Excel sin que lo pidan.

## Valores esperados desde la BD (Fase B)

`queries/valores_esperados.sql` (Oracle, solo `SELECT`, parámetros `:codigo_generacion`, `:desde`, `:hasta`, `:max_docs`) devuelve
`CODIGO_GENERACION | TIPO_DOC | RUTA | VALOR`; `load_expected(archivo_exportado)` lo convierte en `{uuid: {ruta: valor}}`.

**Cómo obtenerlos (módulo `db_fetch.py`, driver `oracledb` en modo thin):**

1. Verificar la conexión antes de consultar:
   - `config/oracle.local.json` (ignorado por git) con `host`, `port` y `sid` o `service` (sin usuario ni contraseña).
     Si no existe, copiar `config/oracle.example.json` y pedir al usuario los datos que falten (nunca inventarlos).
   - Usuario y contraseña en el archivo **`.env` de la raíz del proyecto** (ignorado por git; plantilla `.env.example`):
     `CV_ORACLE_USER=...` y `CV_ORACLE_PASSWORD=...`. Las variables reales del proceso tienen prioridad sobre `.env`.
   - Comprobar que estén completos **sin mostrar valores**:
     `python -c "import sys; sys.path.insert(0,'src'); from contract_validator.db_fetch import load_db_config as L; c=L(); print('user:', bool(c.get('user')), 'password:', bool(c.get('password')))"`.
     Si faltan, pedir al usuario que complete `.env` él mismo; **nunca** pedir la contraseña en el chat, no leer ni
     mostrar el contenido de `.env`, no escribir la contraseña en archivos ni en comandos.
   - Requiere red/VPN hacia el servidor; si la conexión falla, la información queda `NO VERIFICADA` con el motivo.
2. Elegir el modo:
   - Documentos en mano (lo normal): `python -m contract_validator verificar <archivos> [--save-expected reports/esperados.tsv]`
     (consulta solo los UUID de esos archivos, uno por uno).
   - Solo extraer: `python -m contract_validator fetch-expected --codigo-generacion <UUID> [<UUID> ...] --out reports/esperados.tsv`
     o por rango `fetch-expected --desde 2026-11-01 --hasta 2026-11-15 --max-docs 100`; luego `validate ... --expected reports/esperados.tsv`.
3. Salvaguardas (no desactivarlas): el SQL se rechaza si no es una única consulta SELECT/WITH; la sesión abre con
   `SET TRANSACTION READ ONLY`; al final rollback y cierre; tope de documentos. Para un rango, confirmar con el usuario
   las fechas y el tope antes de ejecutar contra una base real.
4. Los `.tsv` y reportes contienen datos reales: quedan en `reports/` (ignorado por git); no adjuntarlos a tickets ni
   copiarlos a fixtures.
El mapeo aprobado está en `data/specs/mapeo_bd.xlsx` y las normalizaciones en `data/specs/normalizaciones_bd.json`; no añadir
reglas ni equivalencias fuera de esos archivos sin aprobación. Antes de ejecutar contra una base real, confirmar con el usuario
y limitar documentos. `VENTASTQ.FE_PARAMETROS` guarda credenciales: filtrar siempre por `PARAMETRO` y no leerlas.
Datos reales (NIT, nombres, montos) no van a fixtures ni reportes compartidos; los tests usan datos sintéticos.
Comparar contra la BD: `validate <archivos> --expected <export.tsv>` (módulo `db_check.py`). Al resumir, separar
`DB_MISMATCH_*` (diferencias reales) de `DB_MISSING` aviso (campos sin fuente en BD, no verificables); un documento
emitido con una versión anterior del esquema mostrará campos del contrato como faltantes.

## Códigos de hallazgo

Contrato: `NEW_MISSING`, `REQUIRED_MISSING`, `RENAMED_MISSING`, `XML_OPTIONAL_MISSING` (aviso: tag XML opcional del libro que no llegó), `NULL_NOT_ALLOWED`, `TYPE_MISMATCH`, `REMOVED_STILL_PRESENT`, `OLD_NAME_PRESENT`, `UNEXPECTED_FIELD`,
`VALUE_NOT_IN_ENUM` (aviso), `FORMAT_INVALID`/`CONSISTENCY_ERROR` (número de control), `DOCTYPE_UNKNOWN`, `DOCTYPE_NAME_MISMATCH` (aviso),
`ENCODING` (aviso), `PARSE_ERROR`. PDF: `PDF_EMPTY`, `PDF_TOOL_MISSING`, `PDF_IDENTITY`, `PDF_DOCTYPE`, `PDF_MISMATCH`, `PDF_VALUE_MISSING`, `PDF_COUNT_MISMATCH`, `PDF_NO_PAIR`, `PDF_NOT_PRINTED`, `PDF_VALUE_NOT_IN_XML`, `PDF_IGNORED`. BD: `DB_MISMATCH_JSON`, `DB_MISMATCH_XML`, `DB_MISMATCH_PDF`, `DB_MISSING`, `APENDICE_NOT_RECEIVED` (aviso; cada apéndice se compara completo). Cruce: `XREF_MISMATCH`, `XREF_FORMAT_DIFF` (aviso), `XREF_MISSING`, `XREF_COUNT_MISMATCH`, `XREF_DOCTYPE`.

Si hay `UNEXPECTED_FIELD` que parece un renombrado no declarado, sugerirlo como fila RENOMBRADO en el libro; no modificarlo sin que lo pidan.
