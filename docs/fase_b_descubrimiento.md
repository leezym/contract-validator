# Fase B — Descubrimiento de la fuente de valores esperados

Fecha: 2026-11-16. Exploración **solo lectura** (consultas de catálogo y `SELECT` con `FETCH FIRST`) desde DBeaver,
conexión `svmtest` (Oracle, ambiente de pruebas). Sin datos de clientes en este documento.

## Decisión del usuario

Fuente de referencia = **tablas de negocio** (opción B), no el JSON certificado guardado por el generador
(`VENTASSAL.DOC_XML_ENTIDAD_FACTEL.JSON`). Ese JSON se usa solo como ayuda para construir el mapeo.

## Esquema relevante (El Salvador = esquema `VENTASSAL`, CIA `53`; Teramed = `VENTASSVC`, CIA `41`)

Cada compañía tiene su propio esquema con las mismas tablas; el query usa `{ESQUEMA_VENTAS}` y `:cia`, y la relación
CIA → esquema está en `data/specs/companias.json`. Abajo dice `VENTASSAL`, pero aplica igual al esquema de cada compañía.

Clave natural de documento en todas las tablas: `CIA, SUC, BODEGA, CPTO, FUENTE, DCTO`.

| Tabla | Uso | Notas |
|---|---|---|
| `VENTASSAL.DOC` | Encabezado | 190 columnas. **`PREIMPRESO` = Código de Generación** (coincide en el 100 % de los documentos certificados). `TIPO_DCTO_EMITIR`: F, CCF, FE, NC, ND. Totales: `VALOR`, `IVA`, `RETENCION`, `RESTOPAGO`, `VALOR_GRAVADAS/EXENTAS/NO_SUJETAS`, `IVA_PERCIBIDO`. |
| `VENTASSAL.DET_VENTAS` | Ítems (`cuerpoDocumento` / `Items`) | `ARTICULO`, `CANTPED/CANTDESP`, `PRECIO`, `VALOR`, `IVA`, `DESCUENTO` (NUMBER(19,1)), `ORDEN_FACTURAS`. |
| `VENTASSAL.DOC_EXTENDIDO` | Datos FE | `CONSECUTIVO` (número de control), `AUTORIZACION` (sello MH), `CLAVE_ACCESO` (= Código de Generación), `CLI_COND_PAGO`, `CLI_MEDIO_PAGO`, `CAUSA_ANUL`. |
| `VENTASSAL.CLIENTES` | Receptor | `NIT`, `NRC`, `NOMBRE`, `RAZON`, `TELEFONO`, `GIRO` (`codActividad-descActividad`), `LOCALIDAD` (`DDMMMDDD`), `DIRECCION`, `BARRIO`. Clave `CLIENTE` (+`SUC`). |
| `VENTASSAL.CIA` | Emisor | `NIT`, `NRC`, `NOMBRE`, `DIRECCION`, `TELEFONO`, `EMAIL`. |
| `VENTASSAL.DOC_EXP`, `CLIENTES_EXP` | Exportación (FEX) | Incoterm, puertos, consignatario. |
| `VENTASTQ.DOC_DCTO_AFECTADO` | Documento relacionado (NCR/NDB) | `CUFE_AFECTADO`, `DCTO_AFECTADO`, `FECFACTURA_AFECTADO`. |
| `VENTASTQ.DOC_UUID` | Puente UUID ↔ clave | No necesario: `DOC.PREIMPRESO` ya da la unión. |
| `VENTASSAL.DOC_XML_ENTIDAD_FACTEL` | JSON/XML certificado por MH | Solo ayuda de mapeo; no es la referencia. |

## Tipo de DTE

| `DOC.TIPO_DCTO_EMITIR` | `CPTO` | tipoDte | Tipo contrato |
|---|---|---|---|
| F | 001 | 01 | FAC |
| CCF | 001 | 03 | CCF |
| CCF / F | 055 | 04 | REM |
| NC | 105, 108, 109, 115 | 05 | NCR |
| ND | 010, 015 | 06 | NDB |
| FE | 001, 056 | 11 | FAC_EXP |

**INVALIDACION y RETORNO:** no se encontraron en `VENTASSAL`/`VENTASTQ` (ni en `DOC`, `DOC_DCTO_AFECTADO`,
`DOC_EVENTOS_CENDIS`, `LISTAEMPAQUE_EVENTOS`). El XML de retorno de ejemplo (`53_16_01_105_009867_000002…`) no tiene
fila en `DOC`. Pendiente que el usuario indique la fuente.

## Borrador de mapeo (`data/specs/mapeo_bd.xlsx`)

Generado comparando reglas contra el JSON certificado de 7 documentos reales (CCF ×2, FAC, REM, NCR, NDB, FEX):
**330 de 334** comprobaciones regla×muestra coinciden (las 4 restantes son `identificacion.version`: documentos
emitidos con una versión anterior a la vigente del parámetro). 600 filas (ruta × tipo): 122 DIRECTA,
152 TRANSFORMACION, 82 CATALOGO, 190 SIN_FUENTE, 54 NO_APLICA.

Decisiones del usuario (2026-11-16): INVALIDACION y RETORNO fuera de alcance por ahora; TRANSFORMACION aprobadas en
bloque; tolerancia de 0.01 en montos (`data/specs/normalizaciones_bd.json`); buscar en el esquema las tablas para lo
que no tenía fuente → **CATALOGO** (pendiente de aprobación):

- `VENTASTQ.FE_PARAMETROS` (`PAIS='SV'`, `ENTIDAD_FACTEL='MINHACIENDA'`): parámetros del generador —
  `ACTIVIDAD_ECONOMICA_EMISOR`, `DEPARTAMENTO`/`MUNICIPIO`/`DISTRITO` (por `CODIGO_TQ`=CIA), `COMPLEMENTO_CODESTABLECIMIENTO`
  (M) y `COMPLEMENTO_PUNTOEMISION` (P), `VERSION_JSON`, `TIPO_MODELO`, `TIPO_TRANSMISION`, `TIPO_ITEM`,
  `IMPUESTO_REGISTRA_FE`, `TIPO_IDENTIFICACION`, `TIPO_PERSONA`, `INCOTERMS`, `CODIGO_PLAZO`. **Ojo:** la tabla también
  guarda credenciales (`PASSWORD_TOKEN`, `CLAVE`, `USUARIO_TOKEN`); el query debe filtrar por `PARAMETRO` y nunca leerlas.
- `VENTASSAL.PAIS.COD_MINHACIENDA` / `NOMBRE`: país del receptor (FEX).
- `VENTASSAL.SRI_PUNTO_EMISION.CODIGO_PUNTO_EMIS`: `emisor.codPuntoVenta`.

Sin fuente en BD (quedan como `DB_MISSING`): `emisor.code`, `apendice05+`, descripción larga de ítems FAC y los
campos nuevos de NCR/NDB sin muestra. Resueltos el 2026-11-17 desde el código del generador (7/7 JSON certificados):
`ambiente` (`DB_NAME = 'SVMPROD'` → `01`), `horEmi` (`FEC_AUTORIZA − 1 h`, tolerancia 120 s), `tipoMoneda`
(`CONTROL.MONEDA`), `totalLetras` (`VALOR_EN_LETRAS_F` + moneda de `GRUPOS_COMPANIAS`), `uniMedida` (`59`), y los
fijos en 0 `descuNoSuj`, `descuExenta`, `descuGravada`, `porcentajeDescuento`, `totalNoGravado`, `saldoFavor`,
`noGravado`, `psv` (por tipo).

Reglas confirmadas con evidencia:
- **Ítems:** `DET_VENTAS` con `CANTDESP <> 0`, en el orden del generador (`VENTASAPL.IN_VEN_FACTURA_ELECTRONIC3_PKG`):
  exportación (`XML_ITEMS_FACEXP_SV_P`) por `ORDEN_FACTURAS, ARTICULO`; demás tipos (`XML_ITEMS_SV_P`) por `PADRE, ORDEN_FACTURAS`
  (hasta 99 ítems). Antes se ordenaba por `ARTICULO` y en Teramed (CIA 41) cruzaba los ítems.
- **fecEmi:** `DOC_EXTENDIDO.FEC_AUTORIZA` (en REM `DOC.FECFACTURA` es otra fecha).
- **Dirección del receptor:** `CLIENTES.LOCALIDAD`/`DIRECCION`/`BARRIO` (`DOC.LOCALIDAD`/`DIRECCION` es el lugar de entrega).
- **NRC:** la BD lo guarda con guion (`289648-0`); el documento lo lleva sin guion.
- **Descripción del ítem:** `ART.DESCRIPCION || ' ' || ART.PRESENT` (FAC usa otro texto, sin ubicar).
- **FAC:** montos con IVA incluido, recalculados desde `*_ENT` (`precioUni = ROUND(PRECIO_ENT×1.13, 4)`).
- **Totales:** `totalPagar = montoTotalOperacion − RETENCION`; `DOC.RESTOPAGO` es saldo de cartera, no sirve.
  NCR/NDB: `montoTotalOperacion` ya va neto de retención; en NDB difiere 0.01 porque el generador usa la retención
  sin redondear.
- **Apéndice:** por campo (`apendice01`…`04`) con campos de `DOC`/`DOC_EXP`; `apendice05+` ("Mensaje") sin fuente.
  Fuente confirmada (2026-11-17) en el código del generador, leído de `ALL_SOURCE` (solo lectura):
  `VENTASAPL.IN_VEN_FACTURA_ELECTRONIC3_PKG.XML_NOTAS_SV_P` (exportación: `XML_NOTAS_EXP_SV_P`) define
  `ETIQUETA01..04` (texto fijo) y `APENDICE01..04` (concatenaciones de `FACTURA_ENC_TTEMP`, copia de `DOC`/`DOC_EXP`);
  `LN_VEN_FACTURA_ELECTRONIC4_PKG` arma `<Notas><campo>apendiceNN</campo><etiqueta>…</etiqueta><notasAdicionales>…`.
  El primer `<Notas>` sin campo lleva `DOC_EXTENDIDO`/observaciones o `.`; `apendice05+` sale de
  `IN_VEN_FACTURA_VENTA2_PKG.PBS_MENSAJES_FACTURA_P` (sin mapear). `VENTASTQ.XML_CONFIGURACION` solo tiene la
  configuración de Guatemala (GT). El query genera ahora `apendice.campo`, `apendice.etiqueta` y `apendice.valor`
  (incluido `apendice01`); coinciden 4/4 con el XML real y 3/3 con el JSON real.

Hallazgo aparte: el JSON certificado guardado en `DOC_XML_ENTIDAD_FACTEL.JSON` trae texto doble-codificado
(`Ã‰` en lugar de `É`); la BD de negocio está bien.

## Query y verificación (2026-11-16)

`queries/valores_esperados.sql` ejecutado en DBeaver (`svmtest`) para los 7 documentos: 1 566 filas, sin errores.
Cargado con `load_expected` y comparado con el JSON certificado: **1 393 / 1 514** valores coinciden. Diferencias:

- `cuerpoDocumento.descripcion` (114): el generador añade al texto de `ART` el lote/fechas (`| LOTE: … F: … V: … C: …`) y en
  exportación usa otra descripción (nombre largo + registro sanitario). Fuente del sufijo no ubicada.
- `identificacion.version` (3): documentos emitidos antes del cambio de versión del parámetro.
- `receptor.correo` (3): documentos viejos con el correo real del cliente; hoy `CLIENTES_ELECTRONICOS` en pruebas tiene el
  correo de prueba.
- `apendice04` peso (1): corregido (`TO_CHAR` sin cero inicial, `.11`), verificado de nuevo en BD.

CCF `13062C60…` contra sus archivos de `data/specs/`: `tipoDte`, `numeroControl`, `fecEmi`, NIT/NRC/actividad/municipio del
emisor, totales y los valores del ítem 2 coinciden con JSON y XML; en el PDF MINH aparecen con formato de impresión
(`13/11/2026`, `1,603.69`, `0501-091219-101-3`). `totalGravada`/`ivaRete` van en el XML a `OtrosTotales.valorOtrosTotales`
(posicional, no se compara 1 a 1) y `montoTotalOperacion` no tiene nodo en ese XML.

### Descripción del ítem: lotes y registro sanitario (resuelto)

Fuente del sufijo encontrada: **`VENTASSAL.LOTES_ART`** (clave del documento + `ARTICULO` → `NRO_LOTE`,
`FEC_FABRICACION`, `FEC_VENCIMIENTO`, `CANTIDAD`). Regla (en el query):

```
base  = ART.DESCRIPCION || ' ' || ART.PRESENT
        FAC_EXP (como IN_VEN_FACTURA_ELECTRONIC3_PKG.XML_ITEMS_FACEXP_SV_P + OBTENER_DESCRIPCION_ART_F):
          ART_PAISES con PAIS = DOC.PAIS_DESTINO (si es nulo, SUCURSALES.PAIS); descripción y presentación de
          ART_PAISES si está vigente a la fecha de autorización (o vencido con REGISTRO_EN_RENOVACION = 'X', salvo
          artículos PLU del cliente), si no de ART; + ' | Peso neto: <kg> Kilos' si el destino es Nicaragua (29)
lotes = por cada fila de LOTES_ART, en orden de NRO_LOTE:
        ' | LOTE: ' || UPPER(NRO_LOTE) || ['  F: ' || FEC_FABRICACION] || '   V: ' || FEC_VENCIMIENTO || '  C:' || CANTIDAD
        (fechas YYYY/MM/DD; el tramo F se omite si no hay fecha de fabricación)
FAC_EXP: || ' | Registro sanitario: ' || ART_PAISES.REGISTRO          (solo si hay registro)
         || ' | Titular: ' || titular, dirección y país de origen   (solo si hay)
         || ' | Fabricante: ' || NVL(ART_PAISES.FABRICANTE, ART.NOMBRE_FABRICANTE)
         || ' | ' || dirección del fabricante || ' | ' || país del fabricante   (solo si hay)
```

Medición 2026-11-19 (últimas 60 exportaciones por compañía, contra el JSON certificado): El Salvador 310/349 ítems,
Teramed 102/147 (antes 263 y 35). Lo que falta es histórico: maestro de artículos editado después de emitir, documentos
anteriores al peso neto y un prefijo `001; ` en algunas exportaciones a Nicaragua.

Medido en BD contra la descripción de **todos** los ítems certificados (sin contar tildes, por el doble encoding del
JSON guardado): CCF 15 541/17 160, NCR 23 149/24 343, NDB 133/142, REM 21 051/21 637, FAC_EXP 3 900/5 473; los 15 ítems
con varios lotes salen en el orden correcto. Las diferencias restantes son: maestro de artículos editado después de
emitir (≈3 600, p. ej. `CAJX30CAP` → `CJX30CAP`), segmentos opcionales sin fuente ubicada (`| Descuentos Ocasional`,
16 en CCF; prefijo `001; ` y `| Peso neto: … Kilos`, 433 en FAC_EXP), 904 FAC_EXP aún sin clasificar y textos de
licitación (FAC y algunos REM). En los 7 documentos de muestra, tras el cambio, el query coincide en **1 506 / 1 514**
valores; los 8 restantes son `version` y `receptor.correo` (datos que cambiaron después de emitir).

## Comparación implementada (`db_check.py`, `validate --expected`)

Corrida sobre los JSON certificados de los 7 documentos, el XML del CCF y los PDF MINH de CCF/REM/FEX:

- CCF y REM (versión 4 vigente): solo difiere `receptor.descActividad` (y en FEX `cuerpoDocumento.descripcion`) por el
  texto doble-codificado del JSON certificado guardado en BD (`FARMACÃ‰UTICOS`): defecto real de esos archivos.
- Documentos v1/v3: campos del contrato ausentes (`ivaRete` era `ivaRete1`; sin `distrito`, `observaciones`,
  `numDocumento`), además de `version` y `receptor.correo` (datos que cambiaron después de emitir).
- XML: `resumen.montoTotalOperacion` mapea a `ImportesTotal.totalFacturaImpuesto`, que el XML no trae. (Resuelto 2026-11-17: con la hoja `XML` como fuente del plano mapea a `totalBaseImpuesto`.) El apéndice no se
  compara: el contrato mapea `apendice.campo` ↔ `Notas.etiqueta`, pero el XML real lleva el nombre del campo en
  `Notas.campo` (y las etiquetas en `Notas.etiqueta`). Decisión 2026-11-18: manda el libro; esa diferencia se reporta como hallazgo.
- PDF MINH: CCF y REM coinciden. FEX imprime `recintoFiscal` y `regimen` como texto y la BD da el código: resuelto
  (2026-11-17) con catálogos en `pdf_presentacion.json` — `recintoFiscal` desde `FE_PARAMETROS.RECINTO_FISCAL` y
  `regimen` solo con el código visto impreso (`1000.000` → "Exportación Definitiva, Régimen Común"); la comparación de
  anclas acepta textos de varias palabras (también en `pdf_check`). Con eso el PDF FEX coincide 8/8 con la BD.

## Hallazgos iniciales (CCF de ejemplo, comparación de valores)

- Directos: emisor (`CIA.*`), receptor (`CLIENTES.*`), totales (`DOC.VALOR`→totalGravada/subTotal,
  `DOC.IVA`→tributos.valor, `DOC.RETENCION`→ivaRete, `DOC.RESTOPAGO`→totalPagar/montoPago, `DOC.PLAZO`→pagos.periodo),
  ítems (`DET_VENTAS.ARTICULO/CANTPED/PRECIO/VALOR`), `DOC_EXTENDIDO.CLI_COND_PAGO/CLI_MEDIO_PAGO`.
- Con transformación (requieren aprobación): `numeroControl` = `DTE-<tipoDte>-M001P001-` + `CONSECUTIVO` a 15
  dígitos; `codActividad`/`descActividad` = `CLIENTES.GIRO` partido por `-`; departamento/municipio/distrito desde
  `LOCALIDAD`; `complemento` = `DIRECCION` + `BARRIO`; `montoTotalOperacion` = `VALOR + IVA`; `montoDescu` =
  `cantidad × precio − valor` (`DET_VENTAS.DESCUENTO` tiene 1 decimal: 104.2 vs 104.16); `observaciones` y `apendice`
  son concatenaciones de campos de `DOC`.
- Sin columna encontrada: `horEmi` (la hora de `FECFACTURA` no coincide), `totalLetras`, `uniMedida`,
  códigos de tributo, `emisor.codActividad/descActividad`, `codEstable/codPuntoVenta`, `receptor.correo`.
- Orden de ítems: en el ejemplo coincide con `ARTICULO` ascendente y con `ORDEN_FACTURAS` descendente; falta confirmar
  con más documentos.

Actualización 2026-11-18 (FEX): se agregan `emisor.tipoItemExpor`, `receptor.tipoDocumento` y la fórmula completa de `resumen.observaciones` (883/883 JSON certificados). `resumen.totalNoOnerosas` sigue sin fuente.
