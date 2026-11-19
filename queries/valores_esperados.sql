-- Valores esperados por documento (Fase B). Oracle. SOLO LECTURA: únicamente SELECT.
-- Generado desde data/specs/mapeo_bd.xlsx (donde dice VENTASSAL: esquema de la compañía) (aprobado 2026-11-16; 13 campos antes sin fuente: 2026-11-17). Ver docs/fase_b_descubrimiento.md.
--
-- Salida (formato largo, una fila por valor):
--   CODIGO_GENERACION | TIPO_DOC | RUTA | VALOR
--   RUTA = ruta JSON del contrato plano ('.' como separador). Arreglos: sufijo [n] 1-based en el orden del documento
--   (ítems: DET_VENTAS con CANTDESP <> 0 en el orden del generador: FEX por ORDEN_FACTURAS; demás por PADRE, ORDEN_FACTURAS); apéndice: sufijo [apendiceNN] (nombre del campo).
--   Fila ausente = la BD no da valor (null o sin fuente -> DB_MISSING según mapeo_bd).
--   Números: punto decimal, sin ceros finales ni separador de miles (TO_CHAR FM).
--
-- Compañía: {ESQUEMA_VENTAS} es el esquema de la CIA del documento (data/specs/companias.json; p. ej. 53 -> VENTASSAL,
-- 41 -> VENTASSVC). db_fetch lo reemplaza antes de ejecutar; VENTASTQ, VENTASADM y VENTASAPL son comunes.
--
-- Parámetros (bind):
--   :cia                código de la compañía (VENTAS<xxx>.DOC.CIA)
--   :codigo_generacion  UUID de un documento (NULL = todos los del rango)
--   :desde, :hasta      fecha de autorización 'YYYY-MM-DD' (NULL = sin límite)
--   :max_docs           tope de documentos (NULL = 100)
--
-- Alcance: CCF, FAC, FAC_EXP, NCR, NDB, REM. INVALIDACION y RETORNO no están cubiertos.
-- Llama a VENTASAPL.LN_VEN_ENTIDADES_PKG.VALOR_EN_LETRAS_F (la misma función del generador; no escribe en la BD)
-- para resumen.totalLetras.
-- VENTASTQ.FE_PARAMETROS también guarda credenciales: el CTE prm filtra por PARAMETRO y nunca las lee.
WITH prm AS (
  SELECT parametro, codigo_tq, codigo_entidad_factel cod, descripcion des
    FROM ventastq.fe_parametros
   WHERE pais = 'SV' AND entidad_factel = 'MINHACIENDA'
     AND parametro IN ('VERSION_JSON', 'TIPO_MODELO', 'TIPO_TRANSMISION', 'ACTIVIDAD_ECONOMICA_EMISOR', 'DEPARTAMENTO',
                       'MUNICIPIO', 'DISTRITO', 'COMPLEMENTO_CODESTABLECIMIENTO', 'COMPLEMENTO_PUNTOEMISION',
                       'TIPO_IDENTIFICACION', 'TIPO_PERSONA', 'TIPO_ITEM', 'IMPUESTO_REGISTRA_FE', 'CODIGO_PLAZO', 'INCOTERMS')
),
k AS (
  SELECT d.cia, d.suc, d.bodega, d.cpto, d.fuente, d.dcto, d.preimpreso uuid, e.fec_autoriza fec_aut,
         CASE WHEN d.cpto = '055' THEN 'REM'
              ELSE DECODE(d.tipo_dcto_emitir, 'F', 'FAC', 'CCF', 'CCF', 'NC', 'NCR', 'ND', 'NDB', 'FE', 'FAC_EXP') END tipo_doc
    FROM {ESQUEMA_VENTAS}.doc d
    JOIN {ESQUEMA_VENTAS}.doc_extendido e ON e.cia = d.cia AND e.suc = d.suc AND e.bodega = d.bodega AND e.cpto = d.cpto AND e.fuente = d.fuente AND e.dcto = d.dcto
   WHERE d.preimpreso IS NOT NULL
     AND d.tipo_dcto_emitir IN ('F', 'CCF', 'NC', 'ND', 'FE')
     AND (:codigo_generacion IS NULL OR d.preimpreso = :codigo_generacion)
     AND d.cia = :cia
     AND e.fec_autoriza >= NVL(TO_DATE(:desde, 'YYYY-MM-DD'), DATE '2000-01-01')
     AND e.fec_autoriza <  NVL(TO_DATE(:hasta, 'YYYY-MM-DD'), DATE '2999-01-01') + 1
   ORDER BY e.fec_autoriza DESC
   FETCH FIRST NVL(:max_docs, 100) ROWS ONLY
),
it AS (
  SELECT k.uuid, k.tipo_doc, d.nrocompcrfiscal,
         -- orden de los ítems igual al generador (VENTASAPL.IN_VEN_FACTURA_ELECTRONIC3_PKG):
         --   exportación (XML_ITEMS_FACEXP_SV_P): ORDEN_FACTURAS, ARTICULO; demás (XML_ITEMS_SV_P): PADRE, ORDEN_FACTURAS
         ROW_NUMBER() OVER (PARTITION BY k.uuid ORDER BY CASE WHEN k.tipo_doc = 'FAC_EXP' THEN NULL ELSE v.padre END,
                                                         v.orden_facturas, v.articulo) n,
         v.articulo, v.valor, v.valor_exentas, v.valor_no_sujetas, v.iva_ent,
         CASE WHEN k.tipo_doc = 'FAC' THEN v.cantdesp_ent ELSE v.cantdesp END cant,
         CASE WHEN k.tipo_doc = 'FAC' THEN ROUND(v.precio_ent * (1 + v.porciva / 100), 4) ELSE v.precio END pu,
         -- descripción: base + lotes (LOTES_ART, por NRO_LOTE). Exportación, como el generador
         -- (IN_VEN_FACTURA_ELECTRONIC3_PKG.XML_ITEMS_FACEXP_SV_P + LN_VEN_FACTURA_ELECTRONIC4_PKG.OBTENER_DESCRIPCION_ART_F):
         --   descripción/presentación de ART_PAISES (país de destino del documento; si no hay, el de la sucursal) si está vigente a la fecha de autorización
         --   (o vencido con REGISTRO_EN_RENOVACION = 'X', salvo artículos PLU del cliente); si no, de ART;
         --   + peso neto (destino Nicaragua) + lotes + 'Registro sanitario', 'Titular', 'Fabricante', dirección y país del
         --   fabricante, cada uno solo si tiene valor.
         CASE WHEN k.tipo_doc = 'FAC_EXP' THEN
                CASE WHEN ap.descripcion IS NOT NULL AND (ap.fecvenc > k.fec_aut OR (ap.fecvenc <= k.fec_aut AND NVL(ap.registro_en_renovacion, 'NULO') = 'X'
                    AND NOT EXISTS (SELECT 1 FROM {ESQUEMA_VENTAS}.art_plu_clientes p WHERE p.articulo = v.articulo AND p.cliente = d.cliente AND p.fecvenc > k.fec_aut))) THEN ap.descripcion ELSE art.descripcion END
                || NVL2(CASE WHEN ap.present IS NOT NULL AND (ap.fecvenc > k.fec_aut OR (ap.fecvenc <= k.fec_aut AND NVL(ap.registro_en_renovacion, 'NULO') = 'X'
                    AND NOT EXISTS (SELECT 1 FROM {ESQUEMA_VENTAS}.art_plu_clientes p WHERE p.articulo = v.articulo AND p.cliente = d.cliente AND p.fecvenc > k.fec_aut))) THEN ap.present ELSE art.present END,
                        ' ' || CASE WHEN ap.present IS NOT NULL AND (ap.fecvenc > k.fec_aut OR (ap.fecvenc <= k.fec_aut AND NVL(ap.registro_en_renovacion, 'NULO') = 'X'
                    AND NOT EXISTS (SELECT 1 FROM {ESQUEMA_VENTAS}.art_plu_clientes p WHERE p.articulo = v.articulo AND p.cliente = d.cliente AND p.fecvenc > k.fec_aut))) THEN ap.present ELSE art.present END, NULL)
                || CASE WHEN d.pais_destino = '29' THEN ' | Peso neto: ' || TRIM(TO_CHAR(v.cantdesp * art.peso_neto, '999999990.000')) || ' Kilos' END
              ELSE art.descripcion || ' ' || art.present END
         || (SELECT LISTAGG(' | LOTE: ' || UPPER(l.nro_lote)
                            || NVL2(l.fec_fabricacion, '  F: ' || TO_CHAR(l.fec_fabricacion, 'YYYY/MM/DD'), '')
                            || '   V: ' || TO_CHAR(l.fec_vencimiento, 'YYYY/MM/DD') || '  C:' || l.cantidad, '')
                    WITHIN GROUP (ORDER BY l.nro_lote)
               FROM {ESQUEMA_VENTAS}.lotes_art l
              WHERE l.cia = v.cia AND l.suc = v.suc AND l.bodega = v.bodega AND l.cpto = v.cpto AND l.fuente = v.fuente AND l.dcto = v.dcto AND l.articulo = v.articulo)
         || CASE WHEN k.tipo_doc = 'FAC_EXP' THEN
                NVL2(ap.registro, ' | Registro sanitario: ' || ap.registro, NULL)
                || NVL2(TRIM(ap.nombre_titular_reg || ' ' || ap.direccion_titular_reg || ' ' || ap.pais_origen),
                        ' | Titular: ' || TRIM(ap.nombre_titular_reg || ' ' || ap.direccion_titular_reg || ' ' || ap.pais_origen), NULL)
                || NVL2(NVL(ap.fabricante, art.nombre_fabricante), ' | Fabricante: ' || NVL(ap.fabricante, art.nombre_fabricante), NULL)
                || NVL2(ap.direccion_fabricante, ' | ' || ap.direccion_fabricante, NULL)
                || NVL2(ap.pais_fabricante, ' | ' || ap.pais_fabricante, NULL) END descr
    FROM k
    JOIN {ESQUEMA_VENTAS}.doc d ON d.cia = k.cia AND d.suc = k.suc AND d.bodega = k.bodega AND d.cpto = k.cpto AND d.fuente = k.fuente AND d.dcto = k.dcto
    JOIN {ESQUEMA_VENTAS}.det_ventas v ON v.cia = k.cia AND v.suc = k.suc AND v.bodega = k.bodega AND v.cpto = k.cpto AND v.fuente = k.fuente AND v.dcto = k.dcto
    LEFT JOIN {ESQUEMA_VENTAS}.clientes c ON c.suc = d.suc AND c.cliente = d.cliente
    LEFT JOIN {ESQUEMA_VENTAS}.art art ON art.articulo = v.articulo
    LEFT JOIN {ESQUEMA_VENTAS}.sucursales su ON su.suc = v.suc
    LEFT JOIN {ESQUEMA_VENTAS}.art_paises ap ON ap.articulo = v.articulo AND ap.pais = NVL(d.pais_destino, su.pais)
   WHERE v.cantdesp <> 0
),
i AS (
  SELECT it.*, ROUND(cant * pu - vg, 2) descu
    FROM (SELECT it.*, CASE WHEN tipo_doc = 'FAC' THEN ROUND(cant * pu, 2) ELSE valor END vg FROM it) it
),
ia AS (
  SELECT uuid, SUM(vg) sum_vg, SUM(descu) sum_descu FROM i GROUP BY uuid
),
fx AS (
  SELECT k.uuid,
         SUM(CASE WHEN art.tipo_iva = 1 THEN ROUND(ROUND(NVL(v.precioneto_moneda_sin_prorr, 0), 2) * v.cantdesp, 2) ELSE 0 END) grav,
         SUM(CASE WHEN art.tipo_iva = 4 THEN ROUND(ROUND(NVL(v.precioneto_moneda_sin_prorr, 0), 2) * v.cantdesp, 2) ELSE 0 END) exen,
         SUM(CASE WHEN art.tipo_iva = 2 AND NVL(v.linea, 'x') <> '991'
                  THEN ROUND(ROUND(NVL(v.precioneto_moneda_sin_prorr, 0), 2) * v.cantdesp, 2) ELSE 0 END) nosuj
    FROM k
    JOIN {ESQUEMA_VENTAS}.doc d ON d.cia = k.cia AND d.suc = k.suc AND d.bodega = k.bodega AND d.cpto = k.cpto AND d.fuente = k.fuente AND d.dcto = k.dcto
    JOIN {ESQUEMA_VENTAS}.det_ventas v ON v.cia = k.cia AND v.suc = k.suc AND v.bodega = k.bodega AND v.cpto = k.cpto AND v.fuente = k.fuente AND v.dcto = k.dcto
    JOIN {ESQUEMA_VENTAS}.art art ON art.articulo = v.articulo
   WHERE k.tipo_doc = 'FAC_EXP' AND v.cantdesp > 0
     AND NOT EXISTS (SELECT 1 FROM {ESQUEMA_VENTAS}.art_plu_clientes p
                      WHERE p.articulo = v.articulo AND p.cliente = d.cliente AND p.fecvenc > SYSDATE)
   GROUP BY k.uuid
),
h AS (
  SELECT k.uuid, k.tipo_doc,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'VERSION_JSON' AND prm.codigo_tq = DECODE(k.tipo_doc, 'CCF', 'CCF_SAL', 'FAC', 'FAC_SAL', 'FAC_EXP', 'EXP_SAL', 'NCR', 'NCR_SAL', 'NDB', 'NDB_SAL', 'REM', 'REM_SAL')) AS VARCHAR2(4000)) c01,
         CAST(DECODE(k.tipo_doc, 'FAC', '01', 'CCF', '03', 'REM', '04', 'NCR', '05', 'NDB', '06', 'FAC_EXP', '11') AS VARCHAR2(4000)) c02,
         CAST('DTE-' || DECODE(k.tipo_doc, 'FAC', '01', 'CCF', '03', 'REM', '04', 'NCR', '05', 'NDB', '06', 'FAC_EXP', '11') || '-' || (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'COMPLEMENTO_CODESTABLECIMIENTO' AND prm.codigo_tq = 'M') || e.establecimiento || (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'COMPLEMENTO_PUNTOEMISION' AND prm.codigo_tq = 'P') || e.punto_emision_sri || '-' || LPAD(e.consecutivo, 15, '0') AS VARCHAR2(4000)) c03,
         CAST(d.preimpreso AS VARCHAR2(4000)) c04,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_MODELO' AND prm.codigo_tq = 'NORMAL') AS VARCHAR2(4000)) c05,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_TRANSMISION' AND prm.codigo_tq = 'NORMAL') AS VARCHAR2(4000)) c06,
         CAST(TO_CHAR(e.fec_autoriza, 'YYYY-MM-DD') AS VARCHAR2(4000)) c07,
         CAST(DECODE(SYS_CONTEXT('USERENV', 'DB_NAME'), 'SVMPROD', '01', '00') AS VARCHAR2(4000)) c08,
         CAST(TO_CHAR(e.fec_autoriza - 1 / 24, 'HH24:MI:SS') AS VARCHAR2(4000)) c09,
         CAST((SELECT MAX(co.moneda) FROM {ESQUEMA_VENTAS}.control co WHERE co.tipo = '00' AND co.cia = d.cia) AS VARCHAR2(4000)) c10,
         CAST(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN d.nrocompcrfiscal END AS VARCHAR2(4000)) c11,
         CAST(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN TO_CHAR(d.feccompcrfiscal, 'YYYY-MM-DD') END AS VARCHAR2(4000)) c12,
         CAST(a.nit AS VARCHAR2(4000)) c13,
         CAST(REPLACE(a.nrc, '-', '') AS VARCHAR2(4000)) c14,
         CAST(a.nombre AS VARCHAR2(4000)) c15,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'ACTIVIDAD_ECONOMICA_EMISOR' AND prm.codigo_tq = d.cia) AS VARCHAR2(4000)) c16,
         CAST((SELECT MAX(des) FROM prm WHERE prm.parametro = 'ACTIVIDAD_ECONOMICA_EMISOR' AND prm.codigo_tq = d.cia) AS VARCHAR2(4000)) c17,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'DEPARTAMENTO' AND prm.codigo_tq = d.cia) AS VARCHAR2(4000)) c18,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'MUNICIPIO' AND prm.codigo_tq = d.cia) AS VARCHAR2(4000)) c19,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'DISTRITO' AND prm.codigo_tq = d.cia) AS VARCHAR2(4000)) c20,
         CAST(a.direccion AS VARCHAR2(4000)) c21,
         CAST(a.telefono AS VARCHAR2(4000)) c22,
         CAST(a.email AS VARCHAR2(4000)) c23,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'COMPLEMENTO_CODESTABLECIMIENTO' AND prm.codigo_tq = 'M') || e.establecimiento AS VARCHAR2(4000)) c24,
         CAST(sp.codigo_punto_emis AS VARCHAR2(4000)) c25,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN x.recinto END AS VARCHAR2(4000)) c26,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN SUBSTR(x.regimen, 1, INSTR(x.regimen, '.') - 1) END AS VARCHAR2(4000)) c27,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN SUBSTR(x.regimen, INSTR(x.regimen, '.') + 1) END AS VARCHAR2(4000)) c28,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_ITEM' AND prm.codigo_tq = '1') END AS VARCHAR2(4000)) c29,
         CAST(d.nit AS VARCHAR2(4000)) c30,
         CAST(d.nit AS VARCHAR2(4000)) c31,
         CAST(CASE WHEN k.tipo_doc = 'FAC_EXP' THEN '37' ELSE (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_IDENTIFICACION' AND prm.codigo_tq = c.tiponit) END AS VARCHAR2(4000)) c32,
         CAST(REPLACE(d.nrc, '-', '') AS VARCHAR2(4000)) c33,
         CAST(d.nombre AS VARCHAR2(4000)) c34,
         CAST(REGEXP_SUBSTR(c.giro, '^[^-]+') AS VARCHAR2(4000)) c35,
         CAST(CASE WHEN INSTR(c.giro, '-') > 0 THEN SUBSTR(c.giro, INSTR(c.giro, '-') + 1) END AS VARCHAR2(4000)) c36,
         CAST(d.razon AS VARCHAR2(4000)) c37,
         CAST(CASE WHEN REGEXP_LIKE(c.localidad, '^[0-9]{8}$') THEN SUBSTR(c.localidad, 1, 2) END AS VARCHAR2(4000)) c38,
         CAST(CASE WHEN REGEXP_LIKE(c.localidad, '^[0-9]{8}$') THEN LPAD(TO_NUMBER(SUBSTR(c.localidad, 3, 3)), 2, '0') END AS VARCHAR2(4000)) c39,
         CAST(CASE WHEN REGEXP_LIKE(c.localidad, '^[0-9]{8}$') THEN LPAD(TO_NUMBER(SUBSTR(c.localidad, 6, 3)), 2, '0') END AS VARCHAR2(4000)) c40,
         CAST(TRIM(c.direccion || ' ' || c.barrio) AS VARCHAR2(4000)) c41,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN TRIM(d.direccion || ' ' || c.barrio) END AS VARCHAR2(4000)) c42,
         CAST(c.telefono AS VARCHAR2(4000)) c43,
         CAST((SELECT MIN(ce.e_mail) FROM {ESQUEMA_VENTAS}.clientes_electronicos ce WHERE ce.cliente = d.cliente) AS VARCHAR2(4000)) c44,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN pa.cod_minhacienda END AS VARCHAR2(4000)) c45,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN pa.nombre END AS VARCHAR2(4000)) c46,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_PERSONA' AND prm.codigo_tq = c.tipo_persona) END AS VARCHAR2(4000)) c47,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(d.valor_no_sujetas, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c48,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(d.valor_exentas, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c49,
         CAST(RTRIM(RTRIM(TO_CHAR(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c50,
         CAST(RTRIM(RTRIM(TO_CHAR(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c51,
         CAST(RTRIM(RTRIM(TO_CHAR(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c52,
         CAST(RTRIM(RTRIM(TO_CHAR(ROUND(NVL(ia.sum_descu, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c53,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'REM', 'NCR', 'NDB') THEN '0' END AS VARCHAR2(4000)) c54,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'REM', 'NCR', 'NDB') THEN '0' END AS VARCHAR2(4000)) c55,
         CAST('0' AS VARCHAR2(4000)) c56,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP', 'REM') THEN '0' END AS VARCHAR2(4000)) c57,
         CAST(CASE WHEN k.tipo_doc <> 'REM' THEN '0' END AS VARCHAR2(4000)) c58,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP') THEN '0' END AS VARCHAR2(4000)) c59,
         CAST(ventasapl.ln_ven_entidades_pkg.valor_en_letras_f(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - CASE WHEN NVL(d.retencion, 0) <> 0 THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END * 0.01 ELSE 0 END, 2) ELSE ROUND(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - NVL(d.retencion, 0), 2) WHEN k.tipo_doc = 'FAC' THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END ELSE ROUND(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2) + ROUND(NVL(d.iva, 0), 2), 2) END - NVL(d.retencion, 0), 2) END) || ' CON ' || SUBSTR(TO_CHAR(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - CASE WHEN NVL(d.retencion, 0) <> 0 THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END * 0.01 ELSE 0 END, 2) ELSE ROUND(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - NVL(d.retencion, 0), 2) WHEN k.tipo_doc = 'FAC' THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END ELSE ROUND(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2) + ROUND(NVL(d.iva, 0), 2), 2) END - NVL(d.retencion, 0), 2) END, 'FM99999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,'''), -2) || '/100' || UPPER((SELECT MAX(gc.moneda) FROM ventasadm.md_companias mc JOIN ventasadm.grupos_companias gc ON gc.cod_grupo_compania = mc.cod_grupo_compania WHERE mc.codigo_compania = d.cia)) AS VARCHAR2(4000)) c60,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'NCR', 'NDB', 'FAC_EXP') THEN (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'IMPUESTO_REGISTRA_FE' AND prm.codigo_tq = DECODE(k.tipo_doc, 'FAC_EXP', 'IVA_EXP', 'IVA')) END AS VARCHAR2(4000)) c61,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'NCR', 'NDB', 'FAC_EXP') THEN (SELECT MAX(des) FROM prm WHERE prm.parametro = 'IMPUESTO_REGISTRA_FE' AND prm.codigo_tq = DECODE(k.tipo_doc, 'FAC_EXP', 'IVA_EXP', 'IVA')) END AS VARCHAR2(4000)) c62,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'NCR', 'NDB', 'FAC_EXP') THEN RTRIM(RTRIM(TO_CHAR(ROUND(NVL(d.iva, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c63,
         CAST(CASE WHEN k.tipo_doc IN ('FAC') THEN RTRIM(RTRIM(TO_CHAR(ROUND(NVL(d.iva_ent, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c64,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(d.iva_percibido, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c65,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(d.retencion, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c66,
         CAST(RTRIM(RTRIM(TO_CHAR(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - NVL(d.retencion, 0), 2) WHEN k.tipo_doc = 'FAC' THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END ELSE ROUND(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2) + ROUND(NVL(d.iva, 0), 2), 2) END, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c67,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP', 'REM') THEN RTRIM(RTRIM(TO_CHAR(ROUND(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - NVL(d.retencion, 0), 2) WHEN k.tipo_doc = 'FAC' THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END ELSE ROUND(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2) + ROUND(NVL(d.iva, 0), 2), 2) END - NVL(d.retencion, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c68,
         CAST(e.cli_cond_pago AS VARCHAR2(4000)) c69,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP') THEN e.cli_medio_pago END AS VARCHAR2(4000)) c70,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP') THEN RTRIM(RTRIM(TO_CHAR(ROUND(CASE WHEN k.tipo_doc IN ('NCR', 'NDB') THEN ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.iva, 0) - NVL(d.retencion, 0), 2) WHEN k.tipo_doc = 'FAC' THEN CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END ELSE ROUND(ROUND(CASE WHEN k.tipo_doc = 'FAC' THEN NVL(ia.sum_vg, 0) ELSE NVL(d.valor_gravadas, 0) END + NVL(d.valor_exentas, 0) + NVL(d.valor_no_sujetas, 0), 2) + ROUND(NVL(d.iva, 0), 2), 2) END - NVL(d.retencion, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c71,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP') THEN (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'CODIGO_PLAZO' AND prm.codigo_tq = '01') END AS VARCHAR2(4000)) c72,
         CAST(CASE WHEN k.tipo_doc IN ('CCF', 'FAC', 'FAC_EXP') THEN TO_CHAR(d.plazo) END AS VARCHAR2(4000)) c73,
         CAST(CASE WHEN k.tipo_doc = 'FAC_EXP' THEN NVL(x.observaciones || CASE WHEN CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN ROUND(NVL(d.fletes1_exp, 0), 2) ELSE 0 END > 0 OR CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN ROUND(NVL(d.seguros1_exp, 0), 2) ELSE 0 END > 0 THEN CASE WHEN x.observaciones IS NOT NULL THEN ' | ' END || 'Para efectos de aduana: valor FOB: ' || TO_CHAR(ROUND(CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN fx.grav - (NVL(d.gastos1_exp, 0) + NVL(d.fletes1_exp, 0) + NVL(d.seguros1_exp, 0)) ELSE fx.grav END, 2), 'FM999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,''') || ', Gastos ' || TO_CHAR(CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN ROUND(NVL(d.gastos1_exp, 0), 2) ELSE 0 END, 'FM999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,''') || ', Fletes ' || TO_CHAR(CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN ROUND(NVL(d.fletes1_exp, 0), 2) ELSE 0 END, 'FM999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,''') || ', Seguro ' || TO_CHAR(CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN ROUND(NVL(d.seguros1_exp, 0), 2) ELSE 0 END, 'FM999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,''') || '.  Total ' || TO_CHAR(ROUND(fx.exen + CASE WHEN x.incoterm IN ('CIF', 'CIP', 'DAP') THEN fx.grav - (NVL(d.gastos1_exp, 0) + NVL(d.fletes1_exp, 0) + NVL(d.seguros1_exp, 0)) ELSE fx.grav END + fx.nosuj + NVL(d.gastos1_exp, 0) + NVL(d.fletes1_exp, 0) + NVL(d.seguros1_exp, 0), 2), 'FM999999990.00', 'NLS_NUMERIC_CHARACTERS=''.,''') END, '.') ELSE d.suc || '|' || d.bodega || '|' || d.cpto || '|' || d.fuente || '|' || NVL(TO_CHAR(d.plazo), '0') END AS VARCHAR2(4000)) c74,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN RTRIM(RTRIM(TO_CHAR(NVL(d.seguros_exp, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c75,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN RTRIM(RTRIM(TO_CHAR(NVL(d.fletes_exp, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c76,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN (SELECT MAX(cod) FROM prm WHERE prm.parametro = 'INCOTERMS' AND prm.codigo_tq = x.incoterm) END AS VARCHAR2(4000)) c77,
         CAST(CASE WHEN k.tipo_doc IN ('FAC_EXP') THEN (SELECT MAX(des) FROM prm WHERE prm.parametro = 'INCOTERMS' AND prm.codigo_tq = x.incoterm) END AS VARCHAR2(4000)) c78,
         CAST('apendice01' AS VARCHAR2(4000)) c79,
         CAST('apendice02' AS VARCHAR2(4000)) c80,
         CAST('apendice03' AS VARCHAR2(4000)) c81,
         CAST('apendice04' AS VARCHAR2(4000)) c82,
         CAST('Regional:|Bodega:|Concepto:|Pedido:|Plazo:' AS VARCHAR2(4000)) c83,
         CAST(CASE WHEN k.tipo_doc = 'FAC_EXP' THEN 'Lista:|Documento interno:|Instrucciones:' ELSE 'Lista:|Documento interno:' END AS VARCHAR2(4000)) c84,
         CAST(CASE WHEN k.tipo_doc = 'FAC_EXP' THEN 'Código:|Orden de compra:|Fecha Vencimiento:|Moneda' ELSE 'Código:|Orden de compra:|Fecha Vencimiento:' END AS VARCHAR2(4000)) c85,
         CAST(CASE WHEN k.tipo_doc = 'FAC_EXP' THEN 'R.V.:|No. Cajas:|Peso:|Pais Origen:|Puerto Destino' ELSE 'R.V.:|No. Cajas:|Peso:' END AS VARCHAR2(4000)) c86,
         CAST(d.suc || '|' || d.bodega || '|' || d.cpto || '|' || d.fuente || '|' || TO_CHAR(d.plazo) AS VARCHAR2(4000)) c87,
         CAST(d.lista || '|' || d.dcto || CASE WHEN k.tipo_doc = 'FAC_EXP' THEN '|' || x.medio_transporte END AS VARCHAR2(4000)) c88,
         CAST(d.cliente || '|' || d.orden || '|' || TO_CHAR(d.fecvenc, 'YYYY/MM/DD') || CASE WHEN k.tipo_doc = 'FAC_EXP' THEN '|' || d.moneda_destino END AS VARCHAR2(4000)) c89,
         CAST(d.vendedor || '|' || d.cajas || '|' || CASE WHEN k.tipo_doc = 'FAC_EXP' AND d.pais_destino = '29' THEN NULL ELSE TO_CHAR(d.peso, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''') END || CASE WHEN k.tipo_doc = 'FAC_EXP' THEN '|' || x.pais_origen_art || '|' || x.puerto_destino END AS VARCHAR2(4000)) c90
    FROM k
    JOIN {ESQUEMA_VENTAS}.doc d ON d.cia = k.cia AND d.suc = k.suc AND d.bodega = k.bodega AND d.cpto = k.cpto AND d.fuente = k.fuente AND d.dcto = k.dcto
    JOIN {ESQUEMA_VENTAS}.doc_extendido e ON e.cia = k.cia AND e.suc = k.suc AND e.bodega = k.bodega AND e.cpto = k.cpto AND e.fuente = k.fuente AND e.dcto = k.dcto
    JOIN {ESQUEMA_VENTAS}.cia a ON a.cia = d.cia
    LEFT JOIN {ESQUEMA_VENTAS}.clientes c ON c.suc = d.suc AND c.cliente = d.cliente
    LEFT JOIN {ESQUEMA_VENTAS}.doc_exp x ON x.cia = k.cia AND x.suc = k.suc AND x.bodega = k.bodega AND x.cpto = k.cpto AND x.fuente = k.fuente AND x.dcto = k.dcto
    LEFT JOIN {ESQUEMA_VENTAS}.pais pa ON pa.pais = c.pais
    LEFT JOIN {ESQUEMA_VENTAS}.sri_punto_emision sp ON sp.punto_emision_sri = e.punto_emision_sri
    LEFT JOIN ia ON ia.uuid = k.uuid
    LEFT JOIN fx ON fx.uuid = k.uuid
),
hi AS (
  SELECT i.uuid, i.tipo_doc, i.n,
         CAST(TO_CHAR(i.n) AS VARCHAR2(4000)) c01,
         CAST((SELECT MAX(cod) FROM prm WHERE prm.parametro = 'TIPO_ITEM' AND prm.codigo_tq = '1') AS VARCHAR2(4000)) c02,
         CAST(CASE WHEN i.tipo_doc IN ('NCR', 'NDB') THEN i.nrocompcrfiscal END AS VARCHAR2(4000)) c03,
         CAST(RTRIM(RTRIM(TO_CHAR(i.cant, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c04,
         CAST(i.articulo AS VARCHAR2(4000)) c05,
         CAST(CASE WHEN i.tipo_doc <> 'FAC' THEN i.descr END AS VARCHAR2(4000)) c06,
         CAST(RTRIM(RTRIM(TO_CHAR(i.pu, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c07,
         CAST(RTRIM(RTRIM(TO_CHAR(i.descu, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c08,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(i.valor_no_sujetas, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c09,
         CAST(RTRIM(RTRIM(TO_CHAR(NVL(i.valor_exentas, 0), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c10,
         CAST(RTRIM(RTRIM(TO_CHAR(i.vg, 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') AS VARCHAR2(4000)) c11,
         CAST('59' AS VARCHAR2(4000)) c12,
         CAST(CASE WHEN i.tipo_doc <> 'REM' THEN '0' END AS VARCHAR2(4000)) c13,
         CAST(CASE WHEN i.tipo_doc IN ('CCF', 'FAC') THEN '0' END AS VARCHAR2(4000)) c14,
         CAST(CASE WHEN i.tipo_doc = 'FAC' THEN RTRIM(RTRIM(TO_CHAR(ROUND(NVL(i.iva_ent, 0), 2), 'FM99999999999990.9999999999', 'NLS_NUMERIC_CHARACTERS=''.,'''), '0'), '.') END AS VARCHAR2(4000)) c15
    FROM i
)
SELECT uuid codigo_generacion, tipo_doc, ruta, valor
  FROM h UNPIVOT (valor FOR ruta IN (
      c01 AS 'identificacion.version',
      c02 AS 'identificacion.tipoDte',
      c03 AS 'identificacion.numeroControl',
      c04 AS 'identificacion.codigoGeneracion',
      c05 AS 'identificacion.tipoModelo',
      c06 AS 'identificacion.tipoOperacion',
      c07 AS 'identificacion.fecEmi',
      c08 AS 'identificacion.ambiente',
      c09 AS 'identificacion.horEmi',
      c10 AS 'identificacion.tipoMoneda',
      c11 AS 'documentoRelacionado.numeroDocumento[1]',
      c12 AS 'documentoRelacionado.fechaEmision[1]',
      c13 AS 'emisor.nit',
      c14 AS 'emisor.nrc',
      c15 AS 'emisor.nombre',
      c16 AS 'emisor.codActividad',
      c17 AS 'emisor.descActividad',
      c18 AS 'emisor.direccion.departamento',
      c19 AS 'emisor.direccion.municipio',
      c20 AS 'emisor.direccion.distrito',
      c21 AS 'emisor.direccion.complemento',
      c22 AS 'emisor.telefono',
      c23 AS 'emisor.correo',
      c24 AS 'emisor.codEstable',
      c25 AS 'emisor.codPuntoVenta',
      c26 AS 'emisor.recintoFiscal',
      c27 AS 'emisor.tipoRegimen',
      c28 AS 'emisor.regimen',
      c29 AS 'emisor.tipoItemExpor',
      c30 AS 'receptor.nit',
      c31 AS 'receptor.numDocumento',
      c32 AS 'receptor.tipoDocumento',
      c33 AS 'receptor.nrc',
      c34 AS 'receptor.nombre',
      c35 AS 'receptor.codActividad',
      c36 AS 'receptor.descActividad',
      c37 AS 'receptor.nombreComercial',
      c38 AS 'receptor.direccion.departamento',
      c39 AS 'receptor.direccion.municipio',
      c40 AS 'receptor.direccion.distrito',
      c41 AS 'receptor.direccion.complemento',
      c42 AS 'receptor.complemento',
      c43 AS 'receptor.telefono',
      c44 AS 'receptor.correo',
      c45 AS 'receptor.codPais',
      c46 AS 'receptor.nombrePais',
      c47 AS 'receptor.tipoPersona',
      c48 AS 'resumen.totalNoSuj',
      c49 AS 'resumen.totalExenta',
      c50 AS 'resumen.totalGravada',
      c51 AS 'resumen.subTotalVentas',
      c52 AS 'resumen.subTotal',
      c53 AS 'resumen.totalDescu',
      c54 AS 'resumen.descuNoSuj',
      c55 AS 'resumen.descuExenta',
      c56 AS 'resumen.descuGravada',
      c57 AS 'resumen.porcentajeDescuento',
      c58 AS 'resumen.totalNoGravado',
      c59 AS 'resumen.saldoFavor',
      c60 AS 'resumen.totalLetras',
      c61 AS 'resumen.tributos.codigo[1]',
      c62 AS 'resumen.tributos.descripcion[1]',
      c63 AS 'resumen.tributos.valor[1]',
      c64 AS 'resumen.totalIva',
      c65 AS 'resumen.ivaPerci',
      c66 AS 'resumen.ivaRete',
      c67 AS 'resumen.montoTotalOperacion',
      c68 AS 'resumen.totalPagar',
      c69 AS 'resumen.condicionOperacion',
      c70 AS 'resumen.pagos.codigo[1]',
      c71 AS 'resumen.pagos.montoPago[1]',
      c72 AS 'resumen.pagos.plazo[1]',
      c73 AS 'resumen.pagos.periodo[1]',
      c74 AS 'resumen.observaciones',
      c75 AS 'resumen.seguro',
      c76 AS 'resumen.flete',
      c77 AS 'resumen.codIncoterms',
      c78 AS 'resumen.descIncoterms',
      c79 AS 'apendice.campo[apendice01]',
      c80 AS 'apendice.campo[apendice02]',
      c81 AS 'apendice.campo[apendice03]',
      c82 AS 'apendice.campo[apendice04]',
      c83 AS 'apendice.etiqueta[apendice01]',
      c84 AS 'apendice.etiqueta[apendice02]',
      c85 AS 'apendice.etiqueta[apendice03]',
      c86 AS 'apendice.etiqueta[apendice04]',
      c87 AS 'apendice.valor[apendice01]',
      c88 AS 'apendice.valor[apendice02]',
      c89 AS 'apendice.valor[apendice03]',
      c90 AS 'apendice.valor[apendice04]'))
UNION ALL
SELECT uuid, tipo_doc, ruta || '[' || n || ']', valor
  FROM hi UNPIVOT (valor FOR ruta IN (
      c01 AS 'cuerpoDocumento.numItem',
      c02 AS 'cuerpoDocumento.tipoItem',
      c03 AS 'cuerpoDocumento.numeroDocumento',
      c04 AS 'cuerpoDocumento.cantidad',
      c05 AS 'cuerpoDocumento.codigo',
      c06 AS 'cuerpoDocumento.descripcion',
      c07 AS 'cuerpoDocumento.precioUni',
      c08 AS 'cuerpoDocumento.montoDescu',
      c09 AS 'cuerpoDocumento.ventaNoSuj',
      c10 AS 'cuerpoDocumento.ventaExenta',
      c11 AS 'cuerpoDocumento.ventaGravada',
      c12 AS 'cuerpoDocumento.uniMedida',
      c13 AS 'cuerpoDocumento.noGravado',
      c14 AS 'cuerpoDocumento.psv',
      c15 AS 'cuerpoDocumento.ivaItem'))
ORDER BY 1, 3
