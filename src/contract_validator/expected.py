"""Valores esperados por documento, obtenidos del query de referencia (Fase B).

Fuente: queries/valores_esperados.sql (Oracle, solo lectura) sobre las tablas de negocio; mapeo campo -> columna en
data/specs/mapeo_bd.xlsx y normalizaciones aprobadas en data/specs/normalizaciones_bd.json.

  load_expected(source) -> {codigo_generacion: {ruta: valor}}

  * `source`: archivo exportado del query (.tsv, .txt, .csv o .xlsx, p. ej. desde DBeaver) o un iterable de filas
    (dicts con las columnas del query, o tuplas (codigo_generacion, ruta, valor)), por ejemplo un cursor.
  * Columnas requeridas: CODIGO_GENERACION, RUTA, VALOR (sin distinguir mayúsculas). TIPO_DOC es opcional.
  * `ruta` es la del contrato plano (JSON, separador '.'); en arreglos lleva sufijo [n] (1-based, orden del documento)
    o [apendiceNN] para el apéndice. Ejemplo: 'cuerpoDocumento.cantidad[2]'.
  * Solo texto, sin lógica de negocio: el valor se toma tal cual. Vacío, 'null' o '[NULL]' (DBeaver) = sin valor y
    no se incluye.
  * Códigos de hallazgo reservados: DB_MISMATCH_JSON, DB_MISMATCH_XML, DB_MISMATCH_PDF, DB_MISSING.
"""
from __future__ import annotations

import csv
import io
import os
from collections.abc import Iterable

Expected = dict[str, dict[str, str]]

DB_MISMATCH_JSON, DB_MISMATCH_XML, DB_MISMATCH_PDF, DB_MISSING = (
    "DB_MISMATCH_JSON", "DB_MISMATCH_XML", "DB_MISMATCH_PDF", "DB_MISSING")

COLUMNS = ("CODIGO_GENERACION", "RUTA", "VALOR")
NULLS = ("", "null", "[null]")


def _is_null(v) -> bool:
    return v is None or str(v).strip().lower() in NULLS


def _read_file(path: str) -> list[dict]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        from openpyxl import load_workbook
        ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = ws.iter_rows(values_only=True)
        header = [str(h or "").strip().upper() for h in next(rows, ())]
        return [dict(zip(header, r)) for r in rows]
    with open(path, encoding="utf-8-sig", newline="") as fh:
        text = fh.read()
    if ext in (".tsv", ".txt"):
        delim = "\t"
    else:
        first = text.splitlines()[0] if text else ""
        delim = max(("\t", ";", ","), key=first.count)
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delim)  # respeta saltos de línea dentro de comillas
    header = [h.strip().upper() for h in next(reader, [])]
    return [dict(zip(header, r)) for r in reader if any(c.strip() for c in r)]


def _normalize_rows(rows: Iterable) -> Iterable[tuple]:
    for r in rows:
        if isinstance(r, dict):
            up = {str(k).strip().upper(): v for k, v in r.items()}
            missing = [c for c in COLUMNS if c not in up]
            if missing:
                raise ValueError(f"faltan columnas en la fuente de valores esperados: {', '.join(missing)}")
            yield up["CODIGO_GENERACION"], up["RUTA"], up["VALOR"]
        else:
            uuid, ruta, valor = tuple(r)[:3] if len(r) == 3 else (r[0], r[-2], r[-1])
            yield uuid, ruta, valor


def load_expected(source) -> Expected:
    rows = _read_file(source) if isinstance(source, (str, os.PathLike)) else source
    out: Expected = {}
    for uuid, ruta, valor in _normalize_rows(rows):
        if _is_null(uuid) or _is_null(ruta) or _is_null(valor):
            continue
        uuid, ruta, valor = str(uuid).strip().upper(), str(ruta).strip(), str(valor)
        doc = out.setdefault(uuid, {})
        if ruta in doc and doc[ruta] != valor:
            raise ValueError(f"valor duplicado para {uuid} {ruta}: {doc[ruta]!r} y {valor!r}")
        doc[ruta] = valor
    return out
