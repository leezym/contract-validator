"""Contenido de JSON, XML y PDF MINH contra los valores esperados de la BD (Fase B).

Valores esperados: `expected.load_expected` sobre el resultado de `queries/valores_esperados.sql`
({codigo_generacion: {ruta: valor}}; arreglos con sufijo [n] o [apendiceNN]). Emparejamiento por Código de Generación.

Qué se compara (solo lo que definen los specs y lo aprobado):
  * rutas JSON del contrato para el tipo de documento (sin ELIMINADO) con categoría aprobada en `data/specs/mapeo_bd.xlsx`
    (DIRECTA, TRANSFORMACION, CATALOGO con aprobado = SI);
  * XML: la misma ruta vía `ruta_contraparte`; los mapeos posicionales del libro no se comparan 1 a 1 y se omiten;
    cada apéndice se compara como una unidad (unión de datos de BD), reconocido por su nombre apendiceNN;
  * PDF MINH: identidad (Código de Generación, Número de Control), etiquetas ancladas EXACTO en la `Descripcion` del
    Excel (igual que `pdf_check`) y filas de ítems (cantidad, precio, descuento).
Igualdad: texto exacto; números (sin ceros a la izquierda) se comparan como números con la tolerancia aprobada en
`data/specs/normalizaciones_bd.json`; en el PDF además las normalizaciones de presentación de `pdf_presentacion.json`.

Códigos:
  DB_MISMATCH_JSON / DB_MISMATCH_XML / DB_MISMATCH_PDF  (ERROR) el documento difiere de la BD, o la BD tiene valor y el
                                                        documento no, o distinta cantidad de ítems
  DB_MISSING (ERROR)    el documento no está en el resultado del query
  DB_MISSING (WARNING)  campos del documento sin referencia en BD (categoría SIN_FUENTE o la BD no devuelve valor):
                        no verificables; se informa uno por documento con la lista de rutas
"""
from __future__ import annotations
import json
import os
import re
from decimal import Decimal, InvalidOperation
from .crosscheck import apendice_records, document_id, strict, xml_values
from .expected import DB_MISMATCH_JSON, DB_MISMATCH_XML, DB_MISMATCH_PDF, DB_MISSING, Expected
from .models import Finding, Result, REMOVED
from .pdf_check import (Presentation, _anchors, _label_key, _tok, apendice_pairs, compare_labeled, format_match,
                        item_blocks, label_values,
                        printed_with_label, text_continues)
from .pdf_text import read_pdf

DEFAULT_MAPEO = os.path.join("data", "specs", "mapeo_bd.xlsx")
DEFAULT_CONFIG = os.path.join("data", "specs", "normalizaciones_bd.json")
APPROVED = ("DIRECTA", "TRANSFORMACION", "CATALOGO")
OUT_OF_SCOPE = ("INVALIDACION", "RETORNO")  # sin fuente ubicada en la BD (decisión del usuario, Fase B)
MAX_PER_FIELD = 3
APENDICE = "apendice."
IDX = re.compile(r"^(.*)\[([^\]]+)\]$")
NUMBER = re.compile(r"^-?(0|[1-9]\d*)(\.\d+)?$|^-?\.\d+$")


# ------------------------------------------------------------------ configuración
class BdConfig:
    def __init__(self, mapeo: dict | None = None, tolerance: float = 0.0, time_tolerance: int = 0):
        self.mapeo = mapeo or {}            # (tipo_doc, ruta) -> categoria aprobada o 'SIN_FUENTE'
        self.tolerance = Decimal(str(tolerance))
        self.time_tolerance = int(time_tolerance or 0)  # segundos entre horas HH:MM:SS (horEmi)

    def category(self, doc_type: str, path: str) -> str:
        return self.mapeo.get((doc_type, path), "")


def load_bd_config(mapeo_path: str | None = None, config_path: str | None = None) -> BdConfig:
    mapeo_path, config_path = mapeo_path or DEFAULT_MAPEO, config_path or DEFAULT_CONFIG
    mapeo = {}
    if os.path.exists(mapeo_path):
        from openpyxl import load_workbook
        ws = load_workbook(mapeo_path, read_only=True, data_only=True)["mapeo"]
        rows = ws.iter_rows(values_only=True)
        head = [str(h or "").strip().lower() for h in next(rows)]
        ix = {k: head.index(k) for k in ("tipo_doc", "ruta", "categoria", "aprobado")}
        for r in rows:
            cat, ok = str(r[ix["categoria"]] or ""), str(r[ix["aprobado"]] or "").upper()
            key = (str(r[ix["tipo_doc"]] or ""), str(r[ix["ruta"]] or ""))
            if cat in APPROVED and ok == "SI":
                mapeo[key] = cat
            elif cat == "SIN_FUENTE" or ok == "DB_MISSING":
                mapeo[key] = "SIN_FUENTE"
    elif mapeo_path != DEFAULT_MAPEO:
        raise FileNotFoundError(f"No existe el mapeo de BD: {mapeo_path}")
    tol, tol_hora = 0.0, 0
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8-sig") as fh:
            conf = json.load(fh)
        tol = float(conf.get("tolerancia_montos", 0) or 0)
        tol_hora = int(conf.get("tolerancia_hora_segundos", 0) or 0)
    elif config_path != DEFAULT_CONFIG:
        raise FileNotFoundError(f"No existe el archivo de normalizaciones de BD: {config_path}")
    return BdConfig(mapeo, tol, tol_hora)


# ------------------------------------------------------------------ comparación de valores
def _num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    t = str(v).strip()
    if NUMBER.match(t):
        try:
            return Decimal(t)
        except InvalidOperation:
            return None
    return None


HORA = re.compile(r"^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$")


def _secs(t: str) -> int:
    h, m, s = (int(x) for x in HORA.match(t).groups())
    return h * 3600 + m * 60 + s


def same(expected, got, cfg: BdConfig) -> bool:
    e, g = strict(expected), strict(got)
    if e is None or g is None:
        return e is None and g is None
    if cfg.time_tolerance and HORA.match(e) and HORA.match(g):  # horEmi: el generador la toma segundos antes
        d = abs(_secs(e) - _secs(g))
        return min(d, 86400 - d) <= cfg.time_tolerance
    ne, ng = _num(expected), _num(got)
    if ne is not None and ng is not None:
        return abs(ne - ng) <= cfg.tolerance
    return e == g


def _show(v):
    return "null" if strict(v) is None else repr(v)


def _split(key: str) -> tuple[str, str]:
    m = IDX.match(key)
    return (m.group(1), m.group(2)) if m else (key, "")


# ------------------------------------------------------------------ valores del documento por ruta indexada
def json_indexed(raw) -> dict[str, object]:
    """ruta[n] (posición en el arreglo exterior, 1-based) o ruta[campo] (elementos con 'campo', p.ej. apéndice)."""
    out: dict[str, object] = {}

    def walk(o, path, idx):
        if isinstance(o, dict):
            for k, v in o.items():
                walk(v, f"{path}.{k}" if path else k, idx)
        elif isinstance(o, list):
            for i, v in enumerate(o, 1):
                key = idx or (str(v["campo"]) if isinstance(v, dict) and v.get("campo") is not None else str(i))
                walk(v, path, key)
        else:
            out.setdefault(f"{path}[{idx}]" if idx else path, o)
    walk(raw, "", "")
    return out


# ------------------------------------------------------------------ JSON / XML
def _contract_rows(spec, doc_type):
    return {s.path: s for s in spec if s.doc_type == doc_type and s.fmt == "JSON" and s.change != REMOVED
            and s.type not in ("object", "array")}


def check_doc(file: str, doc, exp: dict | None, spec, cfg: BdConfig, doc_type: str = "") -> Result:
    doc_type = doc_type or doc.doc_type
    fmt = doc.fmt
    code = DB_MISMATCH_JSON if fmt == "JSON" else DB_MISMATCH_XML
    res = Result(file=file, doc_type=doc_type, fmt=f"BD<->{fmt}")

    def add(sev, c, path, msg):
        res.findings.append(Finding(sev, c, path, msg))

    uuid = document_id(doc)
    if exp is None:
        add("ERROR", DB_MISSING, "", f"El Código de Generación {uuid or '(no encontrado)'} no está en el resultado del query de BD")
        return res
    rows = _contract_rows(spec, doc_type)
    if fmt == "JSON":
        got = json_indexed(doc.raw)
    else:
        xv = xml_values(doc.raw)
    sin_fuente, sin_valor_bd, shown = [], [], {}

    # rutas del contrato con valor en el documento: verificar que tengan referencia en BD
    if fmt == "JSON":
        for key, v in got.items():
            base, _ = _split(key)
            if base not in rows or base.startswith(APENDICE) or strict(v) is None:
                continue
            cat = cfg.category(doc_type, base)
            if cat == "SIN_FUENTE":
                sin_fuente.append(base)
            elif cat in APPROVED and key not in exp:
                sin_valor_bd.append(key)

    # valores esperados
    counts = {}
    for key, e in exp.items():
        base, idx = _split(key)
        s = rows.get(base)
        if s is None or base.startswith(APENDICE) or cfg.category(doc_type, base) not in APPROVED:
            continue
        if fmt == "JSON":
            g, present = got.get(key), key in got
            label = key
        else:
            if not s.counterpart or "posicional" in s.note:
                continue
            vals = xv.get(s.counterpart) or []
            i = int(idx) if idx else 1
            if not idx and len(vals) > 1:
                continue  # repetido sin índice: no comparable 1 a 1
            g, present = (vals[i - 1], True) if len(vals) >= i else (None, False)
            label = f"{key} <-> {s.counterpart}{f'[{idx}]' if idx else ''}"
        if idx.isdigit():
            counts.setdefault(base.split(".")[0], set()).add(int(idx))
        res.checked += 1
        if not same(e, g, cfg):
            n = shown.get(base, 0) + 1
            shown[base] = n
            if n <= MAX_PER_FIELD:
                what = "falta en el documento" if not present or strict(g) is None else f"documento={_show(g)}"
                add("ERROR", code, key, f"{label}: BD={_show(e)} {what}")
    for base, n in shown.items():
        if n > MAX_PER_FIELD:
            add("ERROR", code, base, f"{base}: {n - MAX_PER_FIELD} diferencia(s) más")
    _check_apendice(doc, exp, rows, cfg, doc_type, code, res, add, sin_valor_bd)

    # cantidad de ítems
    if fmt == "JSON" and isinstance(doc.raw, dict):
        items = doc.raw.get("cuerpoDocumento")
        n_bd = len(counts.get("cuerpoDocumento", ()))
        if isinstance(items, list) and n_bd and len(items) != n_bd:
            add("ERROR", code, "cuerpoDocumento", f"La BD tiene {n_bd} ítem(s) y el documento {len(items)}")
    elif fmt == "XML":
        n_bd = len(counts.get("cuerpoDocumento", ()))
        n_doc = len(doc.raw.findall("DatosEntrada/Items"))
        if n_bd and n_doc != n_bd:
            add("ERROR", code, "DatosEntrada.Items", f"La BD tiene {n_bd} ítem(s) y el documento {n_doc}")

    _missing_summary(add, sin_fuente, sin_valor_bd)
    return res


def _check_apendice(doc, exp, rows, cfg, doc_type, code, res, add, sin_valor_bd):
    """Cada apéndice es la unión de varios datos de BD: se compara completo, ubicado por su nombre apendiceNN."""
    partes = [p for p in ("etiqueta", "valor") if f"{APENDICE}{p}" in rows and cfg.category(doc_type, f"{APENDICE}{p}") in APPROVED]
    if not partes:
        return
    esperado: dict[str, dict] = {}
    for key, e in exp.items():
        base, idx = _split(key)
        if base.startswith(APENDICE) and idx:
            esperado.setdefault(idx, {})[base[len(APENDICE):]] = e
    recibidos = apendice_records(doc)
    for name, e in sorted(esperado.items()):
        res.checked += 1
        r = recibidos.get(name)
        if r is None:
            if doc.fmt == "JSON":
                add("ERROR", code, f"apendice[{name}]", f"Apéndice {name}: BD={_show(e.get('valor'))} falta en el documento")
            else:
                add("WARNING", "APENDICE_NOT_RECEIVED", f"apendice[{name}]",
                    f"Apéndice {name}: la BD lo arma con {_show(e.get('valor'))} y no llegó en el documento")
            continue
        for p in partes:
            if p in e and not same(e[p], r.get(p), cfg):
                add("ERROR", code, f"apendice[{name}]", f"Apéndice {name}, {p}: BD={_show(e[p])} documento={_show(r.get(p))}")
    sin_valor_bd.extend(f"{APENDICE}valor[{n}]" for n in recibidos if n not in esperado)


def _missing_summary(add, sin_fuente, sin_valor_bd):
    if sin_fuente:
        u = sorted(set(sin_fuente))
        add("WARNING", DB_MISSING, "", f"{len(u)} campo(s) sin fuente en BD (no verificables): {', '.join(u)}")
    if sin_valor_bd:
        u = sorted(set(_split(k)[0] for k in sin_valor_bd))
        add("WARNING", DB_MISSING, "", f"{len(u)} campo(s) con valor en el documento y sin valor en BD: {', '.join(u)}")


# ------------------------------------------------------------------ PDF
def check_pdf_db(info, exp: dict | None, spec, cfg: BdConfig, pres: Presentation, doc_type: str = "") -> Result:
    doc_type = doc_type or info.title
    res = Result(file=info.path, doc_type=doc_type, fmt=f"BD<->PDF({info.kind})")

    def add(sev, c, path, msg):
        res.findings.append(Finding(sev, c, path, msg))

    if info.error:
        add("ERROR", info.error[0], "", info.error[1])
        return res
    if exp is None:
        add("ERROR", DB_MISSING, "", f"El Código de Generación {info.uuid or '(no identificado en el PDF)'} no está en el resultado del query de BD")
        return res

    def eq(pdf_val, e):
        if pres.canon(pdf_val) == pres.canon(e):
            return True
        a, b = _num(pres.canon(pdf_val)), _num(pres.canon(e))
        return a is not None and b is not None and abs(a - b) <= cfg.tolerance

    # identidad
    for path, pdf_val, name in (("identificacion.codigoGeneracion", info.uuid, "Código de Generación"),
                                ("identificacion.numeroControl", info.control, "Número de Control")):
        e = exp.get(path)
        if e is None:
            continue
        res.checked += 1
        if not pdf_val or str(pdf_val).strip().upper() != str(e).strip().upper():
            add("ERROR", DB_MISMATCH_PDF, path, f"{name}: BD={e!r} PDF={pdf_val or '(no encontrado)'!r}")

    # etiquetas ancladas en la Descripcion del Excel: XML -> fila JSON (misma contraparte y misma posición)
    json_rows = [s for s in spec if s.doc_type == doc_type and s.fmt == "JSON" and s.change != REMOVED]
    anchors = _anchors(spec, doc_type)
    for line_no, ln in enumerate(info.lines):
        for m in re.finditer(r"([^:]{2,60}):\s*\$?\s*(\S*)", ln):
            s = anchors.get(_label_key(m.group(1)))
            pdf_val = m.group(2).strip()
            if s is None or not pdf_val:
                continue
            pos = re.search(r"posicional\((\d+)\)", s.note)
            cands = [j for j in json_rows if j.counterpart == s.path and
                     (not pos or re.search(r"posicional\((\d+)\)", j.note) and re.search(r"posicional\((\d+)\)", j.note).group(1) == pos.group(1))]
            if len(cands) != 1 or cfg.category(doc_type, cands[0].path) not in APPROVED:
                continue
            e = exp.get(cands[0].path, exp.get(cands[0].path + "[1]"))
            if strict(e) is None:
                continue
            pre_tag = pres.composite.get(s.path)
            pre_row = next((j for j in json_rows if pre_tag and j.counterpart == pre_tag), None)
            e_txt = pres.catalog_text(s.path, e, exp.get(pre_row.path) if pre_row else None)
            res.checked += 1
            if len(str(e_txt).split()) > 1:  # texto de varias palabras: puede seguir en las líneas siguientes
                pdf_val = ln[m.start(2):].strip()
                if text_continues(pres, info.lines, line_no, m.start(2), e_txt):
                    continue
            if not eq(pdf_val, e_txt):
                add("ERROR", DB_MISMATCH_PDF, cands[0].path, f"'{m.group(1).strip()}': BD={e_txt!r} PDF={pdf_val!r}")

    # etiquetas del PDF configuradas (etiquetas_pdf): tag XML -> fila JSON de su contraparte
    for tag, label in pres.labels.items():
        for j in [j for j in json_rows if j.counterpart == tag]:
            e = exp.get(j.path)
            if strict(e) is None or cfg.category(doc_type, j.path) not in APPROVED:
                continue
            res.checked += 1
            out = compare_labeled(pres, e, label_values(info.lines, label))
            if out is None:
                continue
            sev, code, got = out
            if sev == "WARNING":
                kind = "mismo valor con otro formato" if code == "PDF_FORMAT_DIFF" else "el PDF lo muestra cortado"
                add(sev, code, j.path, f"'{label}': BD={e} PDF={got} ({kind})")
            else:
                add("ERROR", DB_MISMATCH_PDF, j.path, f"'{label}': BD={e!r} PDF={got or '(no impreso)'!r}")

    # apéndices que deben verse en el PDF: cada dato con su etiqueta
    for name in sorted(pres.printed_apendices):
        pairs = apendice_pairs(exp.get(f"{APENDICE}etiqueta[{name}]"), exp.get(f"{APENDICE}valor[{name}]"))
        if cfg.category(doc_type, f"{APENDICE}valor") not in APPROVED:
            continue
        for label, val in pairs:
            res.checked += 1
            got = printed_with_label(pres, info.lines, label, val)
            if got is not None:
                add("ERROR", DB_MISMATCH_PDF, f"apendice[{name}].{label}",
                    f"'{label}': BD={val!r} PDF={got or '(no impreso)'!r}")

    # filas de ítems: bloque desde la línea que empieza con el código hasta el siguiente código (la descripción
    # puede ocupar varias líneas)
    codes, n = [], 1
    while f"cuerpoDocumento.codigo[{n}]" in exp:
        codes.append(exp[f"cuerpoDocumento.codigo[{n}]"])
        n += 1
    blocks = item_blocks(info.lines, codes)
    for n, code in enumerate(codes, 1):
        if code not in blocks:
            add("ERROR", DB_MISMATCH_PDF, f"cuerpoDocumento.codigo[{n}]", f"Ítem {n}: el código {code} de la BD no aparece como fila del PDF")
        else:
            row = [pres.canon(_tok(t)) for ln in blocks[code] for t in ln.split()]
            res.checked += 1
            for name in ("cantidad", "precioUni", "montoDescu"):
                e = exp.get(f"cuerpoDocumento.{name}[{n}]")
                if strict(e) is None or cfg.category(doc_type, f"cuerpoDocumento.{name}") not in APPROVED \
                        or any(eq(t, e) for t in row):
                    continue
                shown = format_match(e, blocks[code])
                if shown:  # mismo valor con otro formato (menos decimales): aviso, no error
                    add("WARNING", "PDF_FORMAT_DIFF", f"cuerpoDocumento.{name}[{n}]",
                        f"Ítem {n} ({code}): {name} BD={e} PDF={shown} (mismo valor con otro formato)")
                else:
                    add("ERROR", DB_MISMATCH_PDF, f"cuerpoDocumento.{name}[{n}]", f"Ítem {n} ({code}): {name} BD={e} no figura en la fila del PDF")
    return res


# ------------------------------------------------------------------ orquestación
def validate_db(expected: Expected, docs: dict, pdf_paths, spec, cfg: BdConfig, pres: Presentation, doc_type: str = "") -> list:
    """docs: {ruta: Doc} (JSON/XML ya cargados); pdf_paths: PDF a revisar (solo MINH)."""
    results = []
    for f, d in docs.items():
        if d.raw is None or d.fmt not in ("JSON", "XML"):
            continue
        dt = doc_type or d.doc_type
        if dt in OUT_OF_SCOPE:
            results.append(Result(file=f, doc_type=dt, fmt=f"BD<->{d.fmt}", findings=[Finding(
                "WARNING", DB_MISSING, "", f"{dt}: tipo de documento fuera del alcance de la comparación con la BD "
                "(su fuente no está ubicada); no se verifica su información")]))
            continue
        results.append(check_doc(f, d, expected.get(document_id(d)), spec, cfg, dt))
    for p in pdf_paths:
        info = read_pdf(p)
        if info.kind != "MINH":
            continue
        dt = doc_type or info.title or next((d.doc_type for d in docs.values() if d.raw is not None and document_id(d) == info.uuid), "")
        results.append(check_pdf_db(info, expected.get(info.uuid), spec, cfg, pres, dt))
    return results
