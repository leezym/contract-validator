"""Auditoría del mapeo JSON <-> XML_TQ que define el contrato (sin archivos reales). Solo reporta hechos del libro.

Códigos:
  NO_XML_SOURCE     campo JSON requerido (hoja hoja) sin tag XML de origen (ERROR si no admite null, WARNING si admite null)
  MULTI_SOURCE      un mismo tag XML alimenta varios campos JSON (WARNING)
  XML_NO_DEST       tag XML requerido que no alimenta ningún campo JSON (WARNING)
  DIFF_ACROSS_TYPES el mismo campo JSON apunta a tags XML distintos según el tipo de documento (WARNING)
  POSITIONAL        mapeo posicional del libro; el cruce XML<->JSON no lo compara (INFO)
"""
from __future__ import annotations
import os
from collections import defaultdict
from .models import REMOVED, SpecField

HEADERS = ["tipo_doc", "nivel", "codigo", "ruta_json", "ruta_xml", "detalle", "hoja_origen", "fila_origen"]


def _leaf(p: str) -> str:
    return p.rsplit(".", 1)[-1]


def audit(spec: list[SpecField], doc_type: str = "") -> list[dict]:
    out: list[dict] = []
    types = sorted({s.doc_type for s in spec if s.doc_type and (not doc_type or s.doc_type == doc_type)})

    def add(t, lvl, code, s, detail, xml=""):
        out.append(dict(tipo_doc=t, nivel=lvl, codigo=code, ruta_json=s.path if s and s.fmt == "JSON" else "",
                        ruta_xml=xml or (s.path if s and s.fmt == "XML" else ""), detalle=detail,
                        hoja_origen=s.sheet if s else "", fila_origen=s.row if s else ""))

    for t in types:
        js = [s for s in spec if s.doc_type == t and s.fmt == "JSON" and s.change != REMOVED]
        xs = [s for s in spec if s.doc_type == t and s.fmt == "XML"]
        paths = {s.path for s in js}
        is_leaf = lambda s: s.type not in ("object", "array") and not any(p.startswith(s.path + ".") for p in paths)
        by_xml: dict[str, list[SpecField]] = defaultdict(list)
        for s in js:
            if s.counterpart and "posicional" not in s.note:
                by_xml[s.counterpart].append(s)

        for s in js:
            if s.required and is_leaf(s) and not s.counterpart:
                add(t, "WARNING" if s.nullable else "ERROR", "NO_XML_SOURCE", s,
                    "Campo JSON requerido sin tag XML_TQ de origen en el libro"
                    + (" (admite null)" if s.nullable else ""))
            if s.counterpart and "posicional" in s.note:
                add(t, "INFO", "POSITIONAL", s, "Mapeo por posición en el libro (no comparable 1 a 1)", s.counterpart)
        for xml, rows in by_xml.items():
            if len(rows) > 1:
                add(t, "WARNING", "MULTI_SOURCE", rows[0],
                    "El mismo tag alimenta: " + ", ".join(r.path for r in rows) + "", xml)
        mapped = {s.counterpart for s in js if s.counterpart}
        for x in xs:
            if (x.required and x.type not in ("object", "array") and not x.path.startswith("CabeceraEntrada")
                    and x.path not in mapped and not any(m.startswith(x.path + ".") for m in mapped)):
                add(t, "WARNING", "XML_NO_DEST", x, "Tag XML requerido que no alimenta ningún campo JSON", x.path)


    by_json: dict[str, dict[str, SpecField]] = defaultdict(dict)
    for s in spec:
        if s.fmt == "JSON" and s.counterpart and s.change != REMOVED and "posicional" not in s.note and s.doc_type in types:
            by_json[s.path][s.doc_type] = s
    for path, per_type in by_json.items():
        variants = {s.counterpart for s in per_type.values()}
        if len(variants) > 1:
            first = next(iter(per_type.values()))
            detail = "; ".join(f"{t}: {s.counterpart}" for t, s in sorted(per_type.items()))
            add("*", "WARNING", "DIFF_ACROSS_TYPES", first, detail)
    order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    return sorted(out, key=lambda r: (order[r["nivel"]], r["codigo"], r["tipo_doc"], str(r["ruta_json"] or r["ruta_xml"])))


def summarize(rows: list[dict]) -> dict:
    s: dict = defaultdict(lambda: defaultdict(int))
    for r in rows:
        s[r["codigo"]][r["tipo_doc"]] += 1
    return {k: dict(v) for k, v in s.items()}


def to_text(rows: list[dict]) -> str:
    lines = [f"Auditoría del mapeo: {len(rows)} hallazgo(s)"]
    for code, per in sorted(summarize(rows).items()):
        lines.append(f"  {code:<18}{sum(per.values()):>4}   " + ", ".join(f"{t}={n}" for t, n in sorted(per.items())))
    return "\n".join(lines)


def to_xlsx(rows: list[dict], path: str):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    summ = summarize(rows)
    types = sorted({t for per in summ.values() for t in per})
    ws.append(["codigo", "total"] + types)
    for code, per in sorted(summ.items()):
        ws.append([code, sum(per.values())] + [per.get(t, 0) for t in types])
    hs = wb.create_sheet("Hallazgos")
    hs.append(HEADERS)
    for r in rows:
        hs.append([r[k] for k in HEADERS])
    hs.freeze_panes = "A2"
    hs.auto_filter.ref = hs.dimensions
    for col, w in zip("ABCDEFGH", (10, 9, 20, 42, 60, 90, 26, 10)):
        hs.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
