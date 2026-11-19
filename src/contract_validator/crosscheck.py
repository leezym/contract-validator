"""Cruce XML (TQ) contra JSON (PBS) del mismo documento.

El mapeo sale del propio contrato plano: cada fila JSON trae su tag XML en `ruta_contraparte`.
Se comparan los valores hoja a hoja con igualdad EXACTA de texto (solo null/vacío = sin valor). Los mapeos
posicionales del libro (`OtrosTotales(4)`, `notasAdicionales[2]`) no se comparan 1 a 1 y se omiten.

Códigos:
  XREF_MISMATCH        el valor difiere entre XML y JSON
  XREF_MISSING         el valor existe en un lado y falta (o es nulo) en el otro
  XREF_FORMAT_DIFF     mismo valor con distinto formato (05 vs 5, DD/MM/YYYY vs ISO, mayúsculas): aviso. Con la regla
                       'decimales_equivalentes' (ajustes_contrato.json) los números con el mismo valor decimal
                       ('285.6' vs '285.60', '2260' vs '2260.00') son IGUALES y no generan hallazgo.
  XREF_COUNT_MISMATCH  distinta cantidad de valores presentes (sin contar null), p.ej. ítems del XML vs cuerpoDocumento del JSON
  XREF_DOCTYPE         los dos archivos son de tipos de documento distintos
"""
from __future__ import annotations
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from .models import Result, Finding, REMOVED

MAX_PER_FIELD = 3


def json_values(raw) -> dict[str, list]:
    out: dict[str, list] = {}

    def walk(o, path):
        if isinstance(o, dict):
            for k, v in o.items():
                walk(v, f"{path}.{k}" if path else k)
        elif isinstance(o, list):
            for v in o:
                walk(v, path)
        else:
            out.setdefault(path, []).append(o)
    walk(raw, "")
    return out


def xml_values(root) -> dict[str, list]:
    out: dict[str, list] = {}

    def walk(e, path):
        for k, v in e.attrib.items():
            if not k.startswith("xmlns"):
                out.setdefault(f"{path}.@{k.split('}')[-1]}", []).append(v)
        for c in e:
            p = f"{path}.{c.tag.split('}')[-1]}" if path else c.tag.split("}")[-1]
            if len(c):
                walk(c, p)
            else:
                out.setdefault(p, []).append((c.text or "").strip())
    walk(root, "")
    return out


def strict(v):
    """None = sin valor ('null' o vacío, según el contrato). Todo lo demás se compara como texto exacto."""
    if v is None:
        return None
    if isinstance(v, bool):
        return "true" if v else "false"
    t = str(v).strip()
    return None if t == "" or t.lower() == "null" else t


def loose(v):
    """Solo para detectar diferencias de FORMATO (no de valor): número, fecha DD/MM/YYYY, mayúsculas, ceros."""
    t = strict(v)
    if t is None:
        return None
    if re.fullmatch(r"[+-]?\d+(\.\d+)?", t):
        try:
            return Decimal(t)
        except InvalidOperation:
            return t
    for fmt in ("%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(t, fmt).date().isoformat()
        except ValueError:
            pass
    return t.upper()


_DECIMAL = re.compile(r"[+-]?(0|[1-9]\d*)(\.\d+)?")  # sin ceros a la izquierda: '05' vs '5' sigue siendo formato


def same_decimal(a, b) -> bool:
    """'285.6' == '285.60' == '285.600'; '2260' == '2260.00'. Códigos con ceros a la izquierda no cuentan."""
    if a is None or b is None or not (_DECIMAL.fullmatch(a) and _DECIMAL.fullmatch(b)):
        return False
    return Decimal(a) == Decimal(b)


def _show(v):
    return "null" if v is None else (format(v, "f") if isinstance(v, Decimal) else repr(v))


APENDICE_NAME = re.compile(r"^apendice\d+$", re.I)


def apendice_records(doc) -> dict[str, dict]:
    """Apéndices del documento como unidades: {apendiceNN: {'etiqueta': texto, 'valor': datos unidos}}.
    Cada apéndice es la unión de varios datos; el nombre se reconoce en campo o en etiqueta."""
    out: dict[str, dict] = {}
    if doc.fmt == "JSON":
        items = doc.raw.get("apendice") if isinstance(doc.raw, dict) else None
        for it in items if isinstance(items, list) else []:
            if isinstance(it, dict):
                for key, other in (("campo", "etiqueta"), ("etiqueta", "campo")):
                    name = str(it.get(key) or "").strip()
                    if APENDICE_NAME.match(name):
                        out.setdefault(name, {"etiqueta": it.get(other), "valor": it.get("valor")})
                        break
        return out
    for n in doc.raw.findall("DatosEntrada/Notas"):
        tx = {c.tag.split("}")[-1]: (c.text or "").strip() for c in n}
        for key, other in (("campo", "etiqueta"), ("etiqueta", "campo")):
            if APENDICE_NAME.match(tx.get(key, "")):
                out.setdefault(tx[key], {"etiqueta": tx.get(other), "valor": tx.get("notasAdicionales")})
                break
    return out


def document_id(doc) -> str:
    """Identificador común para emparejar: código de generación (JSON) / uuid de CabeceraEntrada (XML)."""
    if doc.fmt == "XML":
        n = doc.raw.find("CabeceraEntrada/uuid")
        return (n.text or "").strip().upper() if n is not None else ""
    return str(((doc.raw or {}).get("identificacion") or {}).get("codigoGeneracion") or "").strip().upper()


def pair_documents(docs: dict) -> tuple[list[tuple[str, str]], list[str]]:
    """docs: {ruta: Doc}. Devuelve ([(xml, json)], rutas sin pareja) emparejando por UUID."""
    by_id: dict[str, dict] = {}
    for f, d in docs.items():
        if d.raw is None:
            continue
        i = document_id(d)
        if i:
            by_id.setdefault(i, {})[d.fmt] = f
    pairs = [(v["XML"], v["JSON"]) for v in by_id.values() if "XML" in v and "JSON" in v]
    paired = {f for p in pairs for f in p}
    return sorted(pairs), sorted(f for f in docs if f not in paired)


def cross_check(xml_file, xml_doc, json_file, json_doc, spec, doc_type: str = "", reglas: dict | None = None) -> Result:
    """reglas: reglas generales activas (load_reglas); 'decimales_equivalentes' iguala '285.6' y '285.60'."""
    doc_type = doc_type or json_doc.doc_type or xml_doc.doc_type
    res = Result(file=f"{xml_file} <-> {json_file}", doc_type=doc_type, fmt="XML<->JSON")

    def add(sev, code, path, msg):
        res.findings.append(Finding(sev, code, path, msg))

    if xml_doc.doc_type and json_doc.doc_type and xml_doc.doc_type != json_doc.doc_type:
        add("ERROR", "XREF_DOCTYPE", "", f"El XML es {xml_doc.doc_type} y el JSON es {json_doc.doc_type}")
        return res
    xv, jv = xml_values(xml_doc.raw), json_values(json_doc.raw)
    rows = [s for s in spec if s.doc_type == doc_type and s.fmt == "JSON" and s.counterpart
            and s.change != REMOVED and "posicional" not in s.note and not s.path.startswith("apendice.")]
    seen_counts = set()
    decimales = bool((reglas or {}).get("decimales_equivalentes"))
    for s in rows:
        j, x = jv.get(s.path), xv.get(s.counterpart)
        jn = [strict(v) for v in j] if j else []
        xn = [strict(v) for v in x] if x else []
        jhas, xhas = any(v is not None for v in jn), any(v is not None for v in xn)
        if not jhas and not xhas:
            continue
        res.checked += 1
        label = f"{s.path} <-> {s.counterpart}"
        if jhas != xhas:
            side = "JSON" if xhas else "XML"
            add("ERROR", "XREF_MISSING", s.path, f"{label}: el valor existe en {'XML' if xhas else 'JSON'} "
                f"({_show(next(v for v in (xn if xhas else jn) if v is not None))}) y falta en {side}")
            continue
        if len(jn) != len(xn):  # un tag repetido puede venir en null en algunas posiciones: cuentan los valores presentes
            jn, xn = [v for v in jn if v is not None], [v for v in xn if v is not None]
        if len(jn) != len(xn):
            key = (s.path.rsplit(".", 1)[0], s.counterpart.rsplit(".", 1)[0])
            if key not in seen_counts:  # una vez por contenedor
                seen_counts.add(key)
                add("ERROR", "XREF_COUNT_MISMATCH", s.path, f"{label}: JSON tiene {len(jn)} valor(es) y XML {len(xn)}")
            continue
        shown = 0
        for i, (a, b) in enumerate(zip(jn, xn), 1):
            if a == b or (decimales and same_decimal(a, b)):
                continue
            if loose(a) == loose(b):
                add("WARNING", "XREF_FORMAT_DIFF", f"{s.path}{'[%d]' % i if len(jn) > 1 else ''}",
                    f"{label}: mismo valor con distinto formato JSON={_show(a)} XML={_show(b)}")
            elif a != b:
                shown += 1
                if shown <= MAX_PER_FIELD:
                    where = f"[{i}]" if len(jn) > 1 else ""
                    add("ERROR", "XREF_MISMATCH", f"{s.path}{where}", f"{label}: JSON={_show(a)} XML={_show(b)}")
        if shown > MAX_PER_FIELD:
            add("ERROR", "XREF_MISMATCH", s.path, f"{label}: {shown - MAX_PER_FIELD} diferencia(s) más")
    if any(s.doc_type == doc_type and s.fmt == "JSON" and s.path == "apendice.valor" for s in spec):
        _cross_apendice(xml_doc, json_doc, res, add)
    return res


def _cross_apendice(xml_doc, json_doc, res, add):
    """Compara cada apéndice completo, ubicado por su nombre apendiceNN."""
    jr, xr = apendice_records(json_doc), apendice_records(xml_doc)
    for name in sorted(set(jr) | set(xr)):
        res.checked += 1
        path = f"apendice[{name}]"
        j, x = jr.get(name), xr.get(name)
        if j is None:
            add("ERROR", "XREF_MISSING", path, f"Apéndice {name}: existe en XML y falta en JSON")
        elif x is None:
            add("ERROR", "XREF_MISSING", path, f"Apéndice {name}: existe en JSON y falta en XML")
        else:
            for part in ("etiqueta", "valor"):
                a, b = strict(j[part]), strict(x[part])
                if a != b:
                    add("ERROR", "XREF_MISMATCH", path, f"Apéndice {name}, {part}: JSON={_show(a)} XML={_show(b)}")
