"""Correctitud de la información, limitada a lo que definen los specs (Excel y ejemplos entregados).

Además de las reglas del comparador (obligatoriedad, nulos, catálogos `valores` del Excel, campos eliminados/renombrados),
el Excel define el número de control como `DTE-##-#####-###############`, donde `##` es el tipo de documento
(CAT-002) y el último bloque son 15 dígitos. La longitud del bloque central NO se valida: el Excel indica 5 caracteres
en la fila y 'DTE-11-00101-...' en un ejemplo, mientras los JSON de ejemplo usan 8 ('M001P001').

Los tipos de dato los valida el comparador (columna Tipo). No se validan formatos de fecha/hora, NIT/DUI, teléfonos, correos ni totales: los specs no los definen.

Códigos:
  FORMAT_INVALID      el número de control no cumple 'DTE-##-<bloque>-<15 dígitos>'
  CONSISTENCY_ERROR   el tipo del número de control no coincide con tipoDte/tipoFactura o con el tipo de documento
"""
from __future__ import annotations
import re

CONTROL_RE = re.compile(r"^DTE-(\d{2})-[^-]+-(\d{15})$")
# tipoDte / tipoFactura de cada documento (columna TIPO del libro: CCF-03, F-01, FE-11, NC-05, ND-06, REM-04)
DTE_BY_TYPE = {"FAC": "01", "CCF": "03", "REM": "04", "NCR": "05", "NDB": "06", "FAC_EXP": "11"}


def _null(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip().lower() in ("", "null"))


def _control(out, path, value, tipo_dte, doc_type):
    if _null(value):
        return
    m = CONTROL_RE.match(str(value))
    if not m:
        out.append(("ERROR", "FORMAT_INVALID", path, f"Número de control '{value}' debe ser DTE-##-<bloque>-<15 dígitos>"))
        return
    for label, expected in (("tipoDte del documento", tipo_dte), (f"tipo {doc_type}", DTE_BY_TYPE.get(doc_type))):
        if expected and not _null(expected) and str(expected) != m.group(1):
            out.append(("ERROR", "CONSISTENCY_ERROR", path,
                        f"Número de control tiene tipo '{m.group(1)}' pero el {label} es '{expected}'"))
            return


def check(doc, doc_type: str) -> list[tuple]:
    out: list[tuple] = []
    if doc.fmt == "JSON" and isinstance(doc.raw, dict):
        ident = doc.raw.get("identificacion") or {}
        _control(out, "identificacion.numeroControl", ident.get("numeroControl"), ident.get("tipoDte"), doc_type)
    elif doc.fmt == "XML":
        def t(p):
            n = doc.raw.find(p)
            return None if n is None else (n.text or "").strip()
        _control(out, "DatosEntrada.Encabezado.secuenciaDocumento", t("DatosEntrada/Encabezado/secuenciaDocumento"),
                 t("DatosEntrada/Encabezado/tipoFactura"), doc_type)
    return out


def apendice_names(doc, key_path: str = "DatosEntrada.Notas.etiqueta") -> list[tuple[str, str]] | None:
    """[(nombre del apéndice, valor)] en el orden del documento. JSON: apendice[].campo/valor; XML: <Notas> por key_path.
    None si el documento no es JSON/XML legible."""
    if doc.fmt == "JSON" and isinstance(doc.raw, dict):
        arr = doc.raw.get("apendice")
        return [(str(e.get("campo") or ""), str(e.get("valor") or "")) for e in arr if isinstance(e, dict)] \
            if isinstance(arr, list) else []
    if doc.fmt == "XML" and doc.raw is not None:
        parent, key = key_path.rsplit(".", 1)
        out = []
        for el in doc.raw.findall(parent.replace(".", "/")):
            k = el.find(key)
            if k is not None and (k.text or "").strip():
                v = el.find("notasAdicionales")
                out.append(((k.text or "").strip(), "" if v is None else (v.text or "").strip()))
        return out
    return None


XML_REPEATED = ("OtrosTotales", "DescuentoCargo", "Notas", "TotalImpuestoItem")  # bloques que el libro indexa por posición
XML_ITEMS = "DatosEntrada.Items"


def _xml_walk(elements, segments):
    for seg in segments:
        elements = [c for e in elements for c in list(e) if c.tag.split("}")[-1] == seg]
    return elements


def _has_value(elements) -> bool:
    return any((e.text or "").strip() or len(e) for e in elements)


def _where(scopes: list, total: int) -> str:
    return "" if total <= 1 else f" (ítems {', '.join(str(i) for i in scopes)} de {total})"


def check_xml_occurrences(doc, doc_type: str, spec) -> list[tuple]:
    """Tags XML que el libro exige y que la presencia de la ruta no demuestra:
    * posicionales (p. ej. valorOtrosTotales(4)): debe existir la posición indicada, no basta otra repetición del tag;
    * tags de ítem: deben venir en TODOS los ítems, no solo en alguno.
    """
    root = getattr(doc, "raw", None)
    if root is None or getattr(doc, "fmt", "") != "XML":
        return []
    out, positional = [], {}
    for j in spec:
        m = re.search(r"posicional\((\d+)\)", j.note or "")
        if j.doc_type == doc_type and j.fmt == "JSON" and j.counterpart and m and j.change != "ELIMINADO":
            cur = positional.get((j.counterpart, int(m.group(1))))
            rank = lambda x: (x.required, not x.nullable)
            if cur is None or rank(j) > rank(cur):
                positional[(j.counterpart, int(m.group(1)))] = j
    for (path, k), j in sorted(positional.items()):
        segs = path.split(".")
        r = next((i for i, x in enumerate(segs) if x in XML_REPEATED), None)
        if r is None:
            continue
        scopes = _xml_walk([root], segs[:r])
        missing, empty = [], []
        for i, scope in enumerate(scopes or [None], 1):
            blocks = _xml_walk([scope], segs[r:r + 1]) if scope is not None else []
            found = _xml_walk([blocks[k - 1]], segs[r + 1:]) if len(blocks) >= k else []
            if not found:
                missing.append(i)
            elif not _has_value(found):
                empty.append(i)
        src = f"{j.sheet} fila {j.row}" if j.row else j.sheet
        label, total = f"{path}[{k}]", len(scopes)
        if missing and j.required:
            out.append(("ERROR", "REQUIRED_MISSING", label, f"Campo requerido no encontrado en la posición {k}{_where(missing, total)} "
                        f"(cambio: {j.change}, {src})"))
        elif missing:
            out.append(("WARNING", "XML_OPTIONAL_MISSING", label, f"Tag del contrato no recibido en la posición {k}{_where(missing, total)}; "
                        f"el libro no lo marca como requerido ({src})"))
        if empty and j.required and not j.nullable:
            out.append(("ERROR", "NULL_NOT_ALLOWED", label, f"Campo requerido vacío en la posición {k}{_where(empty, total)} ({src})"))
    items = _xml_walk([root], XML_ITEMS.split("."))
    if len(items) > 1:
        for s in spec:
            if s.doc_type != doc_type or s.fmt != "XML" or not s.required or s.change == "ELIMINADO" or "posicional" in (s.note or ""):
                continue
            if not s.path.startswith(XML_ITEMS + "."):
                continue
            sub = s.path[len(XML_ITEMS) + 1:].split(".")
            if any(x in XML_REPEATED for x in sub):
                continue
            per_item = [_xml_walk([it], sub) for it in items]
            if not any(per_item):
                continue  # falta en todo el documento: ya lo reporta la validación general
            src = f"{s.sheet} fila {s.row}" if s.row else s.sheet
            missing = [i for i, f in enumerate(per_item, 1) if not f]
            empty = [i for i, f in enumerate(per_item, 1) if f and not _has_value(f)]
            if missing:
                out.append(("ERROR", "REQUIRED_MISSING", s.path, f"Campo requerido no encontrado{_where(missing, len(items))} "
                            f"(cambio: {s.change}, {src})"))
            if empty and not s.nullable and len(empty) < len(items):
                out.append(("ERROR", "NULL_NOT_ALLOWED", s.path, f"Campo requerido vacío{_where(empty, len(items))} ({src})"))
    return out


def check_apendice_attributes(doc, doc_type: str, spec) -> list[tuple]:
    """Si el libro exige apendice/campo, etiqueta y valor, cada apéndice recibido debe traer los tres con dato."""
    req = {s.path for s in spec if s.doc_type == doc_type and s.fmt == "JSON" and s.required and s.change != "ELIMINADO"}
    if not {"apendice.campo", "apendice.etiqueta", "apendice.valor"} <= req or getattr(doc, "raw", None) is None:
        return []
    items = []
    if doc.fmt == "JSON":
        raw = doc.raw.get("apendice") if isinstance(doc.raw, dict) else None
        for it in raw if isinstance(raw, list) else []:
            if isinstance(it, dict):
                items.append({"campo": it.get("campo"), "etiqueta": it.get("etiqueta"), "valor": it.get("valor")})
                items[-1]["_keys"] = set(it)
    else:
        for n in doc.raw.findall("DatosEntrada/Notas"):
            tx = {c.tag.split("}")[-1]: (c.text or "").strip() for c in n}
            if "campo" in tx or "etiqueta" in tx:  # un <Notas> solo con notasAdicionales no es un apéndice
                items.append({"campo": tx.get("campo"), "etiqueta": tx.get("etiqueta"), "valor": tx.get("notasAdicionales"),
                              "_keys": {k for k, t in (("campo", "campo"), ("etiqueta", "etiqueta"), ("valor", "notasAdicionales")) if t in tx}})
    out = []
    for i, it in enumerate(items, 1):
        names = [str(it[k]).strip() for k in ("campo", "etiqueta") if re.fullmatch(r"apendice\d+", str(it[k] or "").strip(), re.I)]
        name = names[0] if names else str(i)
        for attr in ("campo", "etiqueta", "valor"):
            path = f"apendice[{name}].{attr}"
            if attr not in it["_keys"]:
                out.append(("ERROR", "REQUIRED_MISSING", path, f"Falta el atributo '{attr}' del apéndice {name}: los 3 atributos son obligatorios"))
            elif _null(it[attr]):
                out.append(("ERROR", "NULL_NOT_ALLOWED", path, f"El atributo '{attr}' del apéndice {name} viene vacío: los 3 atributos son obligatorios"))
    return out


def check_apendice(doc, doc_type: str, regla: dict | None, key_path: str = "DatosEntrada.Notas.etiqueta") -> list[tuple]:
    """Regla general 'apendice_obligatorio' (opcional en data/specs/ajustes_contrato.json; hoy inactiva):
    APENDICE_MISSING (ERROR)  falta uno de los apéndices obligatorios (apendice01..04)
    APENDICE_EXTRA (WARNING)  apéndice adicional (p. ej. apendice05): revisar manualmente
    """
    if not regla or doc_type not in regla.get("tipos", []):
        return []
    found = apendice_names(doc, key_path)
    if found is None:
        return []
    where = "apendice" if doc.fmt == "JSON" else key_path.rsplit(".", 1)[0]
    names = [n for n, _ in found]
    obligatorios = list(regla.get("obligatorios", []))
    out = []
    for n in obligatorios:
        if n not in names:
            out.append(("ERROR", "APENDICE_MISSING", f"{where}[{n}]",
                        f"Falta el apéndice obligatorio '{n}' (obligatorios: {', '.join(obligatorios)})"))
    for n, v in found:
        if n not in obligatorios:
            out.append(("WARNING", "APENDICE_EXTRA", f"{where}[{n}]",
                        f"Apéndice adicional '{n}' con valor '{v}': no está entre los obligatorios; revisar manualmente"))
    return out
