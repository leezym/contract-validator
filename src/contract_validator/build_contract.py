"""Genera el contrato plano (hoja 'contrato' + hoja 'lint') a partir del libro de especificación original.

Fuente de los tags (confirmada por el cliente): hojas consolidadas 'XML' y 'JSON'.
  * 'XML': una fila por campo con TAG_MIN_HACIENDA (genérico), Descripcion, Tipo, XML_TQ, Requerido y una columna
    por tipo de documento (CCF - 03, F -01, FE - 11, NC - 05, ND - 06, NOTA REMISION - 04, EVENTO INVALIDACION,
    EVENTO RETORNO) con la ruta JSON v4 de ese tipo ('Adicionar/Eliminar/Renombrar' marcan el cambio). La ruta del
    tipo se empareja con el XML_TQ de SU fila. EVENTO RETORNO trae su propio tag XML en la columna 'XML-ER' y sus
    atributos (Tipo/Requerido) se toman de la fila cuyo TAG_MIN_HACIENDA coincide con la ruta.
  * 'JSON': ruta JSON v3 (base, antes de los cambios) por tipo; se usa para verificar la columna del tipo en 'XML'
    (diferencias -> 'lint').
  * La hoja 'XML' no trae tags XML propios para EVENTO INVALIDACION: ese tipo se sigue leyendo de su hoja
    'Invalidacion' (layout por tipo: TAG v3, TAG v4, XML_TQ, Tipo, Accion, Obligatorio).
"""
from __future__ import annotations
import re
import unicodedata
from openpyxl import Workbook, load_workbook
from .models import (CCF, FAC, FAC_EXP, NCR, NDB, REM, INVALIDACION, RETORNO, NEW, REMOVED, RENAMED, UNCHANGED)
from .spec_loader import TYPE_MAP

COLUMNS = ["tipo_doc", "formato", "ruta", "tipo", "valores", "requerido", "cambio", "ruta_anterior",
           "ruta_contraparte", "descripcion", "hoja_origen", "fila_origen", "nota"]

SHEET_CODES = {
    "comprobante de credito fiscal": CCF, "factura electronica": FAC, "factura de exportacion": FAC_EXP,
    "nota de credito": NCR, "nota de debito": NDB, "nota de remision": REM,
    "invalidacion": INVALIDACION, "retorno": RETORNO,
}
XML_SHEET = "XML"
req_text: dict = {}  # texto original de Requerido por (tipo, ruta), para avisar ambiguedades
MAX_COL = 30
KNOWN_REQ = {"si", "si (null)", "no", "n/a", "si para fex", "requerido", "si/null/ por tipodte"}  # la hoja XML declara dimensiones hasta XFA; solo importan las primeras columnas


def _ascii(s) -> str:
    s = unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip()


def clean(s) -> str:
    """Texto de celda: sin \\xa0, espacios colapsados, sin '*' final."""
    t = re.sub(r"\s+", " ", str(s if s is not None else "").replace("\xa0", " ")).strip()
    return t.rstrip("*").strip()


def clean_tag(s, strip_items=True) -> str:
    """Ruta con '/': sin '[]', sin segmento 'items' (los arrays no llevan nivel intermedio en el aplanado)."""
    t = re.sub(r"\[\d*\]|\(.*?\)", "", clean(s))  # '[]', '[4]' y notas entre paréntesis
    toks = t.split()
    if len(toks) > 1 and toks[1].lower().startswith("normativa"):
        t = toks[0]
    parts = [p.strip() for p in t.split("/") if p.strip() and not (strip_items and p.strip() == "items")]
    return "/".join(parts)


_CHANGE_RE = [
    (re.compile(r"^renombrar\s+(.+?)\s*(?:->|→)\s*(.+)$", re.I), RENAMED),
    (re.compile(r"^adicionar\s+(.+)$", re.I), NEW),
    (re.compile(r"^elim[a-z]*\s+(.+)$", re.I), REMOVED),  # acepta el typo 'Elimnar' del libro
]


def parse_change(text) -> tuple[str, str, str]:
    """'Adicionar X' -> (NUEVO, X, ''); 'Renombrar A -> B' -> (RENOMBRADO, B, A); 'X' -> (SIN_CAMBIO, X, '')."""
    t = clean(text)
    for rx, kind in _CHANGE_RE:
        m = rx.match(t)
        if m:
            if kind == RENAMED:
                return kind, clean_tag(m.group(2)), clean_tag(m.group(1))
            return kind, clean_tag(m.group(1)), ""
    return UNCHANGED, clean_tag(t), ""


def parse_type(raw) -> tuple[str, list[str], str | None]:
    """-> (tipo normalizado, valores permitidos, aviso)."""
    t = clean(raw)
    if not t:
        return "", [], None
    if re.fullmatch(r"[\w.]+(\s*,\s*[\w.]+)+", t):
        return "", [v.strip() for v in t.split(",")], None
    if t == "USD":
        return "string", ["USD"], None
    key = _ascii(t).lower()
    if key in TYPE_MAP:
        return TYPE_MAP[key], [], None
    if re.fullmatch(r"\d{2}", t):  # catálogo de un único valor (p.ej. '09')
        return "", [t], None
    return "", [], f"Tipo no reconocido '{t}' (no se valida tipo)"


def parse_required(raw, code: str) -> str | None:
    s = _ascii(raw).lower()
    if not s or s == "requerido":
        return None
    if s.startswith("si para fex"):
        return "SI" if code == FAC_EXP else "NO"
    if s.startswith("si") and "null" in s:  # 'Si (null)' y 'Si/null/ por tipoDTE': requerido, admite null
        return "SI_NULL"
    if s == "si":
        return "SI"
    if s in ("no", "n/a"):
        return "NO"
    return None


def normalize_xml_path(p: str) -> str:
    """Unifica las variantes del libro con la interfaz TQ real: el tag correcto es 'InformacioGeneral' (sin 'n')."""
    p = p.replace("/InformacionGeneral/", "/InformacioGeneral/").replace("/Adquiriente/", "/InformacionAdquiriente/")
    m = re.match(r"^DatosEntrada/(InformacionEmisor|InformacionAdquiriente)/(.*)$", p)
    return f"DatosEntrada/InformacioGeneral/{m.group(1)}/{m.group(2)}" if m else p


def xml_path(raw) -> str:
    """Ruta XML_TQ válida ('DatosEntrada/...' o 'CabeceraEntrada/...'); '' si es 'x', NULL, nota libre, etc."""
    t = re.sub(r"\([^()]*\)(?=/)", "", clean(raw))      # 'Notas(3..N)/etiqueta' -> 'Notas/etiqueta'
    t = re.sub(r"\(.*$", "", t).strip()                # 'valorOtrosTotales(4)', notas libres al final
    if re.match(r"[A-Z]\w+/\w", t) and not t.startswith(("DatosEntrada/", "CabeceraEntrada/")):
        t = "DatosEntrada/" + t  # el libro omite a veces el prefijo (p.ej. 'ImportesTotal/totalFactura')
    return normalize_xml_path(clean_tag(t, strip_items=False)) if t.startswith(("DatosEntrada/", "CabeceraEntrada/")) else ""


def _find_header(rows):
    """Primera fila (de las primeras 10) que contiene un encabezado TAG_MIN_HACIENDA."""
    for i, r in enumerate(rows[:10]):
        if any(str(c or "").upper().startswith("TAG_MIN_HACIENDA") for c in r):
            return i
    raise ValueError("No se encontró fila de encabezados con TAG_MIN_HACIENDA")


def _cols(header) -> dict:
    col, tags = {}, []
    for i, h in enumerate(header):
        k = _ascii(h).lower()
        if k.startswith("tag_min_hacienda"):
            tags.append(i)
        elif k == "xml_tq" and "xml" not in col:
            col["xml"] = i
        elif k == "tipo" and (str(h).strip() == "Tipo" or "tipo" not in col):
            col["tipo"] = i  # 'Tipo' (tipo de dato) y no 'TIPO' (código numérico del documento)
        elif k == "descripcion":
            col["desc"] = i
        elif k in ("accion",) and "accion" not in col:
            col["accion"] = i
        elif k in ("obligatorio",) and "oblig" not in col:
            col["oblig"] = i
    col["tag_v3"] = tags[0]
    col["tag_v4"] = tags[1] if len(tags) > 1 else tags[0]
    return col


def _xml_requirements(wb, lint):
    """{(codigo, ruta_json): 'SI'|'SI_NULL'|'NO'} desde la hoja consolidada 'XML'."""
    req: dict = {}
    req_text.clear()
    if XML_SHEET not in wb.sheetnames:
        lint.append((XML_SHEET, "", "WARNING", "No existe la hoja 'XML'; no hay obligatoriedad: todo queda como NO"))
        return req
    rows = list(wb[XML_SHEET].iter_rows(values_only=True, max_col=MAX_COL))
    head = rows[_find_header(rows)]
    req_col = next((i for i, h in enumerate(head) if _ascii(h).lower() == "requerido"), 4)
    type_cols = {}
    for i, h in enumerate(head):
        k = _ascii(h).upper()
        m = re.search(r"-\s*(\d{2})$", k)
        code = {"03": CCF, "01": FAC, "11": FAC_EXP, "05": NCR, "06": NDB, "04": REM}.get(m.group(1)) if m else None
        if "INVALIDACION" in k:
            code = INVALIDACION
        elif k == "EVENTO RETORNO":
            code = RETORNO
        if code and i > req_col:
            type_cols[i] = code
    for n, r in enumerate(rows[_find_header(rows) + 1:], start=_find_header(rows) + 2):
        base = clean_tag(r[0]) if r and r[0] else ""
        if not base and not any(r[i] for i in type_cols if i < len(r)):
            continue
        raw = r[req_col] if req_col < len(r) else None
        for i, code in type_cols.items():
            cell = r[i] if i < len(r) else None
            if not cell:
                continue
            kind, path, _ = parse_change(cell)
            val = parse_required(raw, code)
            if path and val and kind != REMOVED:
                req[(code, path)] = val
                req_text[(code, path)] = clean(raw)
        # La obligatoriedad solo aplica a los tipos cuya columna marca el campo (celda no vacía). No se copia
        # a los demás tipos desde la columna base.
    return req


def _tab_rows(ws, code: str, req: dict, lint: list) -> list[dict]:
    """Filas de una hoja por tipo (layout antiguo). Solo se usa para INVALIDACION (sin tags XML en la hoja 'XML')."""
    out: list[dict] = []
    rows = list(ws.iter_rows(values_only=True, max_col=MAX_COL))
    h = _find_header(rows)
    c = _cols(rows[h])
    seen_json, seen_xml = {}, {}

    def g(r, k):
        i = c.get(k)
        return r[i] if i is not None and i < len(r) else None

    for n, r in enumerate(rows[h + 1:], start=h + 2):
        v4raw, v3raw = g(r, "tag_v4"), g(r, "tag_v3")
        if code == RETORNO and not clean(v4raw):
            continue
        if not clean(v4raw):
            if clean(v3raw) and clean(g(r, "xml")) not in ("", "x"):
                lint.append((ws.title, n, "WARNING", f"Fila con campo v3 '{clean(v3raw)}' sin definición v4: se omite"))
            continue
        change, path, old = parse_change(v4raw)
        if not path:
            continue
        if not re.fullmatch(r"[\w@][\w@/.\-]*", path):
            if code != RETORNO:
                lint.append((ws.title, n, "WARNING", f"Texto no interpretable como ruta, se omite: '{path[:60]}'"))
            continue
        if change == RENAMED and not old:
            lint.append((ws.title, n, "ERROR", f"RENOMBRADO sin ruta anterior: '{clean(v4raw)}'"))
            continue
        if "=" in clean(v4raw)[:1]:
            lint.append((ws.title, n, "WARNING", "Fórmula sin resolver en la columna de campo"))
            continue
        tipo, values, tw = parse_type(g(r, "tipo"))
        if code == INVALIDACION and tipo and tipo != "object":
            tipo = ""  # la columna Tipo de esta hoja describe el request TQ (todo 'String'), no el JSON
        if tw:
            lint.append((ws.title, n, "INFO", tw))
        required = req.get((code, path)) or (req.get((code, old)) if old else None)
        if not required and code in (INVALIDACION, RETORNO):
            required = parse_required(g(r, "oblig"), code)
        is_container = tipo in ("object", "array") or _ascii(g(r, "tipo")).lower() in ("object", "array")
        raw_req = _ascii(g(r, "oblig") if code in (INVALIDACION, RETORNO) else req_text.get((code, path), "")).lower()
        if raw_req and raw_req not in ("si", "si (null)", "no", "n/a", "si para fex", "requerido", "si/null/ por tipodte"):
            lint.append((ws.title, n, "WARNING", f"Obligatoriedad ambigua '{raw_req}' en '{path}': se toma la lectura mas permisiva (SI_NULL)"))
        if not required:
            required = "NO"
            if change != REMOVED and not is_container:
                lint.append((ws.title, n, "INFO", f"Sin dato de obligatoriedad para '{path}': se asume NO"))
        if change == REMOVED:
            required = "NO"
        xml = xml_path(g(r, "xml"))
        accion = _ascii(g(r, "accion")).upper()
        if path in seen_json:
            lint.append((ws.title, n, "WARNING", f"Ruta duplicada '{path}' (primera en fila {seen_json[path]}): se omite"))
            continue
        seen_json[path] = n
        desc = clean(g(r, "desc"))
        note = []
        if clean(g(r, "xml")) not in ("", "x") and not xml:
            note.append(f"XML_TQ sin ruta: {clean(g(r, 'xml'))[:80]}")
        im = _POSITIONAL.search(clean(g(r, "xml"))) if xml else None
        if im:  # el tag XML se repite y el libro indica la posición: no comparable 1 a 1 en el cruce XML<->JSON
            note.append(f"posicional({im.group(1)})")
        out.append(dict(tipo_doc=code, formato="JSON", ruta=path, tipo=tipo, valores="|".join(values),
                        requerido=required, cambio=change, ruta_anterior=old, ruta_contraparte=xml,
                        descripcion=desc, hoja_origen=ws.title.strip(), fila_origen=n, nota="; ".join(note)))
        if xml:
            # El tag TQ se valida como opcional si el campo JSON se elimina en v4 (el BUS ya no lo mapea).
            xreq = "NO" if change == REMOVED else required
            xcambio = NEW if (accion == "CREAR" and change != REMOVED) else UNCHANGED
            if xml in seen_xml:
                prev = out[seen_xml[xml]]
                if xreq in ("SI", "SI_NULL") and prev["requerido"] == "NO":
                    prev["requerido"] = xreq
                continue
            seen_xml[xml] = len(out)
            out.append(dict(tipo_doc=code, formato="XML", ruta=xml, tipo="" if is_container else tipo,
                            valores="", requerido=xreq, cambio=xcambio, ruta_anterior="",
                            ruta_contraparte=path, descripcion=desc, hoja_origen=ws.title.strip(),
                            fila_origen=n, nota=next((x for x in note if x.startswith("posicional")), "")))
    return out


TYPE_CODES = {"03": CCF, "01": FAC, "11": FAC_EXP, "05": NCR, "06": NDB, "04": REM}
JSON_SHEET = "JSON"
ONLY_NULL = "solo_null"  # nota del plano: el campo JSON debe llegar y su valor siempre es null
XML_ONLY = "solo xml"  # valor de la columna del tipo en la hoja 'XML' para un tag que no tiene campo JSON
XML_SHEET_TYPES = (CCF, FAC, FAC_EXP, NCR, NDB, REM, RETORNO)   # tags desde la hoja 'XML'
TAB_TYPES = (INVALIDACION,)                                      # tags desde su propia hoja
_PATH = re.compile(r"[\w@][\w@/.\-]*")
_POSITIONAL = re.compile(r"[(\[]\s*(\d+)\s*[)\]]")


def sheet_label(code: str) -> str:
    """Etiqueta de 'lint' para la columna de un tipo en la hoja 'XML' (la fila sola no identifica el tipo)."""
    return f"{XML_SHEET} [{code}]"


def type_columns(head) -> tuple[dict, int | None]:
    """Encabezado de la hoja 'XML' (o 'JSON') -> ({columna: tipo}, columna 'XML-ER' del retorno)."""
    cols, er = {}, None
    for i, h in enumerate(head):
        k = _ascii(h).upper()
        m = re.search(r"-\s*(\d{2})$", k)
        code = TYPE_CODES.get(m.group(1)) if m else None
        if "INVALIDACION" in k:
            code = INVALIDACION
        elif k == "EVENTO RETORNO":
            code = RETORNO
        elif k == "XML-ER":
            er = i
        if code:
            cols[i] = code
    return cols, er


def _named_cols(head) -> dict:
    c = {}
    for i, h in enumerate(head):
        key = {"tag_min_hacienda": "tag", "descripcion": "desc", "tipo": "tipo", "xml_tq": "xml",
               "requerido": "req"}.get(_ascii(h).lower())
        if key and key not in c:
            c[key] = i
    missing = {"tag", "tipo", "xml", "req"} - set(c)
    if missing:
        raise ValueError(f"Hoja '{XML_SHEET}': faltan columnas {sorted(missing)}")
    return c


def tab_removed(wb) -> set:
    """(tipo, ruta JSON) marcados 'Eliminar' en la columna TAG_MIN_HACIENDA (v3 -> v4) de las hojas por tipo."""
    out = set()
    for ws in wb.worksheets:
        code = SHEET_CODES.get(_ascii(ws.title).lower())
        if code not in XML_SHEET_TYPES:
            continue
        rows = list(ws.iter_rows(values_only=True, max_col=MAX_COL))
        h = _find_header(rows)
        col = _cols(rows[h])["tag_v4"]
        for r in rows[h + 1:]:
            if col < len(r) and clean(r[col]):
                change, path, _ = parse_change(r[col])
                if change == REMOVED and path:
                    out.add((code, path))
    return out


def _xml_sheet_rows(wb, lint: list) -> list[dict]:
    removed_in_tab = tab_removed(wb)
    rows = list(wb[XML_SHEET].iter_rows(values_only=True, max_col=MAX_COL))
    h = _find_header(rows)
    c = _named_cols(rows[h])
    tcols, er = type_columns(rows[h])
    cell = lambda r, i: r[i] if i is not None and i < len(r) else None
    body = list(enumerate(rows[h + 1:], start=h + 2))
    by_generic = {}  # ruta genérica (TAG_MIN_HACIENDA) -> fila, para los atributos del retorno
    for n, r in body:
        g = parse_change(cell(r, c["tag"]))[1] if clean(cell(r, c["tag"])) else ""
        if g:
            by_generic.setdefault(g, r)
    out: list[dict] = []
    for col, code in tcols.items():
        if code not in XML_SHEET_TYPES:
            continue
        label = sheet_label(code)
        xcol = er if code == RETORNO else c["xml"]
        if xcol is None:
            lint.append((label, "", "ERROR", "No se encontró la columna 'XML-ER' del retorno: sin tags XML"))
        seen_json, seen_xml, json_idx = {}, {}, {}

        def add_xml_row(xml, path, change, required, tipo, is_container, desc, note, n):
            # El tag TQ se valida como opcional si el campo JSON se elimina en v4 (el BUS ya no lo mapea).
            xreq = "NO" if change == REMOVED else required
            if xml in seen_xml:
                prev = out[seen_xml[xml]]
                if xreq in ("SI", "SI_NULL") and prev["requerido"] == "NO":
                    prev["requerido"] = xreq
                return
            seen_xml[xml] = len(out)
            out.append(dict(tipo_doc=code, formato="XML", ruta=xml, tipo="" if is_container else tipo,
                            valores="", requerido=xreq, cambio=UNCHANGED, ruta_anterior="",
                            ruta_contraparte=path, descripcion=desc, hoja_origen=XML_SHEET,
                            fila_origen=n, nota=next((x for x in note if x.startswith("posicional")), "")))

        for n, r in body:
            raw = clean(cell(r, col))
            if not raw:
                continue
            if _ascii(raw).lower() == XML_ONLY:  # tag sin campo JSON (p. ej. CabeceraEntrada)
                xml = xml_path(cell(r, xcol))
                if not xml:
                    lint.append((label, n, "ERROR", f"'{raw}' sin ruta válida en XML_TQ"))
                    continue
                tipo = parse_type(cell(r, c["tipo"]))[0]
                add_xml_row(xml, "", UNCHANGED, parse_required(cell(r, c["req"]), code) or "NO", tipo, False,
                            clean(cell(r, c.get("desc"))), [], n)
                continue
            only_null = re.match(r"solo null:\s*(.+)$", raw, re.I)  # el campo debe venir y su valor siempre es null
            if only_null:
                raw = only_null.group(1)
            change, path, old = parse_change(raw)
            if not path:
                continue
            if (code, path) in removed_in_tab:  # 'Eliminar' en TAG_MIN_HACIENDA: no va en el JSON; su tag XML es opcional
                change, old = REMOVED, ""
            if not _PATH.fullmatch(path):
                lint.append((label, n, "WARNING", f"Texto no interpretable como ruta, se omite: '{path[:60]}'"))
                continue
            if change == RENAMED and not old:
                lint.append((label, n, "ERROR", f"RENOMBRADO sin ruta anterior: '{raw}'"))
                continue
            attr = (by_generic.get(path) or by_generic.get(old) or r) if code == RETORNO else r
            tipo, values, tw = parse_type(cell(attr, c["tipo"]))
            if tw:
                lint.append((label, n, "INFO", tw))
            raw_req = clean(cell(attr, c["req"]))
            required = "SI_NULL" if only_null else parse_required(raw_req, code)
            is_container = tipo in ("object", "array") or _ascii(cell(attr, c["tipo"])).lower() in ("object", "array")
            if _ascii(raw_req).lower() and _ascii(raw_req).lower() not in KNOWN_REQ:
                lint.append((label, n, "WARNING",
                             f"Obligatoriedad ambigua '{raw_req}' en '{path}': se toma la lectura mas permisiva (SI_NULL)"))
            if change == REMOVED:
                required = "NO"
            elif not required:
                required = "NO"
                if not is_container:
                    lint.append((label, n, "INFO", f"Sin dato de obligatoriedad para '{path}': se asume NO"))
            if path in seen_json:
                if re.search(r"(/items|\[\])$", raw):  # 'X' y 'X/items' (o 'X[]') son el mismo nivel aplanado
                    prev = out[json_idx[path]]
                    xraw = clean(cell(r, xcol))
                    xml = xml_path(xraw)
                    if xml and not prev["ruta_contraparte"]:
                        # el tag XML está en la fila 'X/items': esa fila define el campo (antes el tag se perdía)
                        lint.append((label, prev["fila_origen"], "INFO",
                                     f"Nivel de array '{path}': lo define la fila {n} ('{raw}'), que trae el tag XML"))
                        note = [f"posicional({m.group(1)})"] if (m := _POSITIONAL.search(xraw)) else []
                        prev.update(ruta_contraparte=xml, requerido=required, cambio=change, ruta_anterior=old,
                                    valores="|".join(values), fila_origen=n, nota="; ".join(note))
                        seen_json[path] = n
                        add_xml_row(xml, path, change, required, tipo, is_container, prev["descripcion"], note, n)
                    else:
                        lint.append((label, n, "INFO", f"Nivel de array '{raw}' = '{path}' (fila {seen_json[path]}): mismo nivel aplanado"))
                else:
                    lint.append((label, n, "WARNING", f"Ruta duplicada '{path}' (primera en fila {seen_json[path]}): se omite"))
                continue
            seen_json[path] = n
            json_idx[path] = len(out)
            xraw = clean(cell(r, xcol))
            xml = xml_path(xraw)
            literal = re.sub(r"\(.*$", "", re.sub(r"\([^()]*\)(?=/)", "", xraw)).strip()
            if xml and clean_tag(literal, strip_items=False) != xml:
                lint.append((label, n, "WARNING", f"Ruta XML no unificada en el libro: '{literal}' se interpreta como '{xml}'"))
            desc = clean(cell(attr, c.get("desc")))
            note = [ONLY_NULL] if only_null else []
            if xraw not in ("", "x") and not xml:
                note.append(f"XML_TQ sin ruta: {xraw[:80]}")
            im = _POSITIONAL.search(xraw) if xml else None
            if im:  # el tag XML se repite y el libro indica la posición: no comparable 1 a 1 en el cruce XML<->JSON
                note.append(f"posicional({im.group(1)})")
            out.append(dict(tipo_doc=code, formato="JSON", ruta=path, tipo=tipo, valores="|".join(values),
                            requerido=required, cambio=change, ruta_anterior=old, ruta_contraparte=xml,
                            descripcion=desc, hoja_origen=XML_SHEET, fila_origen=n, nota="; ".join(note)))
            if xml:
                add_xml_row(xml, path, change, required, tipo, is_container, desc, note, n)
    return out


def _check_json_sheet(wb, rows: list[dict], lint: list):
    """La hoja 'JSON' lista, por tipo, las rutas v3 (antes de los cambios). Debe coincidir con la vista v3 de la
    columna del tipo en la hoja 'XML' (SIN_CAMBIO + ELIMINADO + nombre anterior de los RENOMBRADO)."""
    if JSON_SHEET not in wb.sheetnames:
        lint.append((JSON_SHEET, "", "WARNING", "No existe la hoja 'JSON': no se verifica la base v3"))
        return
    jrows = list(wb[JSON_SHEET].iter_rows(values_only=True, max_col=MAX_COL))
    hi = next((i for i, r in enumerate(jrows[:10]) if type_columns(r)[0]), None)
    if hi is None:
        lint.append((JSON_SHEET, "", "WARNING", "Hoja 'JSON' sin encabezados de tipo de documento"))
        return
    for col, code in type_columns(jrows[hi])[0].items():
        sheet = {}
        for n, r in enumerate(jrows[hi + 1:], start=hi + 2):
            p = clean_tag(r[col]) if col < len(r) and clean(r[col]) else ""
            if p:
                sheet.setdefault(p, n)
        v3 = {}
        for x in rows:
            if x["tipo_doc"] == code and x["formato"] == "JSON" and x["cambio"] != NEW:
                v3.setdefault(x["ruta_anterior"] or x["ruta"], x["fila_origen"])
        for p in sorted(set(sheet) - set(v3)):
            lint.append((JSON_SHEET, sheet[p], "WARNING",
                         f"Hoja JSON [{code}]: '{p}' no aparece en la columna del tipo en la hoja 'XML'"))
        for p in sorted(set(v3) - set(sheet)):
            lint.append((sheet_label(code), v3[p], "WARNING", f"Hoja JSON [{code}]: '{p}' (v3) no aparece en la hoja 'JSON'"))


def build_rows(source: str):
    wb = load_workbook(source, read_only=True, data_only=True)
    if XML_SHEET not in wb.sheetnames:
        raise ValueError(f"El libro no tiene la hoja '{XML_SHEET}' (fuente de los tags)")
    lint: list[tuple] = []
    out = _xml_sheet_rows(wb, lint)
    req = _xml_requirements(wb, lint)
    for ws in wb.worksheets:
        code = SHEET_CODES.get(_ascii(ws.title).lower())
        if code in TAB_TYPES:
            out += _tab_rows(ws, code, req, lint)
    _check_json_sheet(wb, out, lint)
    _fill_descriptions(wb, out)
    return out, lint


def _fill_descriptions(wb, rows: list[dict]):
    """La hoja 'XML' deja vacía la Descripcion de varias filas; el validador de PDF la usa como etiqueta impresa.
    Solo para esas filas se toma la Descripcion de la hoja por tipo (misma ruta JSON). No cambia tags ni reglas."""
    tab: dict = {}
    for ws in wb.worksheets:
        code = SHEET_CODES.get(_ascii(ws.title).lower())
        if not code or code in TAB_TYPES:
            continue
        trs = list(ws.iter_rows(values_only=True, max_col=MAX_COL))
        try:
            c = _cols(trs[_find_header(trs)])
        except ValueError:
            continue
        for r in trs[_find_header(trs) + 1:]:
            get = lambda k: r[c[k]] if c.get(k) is not None and c[k] < len(r) else None
            path = parse_change(get("tag_v4"))[1] if clean(get("tag_v4")) else ""
            desc = clean(get("desc"))
            if path and desc:
                tab.setdefault((code, path), desc)
    by_json = {}
    for r in rows:
        if r["formato"] == "JSON" and r["hoja_origen"] == XML_SHEET:
            if not r["descripcion"]:
                r["descripcion"] = tab.get((r["tipo_doc"], r["ruta"])) or tab.get((r["tipo_doc"], r["ruta_anterior"]), "")
            by_json[(r["tipo_doc"], r["ruta"])] = r["descripcion"]
    for r in rows:
        if r["formato"] == "XML" and r["hoja_origen"] == XML_SHEET and not r["descripcion"]:
            r["descripcion"] = by_json.get((r["tipo_doc"], r["ruta_contraparte"]), "")


DEFAULT_AJUSTES = "data/specs/ajustes_contrato.json"
AJUSTES_SOURCE = "ajustes_contrato.json"  # hoja_origen de las filas agregadas por ajuste
AJUSTE_COLUMNS = ["id", "accion", "tipo_doc", "formato", "ruta", "campo", "valor_libro", "valor_ajustado", "motivo",
                  "evidencia", "aprobado"]


class AjusteError(ValueError):
    pass


KNOWN_RULES = {
    # Campos con catálogo de valores en el libro: se validan solo como texto (tipo string, sin lista de valores).
    "catalogos_como_string",
    # Regla de validación (no cambia filas): apéndices obligatorios por nombre y adicionales como aviso.
    "apendice_obligatorio",
    # Erratas de ruta del libro: {ruta JSON del libro: ruta correcta}; corrige la fila JSON y la contraparte XML.
    "correccion_rutas",
    # Regla de validación (no cambia filas): en el cruce XML<->JSON, '285.6' y '285.60' son el mismo valor.
    "decimales_equivalentes",
}


def _read_config(path: str | None) -> dict:
    import json
    import os
    if not path or not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8-sig") as fh:
        return json.load(fh)


def load_ajustes(path: str | None) -> list[dict]:
    """Lee los ajustes por fila de data/specs/ajustes_contrato.json. Sin archivo: sin ajustes."""
    items = _read_config(path).get("ajustes", [])
    ids = [a.get("id") for a in items]
    if len(ids) != len(set(ids)) or any(not i for i in ids):
        raise AjusteError("Cada ajuste necesita un 'id' único")
    for a in items:
        if a.get("accion") not in ("modificar", "agregar"):
            raise AjusteError(f"{a.get('id')}: 'accion' debe ser 'modificar' o 'agregar'")
        for k in ("tipo_doc", "formato", "ruta", "motivo"):
            if not a.get(k):
                raise AjusteError(f"{a.get('id')}: falta '{k}'")
        if not isinstance(a["tipo_doc"], (str, list)):
            raise AjusteError(f"{a['id']}: 'tipo_doc' debe ser un tipo, una lista de tipos o '*'")
        bad = [k for k in a.get("campos", {}) if k not in COLUMNS or k in ("tipo_doc", "formato", "ruta", "hoja_origen", "fila_origen")]
        if bad:
            raise AjusteError(f"{a['id']}: campos no ajustables: {', '.join(bad)}")
    return items


def load_reglas(path: str | None) -> dict:
    """Reglas generales de generación activas: {nombre: {activo, motivo, aprobado, ...}}."""
    reglas = _read_config(path).get("reglas_generales", {})
    unknown = sorted(set(reglas) - KNOWN_RULES)
    if unknown:
        raise AjusteError(f"Regla(s) general(es) desconocida(s): {', '.join(unknown)}")
    return {k: v for k, v in reglas.items() if isinstance(v, dict) and v.get("activo")}


def apply_reglas(rows: list[dict], reglas: dict) -> tuple[list[dict], list[list]]:
    """Aplica las reglas generales a las filas del libro. Devuelve (filas, detalle para la hoja 'ajustes')."""
    rows = [dict(r) for r in rows]
    detail = []
    regla = reglas.get("correccion_rutas")
    if regla:
        meta = [regla.get("motivo", ""), regla.get("evidencia", ""), regla.get("aprobado", "")]
        rutas = regla.get("rutas", {})
        for r in rows:
            campo = "ruta" if r["formato"] == "JSON" else "ruta_contraparte"
            if r.get(campo) in rutas:
                detail.append(["correccion_rutas", "regla", r["tipo_doc"], r["formato"], r["ruta"], campo, r[campo],
                               rutas[r[campo]], *meta])
                r[campo] = rutas[r[campo]]
                r["nota"] = "; ".join(x for x in (r.get("nota") or "", "regla correccion_rutas") if x)
    regla = reglas.get("catalogos_como_string")
    if regla:
        meta = [regla.get("motivo", ""), regla.get("evidencia", ""), regla.get("aprobado", "")]
        catalogos = {(r["tipo_doc"], r["ruta"]) for r in rows if r.get("valores") and r["formato"] == "JSON"}
        for r in rows:  # el tag XML de un campo con catálogo tampoco tiene otro tipo: también es string
            if r["formato"] == "XML" and not r.get("tipo") and (r["tipo_doc"], r.get("ruta_contraparte")) in catalogos:
                detail.append(["catalogos_como_string", "regla", r["tipo_doc"], "XML", r["ruta"], "tipo", "", "string", *meta])
                r["tipo"] = "string"
                r["nota"] = "; ".join(x for x in (r.get("nota") or "", "regla catalogos_como_string") if x)
        for r in rows:
            if r.get("valores"):
                key = (r["tipo_doc"], r["formato"], r["ruta"])
                detail.append(["catalogos_como_string", "regla", *key, "valores", r["valores"], "", *meta])
                if r.get("tipo") != "string":
                    detail.append(["catalogos_como_string", "regla", *key, "tipo", r.get("tipo") or "", "string", *meta])
                r["valores"], r["tipo"] = "", "string"
                r["nota"] = "; ".join(x for x in (r.get("nota") or "", "regla catalogos_como_string") if x)
    regla = reglas.get("decimales_equivalentes")
    if regla:  # regla de validación: se documenta en la hoja 'ajustes'; no modifica filas
        detail.append(["decimales_equivalentes", "regla", "*", "XML<->JSON", "(valores numéricos)", "comparación",
                       "texto exacto", "mismo valor decimal = igual", regla.get("motivo", ""), regla.get("evidencia", ""),
                       regla.get("aprobado", "")])
    regla = reglas.get("apendice_obligatorio")
    if regla:  # regla de validación: se documenta en la hoja 'ajustes'; no modifica filas
        detail.append(["apendice_obligatorio", "regla", ", ".join(regla.get("tipos", [])), "JSON/XML", "apendice / DatosEntrada.Notas",
                       "obligatorios", "", ", ".join(regla.get("obligatorios", [])),
                       regla.get("motivo", ""), regla.get("evidencia", ""), regla.get("aprobado", "")])
    return rows, detail


ALL_TYPES = "*"


def expand_ajustes(ajustes: list[dict], rows: list[dict]) -> list[dict]:
    """tipo_doc: un tipo, una lista de tipos, o '*' = todos los tipos con filas en ese formato."""
    out = []
    for a in ajustes:
        t = a["tipo_doc"]
        if t == ALL_TYPES:
            tipos = sorted({r["tipo_doc"] for r in rows if r["formato"] == a["formato"]})
        elif isinstance(t, list):
            tipos = t
        else:
            out.append(a)
            continue
        out += [{**a, "tipo_doc": x} for x in tipos]
    return out


def apply_ajustes(rows: list[dict], ajustes: list[dict]) -> tuple[list[dict], list[list]]:
    """Aplica los ajustes sobre las filas generadas desde el libro. Devuelve (filas, detalle para la hoja 'ajustes')."""
    ajustes = expand_ajustes(ajustes, rows)
    rows = [dict(r) for r in rows]
    index = {(r["tipo_doc"], r["formato"], r["ruta"]): r for r in rows}
    detail = []
    for a in ajustes:
        key = (a["tipo_doc"], a["formato"], a["ruta"])
        campos = a.get("campos", {})
        meta = [a.get("motivo", ""), a.get("evidencia", ""), a.get("aprobado", "")]
        if a["accion"] == "modificar":
            r = index.get(key)
            if r is None:
                raise AjusteError(f"{a['id']}: no existe la fila {key} para modificar")
            for k, v in campos.items():
                detail.append([a["id"], "modificar", *key, k, r.get(k), v, *meta])
                r[k] = v
            r["nota"] = "; ".join(x for x in (r.get("nota") or "", f"ajuste {a['id']}") if x)
        else:
            if key in index:
                raise AjusteError(f"{a['id']}: la fila {key} ya existe; use 'modificar'")
            r = {c: "" for c in COLUMNS}
            r.update(tipo_doc=key[0], formato=key[1], ruta=key[2], requerido="NO", cambio=UNCHANGED,
                     hoja_origen=AJUSTES_SOURCE, fila_origen=a["id"])
            r.update(campos)
            r["nota"] = "; ".join(x for x in (r.get("nota") or "", f"ajuste {a['id']}") if x)
            rows.append(r)
            index[key] = r
            detail.append([a["id"], "agregar", *key, "(fila nueva)", "", ", ".join(f"{k}={v}" for k, v in campos.items()), *meta])
    return rows, detail


def write_contract(rows, lint, out_path: str, ajustes_detail=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "contrato"
    ws.append(COLUMNS)
    for r in rows:
        ws.append([r[k] for k in COLUMNS])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col, w in zip("ABCDEFGHIJKLM", (12, 9, 50, 10, 16, 11, 13, 40, 55, 40, 24, 10, 40)):
        ws.column_dimensions[col].width = w
    lw = wb.create_sheet("lint")
    lw.append(["hoja", "fila", "nivel", "mensaje"])
    for item in lint:
        lw.append(list(item))
    for col, w in zip("ABCD", (28, 8, 10, 100)):
        lw.column_dimensions[col].width = w
    aw = wb.create_sheet("ajustes")
    aw.append(AJUSTE_COLUMNS)
    for d in ajustes_detail or []:
        aw.append(["" if v is None else v for v in d])
    for col, w in zip("ABCDEFGHIJK", (6, 10, 9, 8, 55, 14, 14, 30, 70, 45, 11)):
        aw.column_dimensions[col].width = w
    wb.save(out_path)


def build_with_ajustes(source: str, ajustes_path: str | None = None):
    """Filas del libro + reglas generales + ajustes aprobados. Devuelve (filas, lint, detalle_ajustes)."""
    rows, lint = build_rows(source)
    rows, d_reglas = apply_reglas(rows, load_reglas(ajustes_path))
    rows, d_ajustes = apply_ajustes(rows, load_ajustes(ajustes_path))
    return rows, lint, d_reglas + d_ajustes


def build(source: str, out_path: str, ajustes_path: str | None = None) -> dict:
    """ajustes_path: archivo de ajustes a aplicar (la CLI pasa data/specs/ajustes_contrato.json); None = solo el libro."""
    rows, lint, detail = build_with_ajustes(source, ajustes_path)
    write_contract(rows, lint, out_path, detail)
    summary: dict = {}
    for r in rows:
        k = (r["tipo_doc"], r["formato"])
        summary[k] = summary.get(k, 0) + 1
    return {"rows": len(rows), "lint": len(lint), "by_type": summary, "ajustes": sorted({d[0] for d in detail}),
            "lint_levels": {lv: sum(1 for x in lint if x[2] == lv) for lv in ("ERROR", "WARNING", "INFO")}}
