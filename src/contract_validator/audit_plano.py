"""Auditoría de fidelidad: libro original (contrato_ejemplo.xlsx) -> contrato plano (contrato_plano.xlsx).

Lector INDEPENDIENTE: interpreta el original con su propio código (no usa parse_change/clean_tag/xml_path de
build_contract) y lo compara contra lo que quedó en la hoja 'contrato' del plano.

Fuente: hoja 'XML' (una columna por tipo de documento; RETORNO con su tag en 'XML-ER') y, para INVALIDACION, su
hoja propia. Una fila de la hoja 'XML' alimenta a varios tipos: la trazabilidad es (hoja, tipo_doc, fila).

Códigos:
  ROW_LOST              fila del original con campo v4 que no está en el plano ni en 'lint' (ERROR); si el texto no
                        parece una ruta y no está en lint, WARNING (ROW_SKIPPED_SILENT)
  ROW_UNTRACEABLE       fila del plano que no apunta a una fila/celda coincidente del original (ERROR)
  FIELD_DIFF            campo del plano distinto del recalculado desde el original (ERROR)
  REQ_AMBIGUOUS         texto de 'Requerido'/'Obligatorio' fuera del conjunto conocido (WARNING)
  COUNTERPART_ASYMMETRIC  la contraparte JSON<->XML no es recíproca (ERROR)
  TYPE_INCOMPLETE       un tipo de documento sin filas en alguno de los formatos (ERROR)
  PLANO_STALE           regenerar el plano desde el original da un resultado distinto del archivo vigente (ERROR)
  LINT_*                avisos de 'lint' clasificados: decisiones pendientes del contrato (INFO/WARNING)
"""
from __future__ import annotations
import os
import re
import unicodedata
from collections import defaultdict
from openpyxl import Workbook, load_workbook

HEADERS = ["hoja", "nivel", "codigo", "tipo_doc", "ruta", "detalle", "fila_origen"]
MAX_COL = 30
TAB_SHEETS = {"invalidacion": "INVALIDACION"}  # tipos sin tags XML en la hoja 'XML': se leen de su hoja
TYPE_SHEETS = {"comprobante de credito fiscal": "CCF", "factura electronica": "FAC", "factura de exportacion": "FAC_EXP",
               "nota de credito": "NCR", "nota de debito": "NDB", "nota de remision": "REM"}
XML_CODES = {"03": "CCF", "01": "FAC", "11": "FAC_EXP", "05": "NCR", "06": "NDB", "04": "REM"}
_POS = re.compile(r"[(\[]\s*\d+\s*[)\]]")
KNOWN_REQ = {"si", "si (null)", "no", "n/a", "si para fex", "requerido", "si/null/ por tipodte"}
_PATH_RE = re.compile(r"[\w@][\w@/.\-]*")


# ---------- utilidades de texto (propias) ----------
def _txt(v) -> str:
    return re.sub(r"\s+", " ", str("" if v is None else v).replace("\xa0", " ")).strip().rstrip("*").strip()


def _asc(v) -> str:
    s = unicodedata.normalize("NFKD", str("" if v is None else v)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip().lower()


def _route(t: str, keep_items=False) -> str:
    t = re.sub(r"\[\d*\]|\(.*?\)", "", t)
    toks = t.split()
    if len(toks) > 1 and toks[1].lower().startswith("normativa"):
        t = toks[0]
    return "/".join(p.strip() for p in t.split("/") if p.strip() and (keep_items or p.strip() != "items"))


def _change(text: str):
    """-> (cambio, ruta, ruta_anterior)."""
    t = _txt(text)
    m = re.match(r"^renombrar\s+(.+?)\s*(?:->|→)\s*(.+)$", t, re.I)
    if m:
        return "RENOMBRADO", _route(_txt(m.group(2))), _route(_txt(m.group(1)))
    m = re.match(r"^adicionar\s+(.+)$", t, re.I)
    if m:
        return "NUEVO", _route(_txt(m.group(1))), ""
    m = re.match(r"^elim\w*\s+(.+)$", t, re.I)
    if m:
        return "ELIMINADO", _route(_txt(m.group(1))), ""
    return "SIN_CAMBIO", _route(t), ""


def _xml(raw) -> str:
    t = re.sub(r"\(.*$", "", re.sub(r"\([^()]*\)(?=/)", "", _txt(raw))).strip()
    if re.match(r"[A-Z]\w+/\w", t) and not t.startswith(("DatosEntrada/", "CabeceraEntrada/")):
        t = "DatosEntrada/" + t
    if not t.startswith(("DatosEntrada/", "CabeceraEntrada/")):
        return ""
    t = _route(t, keep_items=True)
    t = t.replace("/InformacionGeneral/", "/InformacioGeneral/").replace("/Adquiriente/", "/InformacionAdquiriente/")
    m = re.match(r"^DatosEntrada/(InformacionEmisor|InformacionAdquiriente)/(.*)$", t)
    return f"DatosEntrada/InformacioGeneral/{m.group(1)}/{m.group(2)}" if m else t


def _values(raw) -> str:
    t = _txt(raw)
    if re.fullmatch(r"[\w.]+(\s*,\s*[\w.]+)+", t):
        return "|".join(v.strip() for v in t.split(","))
    if t == "USD":
        return "USD"
    return t if re.fullmatch(r"\d{2}", t) else ""


def _req(raw, code: str):
    s = _asc(raw)
    if not s or s == "requerido":
        return None
    if s.startswith("si para fex"):
        return "SI" if code == "FAC_EXP" else "NO"
    if s.startswith("si") and "null" in s:
        return "SI_NULL"
    if s == "si":
        return "SI"
    if s in ("no", "n/a"):
        return "NO"
    return None


# ---------- lectura del original ----------
def _header(rows):
    for i, r in enumerate(rows[:10]):
        if any(_asc(c).startswith("tag_min_hacienda") for c in r):
            return i
    raise ValueError("sin encabezado TAG_MIN_HACIENDA")


def _columns(head) -> dict:
    c, tags = {}, []
    for i, h in enumerate(head):
        k = _asc(h)
        if k.startswith("tag_min_hacienda"):
            tags.append(i)
        elif k == "xml_tq":
            c.setdefault("xml", i)
        elif k == "tipo" and (str(h).strip() == "Tipo" or "tipo" not in c):
            c["tipo"] = i
        elif k == "accion":
            c.setdefault("accion", i)
        elif k == "obligatorio":
            c.setdefault("oblig", i)
    c["v4"] = tags[1] if len(tags) > 1 else tags[0]
    return c


def _type_cols(head) -> tuple[dict, int | None]:
    cols, er = {}, None
    for i, x in enumerate(head):
        k = _asc(x).upper()
        m = re.search(r"-\s*(\d{2})$", k)
        code = XML_CODES.get(m.group(1)) if m else None
        if "INVALIDACION" in k:
            code = "INVALIDACION"
        elif k == "EVENTO RETORNO":
            code = "RETORNO"
        elif k == "XML-ER":
            er = i
        if code:
            cols[i] = code
    return cols, er


def read_original(source: str) -> dict:
    """-> {'sheets': {(hoja, tipo): {'code', 'rows': {fila: dict}}}, 'req': {(code, ruta): texto}}"""
    wb = load_workbook(source, read_only=True, data_only=True)
    out = {"sheets": {}, "req": {}, "tab_removed": set()}
    for ws in wb.worksheets:
        tipo = TYPE_SHEETS.get(_asc(ws.title))
        if tipo:
            rows = list(ws.iter_rows(values_only=True, max_col=MAX_COL))
            h = _header(rows)
            v4 = _columns(rows[h])["v4"]
            for r in rows[h + 1:]:
                if v4 < len(r) and _txt(r[v4]) and _change(r[v4])[0] == "ELIMINADO":
                    out["tab_removed"].add((tipo, _change(r[v4])[1]))
        code = TAB_SHEETS.get(_asc(ws.title))
        if not code:
            continue
        rows = list(ws.iter_rows(values_only=True, max_col=MAX_COL))
        h = _header(rows)
        c = _columns(rows[h])
        cell = lambda r, k: r[c[k]] if k in c and c[k] < len(r) else None
        data = {}
        for n, r in enumerate(rows[h + 1:], start=h + 2):
            v4 = _txt(cell(r, "v4"))
            if not v4:
                continue
            change, path, old = _change(v4)
            data[n] = dict(v4=v4, change=change, path=path, old=old, xml_raw=_txt(cell(r, "xml")), xml=_xml(cell(r, "xml")),
                           values=_values(cell(r, "tipo")), accion=_asc(cell(r, "accion")).upper(),
                           oblig=_txt(cell(r, "oblig")), req_raw=None)
        out["sheets"][(ws.title.strip(), code)] = {"code": code, "rows": data}
    if "XML" not in wb.sheetnames:
        return out
    rows = list(wb["XML"].iter_rows(values_only=True, max_col=MAX_COL))
    h = _header(rows)
    head = rows[h]
    col = {}
    for i, x in enumerate(head):
        k = {"tag_min_hacienda": "tag", "tipo": "tipo", "xml_tq": "xml", "requerido": "req"}.get(_asc(x))
        if k:
            col.setdefault(k, i)
    cols, er = _type_cols(head)
    get = lambda r, i: r[i] if i is not None and i < len(r) else None
    body = list(enumerate(rows[h + 1:], start=h + 2))
    generic = {}
    for n, r in body:
        if _txt(get(r, col["tag"])):
            generic.setdefault(_change(get(r, col["tag"]))[1], r)
    for i, code in cols.items():
        data = {}
        for n, r in body:
            v = _txt(get(r, i))
            if not v:
                continue
            if _asc(v) == "solo xml" and code != "INVALIDACION":  # tag sin campo JSON
                xraw = _txt(get(r, er if code == "RETORNO" else col["xml"]))
                data[n] = dict(v4=v, change="SIN_CAMBIO", path="", old="", xml_raw=xraw, xml=_xml(xraw), values="",
                               accion="", oblig="", req_raw=_txt(get(r, col["req"])), xml_only=True)
                continue
            only_null = bool(re.match(r"solo null:\s*", v, re.I))
            if only_null:
                v = re.sub(r"(?i)^solo null:\s*", "", v)
            change, path, old = _change(v)
            if (code, path) in out["tab_removed"]:  # 'Eliminar' en la hoja del tipo
                change, old = "ELIMINADO", ""
            if code == "INVALIDACION":  # solo obligatoriedad (los tags salen de su hoja)
                if path and change != "ELIMINADO":
                    out["req"][(code, path)] = _txt(get(r, col["req"]))
                continue
            attr = (generic.get(path) or generic.get(old) or r) if code == "RETORNO" else r
            xraw = _txt(get(r, er if code == "RETORNO" else col["xml"]))
            data[n] = dict(v4=v, change=change, path=path, old=old, xml_raw=xraw, xml=_xml(xraw),
                           values=_values(get(attr, col["tipo"])), accion="", oblig="", req_raw=_txt(get(attr, col["req"])), only_null=only_null)
        if code != "INVALIDACION":
            out["sheets"][("XML", code)] = {"code": code, "rows": data}
    return out


def read_plano(path: str):
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["contrato"].iter_rows(values_only=True))
    head = list(rows[0])
    plano = [dict(zip(head, r)) for r in rows[1:] if any(v not in (None, "") for v in r)]
    lint = []
    if "lint" in wb.sheetnames:
        lr = list(wb["lint"].iter_rows(values_only=True))
        lint = [dict(zip(lr[0], r)) for r in lr[1:] if any(v not in (None, "") for v in r)]
    return plano, lint


# ---------- auditoría ----------
def audit(source: str, plano_path: str, ajustes_path: str | None = None) -> tuple[list[dict], dict]:
    from .build_contract import AJUSTES_SOURCE, expand_ajustes, load_ajustes, load_reglas
    orig = read_original(source)
    plano, lint = read_plano(plano_path)
    out: list[dict] = []

    def add(hoja, nivel, codigo, tipo="", ruta="", detalle="", fila=""):
        out.append(dict(hoja=hoja, nivel=nivel, codigo=codigo, tipo_doc=tipo, ruta=ruta, detalle=detalle, fila_origen=fila))

    # 0. ajustes aprobados por el usuario: se informan y sus campos no cuentan como diferencia del libro
    reglas = load_reglas(ajustes_path)
    for nombre, regla in reglas.items():
        add(AJUSTES_SOURCE, "INFO", "REGLA_GENERAL", "", "", f"{nombre}: {regla.get('motivo', '')}", nombre)
    catalogos_string = "catalogos_como_string" in reglas
    ajustes_raw = load_ajustes(ajustes_path)  # la CLI pasa data/specs/ajustes_contrato.json; None = solo el libro
    ajustes = expand_ajustes(ajustes_raw, [p for p in plano if str(p["hoja_origen"]).strip() != AJUSTES_SOURCE])
    adjusted = defaultdict(set)  # (tipo_doc, formato, ruta) -> campos ajustados
    for a_ in ajustes:
        adjusted[(a_["tipo_doc"], a_["formato"], a_["ruta"])] |= set(a_.get("campos", {}))
        add(AJUSTES_SOURCE, "INFO", "AJUSTE_APLICADO", a_["tipo_doc"], a_["ruta"],
            f"{a_['accion']} {a_.get('campos', {})}: {a_.get('motivo', '')}", a_["id"])
    plano_libro = [p for p in plano if str(p["hoja_origen"]).strip() != AJUSTES_SOURCE]
    agregadas = {(a_["tipo_doc"], a_["formato"], a_["ruta"]) for a_ in ajustes if a_["accion"] == "agregar"}
    for p in plano:
        if str(p["hoja_origen"]).strip() == AJUSTES_SOURCE and (p["tipo_doc"], p["formato"], p["ruta"]) not in agregadas:
            add(AJUSTES_SOURCE, "ERROR", "ROW_UNTRACEABLE", p["tipo_doc"], p["ruta"],
                "Fila marcada como ajuste que no está en ajustes_contrato.json", p["fila_origen"])

    sheet_code = {h: s["code"] for (h, _c), s in orig["sheets"].items() if h != "XML"}

    def lint_key(hoja: str):
        """'XML [CCF]' -> ('XML', 'CCF'); 'Invalidacion' -> ('Invalidacion', 'INVALIDACION')."""
        m = re.fullmatch(r"(.+) \[(\w+)\]", hoja)
        return (m.group(1), m.group(2)) if m else (hoja, sheet_code.get(hoja, ""))

    by_cell = defaultdict(list)  # (hoja, tipo, fila) -> filas del plano
    for p in plano_libro:
        by_cell[(str(p["hoja_origen"]).strip(), p["tipo_doc"], p["fila_origen"])].append(p)
    lint_rows = {(*lint_key(str(x["hoja"]).strip()), x["fila"]) for x in lint if x.get("fila") not in (None, "")}
    rutas_corr = (reglas.get("correccion_rutas") or {}).get("rutas", {})
    stats = defaultdict(lambda: dict(filas_origen=0, filas_plano=0, omitidas_lint=0, omitidas_sin_lint=0, diferencias=0))

    # 1. original -> plano
    label = lambda hoja, code: f"{hoja} [{code}]" if hoja == "XML" else hoja
    for (hoja, code), s in orig["sheets"].items():
        st = stats[label(hoja, code)]
        for n, o in s["rows"].items():
            st["filas_origen"] += 1
            formato = "XML" if o.get("xml_only") else "JSON"
            js = [p for p in by_cell.get((hoja, code, n), []) if p["formato"] == formato]
            if o.get("xml_only"):  # el tag puede estar ya definido por otra fila del mismo tipo
                js = js or [p for p in plano_libro if p["tipo_doc"] == code and p["formato"] == "XML" and p["ruta"] == o["xml"]]
            if js:
                continue
            if (hoja, code, n) in lint_rows:
                st["omitidas_lint"] += 1
                continue
            st["omitidas_sin_lint"] += 1
            if bool(_PATH_RE.fullmatch(o["path"])):
                add(label(hoja, code), "ERROR", "ROW_LOST", code, o["path"], "Fila con campo v4 ausente del plano y de 'lint'", n)
            else:
                add(label(hoja, code), "WARNING", "ROW_SKIPPED_SILENT", code, o["path"][:60],
                    "Fila no interpretable como ruta, omitida sin aviso en 'lint'", n)

    # 2 y 3. plano -> original, campo a campo
    for p in plano_libro:
        hoja, n, code = str(p["hoja_origen"]).strip(), p["fila_origen"], p["tipo_doc"]
        stats[label(hoja, code)]["filas_plano"] += 1
        o = orig["sheets"].get((hoja, code), {}).get("rows", {}).get(n)
        if o is None:
            add(label(hoja, code), "ERROR", "ROW_UNTRACEABLE", code, p["ruta"],
                "La fila de origen no existe, no tiene campo v4 o no es de este tipo de documento", n)
            continue
        ajustados = adjusted.get((code, p["formato"], p["ruta"]), set())

        def diff(campo, esperado, obtenido):
            if campo in ajustados:
                return  # el valor viene de un ajuste aprobado (reportado como AJUSTE_APLICADO)
            if str(esperado or "") != str(obtenido or ""):
                stats[label(hoja, code)]["diferencias"] += 1
                add(label(hoja, code), "ERROR", "FIELD_DIFF", code, p["ruta"], f"{campo}: esperado '{esperado}', plano '{obtenido}'", n)

        if p["formato"] == "JSON":
            diff("ruta", rutas_corr.get(o["path"], o["path"]), p["ruta"])  # regla correccion_rutas
            diff("cambio", o["change"], p["cambio"])
            diff("ruta_anterior", o["old"], p["ruta_anterior"])
            diff("valores", "" if catalogos_string else o["values"], p["valores"])  # regla: catálogos solo como texto
            diff("ruta_contraparte", o["xml"], p["ruta_contraparte"])
            posicional = bool(o["xml"]) and bool(_POS.search(o["xml_raw"]))
            if posicional != ("posicional" in str(p["nota"] or "")):
                diff("nota(posicional)", posicional, "posicional" in str(p["nota"] or ""))
            if o["req_raw"] is not None:  # hoja 'XML': obligatoriedad de la propia fila (retorno: fila genérica)
                raw = o["req_raw"]
            else:
                raw = orig["req"].get((code, o["path"])) or (orig["req"].get((code, o["old"])) if o["old"] else "") or ""
            req = _req(raw, code)
            if not req and code == "INVALIDACION":
                raw = o["oblig"]
                req = _req(raw, code)
            if _asc(raw) and _asc(raw) not in KNOWN_REQ:
                add(label(hoja, code), "WARNING", "REQ_AMBIGUOUS", code, o["path"], f"Requerido/Obligatorio '{raw}' fuera del conjunto conocido", n)
            exp = "NO" if o["change"] == "ELIMINADO" else ("SI_NULL" if o.get("only_null") else (req or "NO"))
            diff("nota(solo_null)", bool(o.get("only_null")), "solo_null" in str(p["nota"] or ""))
            diff("requerido", exp, p["requerido"])
        else:
            diff("ruta", o["xml"], p["ruta"])
            exp_cambio = "NUEVO" if (o["accion"] == "CREAR" and o["change"] != "ELIMINADO") else "SIN_CAMBIO"
            diff("cambio", exp_cambio, p["cambio"])

    # 5. simetría de contrapartes y tipos completos
    per_type = defaultdict(lambda: {"JSON": {}, "XML": {}})
    for p in plano:
        per_type[p["tipo_doc"]][p["formato"]][p["ruta"]] = p
    for code in sorted({c for (_h, c) in orig["sheets"]}):
        for fmt in ("JSON", "XML"):
            if not per_type[code][fmt]:
                add("", "ERROR", "TYPE_INCOMPLETE", code, "", f"Sin filas {fmt} para el tipo de documento")
    for code, d in per_type.items():
        for ruta, p in d["JSON"].items():
            cp = p["ruta_contraparte"]
            if cp:
                x = d["XML"].get(cp)
                if x is None:
                    add(p["hoja_origen"], "ERROR", "COUNTERPART_ASYMMETRIC", code, ruta, f"No hay fila XML '{cp}'", p["fila_origen"])
                elif x["ruta_contraparte"] not in d["JSON"] or d["JSON"][x["ruta_contraparte"]]["ruta_contraparte"] != cp:
                    add(p["hoja_origen"], "ERROR", "COUNTERPART_ASYMMETRIC", code, ruta,
                        f"La fila XML '{cp}' apunta a '{x['ruta_contraparte']}', que no la referencia", p["fila_origen"])
        for ruta, x in d["XML"].items():
            if x["ruta_contraparte"] and x["ruta_contraparte"] not in d["JSON"]:
                add(x["hoja_origen"], "ERROR", "COUNTERPART_ASYMMETRIC", code, ruta,
                    f"La contraparte JSON '{x['ruta_contraparte']}' no existe", x["fila_origen"])

    # 6. lint clasificado
    for x in lint:
        m = str(x["mensaje"])
        if m.startswith("Ruta duplicada"):
            c = "LINT_DUPLICATE"
        elif m.startswith("Texto no interpretable"):
            c = "LINT_NOT_PATH"
        elif m.startswith("Sin dato de obligatoriedad"):
            c = "LINT_NO_REQ"
        elif "ambigua" in m:
            c = "LINT_REQ_AMBIGUOUS"
        else:
            c = "LINT_OTHER"
        add(str(x["hoja"]).strip(), x["nivel"], c, lint_key(str(x["hoja"]).strip())[1], "", m, x["fila"])

    # 7. el plano vigente coincide con el regenerado
    from .build_contract import build_rows, COLUMNS, apply_ajustes, apply_reglas
    regen, _ = build_rows(source)
    regen, _ = apply_reglas(regen, reglas)
    regen, _ = apply_ajustes(regen, ajustes)
    key = lambda r: tuple("" if r[k] is None else str(r[k]) for k in COLUMNS)
    a, b = {key(r) for r in regen}, {key(r) for r in plano}
    if a != b:
        add("", "ERROR", "PLANO_STALE", "", "", f"Regenerado vs vigente: {len(a - b)} fila(s) solo en regenerado, {len(b - a)} solo en vigente")

    order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    out.sort(key=lambda r: (order.get(r["nivel"], 3), r["codigo"], r["hoja"], str(r["fila_origen"])))
    return out, dict(stats)


def summarize(rows: list[dict]) -> dict:
    s: dict = defaultdict(lambda: defaultdict(int))
    for r in rows:
        s[(r["nivel"], r["codigo"])][r["tipo_doc"] or "-"] += 1
    return {k: dict(v) for k, v in s.items()}


def to_text(rows: list[dict], stats: dict) -> str:
    lines = ["Auditoría del contrato plano vs original"]
    lines.append(f"  {'hoja':<34}{'origen':>7}{'plano':>7}{'lint':>6}{'sin aviso':>10}{'difs':>6}")
    for h, s in stats.items():
        lines.append(f"  {h:<34}{s['filas_origen']:>7}{s['filas_plano']:>7}{s['omitidas_lint']:>6}{s['omitidas_sin_lint']:>10}{s['diferencias']:>6}")
    errs = sum(1 for r in rows if r["nivel"] == "ERROR")
    lines.append(f"Hallazgos: {len(rows)}  (ERROR={errs})")
    for (lvl, code), per in sorted(summarize(rows).items(), key=lambda kv: ({"ERROR": 0, "WARNING": 1, "INFO": 2}.get(kv[0][0], 3), kv[0][1])):
        lines.append(f"  {lvl:<8}{code:<24}{sum(per.values()):>5}   " + ", ".join(f"{t}={n}" for t, n in sorted(per.items())))
    return "\n".join(lines)


def to_xlsx(rows: list[dict], stats: dict, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.append(["hoja", "filas_origen", "filas_plano", "omitidas_lint", "omitidas_sin_lint", "diferencias", "veredicto"])
    bad = {r["hoja"] for r in rows if r["nivel"] == "ERROR"}
    for h, s in stats.items():
        ws.append([h, s["filas_origen"], s["filas_plano"], s["omitidas_lint"], s["omitidas_sin_lint"], s["diferencias"],
                   "FAIL" if h in bad else "OK"])
    ws.append([])
    ws.append(["nivel", "codigo", "total"])
    for (lvl, code), per in sorted(summarize(rows).items()):
        ws.append([lvl, code, sum(per.values())])
    hs = wb.create_sheet("Hallazgos")
    hs.append(HEADERS)
    for r in rows:
        hs.append([r[k] for k in HEADERS])
    hs.freeze_panes = "A2"
    hs.auto_filter.ref = hs.dimensions
    for col, w in zip("ABCDEFG", (30, 9, 24, 13, 50, 100, 10)):
        hs.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
