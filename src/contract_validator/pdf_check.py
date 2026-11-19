"""Validación de la información del PDF contra el XML del mismo documento.

Estrategia en 3 capas, sin inventar un mapeo (ver README):
  1. Identidad: Código de Generación, Número de Control y título del PDF contra el XML (PDF_IDENTITY, PDF_DOCTYPE).
  2. Valores con ancla en los specs:
       - etiquetas del PDF que coinciden EXACTO con una `Descripcion` del Excel (se usa su tag XML_TQ, incluida la
         posición de `OtrosTotales(n)`); las filas bajo `Items` y `Notas` se excluyen porque su etiqueta es ambigua;
       - pares etiqueta/valor de `Notas` del propio XML;
       - filas de ítems (código, cantidad, precio, descuento) contra `Items[i]` y cantidad de filas.
  3. Contención: cada valor del XML debe aparecer en el PDF (PDF_NOT_PRINTED, aviso) y cada importe del PDF debe existir
     en el XML (PDF_VALUE_NOT_IN_XML, aviso).
La igualdad es exacta salvo las normalizaciones de `data/specs/pdf_presentacion.json` (lista explícita y editable).

Códigos: PDF_EMPTY, PDF_TOOL_MISSING, PDF_NO_PAIR, PDF_IDENTITY, PDF_DOCTYPE, PDF_MISMATCH, PDF_VALUE_MISSING,
PDF_COUNT_MISMATCH (ERROR); PDF_NOT_PRINTED, PDF_VALUE_NOT_IN_XML, PDF_NO_PAIR auto, PDF_IGNORED (WARNING).
Solo se valida el PDF MINH (oficial de Hacienda); el PDF interno de TQ se ignora.
"""
from __future__ import annotations
import json
import os
import re
from .crosscheck import apendice_records, document_id, xml_values
from .models import Finding, Result
from .pdf_text import PdfInfo, deaccent, read_pdf

DEFAULT_CONFIG = os.path.join("data", "specs", "pdf_presentacion.json")
MAX_NOT_PRINTED = 30
LETTER = re.compile(r"[A-Za-zÁÉÍÓÚÑáéíóúñÜü]")
DATE_DMY = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
DATE_YMD_SLASH = re.compile(r"^(\d{4})/(\d{2})/(\d{2})$")
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
NUM = re.compile(r"^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$|^[+-]?\d*\.\d+$")
ID_DASH = re.compile(r"^\d+([- ]\d+)+$")
DECIMAL_TOKEN = re.compile(r"^\d[\d,]*\.\d{2,4}$")


class Presentation:
    """Normalizaciones de presentación aceptadas (data/specs/pdf_presentacion.json). Sin archivo: igualdad exacta."""

    def __init__(self, data: dict | None = None):
        data = data or {}
        n = data.get("normalizaciones", {})
        self.fechas = bool(n.get("fechas", False))
        self.numeros = bool(n.get("numeros", False))
        self.ids = bool(n.get("identificadores", False))
        self.acentos = bool(n.get("texto_ignorar_acentos", False))
        self.mayus = bool(n.get("texto_ignorar_mayusculas", False))
        self.catalogs = {k: v for k, v in data.get("catalogos", {}).items() if isinstance(v, dict)}
        self.composite = {k: v for k, v in data.get("catalogos_compuestos", {}).items() if k != "descripcion"}
        self.printed_apendices = set(data.get("apendices_impresos", {}).get("apendices", []))
        ni = data.get("no_impreso", {})
        self.no_prefix = tuple(ni.get("prefijos", []))
        self.no_tags = set(ni.get("tags", []))

    def text(self, s: str) -> str:
        t = " ".join(str(s).split())
        if self.acentos:
            t = deaccent(t)
        return t.upper() if self.mayus else t

    def canon(self, v) -> str:
        t = self.text(v)
        if not t:
            return ""
        if self.fechas:
            m = DATE_DMY.match(t)
            if m:
                return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
            m = DATE_YMD_SLASH.match(t)
            if m:
                return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        if ISO_DATE.match(t):
            return t
        if self.numeros and NUM.match(t):
            t = t.replace(",", "")
            if t.startswith("."):
                t = "0" + t
            if "." in t:
                t = t.rstrip("0").rstrip(".") or "0"
            return t
        if self.ids and ID_DASH.match(t):
            return re.sub(r"[- ]", "", t)
        return t

    def catalog_text(self, path: str, value, prefix=None):
        """Texto impreso de un código. En catálogos compuestos el código es '<prefijo>.<valor>'."""
        code = str(value).strip()
        if path in self.composite and not _null(prefix):
            code = f"{str(prefix).strip()}.{code}"
        return self.catalogs.get(path, {}).get(code, value)

    def printed(self, path: str) -> bool:
        return not (path.startswith(self.no_prefix) or path in self.no_tags)


def load_presentation(path: str | None = None) -> Presentation:
    path = path or DEFAULT_CONFIG
    if not os.path.exists(path):
        if path != DEFAULT_CONFIG:
            raise FileNotFoundError(f"No existe el archivo de presentación: {path}")
        return Presentation()
    with open(path, encoding="utf-8-sig") as f:
        return Presentation(json.load(f))


def _null(v) -> bool:
    return v is None or str(v).strip().lower() in ("", "null")


def _label_key(s: str) -> str:
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", deaccent(s).lower()).split())


def text_continues(cfg, lines: list, index: int, start: int, expected) -> bool:
    """El texto esperado empieza en lines[index][start:] y puede continuar en las líneas siguientes."""
    txt = lines[index][start:]
    for extra in [""] + list(lines[index + 1:index + 4]):
        txt = f"{txt} {extra}".strip()
        if cfg.text(txt).startswith(cfg.text(expected)):
            return True
    return False


def apendice_pairs(labels, values) -> list[tuple[str, str]]:
    """Etiquetas y valores unidos con '|' -> [(etiqueta, valor)]."""
    ls = [x.strip().rstrip(":").strip() for x in str(labels or "").split("|")]
    vs = [x.strip() for x in str(values or "").split("|")]
    return [(l, v) for l, v in zip(ls, vs) if l and v] if len(ls) == len(vs) else []


def printed_with_label(cfg, lines: list, label: str, value: str) -> str | None:
    """None si el PDF muestra 'etiqueta: valor'; si no, lo que encontró ('' = la etiqueta no está impresa)."""
    found = ""
    rx = re.compile(r"(?<![\w])" + re.escape(label) + r"\s*:", re.I)
    for i, ln in enumerate(lines):
        for m in rx.finditer(ln):
            window = " ".join([ln[m.end():], *lines[i + 1:i + 3]])
            several = len(str(value).split()) > 1 and cfg.text(value) in cfg.text(window)
            if several or cfg.canon(value) in {cfg.canon(_tok(t)) for t in window.split()}:
                return None
            found = found or (ln[m.end():].split() or ["(sin valor)"])[0]
    return found


def _tok(t: str) -> str:
    return t.strip("$:;()").rstrip(",")


class _Ctx:
    def __init__(self, info: PdfInfo, cfg: Presentation):
        self.cfg = cfg
        self.flat = cfg.text(info.text)
        raw = [t for t in info.text.split() if t]
        self.tokset = {cfg.canon(_tok(t)) for t in raw}
        self.words = {cfg.text(w.strip(",.;:()")) for w in raw}

    def present(self, value) -> bool:
        cfg = self.cfg
        t = cfg.text(value)
        if not t:
            return True
        if LETTER.search(t):
            if t in self.flat:
                return True
            return all(cfg.text(w.strip(",.;:()")) in self.words for w in t.split())  # palabras presentes (texto partido en líneas)
        if t in self.flat or cfg.canon(t) in self.tokset:
            return True
        if "|" in t:
            return all(not p.strip() or cfg.canon(p.strip()) in self.tokset for p in t.split("|"))
        return False


def _text(el, path):
    n = el.find(path)
    return "" if n is None else (n.text or "").strip()


def _anchors(spec, doc_type):
    """etiqueta normalizada -> fila del contrato, solo si la etiqueta coincide exacto con UNA Descripcion del libro."""
    found: dict = {}
    for s in spec:
        if s.doc_type != doc_type or s.fmt != "XML" or not s.description:
            continue
        if s.path.startswith(("CabeceraEntrada.", "DatosEntrada.Notas")) or "Items" in s.path or "TotalImpuestoItem" in s.path:
            continue  # etiquetas ambiguas a nivel de ítem (p.ej. "Código", "Ventas No Sujetas")
        found.setdefault(_label_key(s.description), []).append(s)
    return {k: v[0] for k, v in found.items() if len(v) == 1 and k}


def check_pdf(info: PdfInfo, xml_file: str, xml_doc, spec, cfg: Presentation, doc_type: str = "") -> Result:
    doc_type = doc_type or xml_doc.doc_type
    res = Result(file=f"{info.path} <-> {xml_file}", doc_type=doc_type, fmt=f"PDF({info.kind})<->XML")

    def add(sev, code, path, msg):
        res.findings.append(Finding(sev, code, path, msg))

    if info.error:
        add("ERROR", info.error[0], "", info.error[1])
        return res
    root, xv, ctx = xml_doc.raw, xml_values(xml_doc.raw), _Ctx(info, cfg)

    # ---- capa 1: identidad
    res.checked += 1
    xml_uuid = (xv.get("CabeceraEntrada.uuid") or [""])[0].strip().upper()
    if info.uuid and xml_uuid and info.uuid != xml_uuid:
        add("ERROR", "PDF_IDENTITY", "CabeceraEntrada.uuid", f"Código de Generación PDF={info.uuid} XML={xml_uuid}")
    xml_ctrl = (xv.get("DatosEntrada.Encabezado.secuenciaDocumento") or [""])[0].strip()
    if xml_ctrl and info.control != xml_ctrl:
        add("ERROR", "PDF_IDENTITY", "DatosEntrada.Encabezado.secuenciaDocumento",
            f"Número de Control PDF={info.control or '(no encontrado)'} XML={xml_ctrl}")
    if doc_type and info.title and info.title != doc_type:
        add("ERROR", "PDF_DOCTYPE", "", f"El título del PDF corresponde a {info.title} y el XML es {doc_type}")
    elif doc_type and not info.title:
        add("WARNING", "PDF_DOCTYPE", "", "No se pudo identificar el título del documento en el PDF")

    # ---- capa 2a: etiquetas ancladas en las Descripcion del Excel
    anchors = _anchors(spec, doc_type)
    for line_no, ln in enumerate(info.lines):
        for m in re.finditer(r"([^:]{2,60}):\s*\$?\s*(\S*)", ln):
            key = _label_key(m.group(1))
            s = anchors.get(key)  # solo coincidencia EXACTA con la Descripcion del Excel
            pdf_val = m.group(2).strip()
            if s is None or not pdf_val:
                continue
            vals = xv.get(s.path) or []
            pm = re.search(r"posicional\((\d+)\)", s.note)
            exp = (vals[int(pm.group(1)) - 1] if pm and len(vals) >= int(pm.group(1)) else None) if pm else (vals[0] if len(vals) == 1 else None)
            if _null(exp):
                continue
            prefix_vals = xv.get(cfg.composite.get(s.path, "")) or [None]
            exp_txt = cfg.catalog_text(s.path, exp, prefix_vals[0])
            res.checked += 1
            if len(str(exp_txt).split()) > 1:  # texto de varias palabras: puede seguir en las líneas siguientes
                pdf_val = ln[m.start(2):].strip()
                if text_continues(cfg, info.lines, line_no, m.start(2), exp_txt):
                    continue
            if cfg.canon(pdf_val) != cfg.canon(exp_txt):
                add("ERROR", "PDF_MISMATCH", s.path, f"'{m.group(1).strip()}': PDF={pdf_val!r} XML={exp_txt!r}")

    # ---- capa 2b: Notas (etiqueta y valor viajan en el propio XML)
    records = apendice_records(xml_doc)
    for name in sorted(cfg.printed_apendices & set(records)):
        for label, val in apendice_pairs(records[name].get("etiqueta"), records[name].get("valor")):
            res.checked += 1
            got = printed_with_label(cfg, info.lines, label, val)
            if got == "":
                add("ERROR", "PDF_VALUE_MISSING", f"DatosEntrada.Notas.{label}",
                    f"Valor {val!r} de Notas ({label}, {name}) no aparece en el PDF: falta la etiqueta '{label}:'")
            elif got is not None:
                add("ERROR", "PDF_MISMATCH", f"DatosEntrada.Notas.{label}", f"'{label}': PDF={got!r} XML={val!r}")
    for n in root.findall("DatosEntrada/Notas"):
        if {_text(n, "campo"), _text(n, "etiqueta")} & cfg.printed_apendices:
            continue  # ya validado arriba, dato por dato con su etiqueta
        labels = [x.strip().rstrip(":").strip() for x in _text(n, "etiqueta").split("|")] if _text(n, "etiqueta") else []
        full = _text(n, "notasAdicionales")
        values = full.split("|")
        if _null(full) or full.strip(".") == "":
            continue
        res.checked += 1
        if full in ctx.flat:
            continue
        for i, val in enumerate(values):
            if not val.strip():
                continue
            label = labels[i] if len(labels) == len(values) else ""
            pdf_val = ""
            if label:
                for ln in info.lines:
                    m = re.match(r"^" + re.escape(label) + r"\s*:\s*(.*)$", ln, re.I)
                    # si lo que sigue es otra etiqueta ("Lista: Fecha Nota de Remisión:") el valor está en otra línea
                    if m and m.group(1).strip() and ":" not in m.group(1):
                        pdf_val = m.group(1).strip()
                        break
            if pdf_val:
                if cfg.canon(val) not in cfg.canon(pdf_val):
                    add("ERROR", "PDF_MISMATCH", f"DatosEntrada.Notas.{label}", f"'{label}': PDF={pdf_val!r} XML={val!r}")
            elif not ctx.present(val):
                add("ERROR", "PDF_VALUE_MISSING", f"DatosEntrada.Notas.{label or 'notasAdicionales'}",
                    f"Valor {val!r} de Notas ({label or 'sin etiqueta'}) no aparece en el PDF")

    # ---- capa 2c: ítems
    items = root.findall("DatosEntrada/Items")
    for it in items:
        code = _text(it, "codigoArticulo")
        if _null(code):
            continue
        line = next((ln for ln in info.lines if ln.split(" ", 1)[0] == code and len(ln.split()) >= 4), None)
        if line is None:
            if info.kind == "MINH":
                add("ERROR", "PDF_VALUE_MISSING", "DatosEntrada.Items.codigoArticulo", f"Ítem {_text(it, 'numero')}: el código {code} no aparece como fila del PDF")
            continue
        row = {cfg.canon(_tok(t)) for t in line.split()}
        res.checked += 1
        for name, tag in (("cantidad", "cantidadArticulo"), ("precio", "monedaPrecioUnitario"), ("descuento", "CargoDescuentos/valor")):
            v = _text(it, tag)
            if not _null(v) and cfg.canon(v) not in row:
                add("ERROR", "PDF_MISMATCH", f"DatosEntrada.Items.{tag}", f"Ítem {_text(it, 'numero')} ({code}): {name} XML={v} no figura en la fila del PDF")
    if info.kind == "MINH" and items:
        rows = [ln for ln in info.lines if re.match(r"^\d{5,}\s+\S+", ln) and len(ln.split()) >= 5 and re.search(r"\d\.\d{2}", ln)]
        if len(rows) != len(items):
            add("ERROR", "PDF_COUNT_MISMATCH", "DatosEntrada.Items", f"El PDF tiene {len(rows)} fila(s) de ítems y el XML {len(items)}")

    # ---- capa 3: contención (adelante)
    shown = 0
    seen = set()
    for path, vals in xv.items():
        if path.startswith("DatosEntrada.Notas") or path == "CabeceraEntrada.uuid" or not cfg.printed(path):
            continue
        for v in vals:
            if _null(v) or (path, v) in seen:
                continue
            seen.add((path, v))
            res.checked += 1
            if not ctx.present(v):
                shown += 1
                if shown <= MAX_NOT_PRINTED:
                    add("WARNING", "PDF_NOT_PRINTED", path, f"Valor del XML {v!r} no aparece en el PDF (si no se imprime, agréguelo a 'no_impreso')")
    if shown > MAX_NOT_PRINTED:
        add("WARNING", "PDF_NOT_PRINTED", "", f"{shown - MAX_NOT_PRINTED} valor(es) más del XML no aparecen en el PDF")

    # ---- capa 3: contención (atrás): importes del PDF que no existen en el XML
    xml_nums = set()
    for vals in xv.values():
        for v in vals:
            for part in str(v).split("|"):
                c = cfg.canon(part.strip())
                if re.fullmatch(r"-?\d+(\.\d+)?", c):
                    xml_nums.add(c)
    reported = set()
    for t in info.text.split():
        t = _tok(t)
        if DECIMAL_TOKEN.match(t):
            c = cfg.canon(t)
            if c not in xml_nums and c not in reported:
                reported.add(c)
                add("WARNING", "PDF_VALUE_NOT_IN_XML", "", f"El importe {t} del PDF no existe entre los valores del XML")
    return res


def validate_pdfs(pdf_paths, docs: dict, spec, cfg: Presentation, explicit=None, doc_type: str = "") -> list:
    """Valida cada PDF **MINH** contra su XML (emparejado por Código de Generación o explícito).

    Los PDF internos de TQ (sin MINH en el nombre) no se validan. docs: {ruta: Doc} de los XML ya cargados.
    explicit: [(xml, pdf)]. Un PDF sin XML pareja es WARNING si el emparejamiento era automático y ERROR si se pidió explícito.
    """
    key = lambda p: os.path.normcase(os.path.abspath(p))
    forced = {key(pdf): xml for xml, pdf in (explicit or [])}
    xml_by_uuid = {document_id(d): f for f, d in docs.items() if d.fmt == "XML" and d.raw is not None and document_id(d)}
    xml_keys = {key(f): f for f, d in docs.items() if d.fmt == "XML"}
    results = []
    for path in pdf_paths:
        info = read_pdf(path)
        want = forced.get(key(path))
        if info.kind != "MINH":
            if want:  # lo pidieron explícitamente: avisar que se ignora
                results.append(Result(file=path, fmt="PDF(INTERNO)", findings=[
                    Finding("WARNING", "PDF_IGNORED", "", "Solo se validan los PDF MINH; el PDF interno de TQ se ignora")]))
            continue
        xml_file = xml_keys.get(key(want)) if want else (xml_by_uuid.get(info.uuid) if info.uuid else None)
        if info.error:
            results.append(Result(file=path, doc_type=info.title, fmt="PDF(MINH)",
                                  findings=[Finding("ERROR", info.error[0], "", info.error[1])]))
        elif xml_file is None:
            sev = "ERROR" if want else "WARNING"
            msg = (f"El XML indicado ({want}) no está entre los archivos cargados" if want else
                   f"No hay XML con Código de Generación {info.uuid or '(no identificado en el PDF)'} entre los archivos; no se valida su contenido")
            results.append(Result(file=path, doc_type=info.title, fmt="PDF(MINH)", findings=[Finding(sev, "PDF_NO_PAIR", "", msg)]))
        else:
            results.append(check_pdf(info, xml_file, docs[xml_file], spec, cfg, doc_type=doc_type))
    return results
