"""Aplana JSON/XML a {ruta: {tipos}}.

Rutas con '.', sin '[]'. XML: se retira la envoltura SOAP (Envelope/Body/<Operacion>/<...Entrada>)
para que las rutas inicien en CabeceraEntrada / DatosEntrada; atributos como '@attr'; namespaces ignorados;
el texto 'null' cuenta como valor nulo.
"""
from __future__ import annotations
import json
import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from .models import DTE_CODES, FAC, CCF, REM, NCR, NDB, FAC_EXP, INVALIDACION, RETORNO

MAX_SAMPLES = 50


class Doc(dict):
    """{ruta: {tipos}} + muestras de valores, tipo de documento detectado y avisos de lectura."""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.samples: dict[str, set[str]] = {}
        self.doc_type = ""
        self.fmt = ""
        self.raw = None  # JSON parseado o elemento XML (sin envoltura SOAP), para las reglas de correctitud
        self.warnings: list[tuple[str, str]] = []


def _scalar_type(v) -> str:
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, int):
        return "integer"
    if isinstance(v, float):
        return "number"
    return "string"


def _text_type(t: str) -> str:
    t = t.strip()
    if t == "" or t.lower() == "null":  # elemento vacío = sin valor
        return "null"
    if re.fullmatch(r"[+-]?\d+", t):
        return "integer"
    if re.fullmatch(r"[+-]?\d*\.\d+", t):
        return "number"
    if t.lower() in ("true", "false"):
        return "boolean"
    return "string"


def _sample(out, path, value):
    if isinstance(out, Doc):
        s = out.samples.setdefault(path, set())
        if len(s) < MAX_SAMPLES:
            s.add(str(value))


def flatten_json(data, prefix="", out=None) -> Doc:
    out = Doc() if out is None else out
    if isinstance(data, dict):
        for k, v in data.items():
            p = f"{prefix}.{k}" if prefix else k
            kind = "object" if isinstance(v, dict) else "array" if isinstance(v, list) else _scalar_type(v)
            out.setdefault(p, set()).add(kind)
            if kind not in ("object", "array", "null"):
                _sample(out, p, v)
            flatten_json(v, p, out)
    elif isinstance(data, list):
        for item in data:
            if prefix and not isinstance(item, (dict, list)):
                _sample(out, prefix, item)
            flatten_json(item, prefix, out)
    return out


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1]


def flatten_xml(elem, prefix="", out=None) -> Doc:
    out = Doc() if out is None else out
    for k, v in elem.attrib.items():
        k = _local(k)
        if k.startswith("xmlns"):
            continue
        p = f"{prefix}.@{k}" if prefix else f"@{k}"
        out.setdefault(p, set()).add(_text_type(v))
        _sample(out, p, v)
    for child in elem:
        p = f"{prefix}.{_local(child.tag)}" if prefix else _local(child.tag)
        if len(child):
            kind = "object"
        else:
            kind = _text_type(child.text or "")
            if kind != "null":
                _sample(out, p, (child.text or "").strip())
        out.setdefault(p, set()).add(kind)
        flatten_xml(child, p, out)
    return out


def strip_soap(root):
    """Envelope/Body/<Operacion>/<...Entrada> -> <...Entrada>. Si no hay envoltura devuelve la raíz."""
    if _local(root.tag) != "Envelope":
        return root
    body = next((c for c in root if _local(c.tag) == "Body"), None)
    if body is None or not len(body):
        return root
    op = body[0]
    return op[0] if len(op) == 1 else op


def decode_bytes(raw: bytes) -> tuple[str, str | None]:
    """UTF-8 y, si falla, cp1252. Devuelve (texto, aviso)."""
    try:
        return raw.decode("utf-8-sig"), None
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace"), (
            "El archivo no es UTF-8 válido; se leyó como cp1252/Latin-1 (los acentos y ñ se corrompen "
            "si el emisor no declara el encoding)")


def _detect(doc: Doc) -> str:
    if doc.fmt == "JSON":
        tipo = next(iter(doc.samples.get("identificacion.tipoDte", [])), "")
        if tipo in DTE_CODES:
            return DTE_CODES[tipo]
        if "identificacion.tipoEvento" in doc:
            return RETORNO
        if "motivo" in doc and "documento" in doc and "cuerpoDocumento" not in doc:
            return INVALIDACION
    else:
        tipo = next(iter(doc.samples.get("DatosEntrada.Encabezado.tipoFactura", [])), "")
        if tipo in DTE_CODES:
            return DTE_CODES[tipo]
        # Eventos (sin tipoFactura): la operación SOAP lo identifica (<EventoRetornoEntrada>, <EventoInvalidacionEntrada>)
        tag = _local(doc.raw.tag) if doc.raw is not None else ""
        if "EventoRetorno" in tag:
            return RETORNO
        if "EventoInvalidacion" in tag:
            return INVALIDACION
    return ""


def load_document(path: str) -> Doc:
    low = path.lower()
    with open(path, "rb") as f:
        raw = f.read()
    text, warn = decode_bytes(raw)
    if low.endswith(".json"):
        raw_obj = json.loads(text)
        doc = flatten_json(raw_obj)
        doc.fmt, doc.raw = "JSON", raw_obj
    elif low.endswith(".xml"):
        text = re.sub(r"^\s*<\?xml[^>]*\?>", "", text)  # la declaración de encoding ya no aplica al texto decodificado
        entrada = strip_soap(ET.fromstring(text.encode("utf-8")))
        doc = flatten_xml(entrada)
        doc.fmt, doc.raw = "XML", entrada
    else:
        raise ValueError(f"Extensión no soportada: {path}")
    if warn:
        doc.warnings.append(("ENCODING", warn))
    doc.doc_type = _detect(doc)
    return doc


_NAME_HINTS = [  # (fragmento en el nombre de archivo normalizado, código)
    ("exportacion", FAC_EXP), ("credito fiscal", CCF), ("ccf", CCF), ("debito", NDB), ("remision", REM),
    ("invalidacion", INVALIDACION), ("retorno", RETORNO), ("credito", NCR), ("factura", FAC),
]


def doc_type_from_filename(path: str) -> str:
    name = unicodedata.normalize("NFKD", os.path.splitext(os.path.basename(path))[0])
    name = name.encode("ascii", "ignore").decode().lower()
    name = re.sub(r"[-_.]+", " ", name)
    return next((code for frag, code in _NAME_HINTS if frag in name), "")
