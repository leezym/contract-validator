"""Compañía (CIA) de cada documento y esquema Oracle donde se consultan sus valores esperados.

Configuración: data/specs/companias.json
    {"companias": {"53": {"nombre": "...", "esquema": "VENTASSAL", "nit": ["..."], "codigo_empresa": ["..."]}, ...}}

La CIA se identifica con lo que traen los archivos del documento (todos los que comparten el Código de Generación):
  * NIT del emisor: JSON `emisor.nit` / XML `DatosEntrada/InformacioGeneral/InformacionEmisor/nitCompania`;
  * JSON `codigoEmpresa`;
  * prefijo del nombre del XML (`41_24_01_...`) y del PDF oficial (`..._MINH_41...`).
Si las señales no coinciden o la CIA no está configurada, el documento no se consulta en la BD (queda NO VERIFICADA
con el motivo). Agregar una compañía = agregar una entrada en el .json; el query es el mismo para todas.
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field

DEFAULT_COMPANIAS = os.path.join("data", "specs", "companias.json")
SCHEMA_RX = re.compile(r"^[A-Z][A-Z0-9_$#]{0,29}$")
XML_NAME_RX = re.compile(r"^(\d{2})_\d{2}_\d{2}_")
PDF_NAME_RX = re.compile(r"MINH_(\d{2})\d", re.I)
NIT_XML = "DatosEntrada/InformacioGeneral/InformacionEmisor/nitCompania"


class CompaniaError(ValueError):
    pass


@dataclass
class Compania:
    cia: str
    nombre: str
    esquema: str
    nit: list = field(default_factory=list)
    codigo_empresa: list = field(default_factory=list)

    @property
    def etiqueta(self) -> str:
        return f"CIA {self.cia} ({self.nombre}, esquema {self.esquema})"


def load_companias(path: str | None = None) -> dict[str, Compania]:
    p = path or DEFAULT_COMPANIAS
    with open(p, encoding="utf-8-sig") as fh:
        data = json.load(fh).get("companias", {})
    out = {}
    for cia, c in data.items():
        esquema = str(c.get("esquema", "")).strip().upper()
        if not SCHEMA_RX.match(esquema):
            raise CompaniaError(f"{p}: esquema inválido para la CIA {cia}: {esquema!r}")
        out[str(cia).strip()] = Compania(str(cia).strip(), str(c.get("nombre", "")).strip(), esquema,
                                         [_digits(n) for n in c.get("nit", []) if _digits(n)],
                                         [str(x).strip().upper() for x in c.get("codigo_empresa", []) if str(x).strip()])
    return out


def _digits(v) -> str:
    return re.sub(r"\D", "", str(v or ""))


def signals(path: str, doc=None, companias: dict[str, Compania] | None = None) -> list[tuple[str, str]]:
    """[(cia, evidencia)] que aporta un archivo. `doc` = documento cargado (JSON/XML) o None (PDF)."""
    companias = companias or {}
    name = os.path.basename(path)
    out = []
    m = (XML_NAME_RX.match(name) if name.lower().endswith(".xml") else
         PDF_NAME_RX.search(name) if name.lower().endswith(".pdf") else None)
    if m:
        out.append((m.group(1), f"nombre del archivo {name}"))
    raw = getattr(doc, "raw", None)
    if raw is None:
        return out
    if getattr(doc, "fmt", "") == "XML":
        n = raw.find(NIT_XML)
        nit, empresa = _digits(n.text if n is not None else ""), ""
    else:
        nit = _digits(((raw.get("emisor") or {}).get("nit")) if isinstance(raw, dict) else "")
        empresa = str(raw.get("codigoEmpresa") or "").strip().upper() if isinstance(raw, dict) else ""
    for c in companias.values():
        if nit and nit in c.nit:
            out.append((c.cia, f"NIT del emisor en {name}"))
        if empresa and empresa in c.codigo_empresa:
            out.append((c.cia, f"codigoEmpresa={empresa} en {name}"))
    return out


def identify(files_by_uuid: dict[str, list], docs: dict, companias: dict[str, Compania]):
    """files_by_uuid: {uuid: [rutas]}; docs: {ruta: Doc}. Devuelve ({uuid: Compania}, {uuid: motivo sin CIA})."""
    found, problems = {}, {}
    for uuid, files in files_by_uuid.items():
        sig = [s for f in files for s in signals(f, docs.get(f), companias)]
        cias = sorted({c for c, _ in sig})
        if not cias:
            problems[uuid] = ("no se pudo identificar la compañía (CIA) del documento: el NIT del emisor, codigoEmpresa y "
                              f"el nombre del archivo no corresponden a ninguna de {_known(companias)} "
                              "(data/specs/companias.json)")
        elif len(cias) > 1:
            detalle = "; ".join(f"CIA {c}: {', '.join(e for x, e in sig if x == c)}" for c in cias)
            problems[uuid] = f"los archivos del documento indican compañías distintas ({detalle})"
        elif cias[0] not in companias:
            problems[uuid] = (f"la CIA {cias[0]} ({', '.join(e for _, e in sig)}) no está configurada en "
                              f"data/specs/companias.json; configuradas: {_known(companias)}")
        else:
            found[uuid] = companias[cias[0]]
    return found, problems


def _known(companias) -> str:
    return ", ".join(f"{c.cia} ({c.esquema})" for c in companias.values()) or "ninguna"
