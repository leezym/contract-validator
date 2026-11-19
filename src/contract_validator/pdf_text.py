"""Extracción de texto de PDF con `pdftotext` (Xpdf/Poppler), sin dependencias de Python.

Se usa `-raw` (orden del flujo de contenido): en el PDF oficial de Hacienda deja líneas `Etiqueta: valor` y filas de ítems
completas; `-layout` mezcla columnas y no sirve. La ruta del ejecutable sale de la variable de entorno PDFTOTEXT o del PATH.
"""
from __future__ import annotations
import os
import re
import shutil
import subprocess
import unicodedata
from dataclasses import dataclass, field

UUID_RE = re.compile(r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")
CONTROL_RE = re.compile(r"DTE-\d{2}-[A-Z0-9]+-\d{15}")
SEAL_RE = re.compile(r"\b20\d{2}[A-Z0-9]{30,}\b")
MIN_CHARS = 40  # menos texto que esto = PDF sin contenido (el NCR interno de ejemplo solo trae la cabecera de Oracle)

# Título del documento (hoja del Excel) -> código interno; se busca la primera aparición en el texto.
TITLES = [
    ("factura de exportacion", "FAC_EXP"), ("comprobante de credito fiscal", "CCF"), ("factura electronica", "FAC"),
    ("nota de remision", "REM"), ("nota de credito", "NCR"), ("nota de debito", "NDB"), ("factura", "FAC"),
]


def deaccent(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def find_tool() -> str | None:
    env = os.environ.get("PDFTOTEXT")
    if env and os.path.exists(env):
        return env
    found = shutil.which("pdftotext")
    if found or os.name != "nt":
        return found
    # Windows: Git for Windows trae pdftotext, pero su carpeta no siempre está en el PATH de PowerShell
    roots = [os.path.dirname(os.path.dirname(g)) for g in [shutil.which("git")] if g]
    roots += [os.path.join(os.environ.get(v, ""), "Git") for v in ("ProgramFiles", "ProgramW6432", "LOCALAPPDATA") if os.environ.get(v)]
    for root in roots:
        for sub in ("ucrt64", "mingw64", "clang64"):
            exe = os.path.join(root, sub, "bin", "pdftotext.exe")
            if os.path.exists(exe):
                return exe
    return None


@dataclass
class PdfInfo:
    path: str
    text: str = ""
    kind: str = ""            # MINH (oficial Hacienda, se valida) | INTERNO (TQ, se ignora)
    error: tuple | None = None  # (codigo, mensaje)
    uuid: str = ""
    control: str = ""
    title: str = ""           # código de tipo de documento según el título impreso
    seal: str = ""
    lines: list = field(default_factory=list)


def read_pdf(path: str) -> PdfInfo:
    info = PdfInfo(path=path)
    info.kind = "MINH" if "MINH" in os.path.basename(path).upper() else "INTERNO"  # el tipo depende solo del nombre
    tool = find_tool()
    if not tool:
        info.error = ("PDF_TOOL_MISSING", "No se encontró 'pdftotext' (Xpdf/Poppler). Instálelo o defina PDFTOTEXT=<ruta>")
        return info
    try:
        proc = subprocess.run([tool, "-raw", "-enc", "UTF-8", path, "-"], capture_output=True, timeout=60)
    except (OSError, subprocess.TimeoutExpired) as e:
        info.error = ("PARSE_ERROR", f"pdftotext falló: {e}")
        return info
    if proc.returncode != 0:
        info.error = ("PARSE_ERROR", f"pdftotext devolvió {proc.returncode}: {proc.stderr.decode('utf-8', 'replace').strip()[:200]}")
        return info
    info.text = proc.stdout.decode("utf-8", "replace").replace("\x0c", "\n")
    info.lines = [ln.strip() for ln in info.text.split("\n") if ln.strip()]
    if len(re.sub(r"\s", "", info.text)) < MIN_CHARS:
        info.error = ("PDF_EMPTY", "El PDF no tiene contenido (sin texto extraíble): posible falla de generación")
    info.seal = (SEAL_RE.search(info.text) or [""])[0] if SEAL_RE.search(info.text) else ""
    info.uuid, info.control, info.title = parse_identity(info.text)
    return info


def parse_identity(text: str) -> tuple[str, str, str]:
    """(UUID del código de generación, número de control, tipo según el título impreso)."""
    uuid = ""
    m = re.search(r"C.digo de Generaci.n\s*(" + UUID_RE.pattern + ")", text)
    if m:
        uuid = m.group(1)
    elif UUID_RE.search(text):
        uuid = UUID_RE.search(text).group(0)
    ctrl = (CONTROL_RE.search(text) or [""])[0] if CONTROL_RE.search(text) else ""
    # El título es una línea que contiene solo el nombre del documento ("NOTA DE CRÉDITO"); así no se confunde con
    # etiquetas como "Fecha Nota de Remisión:" ni con el tipo del documento relacionado (que aparece después).
    title = ""
    for ln in text.splitlines():
        flat = " ".join(deaccent(ln).lower().split())  # igualdad exacta: "Nota de Remision #:" no es un título
        hit = next((code for name, code in TITLES if flat == name), "")
        if hit:
            title = hit
            break
    return uuid.upper(), ctrl, title
