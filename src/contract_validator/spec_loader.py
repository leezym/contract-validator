"""Lee el .xlsx de especificación (contrato) y lo normaliza a SpecField."""
from __future__ import annotations
import re
import unicodedata
from openpyxl import load_workbook
from .models import SpecField, NEW, REMOVED, RENAMED, UNCHANGED

ALIASES = {
    "path": {"campo", "field", "path", "ruta", "nombre", "nombre_campo"},
    "type": {"tipo", "type", "tipo_dato", "datatype"},
    "required": {"requerido", "required", "obligatorio", "mandatory"},
    "change": {"cambio", "change", "estado", "status", "accion"},
    "old_path": {"campo_anterior", "old_field", "old_path", "nombre_anterior", "ruta_anterior"},
    "doc_type": {"tipo_doc", "tipo_documento", "doc_type"},
    "fmt": {"formato", "format"},
    "values": {"valores", "valores_permitidos", "values", "enum"},
    "sheet": {"hoja_origen"},
    "counterpart": {"ruta_contraparte"},
    "note": {"nota"},
    "description": {"descripcion", "description"},
    "src_row": {"fila_origen"},
}
CHANGE_MAP = {
    "nuevo": NEW, "new": NEW, "alta": NEW, "agregado": NEW,
    "eliminado": REMOVED, "removed": REMOVED, "baja": REMOVED, "borrado": REMOVED,
    "renombrado": RENAMED, "renamed": RENAMED, "rename": RENAMED,
    "": UNCHANGED, "sin_cambio": UNCHANGED, "igual": UNCHANGED, "unchanged": UNCHANGED,
}
TRUE = {"si", "s", "y", "yes", "true", "1", "x", "obligatorio"}
NULLABLE = {"si_null", "si_(null)", "si(null)"}
TYPE_MAP = {
    "string": "string", "str": "string", "texto": "string", "varchar": "string", "date": "string", "datetime": "string",
    "int": "integer", "integer": "integer", "entero": "integer", "long": "integer",
    "number": "number", "decimal": "number", "float": "number", "double": "number", "numerico": "number",
    "bool": "boolean", "boolean": "boolean", "booleano": "boolean",
    "hour": "string", "time": "string", "timestamp": "string",
    "object": "object", "objeto": "object", "array": "array", "lista": "array",
}


def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", "_", s.strip().lower())


def norm_path(p) -> str:
    """Normaliza rutas: separador '.', sin '[]' ni separador inicial."""
    p = str(p if p is not None else "").replace(" ", " ").strip().rstrip("*").strip().replace("/", ".").replace("[]", "")
    return re.sub(r"\.+", ".", p).strip(".")


DEFAULT_SHEET = "contrato"


def load_spec(xlsx: str, sheet: str | None = None) -> list[SpecField]:
    """Lee la hoja plana 'contrato' (generada por `build`) o, si no existe, la primera hoja / `sheet`."""
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else (wb[DEFAULT_SHEET] if DEFAULT_SHEET in wb.sheetnames else wb.worksheets[0])
    rows = ws.iter_rows(values_only=True)
    header = [_norm(c) for c in next(rows)]
    col = {}
    for key, names in ALIASES.items():
        for i, h in enumerate(header):
            if h in names:
                col[key] = i
                break
    if "path" not in col:
        raise ValueError(f"El Excel debe tener una columna de campo {sorted(ALIASES['path'])}. Encabezados: {header}")

    def get(r, k):
        i = col.get(k)
        return r[i] if i is not None and i < len(r) else None

    fields = []
    for n, r in enumerate(rows, start=2):
        path = norm_path(get(r, "path"))
        if not path:
            continue
        change = CHANGE_MAP.get(_norm(get(r, "change")))
        if change is None:
            raise ValueError(f"Fila {n}: valor de cambio desconocido '{get(r, 'change')}'")
        req = get(r, "required")
        nullable = _norm(req) in NULLABLE
        required = True if req in (None, "") else (req is True or nullable or _norm(req) in TRUE)
        values = tuple(v for v in str(get(r, "values") or "").split("|") if v)
        src_row = get(r, "src_row")
        sheet = str(get(r, "sheet") or "")
        if isinstance(src_row, str) and src_row.strip() and not src_row.strip().isdigit():
            sheet = f"{sheet} ajuste {src_row.strip()}"  # fila agregada por ajuste: se identifica por su id
            src_row = 0
        sf = SpecField(
            path=path, type=TYPE_MAP.get(_norm(get(r, "type")), _norm(get(r, "type"))),
            required=required, change=change, old_path=norm_path(get(r, "old_path")),
            row=int(src_row) if isinstance(src_row, (int, float)) or str(src_row).isdigit() else n,
            doc_type=str(get(r, "doc_type") or "").strip().upper(), fmt=str(get(r, "fmt") or "").strip().upper(),
            nullable=nullable, values=values, sheet=sheet,
            counterpart=norm_path(get(r, "counterpart")), note=str(get(r, "note") or ""),
            description=str(get(r, "description") or ""),
        )
        if change == RENAMED and not sf.old_path:
            raise ValueError(f"Fila {n}: RENOMBRADO requiere 'campo_anterior'")
        fields.append(sf)
    return fields
