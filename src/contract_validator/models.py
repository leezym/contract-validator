from __future__ import annotations
from dataclasses import dataclass, field

NEW, REMOVED, RENAMED, UNCHANGED = "NUEVO", "ELIMINADO", "RENOMBRADO", "SIN_CAMBIO"

# Código interno de tipo de documento (columna TIPO_REPORTE del Excel).
FAC, CCF, REM, NCR, NDB, FAC_EXP, INVALIDACION, RETORNO = (
    "FAC", "CCF", "REM", "NCR", "NDB", "FAC_EXP", "INVALIDACION", "RETORNO")
# identificacion.tipoDte (Ministerio de Hacienda) -> código interno
DTE_CODES = {"01": FAC, "03": CCF, "04": REM, "05": NCR, "06": NDB, "11": FAC_EXP}


@dataclass
class SpecField:
    path: str
    type: str = ""
    required: bool = True
    change: str = UNCHANGED
    old_path: str = ""
    row: int = 0
    doc_type: str = ""      # "" = aplica a cualquier tipo de documento
    fmt: str = ""           # "" | JSON | XML
    nullable: bool = False  # requerido pero el valor puede ser null ("Si (null)")
    values: tuple = ()      # valores permitidos (catálogo corto), si se conocen
    sheet: str = ""
    counterpart: str = ""   # ruta equivalente en el otro formato (JSON <-> XML_TQ)
    note: str = ""          # 'posicional' / 'posicional(7)' = mapeo por posición, no comparable 1 a 1
    description: str = ""   # columna Descripcion del libro (coincide con etiquetas de los PDF)


@dataclass
class Finding:
    severity: str  # ERROR | WARNING
    code: str
    path: str
    message: str


@dataclass
class Result:
    file: str
    findings: list[Finding] = field(default_factory=list)
    checked: int = 0
    doc_type: str = ""
    fmt: str = ""

    @property
    def errors(self):
        return [f for f in self.findings if f.severity == "ERROR"]

    @property
    def ok(self):
        return not self.errors
