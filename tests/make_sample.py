"""Genera un contrato simple (contrato_simple.xlsx) y archivos de ejemplo en data/inputs de `base`.

No usa el nombre contrato_simple.xlsx: ese es el libro real de especificación."""
import os
from openpyxl import Workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ROWS = [
    ["campo", "tipo", "requerido", "cambio", "campo_anterior"],
    ["cliente.id", "integer", "SI", "SIN_CAMBIO", ""],
    ["cliente.email", "string", "SI", "NUEVO", ""],
    ["cliente.telefono", "string", "NO", "ELIMINADO", ""],
    ["cliente.nombreCompleto", "string", "SI", "RENOMBRADO", "cliente.nombre"],
    ["cliente.saldo", "number", "SI", "SIN_CAMBIO", ""],
]

GOOD_JSON = '{"cliente": {"id": 1, "email": "a@b.com", "nombreCompleto": "Ana", "saldo": 10.5}}'
BAD_JSON = '{"cliente": {"id": "x", "nombre": "Ana", "telefono": "555", "saldo": 10, "extra": true}}'
GOOD_XML = "<root><cliente><id>1</id><email>a@b.com</email><nombreCompleto>Ana</nombreCompleto><saldo>10.5</saldo></cliente></root>"
BAD_XML = "<root><cliente><id>1</id><nombre>Ana</nombre><telefono>555</telefono></cliente></root>"


def build(base=ROOT):
    os.makedirs(f"{base}/data/specs", exist_ok=True)
    os.makedirs(f"{base}/data/inputs", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    for r in ROWS:
        ws.append(r)
    wb.save(f"{base}/data/specs/contrato_simple.xlsx")
    for name, content in {"ok.json": GOOD_JSON, "bad.json": BAD_JSON, "ok.xml": GOOD_XML, "bad.xml": BAD_XML}.items():
        with open(f"{base}/data/inputs/{name}", "w", encoding="utf-8") as f:
            f.write(content)


if __name__ == "__main__":
    build()
