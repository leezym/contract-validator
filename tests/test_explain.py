"""Explicación en lenguaje natural de los hallazgos (columna 'explicacion' del reporte). Datos sintéticos."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.explain import explain, field_names
from contract_validator.models import Finding, Result, SpecField

SPEC = [SpecField(path="resumen.totalPagar", doc_type="CCF", fmt="JSON", description="Total a Pagar"),
        SpecField(path="cuerpoDocumento.cantidad", doc_type="CCF", fmt="JSON", description="Cantidad"),
        SpecField(path="DatosEntrada.Encabezado.nitCompania", doc_type="CCF", fmt="XML", description="NIT Emisor")]
NAMES = field_names(SPEC)


def ex(code, path, msg, fmt="JSON", sev="ERROR", dt="CCF"):
    return explain(Finding(sev, code, path, msg), Result("f", doc_type=dt, fmt=fmt), NAMES)


@pytest.mark.parametrize("args, expected", [
    (("REQUIRED_MISSING", "resumen.totalPagar", "Campo requerido no encontrado (cambio: SIN_CAMBIO, Comprobante fila 9)"),
     "Falta «Total a Pagar» (resumen.totalPagar) en el JSON. El contrato lo exige para el Comprobante de Crédito Fiscal. Definido en: Comprobante fila 9."),
    (("NULL_NOT_ALLOWED", "resumen.totalPagar", "Campo requerido con valor nulo (x fila 1)"),
     "«Total a Pagar» (resumen.totalPagar) viene vacío (null) en el JSON, pero el contrato exige que traiga un valor."),
    (("UNEXPECTED_FIELD", "DatosEntrada.Otro", "Campo presente en el archivo pero no definido en el contrato", "XML"),
     "El XML trae el campo DatosEntrada.Otro, que el contrato no define para el Comprobante de Crédito Fiscal. "
     "Si debe aceptarse, hay que agregarlo al contrato (ajuste aprobado); si no, el generador debe dejar de enviarlo."),
    (("XREF_MISMATCH", "resumen.totalPagar", "resumen.totalPagar <-> X: JSON='10.00' XML='11.00'", "XML<->JSON"),
     "El JSON y el XML del mismo documento no coinciden en «Total a Pagar» (resumen.totalPagar): el JSON dice «10.00» y el XML dice «11.00»."),
    (("DB_MISMATCH_JSON", "cuerpoDocumento.cantidad[2]", "cuerpoDocumento.cantidad[2]: BD='35' documento=36", "BD<->JSON"),
     "Según la base de datos, «Cantidad» (cuerpoDocumento.cantidad) del ítem 2 debería ser «35», pero el JSON trae «36»."),
    (("DB_MISMATCH_XML", "resumen.totalPagar", "resumen.totalPagar <-> X: BD='1812.17' falta en el documento", "BD<->XML"),
     "Según la base de datos, «Total a Pagar» (resumen.totalPagar) debería ser «1812.17», pero el XML no lo trae."),
    (("DB_MISMATCH_PDF", "emisor.recintoFiscal", "'Recinto fiscal': BD='Terrestre San Bartolo' PDF='Maritima'", "BD<->PDF(MINH)"),
     "En el PDF, «Recinto fiscal» muestra «Maritima», pero según la base de datos debería ser «Terrestre San Bartolo»."),
    (("PDF_MISMATCH", "X", "'Modelo de Facturación': PDF='Diferido' XML='Previo'", "PDF(MINH)<->XML"),
     "En el PDF, «Modelo de Facturación» muestra «Diferido», pero el XML tiene «Previo»."),
])
def test_templates(args, expected):
    assert ex(*args) == expected


def test_db_missing_variants():
    assert "no se encontró en la base de datos" in ex("DB_MISSING", "", "no está en el resultado del query de BD", "BD<->JSON")
    w = ex("DB_MISSING", "", "2 campo(s) sin fuente en BD (no verificables): a, b", "BD<->JSON", "WARNING")
    assert w.startswith("2 campo(s) del JSON no se pueden verificar") and "informativo" in w   # contracción 'del'
    assert "invalidaciones y los retornos" in ex("DB_MISSING", "", "INVALIDACION: tipo de documento fuera del alcance", "BD<->JSON", "WARNING")


def test_unknown_code_falls_back_to_message():
    assert ex("ALGO_NUEVO", "x", "mensaje técnico") == "mensaje técnico"


def test_report_has_explanation_column(tmp_path):
    from openpyxl import load_workbook
    from contract_validator import report
    r = Result("doc.json", [Finding("ERROR", "NULL_NOT_ALLOWED", "resumen.totalPagar", "Campo requerido con valor nulo (x fila 1)")],
               5, "CCF", "JSON")
    out = tmp_path / "r.xlsx"
    report.to_xlsx([r, Result("ok.json", [], 3, "CCF", "JSON")], str(out), spec=SPEC)
    rows = list(load_workbook(out)["Hallazgos"].iter_rows(values_only=True))
    assert rows[0] == ("archivo", "tipo_doc", "formato", "validacion", "severidad", "codigo", "campo", "explicacion", "mensaje")
    assert rows[1][7].startswith("«Total a Pagar»") and rows[1][8].startswith("Campo requerido")   # técnico se conserva
    assert rows[2][7].startswith("Sin hallazgos")
    data = json.loads(report.to_json([r], spec=SPEC))
    assert data[0]["findings"][0]["explicacion"].startswith("«Total a Pagar»")
