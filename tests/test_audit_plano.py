import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import pytest
from openpyxl import Workbook, load_workbook
from contract_validator.build_contract import build
from contract_validator import audit_plano

V3V4 = "TAG_MIN_HACIENDA - (v3 → v4)"


def _libro(path):
    """Libro sintético con el layout real: hoja 'XML' (una columna por tipo) + hoja 'JSON' (base v3)."""
    wb = Workbook()
    x = wb.active
    x.title = "XML"
    x.append(["TAG_MIN_HACIENDA", "Descripcion", "Tipo", "XML_TQ", "Requerido", None, "ND - 06"])
    x.append(["a", "A", "object", "x", None, None, "a"])
    x.append(["a/b", "B", "string", "DatosEntrada/Encabezado/b", "Si", None, "a/b"])
    x.append(["a/c", "C", "00, 01", "DatosEntrada/Encabezado/c", "Si (null)", None, "Adicionar a/c"])
    x.append(["a/d", "D", "string", None, "No", None, "Eliminar a/d"])
    x.append(["a/e", "E", "string", "DatosEntrada/Encabezado/e", "No", None, "Renombrar a/old -> a/e"])
    x.append(["a/items/f", "F", "string", "DatosEntrada/Items/f (2)", "Si", None, "a/items/f"])
    x.append(["a/b", "B", "string", None, "Si", None, "a/b"])  # duplicada -> lint
    j = wb.create_sheet("JSON")
    j.append([None, "ND - 06"])
    for r in ("a", "a/b", "a/d", "a/old", "a/items/f"):
        j.append([None, r])
    wb.save(path)


@pytest.fixture()
def env(tmp_path):
    src, plano = str(tmp_path / "orig.xlsx"), str(tmp_path / "plano.xlsx")
    _libro(src)
    build(src, plano)
    return src, plano


def errors(rows):
    return [(r["codigo"], r["ruta"]) for r in rows if r["nivel"] == "ERROR"]


def _edit(plano, fn):
    wb = load_workbook(plano)
    ws = wb["contrato"]
    head = [c.value for c in ws[1]]
    fn(ws, {h: i + 1 for i, h in enumerate(head)})
    wb.save(plano)


def test_plano_fiel_pasa(env):
    rows, stats = audit_plano.audit(*env)
    assert errors(rows) == []
    assert stats["XML [NDB]"]["omitidas_lint"] == 1
    assert any(r["codigo"] == "LINT_DUPLICATE" for r in rows)


def test_fila_borrada_es_row_lost(env):
    src, plano = env
    _edit(plano, lambda ws, c: [ws.delete_rows(i) for i in range(ws.max_row, 1, -1)
                                if ws.cell(i, c["ruta"]).value == "a/c" and ws.cell(i, c["formato"]).value == "JSON"])
    codes = {e[0] for e in errors(audit_plano.audit(src, plano)[0])}
    assert "ROW_LOST" in codes and "PLANO_STALE" in codes


def test_requerido_alterado(env):
    src, plano = env

    def f(ws, c):
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, c["ruta"]).value == "a/b" and ws.cell(i, c["formato"]).value == "JSON":
                ws.cell(i, c["requerido"]).value = "NO"
    _edit(plano, f)
    assert ("FIELD_DIFF", "a/b") in errors(audit_plano.audit(src, plano)[0])


def test_cambio_invertido(env):
    src, plano = env

    def f(ws, c):
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, c["ruta"]).value == "a/d":
                ws.cell(i, c["cambio"]).value = "NUEVO"
    _edit(plano, f)
    assert ("FIELD_DIFF", "a/d") in errors(audit_plano.audit(src, plano)[0])


def test_fila_inventada_y_contraparte(env):
    src, plano = env

    def f(ws, c):
        ws.append(["NDB", "JSON", "inventado", "", "", "NO", "SIN_CAMBIO", "", "DatosEntrada/Encabezado/zzz", "", "XML", 99, ""])
    _edit(plano, f)
    codes = {e[0] for e in errors(audit_plano.audit(src, plano)[0])}
    assert {"ROW_UNTRACEABLE", "COUNTERPART_ASYMMETRIC"} <= codes


def test_fila_de_otro_tipo_no_es_trazable(env):
    """Una fila de la hoja 'XML' alimenta a varios tipos: la trazabilidad incluye el tipo de documento."""
    src, plano = env

    def f(ws, c):
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, c["ruta"]).value == "a/b" and ws.cell(i, c["formato"]).value == "JSON":
                ws.cell(i, c["tipo_doc"]).value = "NCR"
    _edit(plano, f)
    assert ("ROW_UNTRACEABLE", "a/b") in errors(audit_plano.audit(src, plano)[0])


def test_hoja_json_distinta_va_a_lint(tmp_path):
    from contract_validator.build_contract import build_rows
    src = str(tmp_path / "orig.xlsx")
    _libro(src)
    wb = load_workbook(src)
    wb["JSON"].append([None, "a/zzz"])
    wb.save(src)
    _, lint = build_rows(src)
    assert any("a/zzz" in m and h == "JSON" for h, _f, _n, m in lint)
