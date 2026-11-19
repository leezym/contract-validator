"""Veredicto por documento y flujo 'verificar' de punta a punta (datos sintéticos, BD falsa)."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator import db_fetch, verdict
from contract_validator.models import Finding, Result

U = "AAAAAAAA-0000-4000-8000-000000000001"
E = lambda code: Finding("ERROR", code, "x", "m")
W = lambda code: Finding("WARNING", code, "", "m")


def _ids(*files):
    return {f: U for f in files}


def test_all_ok_is_cumple():
    rs = [Result("a.json", [], 10, "CCF", "JSON"), Result("a.xml", [], 10, "CCF", "XML"),
          Result("a.xml <-> a.json", [], 5, "CCF", "XML<->JSON"),
          Result("a.json", [W("DB_MISSING")], 40, "CCF", "BD<->JSON"), Result("a.xml", [], 30, "CCF", "BD<->XML")]
    [v] = verdict.build(rs, _ids("a.json", "a.xml"))
    assert (v["estructura"], v["informacion_bd"], v["coherencia"], v["resultado"]) == ("CUMPLE", "CORRECTA", "OK", "CUMPLE")
    assert "70 valor(es)" in v["informacion_bd_detalle"] and "1 aviso" in v["informacion_bd_detalle"]
    assert verdict.exit_code([v]) == 0


def test_structure_fails_db_ok():
    rs = [Result("a.json", [E("REQUIRED_MISSING")], 10, "CCF", "JSON"), Result("a.json", [], 40, "CCF", "BD<->JSON")]
    [v] = verdict.build(rs, _ids("a.json"))
    assert (v["estructura"], v["informacion_bd"], v["resultado"]) == ("NO CUMPLE", "CORRECTA", "NO CUMPLE")
    assert "REQUIRED_MISSING x1" in v["estructura_detalle"]
    assert verdict.exit_code([v]) == 1


def test_db_mismatch_is_incorrecta():
    rs = [Result("a.json", [E("DB_MISMATCH_JSON")], 10, "CCF", "BD<->JSON")]
    [v] = verdict.build(rs, _ids("a.json"))
    assert (v["informacion_bd"], v["resultado"]) == ("INCORRECTA", "NO CUMPLE")


def _base():
    return [Result("a.json", [], 10, "CCF", "JSON"), Result("a.json", [], 10, "CCF", "BD<->JSON")]


def test_pdf_has_its_own_verdict():
    ok = _base() + [Result("a.pdf <-> a.xml", [], 20, "CCF", "PDF(MINH)<->XML"), Result("a.pdf", [], 8, "CCF", "BD<->PDF(MINH)")]
    [v] = verdict.build(ok, _ids("a.json", "a.pdf"))
    assert (v["validacion_pdf"], v["informacion_bd"], v["coherencia"], v["resultado"]) == ("CUMPLE", "CORRECTA", "NO APLICA", "CUMPLE")
    assert "vs XML: 20 revisado(s)" in v["validacion_pdf_detalle"] and "vs BD: 8 revisado(s)" in v["validacion_pdf_detalle"]
    bad = _base() + [Result("a.pdf <-> a.xml", [E("PDF_MISMATCH")], 20, "CCF", "PDF(MINH)<->XML"),
                     Result("a.pdf", [E("DB_MISMATCH_PDF")], 8, "CCF", "BD<->PDF(MINH)")]
    [v] = verdict.build(bad, _ids("a.json", "a.pdf"))
    assert (v["validacion_pdf"], v["informacion_bd"], v["resultado"]) == ("NO CUMPLE", "CORRECTA", "NO CUMPLE")


def test_pdf_tool_missing_is_not_verified_and_keeps_minh_kind(tmp_path, monkeypatch):
    from contract_validator import pdf_text
    monkeypatch.setattr(pdf_text, "find_tool", lambda: None)
    info = pdf_text.read_pdf(str(tmp_path / "EXP_SAL_MINH_1.pdf"))
    assert info.kind == "MINH" and info.error[0] == "PDF_TOOL_MISSING"
    assert pdf_text.read_pdf(str(tmp_path / "EXP_SAL_2026_1.pdf")).kind == "INTERNO"
    pdf = [Result("a.pdf", [E("PDF_TOOL_MISSING")], 0, "", "PDF(MINH)"), Result("a.pdf", [E("PDF_TOOL_MISSING")], 0, "", "BD<->PDF(MINH)")]
    [v] = verdict.build(_base() + pdf, _ids("a.json", "a.pdf"))
    assert (v["validacion_pdf"], v["resultado"]) == ("NO VERIFICADO", "INCOMPLETO")
    assert "no se pudo leer" in v["validacion_pdf_detalle"]


def test_pdf_missing_or_ignored_is_reported(tmp_path):
    from openpyxl import load_workbook
    from contract_validator import report
    [v] = verdict.build(_base(), _ids("a.json"))
    assert (v["validacion_pdf"], v["resultado"]) == ("NO RECIBIDO", "CUMPLE")
    [v] = verdict.build(_base(), _ids("a.json"), ignored_pdfs={"interno.pdf": U})
    assert v["validacion_pdf"] == "IGNORADO" and "interno.pdf" in v["validacion_pdf_detalle"] and "PDF" in v["archivos"]
    out = tmp_path / "r.xlsx"
    report.to_xlsx(_base(), str(out), [v])
    wb = load_workbook(out)
    assert wb["Veredicto"][1][9].value == "validacion_pdf" and wb["Veredicto"][2][9].value == "IGNORADO"
    assert [c.value for c in wb["Resumen"][wb["Resumen"].max_row]][3:5] == ["Validación del PDF", "IGNORADO"]
    assert [c.value for c in wb["Hallazgos"][wb["Hallazgos"].max_row]][3:6] == ["Validación del PDF", "WARNING", "PDF_IGNORED"]


def test_db_unavailable_is_incompleto():
    rs = [Result("a.json", [], 10, "CCF", "JSON")]
    [v] = verdict.build(rs, _ids("a.json"), db_status="no se pudo consultar la BD: sin VPN")
    assert (v["informacion_bd"], v["resultado"]) == ("NO VERIFICADA", "INCOMPLETO")
    assert "sin VPN" in v["informacion_bd_detalle"]
    assert verdict.exit_code([v]) == 2


def test_document_not_in_db_is_not_verified():
    rs = [Result("a.json", [], 10, "CCF", "JSON"),
          Result("a.json", [Finding("ERROR", "DB_MISSING", "", "no está en el resultado del query")], 0, "CCF", "BD<->JSON")]
    [v] = verdict.build(rs, _ids("a.json"))
    assert (v["informacion_bd"], v["resultado"]) == ("NO VERIFICADA", "INCOMPLETO")


def test_out_of_scope_type_is_not_verified(tmp_path):
    from contract_validator.db_check import BdConfig, validate_db
    from contract_validator.flatten import load_document
    from contract_validator.pdf_check import Presentation
    p = tmp_path / "inv.json"
    p.write_text(json.dumps({"identificacion": {"codigoGeneracion": U}, "motivo": {"tipoAnulacion": 2}}), encoding="utf-8")
    [r] = validate_db({}, {str(p): load_document(str(p))}, [], [], BdConfig(), Presentation(), "INVALIDACION")
    assert r.ok and r.findings[0].code == "DB_MISSING" and "fuera del alcance" in r.findings[0].message
    [v] = verdict.build([Result(str(p), [], 5, "INVALIDACION", "JSON"), r], {str(p): U})
    assert (v["informacion_bd"], v["resultado"]) == ("NO VERIFICADA", "INCOMPLETO")
    assert "fuera del alcance" in v["informacion_bd_detalle"]


def test_folder_scan_skips_validator_config(tmp_path):
    from contract_validator.cli import collect
    for n in ("doc.json", "pdf_presentacion.json", "normalizaciones_bd.json", "x.local.json", "a.xml"):
        (tmp_path / n).write_text("{}", encoding="utf-8")
    assert sorted(os.path.basename(f) for f in collect([str(tmp_path)])) == ["a.xml", "doc.json"]


def test_groups_by_document_and_orders_failures_first():
    u2 = "BBBBBBBB-0000-4000-8000-000000000002"
    rs = [Result("ok.json", [], 1, "FAC", "JSON"), Result("ok.json", [], 1, "FAC", "BD<->JSON"),
          Result("bad.json", [E("UNEXPECTED_FIELD")], 1, "CCF", "JSON"), Result("bad.json", [], 1, "CCF", "BD<->JSON"),
          Result("roto.json", [E("PARSE_ERROR")])]
    vs = verdict.build(rs, {"ok.json": U, "bad.json": u2})
    assert [v["resultado"] for v in vs] == ["NO CUMPLE", "NO CUMPLE", "CUMPLE"]
    assert any(v["codigo_generacion"].startswith(verdict.NO_DOC) for v in vs)
    assert "TOTAL: 3 documento(s)" in verdict.to_text(vs)


def test_verificar_end_to_end_with_fake_db(tmp_path, monkeypatch, capsys):
    """Un JSON sintético -> 'verificar' consulta la BD (falsa) por su UUID, compara y deja el veredicto en el .xlsx."""
    from contract_validator import cli
    from openpyxl import load_workbook
    doc = tmp_path / "doc.json"
    doc.write_text(json.dumps({"identificacion": {"tipoDte": "03", "codigoGeneracion": U, "numeroControl": "DTE-03-M001P001-000000000000001"},
                               "emisor": {"nit": "00000000000091"}}), encoding="utf-8")
    asked = {}

    def fake_fetch(uuids, db_config=None, **kw):
        asked["uuids"] = uuids
        return [(U, "CCF", "identificacion.tipoDte", "03"), (U, "CCF", "identificacion.numeroControl", "DTE-03-M001P001-000000000000009")]
    monkeypatch.setattr(db_fetch, "fetch", fake_fetch)
    out = tmp_path / "rep.xlsx"
    code = cli.main(["verificar", str(doc), "--out", str(out), "--companias", _cias(tmp_path)])
    text = capsys.readouterr().out
    assert asked["uuids"] == [U]
    assert "[CIA 91 (Prueba, esquema VENTASAAA)]" in text
    assert "VEREDICTO POR DOCUMENTO" in text and "Información vs base de datos   : INCORRECTA" in text
    assert code == 1
    wb = load_workbook(out)
    assert wb.sheetnames[0] == "Veredicto"
    row = [c.value for c in wb["Veredicto"][2]]
    assert row[0] == U and row[5] == "INCORRECTA" and row[-1] == "NO CUMPLE"


def test_verificar_without_db_still_validates_structure(tmp_path, monkeypatch, capsys):
    from contract_validator import cli

    def no_db(uuids, db_config=None, **kw):
        raise db_fetch.DbFetchError("Falta la contraseña: defina la variable de entorno CV_ORACLE_PASSWORD")
    monkeypatch.setattr(db_fetch, "fetch", no_db)
    doc = tmp_path / "doc.json"
    doc.write_text(json.dumps({"identificacion": {"tipoDte": "03", "codigoGeneracion": U}, "emisor": {"nit": "00000000000091"}}),
                   encoding="utf-8")
    code = cli.main(["verificar", str(doc), "--out", str(tmp_path / "r.xlsx"), "--companias", _cias(tmp_path)])
    text = capsys.readouterr().out
    assert "NO VERIFICADA" in text and "CV_ORACLE_PASSWORD" in text
    assert "Estructura vs contrato (.xlsx) : NO CUMPLE" in text   # el JSON sintético incompleto no cumple el contrato
    assert code == 1


def _cias(tmp_path):
    p = tmp_path / "companias.json"
    p.write_text(json.dumps({"companias": {"91": {"nombre": "Prueba", "esquema": "VENTASAAA", "nit": ["00000000000091"]}}}),
                 encoding="utf-8")
    return str(p)


def test_unknown_company_is_not_queried(tmp_path, monkeypatch, capsys):
    """Sin CIA identificable el documento no se consulta en ningún esquema y la BD queda NO VERIFICADA con el motivo."""
    from contract_validator import cli
    asked = {}
    monkeypatch.setattr(db_fetch, "fetch", lambda uuids, **kw: asked.setdefault("uuids", uuids) and [])
    doc = tmp_path / "doc.json"
    doc.write_text(json.dumps({"identificacion": {"tipoDte": "03", "codigoGeneracion": U}, "emisor": {"nit": "123"}}), encoding="utf-8")
    code = cli.main(["verificar", str(doc), "--out", str(tmp_path / "r.xlsx"), "--companias", _cias(tmp_path)])
    text = capsys.readouterr().out
    assert "uuids" not in asked
    assert "NO VERIFICADA" in text and "no se pudo identificar la compañía" in text
    assert code == 1


def test_db_status_per_document():
    [v] = verdict.build(_base(), _ids("a.json", "a.xml"), {U: "sin CIA"})
    assert (v["informacion_bd"], v["informacion_bd_detalle"]) == ("NO VERIFICADA", "sin CIA")
