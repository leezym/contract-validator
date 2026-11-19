"""Mecanismo 'apendice_obligatorio' (hoy inactivo). Datos sintéticos."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.explain import explain
from contract_validator.flatten import load_document
from contract_validator.models import Finding, Result
from contract_validator.rules import apendice_names, check_apendice

REGLA = {"activo": True, "tipos": ["CCF", "FAC"], "obligatorios": ["apendice01", "apendice02", "apendice03", "apendice04"]}


def _json(tmp_path, campos):
    p = tmp_path / "d.json"
    p.write_text(json.dumps({"identificacion": {"tipoDte": "03"},
                             "apendice": [{"campo": c, "etiqueta": "E:", "valor": f"v{c[-2:]}"} for c in campos]}), encoding="utf-8")
    return load_document(str(p))


def _xml(tmp_path, campos, key="etiqueta"):
    notas = "".join(f"<Notas><{key}>{c}</{key}><notasAdicionales>v{c[-2:]}</notasAdicionales></Notas>" for c in campos)
    p = tmp_path / "d.xml"
    p.write_text('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body>'
                 "<EnviarDocumento><EnviarDocumentoEntrada><CabeceraEntrada><uuid>X</uuid></CabeceraEntrada><DatosEntrada>"
                 f"<Notas><notasAdicionales>.</notasAdicionales></Notas>{notas}</DatosEntrada>"
                 "</EnviarDocumentoEntrada></EnviarDocumento></soapenv:Body></soapenv:Envelope>", encoding="utf-8")
    return load_document(str(p))


def codes(out):
    return [(s, c, p) for s, c, p, _ in out]


def test_complete_json_and_xml_pass(tmp_path):
    full = ["apendice01", "apendice02", "apendice03", "apendice04"]
    assert check_apendice(_json(tmp_path, full), "CCF", REGLA) == []
    assert check_apendice(_xml(tmp_path, full), "CCF", REGLA) == []


def test_missing_is_error_and_extra_is_warning(tmp_path):
    out = check_apendice(_json(tmp_path, ["apendice02", "apendice03", "apendice04", "apendice05"]), "CCF", REGLA)
    assert codes(out) == [("ERROR", "APENDICE_MISSING", "apendice[apendice01]"),
                          ("WARNING", "APENDICE_EXTRA", "apendice[apendice05]")]
    out = check_apendice(_xml(tmp_path, ["apendice01", "apendice02", "apendice03", "apendice04", "apendice06"]), "CCF", REGLA)
    assert codes(out) == [("WARNING", "APENDICE_EXTRA", "DatosEntrada.Notas[apendice06]")]


def test_xml_key_follows_contract(tmp_path):
    doc = _xml(tmp_path, ["apendice01"])
    assert [n for n, _ in apendice_names(doc, "DatosEntrada.Notas.etiqueta")] == ["apendice01"]
    assert apendice_names(doc, "DatosEntrada.Notas.campo") == []          # el <Notas> sin campo (".") no cuenta


def test_scope_and_inactive(tmp_path):
    doc = _json(tmp_path, [])
    assert check_apendice(doc, "RETORNO", REGLA) == []                    # tipo fuera de la regla
    assert check_apendice(doc, "CCF", None) == []                         # regla inactiva
    assert len(check_apendice(doc, "CCF", REGLA)) == 4                    # sin apéndice: faltan los 4


def test_explanations():
    r = Result("d.json", doc_type="CCF", fmt="JSON")
    m = explain(Finding("ERROR", "APENDICE_MISSING", "apendice[apendice01]",
                        "Falta el apéndice obligatorio 'apendice01' (obligatorios: apendice01, apendice02, apendice03, apendice04)"), r)
    assert m.startswith("Falta el apéndice obligatorio «apendice01» en el JSON")
    w = explain(Finding("WARNING", "APENDICE_EXTRA", "apendice[apendice05]",
                        "Apéndice adicional 'apendice05' con valor 'OC123': no está entre los obligatorios; revisar manualmente"), r)
    assert "«apendice05» (valor «OC123»)" in w and "revíselo manualmente" in w


@pytest.mark.skipif(not os.path.exists(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx")), reason="falta el libro")
def test_mapping_follows_workbook_and_rule_is_inactive(tmp_path):
    from contract_validator.build_contract import build, load_reglas
    from contract_validator.spec_loader import load_spec
    real = os.path.join(ROOT, "data", "specs", "ajustes_contrato.json")
    assert "apendice_obligatorio" not in load_reglas(real)
    out = tmp_path / "plano.xlsx"
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), str(out), real)
    spec = load_spec(str(out))
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM", "RETORNO"):
        cp = {(s.fmt, s.path): s.counterpart for s in spec if s.doc_type == t}
        assert cp[("JSON", "apendice.campo")] == "DatosEntrada.Notas.etiqueta", t
        assert cp[("JSON", "apendice.etiqueta")] == "DatosEntrada.Notas.campo", t
        assert cp[("XML", "DatosEntrada.Notas.etiqueta")] == "apendice.campo", t




def test_cross_check_compares_whole_apendice(tmp_path):
    from contract_validator.crosscheck import cross_check
    from contract_validator.models import SpecField
    spec = [SpecField("apendice.valor", doc_type="CCF", fmt="JSON", counterpart="DatosEntrada.Notas.notasAdicionales")]
    xml = _xml(tmp_path, ["apendice01", "apendice02"], key="campo")
    p = tmp_path / "j.json"
    p.write_text(json.dumps({"identificacion": {"tipoDte": "03"}, "resumen": {"observaciones": "v01"},
                             "apendice": [{"campo": "apendice02", "etiqueta": None, "valor": "v02"}]}), encoding="utf-8")
    assert [(f.code, f.path) for f in cross_check("x", xml, "j", load_document(str(p)), spec, "CCF").findings] == [
        ("XREF_MISSING", "apendice[apendice01]")]   # aunque resumen.observaciones traiga sus datos
    p.write_text(json.dumps({"identificacion": {"tipoDte": "03"}, "resumen": {"observaciones": "otro"},
                             "apendice": [{"campo": "apendice02", "etiqueta": None, "valor": "zz"}]}), encoding="utf-8")
    got = {(f.code, f.path) for f in cross_check("x", xml, "j", load_document(str(p)), spec, "CCF").findings}
    assert got == {("XREF_MISSING", "apendice[apendice01]"), ("XREF_MISMATCH", "apendice[apendice02]")}


ATTR_SPEC = None


def _attr_spec():
    from contract_validator.models import SpecField
    return [SpecField(f"apendice.{a}", doc_type="CCF", fmt="JSON") for a in ("campo", "etiqueta", "valor")]


def test_each_apendice_needs_its_three_attributes_in_json(tmp_path):
    from contract_validator.rules import check_apendice_attributes
    p = tmp_path / "d.json"
    p.write_text(json.dumps({"identificacion": {"tipoDte": "03"}, "apendice": [
        {"campo": "apendice01", "etiqueta": "E:", "valor": "1"}, {"campo": "apendice02", "valor": ""}]}), encoding="utf-8")
    assert codes(check_apendice_attributes(load_document(str(p)), "CCF", _attr_spec())) == [
        ("ERROR", "REQUIRED_MISSING", "apendice[apendice02].etiqueta"), ("ERROR", "NULL_NOT_ALLOWED", "apendice[apendice02].valor")]


def test_each_apendice_needs_its_three_attributes_in_xml(tmp_path):
    from contract_validator.rules import check_apendice_attributes
    p = tmp_path / "d.xml"
    p.write_text('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body><E><F><CabeceraEntrada><uuid>X</uuid>'
                 "</CabeceraEntrada><DatosEntrada><Notas><notasAdicionales>.</notasAdicionales></Notas>"
                 "<Notas><campo>apendice01</campo><notasAdicionales>1</notasAdicionales></Notas></DatosEntrada></F></E>"
                 "</soapenv:Body></soapenv:Envelope>", encoding="utf-8")
    out = check_apendice_attributes(load_document(str(p)), "CCF", _attr_spec())
    assert codes(out) == [("ERROR", "REQUIRED_MISSING", "apendice[apendice01].etiqueta")]
    assert "«etiqueta» no llegó" in explain(Finding(*out[0]), Result("d.xml", doc_type="CCF", fmt="XML"))


def test_workbook_requires_apendice_attributes(tmp_path):
    from contract_validator.build_contract import build
    from contract_validator.spec_loader import load_spec
    out = tmp_path / "plano.xlsx"
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), str(out), os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(str(out))
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM"):
        got = {(s.fmt, s.path): s for s in spec if s.doc_type == t}
        for k in (("JSON", "apendice"), ("JSON", "apendice.campo"), ("JSON", "apendice.etiqueta"), ("JSON", "apendice.valor"),
                  ("XML", "DatosEntrada.Notas.campo"), ("XML", "DatosEntrada.Notas.etiqueta")):
            assert got[k].required and not got[k].nullable, (t, k)
