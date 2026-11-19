"""Mecanismo de excepciones al libro (data/specs/ajustes_contrato.json, hoy sin entradas)."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.build_contract import (AJUSTES_SOURCE, AjusteError, COLUMNS, apply_ajustes, build,
                                               load_ajustes)

SOURCE = os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx")
REAL = os.path.join(ROOT, "data", "specs", "ajustes_contrato.json")


def _row(**kw):
    r = {c: "" for c in COLUMNS}
    r.update(tipo_doc="CCF", formato="XML", ruta="A/b", requerido="SI", cambio="SIN_CAMBIO", hoja_origen="H", fila_origen=5)
    r.update(kw)
    return r


def _aj(**kw):
    a = dict(id="X1", accion="modificar", tipo_doc="CCF", formato="XML", ruta="A/b", campos={"requerido": "SI_NULL"}, motivo="m")
    a.update(kw)
    return a


def test_modify_and_add():
    rows, detail = apply_ajustes([_row()], [_aj(), _aj(id="X2", accion="agregar", ruta="A/c", campos={"tipo": "string"})])
    mod, new = rows
    assert mod["requerido"] == "SI_NULL" and "ajuste X1" in mod["nota"]
    assert new["ruta"] == "A/c" and new["requerido"] == "NO" and new["hoja_origen"] == AJUSTES_SOURCE and new["fila_origen"] == "X2"
    assert detail[0][:8] == ["X1", "modificar", "CCF", "XML", "A/b", "requerido", "SI", "SI_NULL"]


def test_original_rows_not_mutated():
    base = [_row()]
    apply_ajustes(base, [_aj()])
    assert base[0]["requerido"] == "SI"


@pytest.mark.parametrize("bad, msg", [
    (_aj(ruta="no/existe"), "no existe"),
    (_aj(accion="agregar"), "ya existe"),
])
def test_invalid_targets(bad, msg):
    with pytest.raises(AjusteError, match=msg):
        apply_ajustes([_row()], [bad])


def test_load_validation(tmp_path):
    p = tmp_path / "a.json"
    for content, msg in [({"ajustes": [_aj(), _aj()]}, "único"),
                         ({"ajustes": [_aj(accion="borrar")]}, "accion"),
                         ({"ajustes": [_aj(motivo="")]}, "motivo"),
                         ({"ajustes": [_aj(campos={"fila_origen": 1})]}, "no ajustables")]:
        p.write_text(json.dumps(content), encoding="utf-8")
        with pytest.raises(AjusteError, match=msg):
            load_ajustes(str(p))
    assert load_ajustes(str(tmp_path / "no_existe.json")) == []


def test_real_adjustments_file_only_has_general_rules():
    from contract_validator.build_contract import load_reglas
    assert load_ajustes(REAL) == [] and sorted(load_reglas(REAL)) == ["catalogos_como_string", "decimales_equivalentes"]


def test_type_list_and_all_types():
    from contract_validator.build_contract import expand_ajustes
    rows = [_row(tipo_doc=t, formato="JSON") for t in ("CCF", "FAC", "NCR")] + [_row(tipo_doc="REM")]
    ex = expand_ajustes([_aj(tipo_doc="*", formato="JSON"), _aj(id="X2", tipo_doc=["CCF", "FAC"])], rows)
    assert [(a["id"], a["tipo_doc"]) for a in ex] == [("X1", "CCF"), ("X1", "FAC"), ("X1", "NCR"), ("X2", "CCF"), ("X2", "FAC")]
    out, _ = apply_ajustes(rows, [_aj(id="P1", accion="agregar", tipo_doc="*", formato="JSON", ruta="codigoEmpresa",
                                      campos={"requerido": "SI"})])
    assert sorted(r["tipo_doc"] for r in out if r["ruta"] == "codigoEmpresa") == ["CCF", "FAC", "NCR"]


def test_rule_catalogs_as_string(tmp_path):
    from contract_validator.build_contract import apply_reglas, load_reglas
    rows, detail = apply_reglas([_row(tipo="", valores="00|01"), _row(ruta="A/c", tipo="string", valores="USD"), _row(ruta="A/d")],
                                {"catalogos_como_string": {"activo": True, "motivo": "m"}})
    assert [(r["tipo"], r["valores"]) for r in rows] == [("string", ""), ("string", ""), ("", "")]
    assert "regla catalogos_como_string" in rows[0]["nota"] and len(detail) == 3  # valores x2 + tipo x1
    p = tmp_path / "a.json"
    p.write_text(json.dumps({"reglas_generales": {"inventada": {"activo": True}}}), encoding="utf-8")
    with pytest.raises(AjusteError, match="desconocida"):
        load_reglas(str(p))
    p.write_text(json.dumps({"reglas_generales": {"catalogos_como_string": {"activo": False}}}), encoding="utf-8")
    assert load_reglas(str(p)) == {}


def test_catalog_rule_also_marks_xml_counterpart():
    from contract_validator.build_contract import apply_reglas
    rows, _ = apply_reglas([_row(formato="JSON", ruta="a/b", tipo="", valores="1|2", ruta_contraparte="X/b"),
                            _row(formato="XML", ruta="X/b", tipo="", ruta_contraparte="a/b")],
                           {"catalogos_como_string": {"activo": True}})
    assert [(r["tipo"], r["valores"]) for r in rows] == [("string", ""), ("string", "")]


@pytest.mark.skipif(not os.path.exists(SOURCE), reason="falta el libro original")
def test_build_is_faithful_to_workbook(tmp_path):
    from contract_validator.audit_plano import audit
    from contract_validator.spec_loader import load_spec
    out = tmp_path / "plano.xlsx"
    s = build(SOURCE, str(out), REAL)
    assert s["ajustes"] == ["catalogos_como_string", "decimales_equivalentes"]
    rows, _ = audit(SOURCE, str(out), REAL)
    assert not [r for r in rows if r["nivel"] == "ERROR"]
    assert sum(r["codigo"] == "REGLA_GENERAL" for r in rows) == 2 and not [r for r in rows if r["codigo"] == "AJUSTE_APLICADO"]
    spec = load_spec(str(out))
    assert not [f for f in spec if f.values]                      # ningún campo con lista de valores
    amb = {f.fmt: f for f in spec if f.doc_type == "CCF" and f.path in ("identificacion.ambiente", "DatosEntrada.Encabezado.ambiente")}
    assert amb["JSON"].type == "string" and amb["XML"].type == "string"
    proveedor = ("codigoEmpresa", "firmaElectronica", "selloRecibido")
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM"):                    # hojas XML y JSON del libro
        got = {f.path: f for f in spec if f.doc_type == t and f.fmt == "JSON"}
        assert all(got[c].required and not got[c].nullable and got[c].type == "string" and not got[c].counterpart
                   for c in proveedor), t
    assert not [f for f in spec if f.path in proveedor and (f.fmt == "XML" or f.doc_type in ("INVALIDACION", "RETORNO"))]


@pytest.mark.skipif(not os.path.exists(SOURCE), reason="falta el libro original")
def test_mappings_follow_xml_sheet(tmp_path):
    from contract_validator.spec_loader import load_spec
    out = tmp_path / "plano.xlsx"
    build(SOURCE, str(out), REAL)
    spec = load_spec(str(out))
    cp = {(s.doc_type, s.fmt, s.path): s.counterpart for s in spec}
    paths = {(s.doc_type, s.path) for s in spec if s.fmt == "JSON"}
    for t in ("CCF", "FAC_EXP"):
        assert (t, "receptor.nombreComercial") in paths and (t, "recepto.nombreComercial") not in paths
    emi = "DatosEntrada.InformacioGeneral.InformacionEmisor."
    assert cp[("FAC_EXP", "JSON", "emisor.recintoFiscal")] == emi + "recintoFiscal"   # hoja XML, filas 47-48
    assert cp[("FAC_EXP", "JSON", "emisor.regimen")] == emi + "regimen"
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM"):
        assert cp[(t, "JSON", "resumen.montoTotalOperacion")] == "DatosEntrada.ImportesTotal.totalBaseImpuesto", t


