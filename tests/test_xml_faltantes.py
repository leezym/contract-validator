"""Tags XML del contrato que no llegan: ninguno debe pasar en silencio. Datos sintéticos."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from openpyxl import Workbook

from contract_validator.build_contract import build_rows
from contract_validator.comparator import compare
from contract_validator.explain import explain
from contract_validator.flatten import load_document
from contract_validator.models import Finding, Result, SpecField
from contract_validator.rules import check_xml_occurrences


def _doc(tmp_path, body):
    p = tmp_path / "d.xml"
    p.write_text('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body>'
                 "<EnviarDocumento><EnviarDocumentoEntrada><CabeceraEntrada><uuid>X</uuid></CabeceraEntrada>"
                 f"<DatosEntrada>{body}</DatosEntrada></EnviarDocumentoEntrada></EnviarDocumento>"
                 "</soapenv:Body></soapenv:Envelope>", encoding="utf-8")
    return load_document(str(p))


def _json(path, xml, k, required=True, nullable=False):
    return SpecField(path, required=required, nullable=nullable, doc_type="CCF", fmt="JSON", counterpart=xml,
                     note=f"posicional({k})", sheet="XML", row=10 + k)


def _xml(path, required=True, nullable=False, **kw):
    return SpecField(path, required=required, nullable=nullable, doc_type="CCF", fmt="XML", sheet="XML", row=5, **kw)


OT = "DatosEntrada.OtrosTotales.valorOtrosTotales"


def codes(out):
    return [(s, c, p) for s, c, p, _ in out]


def test_required_position_missing_or_empty(tmp_path):
    doc = _doc(tmp_path, "<OtrosTotales><valorOtrosTotales>1</valorOtrosTotales></OtrosTotales>"
                         "<OtrosTotales><valorOtrosTotales></valorOtrosTotales></OtrosTotales>")
    spec = [_json("a", OT, 1), _json("b", OT, 2), _json("c", OT, 3), _json("d", OT, 4, required=False)]
    assert codes(check_xml_occurrences(doc, "CCF", spec)) == [
        ("ERROR", "NULL_NOT_ALLOWED", OT + "[2]"), ("ERROR", "REQUIRED_MISSING", OT + "[3]"),
        ("WARNING", "XML_OPTIONAL_MISSING", OT + "[4]")]


def test_nullable_position_may_be_empty_but_must_exist(tmp_path):
    doc = _doc(tmp_path, "<OtrosTotales><valorOtrosTotales></valorOtrosTotales></OtrosTotales>")
    spec = [_json("a", OT, 1, nullable=True), _json("b", OT, 2, nullable=True)]
    assert codes(check_xml_occurrences(doc, "CCF", spec)) == [("ERROR", "REQUIRED_MISSING", OT + "[2]")]


def test_position_is_checked_in_every_item(tmp_path):
    base = "DatosEntrada.Items.TotalImpuestoItem.Impuestos.baseImponible"
    ok = "<TotalImpuestoItem><Impuestos><baseImponible>1</baseImponible></Impuestos></TotalImpuestoItem>"
    doc = _doc(tmp_path, f"<Items>{ok}{ok}</Items><Items>{ok}</Items>")
    out = check_xml_occurrences(doc, "CCF", [_json("a", base, 2)])
    assert codes(out) == [("ERROR", "REQUIRED_MISSING", base + "[2]")] and "ítems 2 de 2" in out[0][3]


def test_item_tag_must_come_in_all_items(tmp_path):
    doc = _doc(tmp_path, "<Items><codigoArticulo>1</codigoArticulo><descripcion>x</descripcion></Items>"
                         "<Items><descripcion></descripcion></Items>")
    spec = [_xml("DatosEntrada.Items.codigoArticulo"), _xml("DatosEntrada.Items.descripcion")]
    out = check_xml_occurrences(doc, "CCF", spec)
    assert codes(out) == [("ERROR", "REQUIRED_MISSING", "DatosEntrada.Items.codigoArticulo"),
                          ("ERROR", "NULL_NOT_ALLOWED", "DatosEntrada.Items.descripcion")]
    r = Result("d.xml", doc_type="CCF", fmt="XML")
    assert "en los ítems 2 de 2" in explain(Finding(*out[0]), r)


def test_optional_tag_missing_is_warning_except_removed_or_added_by_adjustment(tmp_path):
    doc = _doc(tmp_path, "<A>1</A>")
    spec = [_xml("DatosEntrada.A"), _xml("DatosEntrada.B", required=False, counterpart="b"),
            _xml("DatosEntrada.C", required=False, counterpart="c"),
            SpecField("DatosEntrada.D", required=False, doc_type="CCF", fmt="XML", sheet="ajustes_contrato.json ajuste A03"),
            SpecField("c", required=False, change="ELIMINADO", doc_type="CCF", fmt="JSON")]
    res = compare("d.xml", doc, spec, doc_type="CCF", strict_extra=False)
    assert [(f.severity, f.path) for f in res.findings if f.code == "XML_OPTIONAL_MISSING"] == [("WARNING", "DatosEntrada.B")]
    assert res.ok


def test_xml_tag_on_array_level_row_is_not_lost(tmp_path):
    wb = Workbook()
    x = wb.active
    x.title = "XML"
    x.append(["TAG_MIN_HACIENDA", "Descripcion", "Tipo", "XML_TQ", "Requerido", None, "CCF - 03"])
    x.append(["c/tributos", "T", "array", "x", None, None, "c/tributos"])
    x.append(["c/tributos/items", None, None, "DatosEntrada/Items/Total/tipoImpuesto", "Si (null)", None, "c/tributos/items"])
    src = tmp_path / "libro.xlsx"
    wb.save(src)
    rows, _ = build_rows(str(src))
    got = {(r["formato"], r["ruta"]): r for r in rows}
    assert got[("JSON", "c/tributos")]["ruta_contraparte"] == "DatosEntrada/Items/Total/tipoImpuesto"
    assert got[("XML", "DatosEntrada/Items/Total/tipoImpuesto")]["requerido"] == "SI_NULL"


def test_workbook_xml_paths_are_unified():
    """Las columnas de tags XML del libro ya usan la ruta completa y unificada."""
    src = os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx")
    _, lint = build_rows(src)
    assert not [m for *_, m in lint if "no unificada" in m]


def test_non_unified_xml_path_warns(tmp_path):
    wb = Workbook()
    x = wb.active
    x.title = "XML"
    x.append(["TAG_MIN_HACIENDA", "Descripcion", "Tipo", "XML_TQ", "Requerido", None, "CCF - 03"])
    x.append(["emisor/nit", "N", "string", "DatosEntrada/InformacionEmisor/nit", "Si", None, "emisor/nit"])
    src = tmp_path / "libro.xlsx"
    wb.save(src)
    rows, lint = build_rows(str(src))
    assert ("XML", "DatosEntrada/InformacioGeneral/InformacionEmisor/nit") in {(r["formato"], r["ruta"]) for r in rows}
    assert [m for *_, m in lint if "no unificada" in m]


def test_header_tags_are_required_xml_only_for_invoice_types(tmp_path):
    from contract_validator.spec_loader import load_spec
    from contract_validator.build_contract import build
    src = os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx")
    out = str(tmp_path / "plano.xlsx")
    build(src, out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(out)
    tags = {"pais", "compania", "idAplicacion", "fechaTransaccion", "host", "usuario", "password", "objetoNegocio", "operacion", "token"}
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM"):
        got = {s.path.split(".")[1]: s for s in spec if s.doc_type == t and s.path.startswith("CabeceraEntrada.")}
        assert tags <= set(got), t
        assert all(got[x].required and not got[x].nullable and got[x].description and not got[x].counterpart for x in tags), t
    assert not [s for s in spec if s.doc_type == "RETORNO" and s.path == "CabeceraEntrada.pais"]


def test_xml_only_marker_creates_no_json_row(tmp_path):
    wb = Workbook()
    x = wb.active
    x.title = "XML"
    x.append(["TAG_MIN_HACIENDA", "Descripcion", "Tipo", "XML_TQ", "Requerido", None, "CCF - 03"])
    x.append([None, "Valor fijo: DELTA", "string", "CabeceraEntrada/idAplicacion", "Si", None, "Solo XML"])
    src = tmp_path / "libro.xlsx"
    wb.save(src)
    rows, _ = build_rows(str(src))
    assert [(r["formato"], r["ruta"], r["requerido"], r["ruta_contraparte"]) for r in rows] == [
        ("XML", "CabeceraEntrada/idAplicacion", "SI", "")]


def test_tq_tags_are_required_xml_only_for_invoice_types(tmp_path):
    from contract_validator.spec_loader import load_spec
    from contract_validator.build_contract import build
    out = str(tmp_path / "plano.xlsx")
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(out)
    ig = "DatosEntrada.InformacioGeneral."
    for t in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM"):
        got = {s.path: s for s in spec if s.doc_type == t and s.fmt == "XML"}
        for p in ("InformacionEmisor.nitCompania", "codigoEstablecimientoTQ", "puntoEmisionTQ"):
            assert got[ig + p].required and not got[ig + p].nullable and got[ig + p].description, (t, p)
        assert got[ig + "codigoSRI"].required and got[ig + "codigoSRI"].nullable, t


def test_referencia_only_for_credit_and_debit_notes(tmp_path):
    from contract_validator.spec_loader import load_spec
    from contract_validator.build_contract import build
    out = str(tmp_path / "plano.xlsx")
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(out)
    for tag, campo, retorno in (("tipoDocumentoReferencia", "tipoDocumento", True), ("codigoReferencia", "tipoGeneracion", False),
                                ("numeroDocumentoReferencia", "numeroDocumento", False), ("fechaEmisionReferencia", "fechaEmision", True)):
        extra = {"RETORNO"} if retorno else set()
        assert {s.doc_type for s in spec if s.path == f"documentoRelacionado.{campo}"} == {"NCR", "NDB"} | extra, campo
        tipos = {s.doc_type for s in spec if s.path == f"DatosEntrada.Referencia.{tag}"} - {"RETORNO"}
        assert tipos == {"NCR", "NDB"}, tag


def load_document_json(tmp_path, data):
    import json
    p = tmp_path / "d.json"
    p.write_text(json.dumps(data), encoding="utf-8")
    return load_document(str(p))


def _book_with_tab(tmp_path):
    wb = Workbook()
    x = wb.active
    x.title = "XML"
    x.append(["TAG_MIN_HACIENDA", "Descripcion", "Tipo", "XML_TQ", "Requerido", None, "F -01"])
    x.append(["a/b", "B", "string", "DatosEntrada/E/b", "Si", None, "a/b"])
    x.append(["a/c", "C", "string", "DatosEntrada/E/c", "Si", None, "a/c"])
    x.append(["a/d", "D", "number", "DatosEntrada/O/v(1)", "Si", None, "a/d"])
    t = wb.create_sheet("Factura Electrónica")
    t.append(["TIPO", "REP", "TAG_MIN_HACIENDA", "Descripcion", "Tipo", "TAG_MIN_HACIENDA - (v3 → v4)", "XML_TQ"])
    t.append([1, "FAC", "a/b", "B", "string", "Eliminar a/b", "DatosEntrada/E/b"])
    t.append([1, "FAC", "a/d", "D", "number", "Eliminar a/d", "DatosEntrada/O/v(1)"])
    src = tmp_path / "libro.xlsx"
    wb.save(src)
    return str(src)


def test_eliminar_in_type_sheet_makes_xml_tag_optional(tmp_path):
    from contract_validator.build_contract import build
    from contract_validator.spec_loader import load_spec
    src, plano = _book_with_tab(tmp_path), str(tmp_path / "plano.xlsx")
    build(src, plano)
    spec = load_spec(plano)
    got = {(s.fmt, s.path): s for s in spec}
    assert not got[("XML", "DatosEntrada.E.b")].required and got[("XML", "DatosEntrada.E.c")].required
    assert got[("JSON", "a.b")].change == "ELIMINADO" and not got[("JSON", "a.b")].required
    j = load_document_json(tmp_path, {"identificacion": {"tipoDte": "01"}, "a": {"b": "x", "c": "y"}})
    assert ("REMOVED_STILL_PRESENT", "a.b") in {(f.code, f.path) for f in compare("d.json", j, spec, doc_type="FAC").findings}
    res = compare("d.xml", _doc(tmp_path, "<E><c>1</c></E>"), spec, doc_type="FAC", strict_extra=False)
    assert not [f for f in res.findings if "E.b" in f.path or "O.v" in f.path]               # ausentes: sin error ni aviso
    res = compare("d.xml", _doc(tmp_path, "<E><b>1</b><c>1</c></E><O><v>2</v></O>"), spec, doc_type="FAC")
    assert not [f for f in res.errors if f.path.startswith("DatosEntrada")]                 # presentes: válidos


def test_documento_relacionado_is_null_except_notes_and_return(tmp_path):
    from contract_validator.build_contract import build
    from contract_validator.spec_loader import load_spec
    out = str(tmp_path / "plano.xlsx")
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(out)
    got = {s.doc_type: s for s in spec if s.fmt == "JSON" and s.path == "documentoRelacionado"}
    for t in ("CCF", "FAC", "FAC_EXP", "REM"):
        assert got[t].required and got[t].nullable and "solo_null" in got[t].note, t
    for t in ("NCR", "NDB", "RETORNO"):
        assert "solo_null" not in got[t].note, t
    base = {"identificacion": {"tipoDte": "03"}}
    run = lambda d: {(f.code, f.path) for f in compare("d.json", load_document_json(tmp_path, d), spec, doc_type="CCF").findings}
    assert ("REQUIRED_MISSING", "documentoRelacionado") in run(base)
    assert not [c for c in run({**base, "documentoRelacionado": None}) if c[1] == "documentoRelacionado"]
    assert ("VALUE_NOT_NULL", "documentoRelacionado") in run({**base, "documentoRelacionado": [{"x": 1}]})


def test_adicionar_marks_are_consistent_between_sheets(tmp_path):
    from contract_validator.build_contract import build
    from contract_validator.spec_loader import load_spec
    out = str(tmp_path / "plano.xlsx")
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    got = {(s.doc_type, s.path): s for s in load_spec(out) if s.fmt == "JSON"}
    assert got[("FAC_EXP", "cuerpoDocumento.codTributo")].change == "NUEVO"
    for t in ("NCR", "NDB"):
        assert got[(t, "ventaTercero.codDomiciliado")].change == "NUEVO", t


def test_receptor_municipio_tag_is_xml_only_for_fe(tmp_path):
    from contract_validator.build_contract import build
    from contract_validator.spec_loader import load_spec
    out = str(tmp_path / "plano.xlsx")
    build(os.path.join(ROOT, "data", "specs", "contrato_ejemplo.xlsx"), out, os.path.join(ROOT, "data", "specs", "ajustes_contrato.json"))
    spec = load_spec(out)
    tag = "DatosEntrada.InformacioGeneral.InformacionAdquiriente.codigoMunicipio"
    fe = [s for s in spec if s.doc_type == "FAC_EXP" and s.path == tag]
    assert len(fe) == 1 and fe[0].fmt == "XML" and not fe[0].required and not fe[0].counterpart
    assert not [s for s in spec if s.doc_type == "FAC_EXP" and s.path == "receptor.direccion.distrito"]
