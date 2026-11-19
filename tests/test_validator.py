import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import pytest
from make_sample import build
from contract_validator.spec_loader import load_spec
from contract_validator.flatten import load_document, flatten_json
from contract_validator.comparator import compare


@pytest.fixture(scope="module")
def env(tmp_path_factory):
    base = str(tmp_path_factory.mktemp("proj"))
    build(base)
    return base, load_spec(f"{base}/data/specs/contrato_simple.xlsx")


def codes(env, name):
    base, spec = env
    res = compare(name, load_document(f"{base}/data/inputs/{name}"), spec)
    return res, {(f.code, f.path) for f in res.findings}


@pytest.mark.parametrize("name", ["ok.json", "ok.xml"])
def test_valid_files_pass(env, name):
    res, found = codes(env, name)
    assert res.ok, found


def test_bad_json_detects_all_change_types(env):
    res, found = codes(env, "bad.json")
    assert not res.ok
    assert ("NEW_MISSING", "cliente.email") in found
    assert ("REMOVED_STILL_PRESENT", "cliente.telefono") in found
    assert ("OLD_NAME_PRESENT", "cliente.nombre") in found
    assert ("RENAMED_MISSING", "cliente.nombreCompleto") in found
    assert ("UNEXPECTED_FIELD", "cliente.extra") in found


def test_bad_xml(env):
    _, found = codes(env, "bad.xml")
    assert ("NEW_MISSING", "cliente.email") in found
    assert ("REMOVED_STILL_PRESENT", "cliente.telefono") in found
    assert ("OLD_NAME_PRESENT", "cliente.nombre") in found
    assert ("REQUIRED_MISSING", "cliente.saldo") in found


# ---------------------------------------------------------------- contrato real (hoja plana generada)
import glob
from contract_validator.build_contract import build as build_flat, parse_change, parse_type, parse_required, xml_path, clean_tag
from contract_validator.flatten import doc_type_from_filename, decode_bytes

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPECS = os.path.join(ROOT, "data", "specs")
SOURCE = os.path.join(SPECS, "contrato_ejemplo.xlsx")
XML_CCF = (glob.glob(os.path.join(SPECS, "53_16_01_001_131126_020656_*.xml")) or [None])[0]  # XML de CCF entregado por el usuario
needs_xml = pytest.mark.skipif(XML_CCF is None, reason="falta el XML de CCF en data/specs")
needs_source = pytest.mark.skipif(not os.path.exists(SOURCE), reason="falta data/specs/contrato_ejemplo.xlsx")


def test_parse_change():
    assert parse_change("Adicionar   emisor/direccion/distrito") == ("NUEVO", "emisor/direccion/distrito", "")
    assert parse_change("Eliminar emisor/tipoEstablecimiento")[0] == "ELIMINADO"
    assert parse_change("Elimnar extension/observaciones")[0] == "ELIMINADO"  # typo del libro
    assert parse_change("Renombrar   resumen/ivaPerci1   → resumen/ivaPerci") == ("RENOMBRADO", "resumen/ivaPerci", "resumen/ivaPerci1")
    assert parse_change("Renombrar receptor/nit -> receptor/numDocumento")[1:] == ("receptor/numDocumento", "receptor/nit")
    assert parse_change("documentoRelacionado/items/tipoDocumento") == ("SIN_CAMBIO", "documentoRelacionado/tipoDocumento", "")
    assert parse_change("resumen/pagos/items/codigo \nNormativa 154")[1] == "resumen/pagos/codigo"


def test_parse_type_required_and_xml_path():
    assert parse_type("00, 01") == ("", ["00", "01"], None)
    assert parse_type("date")[0] == "string" and parse_type("Integer")[0] == "integer"
    assert parse_type("USD") == ("string", ["USD"], None)
    assert parse_required("Si (null)", "NDB") == "SI_NULL"
    assert parse_required("Si/null/ por tipoDTE", "NDB") == "SI_NULL"  # el libro admite null
    assert parse_required("Si para FEX", "NDB") == "NO" and parse_required("Si para FEX", "FAC_EXP") == "SI"
    assert parse_required("N/A", "NDB") == "NO" and parse_required(None, "NDB") is None
    assert xml_path("DatosEntrada/Notas(3….N)") == "DatosEntrada/Notas"
    assert xml_path("DatosEntrada/Notas/notasAdicionales[2]") == "DatosEntrada/Notas/notasAdicionales"
    assert xml_path("ImportesTotal/totalFactura") == "DatosEntrada/ImportesTotal/totalFactura"
    assert xml_path("x") == "" and xml_path("NULL (NO LO ENVIAMOS)") == ""


def test_doc_type_from_filename_and_decoding():
    assert doc_type_from_filename("a/pbs-notas-debito.json") == "NDB"
    assert doc_type_from_filename("pbs-notas-credito.json") == "NCR"
    assert doc_type_from_filename("pbs-comprobante-credito-fiscal.json") == "CCF"
    assert doc_type_from_filename("pbs-factura-exportacion.json") == "FAC_EXP"
    text, warn = decode_bytes(b"D\xd3LAR")
    assert text == "DÓLAR" and warn
    assert decode_bytes("DÓLAR".encode())[1] is None


@pytest.fixture(scope="module")
def real(tmp_path_factory):
    if not os.path.exists(SOURCE):
        pytest.skip("falta contrato_ejemplo.xlsx")
    out = str(tmp_path_factory.mktemp("plano") / "contrato_plano.xlsx")
    summary = build_flat(SOURCE, out)
    return out, summary, load_spec(out)


@needs_source
def test_build_covers_every_type_and_format(real):
    _, summary, _ = real
    assert summary["lint_levels"]["ERROR"] == 0
    for code in ("CCF", "FAC", "FAC_EXP", "NCR", "NDB", "REM", "INVALIDACION", "RETORNO"):
        for fmt in ("JSON", "XML"):
            assert summary["by_type"].get((code, fmt), 0) > 20, (code, fmt)


@needs_source
def test_ndb_changes_follow_xml_sheet(real):
    """NDB: ivaPerci1 sin renombrar (hoja XML), tipoEstablecimiento eliminado (hoja del tipo), distrito nuevo."""
    _, _, spec = real
    doc = load_document(os.path.join(SPECS, "pbs-notas-debito.json"))
    assert doc.doc_type == "NDB"
    sel = {(s.change, s.path) for s in spec if s.doc_type == "NDB" and s.fmt == "JSON"}
    assert sel >= {("SIN_CAMBIO", "resumen.ivaPerci1"), ("ELIMINADO", "emisor.tipoEstablecimiento"),
                   ("NUEVO", "emisor.direccion.distrito"), ("NUEVO", "resumen.totalPagar")}


@needs_source
def test_ccf_renames_and_removals(real):
    """CCF sí marca cambios en la hoja 'XML': Eliminar codEstableMH y Renombrar ivaPerci1 -> ivaPerci."""
    _, _, spec = real
    sel = {(s.change, s.path, s.old_path) for s in spec if s.doc_type == "CCF" and s.fmt == "JSON"}
    assert ("ELIMINADO", "emisor.codEstableMH", "") in sel
    assert ("RENOMBRADO", "resumen.ivaPerci", "resumen.ivaPerci1") in sel
    doc = load_document(os.path.join(SPECS, "pbs-comprobante-credito-fiscal.json"))
    doc["emisor.codEstableMH"] = {"string"}
    doc["resumen.ivaPerci1"] = {"number"}
    del doc["resumen.ivaPerci"]
    found = {(f.code, f.path) for f in compare("x", doc, spec).findings}
    assert ("REMOVED_STILL_PRESENT", "emisor.codEstableMH") in found
    assert ("OLD_NAME_PRESENT", "resumen.ivaPerci1") in found
    assert ("RENAMED_MISSING", "resumen.ivaPerci") in found


@needs_source
def test_null_required_and_unknown_doc_type(real):
    _, _, spec = real
    doc = load_document(os.path.join(SPECS, "pbs-notas-debito.json"))
    doc["identificacion.version"] = {"null"}
    assert ("NULL_NOT_ALLOWED", "identificacion.version") in {(f.code, f.path) for f in compare("x", doc, spec).findings}
    doc.doc_type = ""
    assert [f.code for f in compare("x", doc, spec).findings] == ["DOCTYPE_UNKNOWN"]


@needs_source
@needs_xml
def test_xml_tq_is_parsed_and_checked(real):
    _, _, spec = real
    doc = load_document(XML_CCF)
    assert doc.fmt == "XML" and doc.doc_type == "CCF"          # tipoFactura=03
    assert "DatosEntrada.Encabezado.versionUBL" in doc and not any("Envelope" in p for p in doc)
    assert "null" in doc["DatosEntrada.Items.TotalImpuestoItem.Impuestos.tipoImpuesto"]  # 'null' = valor nulo
    res = compare(XML_CCF, doc, spec)
    found = {(f.code, f.path) for f in res.findings}
    # los tags de CabeceraEntrada están definidos en el libro (hoja XML, filas 216-225)
    assert not [p for c, p in found if c == "UNEXPECTED_FIELD" and p.startswith("CabeceraEntrada")]
    # InformacioGeneral (sin 'n') es el tag correcto: ningun contenedor de InformacioGeneral debe faltar
    assert not [f for f in res.errors if f.code == "REQUIRED_MISSING" and "InformacioGeneral" in f.path]
    assert ("UNEXPECTED_FIELD", "DatosEntrada.InformacioGeneral.codigoSRI") not in found   # hoja XML, fila 228


def test_catalog_field_accepts_text_or_number():
    from contract_validator.models import SpecField
    nota = "regla catalogos_como_string"
    spec = [SpecField(path=p, type="string", required=True, note=nota) for p in ("a", "b", "c")]
    assert compare("x", flatten_json({"a": "1", "b": 1, "c": 1.5}), spec).ok
    res = compare("x", flatten_json({"a": True, "b": {"z": 1}, "c": "1"}), spec, strict_extra=False)
    assert {f.path for f in res.errors if f.code == "TYPE_MISMATCH"} == {"a", "b"}


def test_types_follow_workbook():
    from contract_validator.models import SpecField
    spec = [SpecField(path="a", type="string", required=True), SpecField(path="b", type="number", required=True),
            SpecField(path="c", type="integer", required=True, nullable=True)]
    res = compare("x", flatten_json({"a": 5, "b": 5, "c": None}), spec)
    assert [(f.code, f.path) for f in res.errors] == [("TYPE_MISMATCH", "a")]   # un entero es válido como number


# ---------------------------------------------------------------- cruce XML <-> JSON
import copy
from contract_validator.crosscheck import cross_check, pair_documents, json_values, xml_values
from decimal import Decimal


def _twin_json(xml_doc, spec, template_raw, doc_type):
    """Construye el JSON 'equivalente' a un XML usando el mapeo del contrato (ida y vuelta, sin tocar el Excel)."""
    xv = xml_values(xml_doc.raw)
    arrays = {k for k, v in template_raw.items() if isinstance(v, list)}
    out: dict = {}
    for s in spec:
        if s.doc_type != doc_type or s.fmt != "JSON" or not s.counterpart or s.change == "ELIMINADO":
            continue
        if "posicional" in s.note or s.counterpart not in xv:
            continue
        parts, vals = s.path.split("."), xv[s.counterpart]
        if (parts[0] in arrays or len(vals) > 1) and len(parts) == 2:
            rows = out.setdefault(parts[0], [])
            rows.extend({} for _ in range(len(vals) - len(rows)))
            for r, v in zip(rows, vals):
                r[parts[1]] = v
        elif parts[0] not in arrays and not any(isinstance(template_raw.get(parts[0]), list) for _ in [0]):
            node = out
            for k in parts[:-1]:
                node = node.setdefault(k, {})
            node[parts[-1]] = vals[0]
    return out


@pytest.fixture(scope="module")
def pair(real):
    if XML_CCF is None:
        pytest.skip("falta el XML de CCF en data/specs")
    _, _, spec = real
    xml_doc = load_document(XML_CCF)
    tmpl = load_document(os.path.join(SPECS, "pbs-comprobante-credito-fiscal.json"))
    twin = _twin_json(xml_doc, spec, tmpl.raw, "CCF")
    twin["identificacion"]["tipoDte"] = "03"
    from contract_validator.crosscheck import apendice_records
    twin["apendice"] = [{"campo": n, **r} for n, r in apendice_records(xml_doc).items()]  # el apéndice se cruza completo
    jd = flatten_json(twin)
    jd.fmt, jd.raw, jd.doc_type = "JSON", twin, "CCF"
    return spec, xml_doc, jd


@needs_source
def test_cross_check_clean_when_json_mirrors_xml(pair):
    spec, xml_doc, jd = pair
    res = cross_check("a.xml", xml_doc, "a.json", jd, spec)
    assert res.checked > 30 and not res.findings, [(f.code, f.path, f.message) for f in res.findings]


@needs_source
def test_cross_check_detects_mismatch_missing_count_and_doctype(pair):
    spec, xml_doc, jd = pair
    j = copy.deepcopy(jd)
    j.raw["emisor"]["nit"] = "99999999999999"
    del j.raw["receptor"]["correo"]
    j.raw["cuerpoDocumento"] = j.raw["cuerpoDocumento"][:2]
    j.raw["resumen"]["totalPagar"] = "1.00"
    found = {(f.code, f.path) for f in cross_check("a.xml", xml_doc, "a.json", j, spec).findings}
    assert ("XREF_MISMATCH", "emisor.nit") in found
    assert ("XREF_MISSING", "receptor.correo") in found
    assert any(c == "XREF_COUNT_MISMATCH" for c, _ in found)
    assert ("XREF_MISMATCH", "resumen.totalPagar") in found
    j2 = copy.deepcopy(jd)
    j2.doc_type = "NDB"
    assert [f.code for f in cross_check("a.xml", xml_doc, "a.json", j2, spec).findings] == ["XREF_DOCTYPE"]


@needs_source
def test_pairing_by_uuid(pair):
    _, xml_doc, jd = pair
    other = load_document(os.path.join(SPECS, "pbs-notas-debito.json"))
    pairs, lone = pair_documents({"x.xml": xml_doc, "x.json": jd, "other.json": other})
    assert pairs == [("x.xml", "x.json")] and lone == ["other.json"]


# ---------------------------------------------------------------- audit-mapping
from contract_validator import audit_mapping
from contract_validator.models import SpecField


def _sf(path, fmt="JSON", cp="", req=True, nullable=False, t="string", note="", dt="FAC"):
    return SpecField(path=path, type=t, required=req, nullable=nullable, doc_type=dt, fmt=fmt, counterpart=cp, note=note, sheet="H", row=1)


def test_audit_detects_each_issue_on_synthetic_spec():
    spec = [
        _sf("a.sinOrigen"), _sf("a.conNull", nullable=True),
        _sf("a.total", cp="X.total"), _sf("a.pago", cp="X.total"),
        _sf("n.campo", cp="N.etiqueta"), _sf("n.etiqueta", cp="N.campo"),
        _sf("a.pos", cp="X.pos", note="posicional"),
        _sf("a.dif", cp="X.uno"), _sf("a.dif", cp="X.dos", dt="NCR"),
        _sf("X.total", fmt="XML", dt="FAC"), _sf("X.huerfano", fmt="XML", dt="FAC"),
        _sf("CabeceraEntrada.uuid", fmt="XML", dt="FAC"),
    ]
    rows = audit_mapping.audit(spec)
    got = {(r["codigo"], r["nivel"]) for r in rows}
    assert ("NO_XML_SOURCE", "ERROR") in got and ("NO_XML_SOURCE", "WARNING") in got
    assert ("MULTI_SOURCE", "WARNING") in got and "CROSSED_MAPPING" not in {c for c, _ in got}
    assert ("POSITIONAL", "INFO") in got and ("DIFF_ACROSS_TYPES", "WARNING") in got
    assert [r["ruta_xml"] for r in rows if r["codigo"] == "XML_NO_DEST"] == ["X.huerfano"]  # uuid de Cabecera se ignora
    assert audit_mapping.audit(spec, "NCR") != rows


@needs_source
def test_audit_cli_on_real_contract(real, tmp_path):
    from contract_validator.cli import main
    plano, _, _ = real
    out = str(tmp_path / "aud.xlsx")
    assert main(["audit-mapping", "--spec", plano, "--out", out]) == 0
    from openpyxl import load_workbook
    wb = load_workbook(out)
    assert wb.sheetnames == ["Resumen", "Hallazgos"] and wb["Hallazgos"].max_row > 5
    assert main(["audit-mapping", "--spec", plano, "--out", out, "--doc-type", "FAC"]) == 0
    rows = audit_mapping.audit(load_spec(plano), "FAC")
    assert {"NO_XML_SOURCE", "MULTI_SOURCE", "POSITIONAL"} <= {r["codigo"] for r in rows}





# ---------------------------------------------------------------- politica: solo lo que definen los specs
@needs_source
def test_only_spec_defined_rules_run(real):
    _, _, spec = real
    doc = load_document(os.path.join(SPECS, "pbs-notas-debito.json"))
    raw = copy.deepcopy(doc.raw)
    # nada de esto esta definido por los specs: no debe producir hallazgos de contenido
    raw["identificacion"].update(fecEmi="no-es-fecha", horEmi="25:61", codigoGeneracion="abc")
    raw["receptor"].update(numDocumento="1", telefono="x", correo="sin-arroba")
    raw["resumen"].update(totalGravada=999.0, totalPagar=1.0)
    doc.raw = raw
    codes_found = {f.code for f in compare("x", doc, spec).findings}
    assert not codes_found & {"FORMAT_INVALID", "CONSISTENCY_ERROR", "TEXT_CORRUPT", "TYPE_MISMATCH", "STRUCTURE_MISMATCH"}
    # lo unico definido: numero de control DTE-##-<bloque>-<15 digitos> con ## = tipo de documento
    raw["identificacion"]["numeroControl"] = "DTE-03-M001P001-000000000001116"
    assert ("CONSISTENCY_ERROR", "identificacion.numeroControl") in {(f.code, f.path) for f in compare("x", doc, spec).findings}
    raw["identificacion"]["numeroControl"] = "DTE-06-M001P001-123"
    assert ("FORMAT_INVALID", "identificacion.numeroControl") in {(f.code, f.path) for f in compare("x", doc, spec).findings}
    raw["identificacion"]["numeroControl"] = "DTE-06-00101-000000500000002"  # bloque central de 5 (ejemplo del Excel): valido
    assert not [f for f in compare("x", doc, spec).findings if f.path == "identificacion.numeroControl"]


@needs_source
def test_goldens_have_no_content_findings(real):
    content = {"FORMAT_INVALID", "CONSISTENCY_ERROR", "TEXT_CORRUPT", "STRUCTURE_MISMATCH"}
    spec = real[2]
    for name in ("pbs-comprobante-credito-fiscal", "pbs-factura-electronica", "pbs-factura-exportacion",
                 "pbs-notas-credito", "pbs-notas-debito", "pbs-notas-remision"):
        res = compare(name, load_document(os.path.join(SPECS, name + ".json")), spec)
        assert not {f.code for f in res.errors} & content, (name, [(f.code, f.path) for f in res.errors])


def test_cross_check_is_strict_but_reports_format_only_differences_as_warning():
    from contract_validator.crosscheck import strict, loose
    assert strict("null") is None and strict("") is None and strict(None) is None
    assert strict(2260.0) == "2260.0" and strict("2260.00") == "2260.00"   # no se equiparan: igualdad exacta
    assert loose(2260.0) == loose("2260.00") and loose("05") == loose(5) and loose("12/11/2026") == "2026-11-12"
    assert loose("abc") != loose("abd")


@needs_source
def test_cross_check_format_diff_is_warning_and_value_diff_is_error(pair):
    spec, xml_doc, jd = pair
    j = copy.deepcopy(jd)
    total = xml_doc.raw.find("DatosEntrada/ImportesTotal/totalFactura").text.strip()   # p.ej. 1796.13
    j.raw["resumen"]["totalPagar"] = total + "0"      # 1796.130: mismo numero, distinto formato
    j.raw["emisor"]["nit"] = "05010912191014"        # valor distinto
    found = {(f.severity, f.code, f.path) for f in cross_check("a.xml", xml_doc, "a.json", j, spec).findings}
    assert ("WARNING", "XREF_FORMAT_DIFF", "resumen.totalPagar") in found
    assert ("ERROR", "XREF_MISMATCH", "emisor.nit") in found
    assert not [f for f in found if f[1] == "XREF_MISMATCH" and f[2] == "resumen.totalPagar"]
    # regla 'decimales_equivalentes': mismo valor decimal = igual (sin hallazgo); el resto sigue igual
    found = {(f.severity, f.code, f.path) for f in cross_check("a.xml", xml_doc, "a.json", j, spec,
                                                               reglas={"decimales_equivalentes": {"activo": True}}).findings}
    assert not [f for f in found if f[2] == "resumen.totalPagar"]
    assert ("ERROR", "XREF_MISMATCH", "emisor.nit") in found


def test_same_decimal():
    from contract_validator.crosscheck import same_decimal
    assert same_decimal("285.6", "285.60") and same_decimal("2260", "2260.00") and same_decimal("-1.5", "-1.50")
    assert not same_decimal("05", "5")            # código con cero a la izquierda: es formato, no decimal
    assert not same_decimal("285.6", "285.61") and not same_decimal("12/11/2026", "2026-11-12")
    assert not same_decimal(None, "1") and not same_decimal("abc", "abc")


@needs_source
def test_si_null_por_tipo_dte_is_required_and_nullable(real):
    _, summary, spec = real
    row = next(x for x in spec if x.doc_type == "NDB" and x.fmt == "JSON" and x.path == "cuerpoDocumento.codTributo")
    assert row.required and row.nullable
    row = next(x for x in spec if x.doc_type == "NDB" and x.fmt == "JSON" and x.path == "identificacion.tipoContingencia")
    assert row.required and row.nullable  # 'Si (null)' si admite null
    import openpyxl
    lint = [r[3] for r in openpyxl.load_workbook(real[0], read_only=True)["lint"].iter_rows(values_only=True)]
    assert not [m for m in lint if m and "ambigua" in m and "por tipodte" in m]


@needs_source
def test_requirement_only_applies_to_types_marked_in_xml_sheet(real):
    """cuerpoDocumento.codTributo: en FE - 11 es un campo nuevo ('Adicionar'), en los demás ya existía."""
    _, _, spec = real
    def row(t):
        return next((x for x in spec if x.doc_type == t and x.fmt == "JSON" and x.path == "cuerpoDocumento.codTributo"), None)
    assert row("NDB").required and row("NDB").nullable
    assert row("FAC_EXP").change == "NUEVO" and row("FAC_EXP").required and row("FAC_EXP").nullable


def test_xml_path_keeps_segments_after_parenthesised_notes():
    assert xml_path("DatosEntrada/Notas(3….N)/etiqueta") == "DatosEntrada/Notas/etiqueta"
    assert xml_path("DatosEntrada/OtrosTotales/valorOtrosTotales(4)") == "DatosEntrada/OtrosTotales/valorOtrosTotales"
    assert xml_path("NULL (NC), Validar con PBS") == ""
    assert clean_tag("resumen/subTotalVentas (normativa 125)") == "resumen/subTotalVentas"


def test_cross_check_count_ignores_null_repetitions(tmp_path):
    from contract_validator.models import SpecField
    p = tmp_path / "d.xml"
    item = "<Items>" + "<T><I>null</I></T>" * 2 + "<T><I>%s</I></T></Items>"
    p.write_text('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body><E><F><CabeceraEntrada><uuid>X</uuid>'
                 f"</CabeceraEntrada><DatosEntrada>{item % '20'}{item % '30'}</DatosEntrada></F></E></soapenv:Body></soapenv:Envelope>", encoding="utf-8")
    xml = load_document(str(p))
    spec = [SpecField("c.t", doc_type="CCF", fmt="JSON", counterpart="DatosEntrada.Items.T.I")]

    def run(vals):
        j = flatten_json({"c": [{"t": [v]} for v in vals]})
        j.fmt, j.raw, j.doc_type = "JSON", {"c": [{"t": [v]} for v in vals]}, "CCF"
        return [f.code for f in cross_check("x", xml, "j", j, spec, "CCF").findings]
    assert run(["20", "30"]) == []
    assert run(["20", "99"]) == ["XREF_MISMATCH"]
    assert run(["20"]) == ["XREF_COUNT_MISMATCH"]
