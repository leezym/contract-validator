"""db_check: contenido JSON/XML/PDF contra valores esperados de BD. Solo datos sintéticos; usa el contrato plano."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.db_check import BdConfig, check_doc, check_pdf_db, load_bd_config, same
from contract_validator.flatten import load_document
from contract_validator.pdf_check import Presentation
from contract_validator.pdf_text import PdfInfo
from contract_validator.spec_loader import load_spec

PLANO = os.path.join(ROOT, "data", "specs", "contrato_plano.xlsx")
pytestmark = pytest.mark.skipif(not os.path.exists(PLANO), reason="falta data/specs/contrato_plano.xlsx")

U = "AAAAAAAA-0000-4000-8000-000000000001"
CTRL = "DTE-03-M001P001-000000000000001"
APPROVED = {p: "DIRECTA" for p in (
    "identificacion.tipoDte", "identificacion.numeroControl", "identificacion.codigoGeneracion", "emisor.nit",
    "cuerpoDocumento.codigo", "cuerpoDocumento.cantidad", "cuerpoDocumento.precioUni", "resumen.totalPagar",
    "apendice.valor", "emisor.direccion.departamento")}
SIN_FUENTE = ("identificacion.horEmi",)
EXPECTED = {
    "identificacion.tipoDte": "03", "identificacion.numeroControl": CTRL, "identificacion.codigoGeneracion": U,
    "emisor.nit": "00000000000001", "resumen.totalPagar": "100.5", "emisor.direccion.departamento": "05",
    "cuerpoDocumento.codigo[1]": "0000001", "cuerpoDocumento.cantidad[1]": "2", "cuerpoDocumento.precioUni[1]": "10.25",
    "cuerpoDocumento.codigo[2]": "0000002", "cuerpoDocumento.cantidad[2]": "1", "cuerpoDocumento.precioUni[2]": "80",
    "apendice.valor[apendice02]": "000001|000002",
}


@pytest.fixture(scope="module")
def spec():
    return load_spec(PLANO)


@pytest.fixture
def cfg():
    m = {("CCF", p): c for p, c in APPROVED.items()}
    m.update({("CCF", p): "SIN_FUENTE" for p in SIN_FUENTE})
    return BdConfig(m, 0.01)


def _json(tmp_path, **over):
    d = {
        "identificacion": {"tipoDte": "03", "numeroControl": CTRL, "codigoGeneracion": U, "horEmi": "10:00:00"},
        "emisor": {"nit": "00000000000001", "direccion": {"departamento": "05"}},
        "cuerpoDocumento": [{"codigo": "0000001", "cantidad": 2, "precioUni": 10.25},
                            {"codigo": "0000002", "cantidad": 1, "precioUni": 80.0}],
        "resumen": {"totalPagar": 100.50},
        "apendice": [{"campo": "apendice02", "etiqueta": "Lista:|Documento:", "valor": "000001|000002"}],
    }
    for k, v in over.items():
        sec, key = k.split("__")
        d[sec][key] = v
    p = tmp_path / "doc.json"
    p.write_text(json.dumps(d), encoding="utf-8")
    return str(p), load_document(str(p))


def _codes(res):
    return [(f.severity, f.code, f.path) for f in res.findings]


def test_same_rules(cfg):
    assert same("100.5", 100.50, cfg) and same("100.5", "100.51", cfg)
    assert not same("100.5", "100.52", cfg)
    assert not same("05", 5, cfg)            # códigos con cero a la izquierda: texto exacto
    assert same(None, "", cfg) and not same("x", None, cfg)
    assert not same("ABC", "abc", cfg)


def test_json_equal(tmp_path, spec, cfg):
    f, doc = _json(tmp_path)
    res = check_doc(f, doc, EXPECTED, spec, cfg, "CCF")
    assert res.ok, _codes(res)
    assert res.checked >= 12
    assert ("WARNING", "DB_MISSING", "") in _codes(res)  # horEmi sin fuente en BD
    assert "identificacion.horEmi" in res.findings[-1].message


def test_json_mismatch_and_missing(tmp_path, spec, cfg):
    f, doc = _json(tmp_path, resumen__totalPagar=101.0, emisor__nit=None)
    res = check_doc(f, doc, EXPECTED, spec, cfg, "CCF")
    got = {(c, p) for _, c, p in _codes(res)}
    assert ("DB_MISMATCH_JSON", "resumen.totalPagar") in got
    assert ("DB_MISMATCH_JSON", "emisor.nit") in got
    assert any("falta en el documento" in x.message for x in res.findings if x.path == "emisor.nit")


def test_json_item_count(tmp_path, spec, cfg):
    f, doc = _json(tmp_path)
    doc.raw["cuerpoDocumento"].pop()
    res = check_doc(f, doc, EXPECTED, spec, cfg, "CCF")
    assert ("ERROR", "DB_MISMATCH_JSON", "cuerpoDocumento") in _codes(res)


def test_unapproved_paths_not_compared(tmp_path, spec):
    f, doc = _json(tmp_path, emisor__nit="99")
    res = check_doc(f, doc, EXPECTED, spec, BdConfig({("CCF", "identificacion.tipoDte"): "DIRECTA"}), "CCF")
    assert res.ok and res.checked == 1


def test_document_not_in_query(tmp_path, spec, cfg):
    f, doc = _json(tmp_path)
    res = check_doc(f, doc, None, spec, cfg, "CCF")
    assert _codes(res) == [("ERROR", "DB_MISSING", "")]


XML = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body><EnviarDocumento>
<EnviarDocumentoEntrada><CabeceraEntrada><uuid>{u}</uuid></CabeceraEntrada><DatosEntrada>
<Encabezado><tipoFactura>03</tipoFactura><secuenciaDocumento>{c}</secuenciaDocumento><nitCompania>{nit}</nitCompania></Encabezado>
<Items><codigoArticulo>0000001</codigoArticulo></Items><Items><codigoArticulo>0000002</codigoArticulo></Items>
</DatosEntrada></EnviarDocumentoEntrada></EnviarDocumento></soapenv:Body></soapenv:Envelope>"""


def _xml(tmp_path, nit="00000000000001", items=2):
    body = XML.format(u=U, c=CTRL, nit=nit)
    if items == 1:
        body = body.replace("<Items><codigoArticulo>0000002</codigoArticulo></Items>", "")
    p = tmp_path / "doc.xml"
    p.write_text(body, encoding="utf-8")
    return str(p), load_document(str(p))


def test_xml_counterparts(tmp_path, spec, cfg):
    f, doc = _xml(tmp_path)
    res = check_doc(f, doc, EXPECTED, spec, cfg, "CCF")
    assert not [x for x in res.findings if x.code == "DB_MISMATCH_XML" and x.path in
                ("emisor.nit", "identificacion.tipoDte", "identificacion.numeroControl", "cuerpoDocumento.codigo[1]")]
    f, doc = _xml(tmp_path, nit="00000000000009", items=1)
    got = {(c, p) for _, c, p in _codes(check_doc(f, doc, EXPECTED, spec, cfg, "CCF"))}
    assert ("DB_MISMATCH_XML", "emisor.nit") in got
    assert ("DB_MISMATCH_XML", "DatosEntrada.Items") in got


def _apendice_doc(tmp_path, body):
    p = tmp_path / "d.xml"
    p.write_text('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"><soapenv:Body><E><F>'
                 f"<CabeceraEntrada><uuid>{U}</uuid></CabeceraEntrada><DatosEntrada>{body}</DatosEntrada></F></E>"
                 "</soapenv:Body></soapenv:Envelope>", encoding="utf-8")
    return load_document(str(p))


APX_EXP = {"apendice.campo[apendice01]": "apendice01", "apendice.etiqueta[apendice01]": "A:|B:", "apendice.valor[apendice01]": "1|2"}
APX_CFG = BdConfig({("CCF", "apendice.etiqueta"): "TRANSFORMACION", ("CCF", "apendice.valor"): "TRANSFORMACION"})


def test_apendice_is_compared_as_a_unit_in_xml(spec, tmp_path):
    for k, o in (("campo", "etiqueta"), ("etiqueta", "campo")):          # el nombre se reconoce en cualquiera de los dos tags
        doc = _apendice_doc(tmp_path, f"<Notas><{k}>apendice01</{k}><{o}>A:|B:</{o}><notasAdicionales>1|2</notasAdicionales></Notas>")
        assert not [f for f in check_doc("d.xml", doc, APX_EXP, spec, APX_CFG, "CCF").findings if "apendice" in f.path]
    doc = _apendice_doc(tmp_path, "<Notas><campo>apendice01</campo><etiqueta>A:|B:</etiqueta><notasAdicionales>9|2</notasAdicionales></Notas>")
    assert [(f.code, f.path) for f in check_doc("d.xml", doc, APX_EXP, spec, APX_CFG, "CCF").findings if "apendice" in f.path] == [
        ("DB_MISMATCH_XML", "apendice[apendice01]")]


def test_apendice_missing_in_json_is_error_even_if_observaciones_matches(spec, tmp_path):
    p = tmp_path / "d.json"
    p.write_text(json.dumps({"identificacion": {"codigoGeneracion": U, "tipoDte": "03"}, "resumen": {"observaciones": "1|2"}}), encoding="utf-8")
    res = check_doc("d.json", load_document(str(p)), APX_EXP, spec, APX_CFG, "CCF")
    assert [(f.severity, f.code, f.path) for f in res.findings if "apendice" in f.path] == [
        ("ERROR", "DB_MISMATCH_JSON", "apendice[apendice01]")]
    xml = _apendice_doc(tmp_path, "<Notas><notasAdicionales>.</notasAdicionales></Notas>")
    assert [(f.severity, f.code) for f in check_doc("d.xml", xml, APX_EXP, spec, APX_CFG, "CCF").findings if "apendice" in f.path] == [
        ("WARNING", "APENDICE_NOT_RECEIVED")]


def _pdf(lines, uuid=U, control=CTRL):
    return PdfInfo(path="x_MINH.pdf", text="\n".join(lines), kind="MINH", uuid=uuid, control=control, title="CCF", lines=lines)


def test_pdf(spec, cfg):
    pres = Presentation({"normalizaciones": {"numeros": True}})
    ok = _pdf(["NIT Emisor: 00000000000001",
               "0000001 PRODUCTO UNO | LOTE: X1", "Unidad 2.00 10.25 0.00",
               "0000002 PRODUCTO DOS Unidad 1.00 80.00 0.00"])
    res = check_pdf_db(ok, EXPECTED, spec, cfg, pres, "CCF")
    assert res.ok, _codes(res)
    bad = _pdf(["NIT Emisor: 00000000000009", "0000001 PRODUCTO UNO Unidad 3.00 10.25 0.00"], control="DTE-03-X")
    got = {(c, p) for _, c, p in _codes(check_pdf_db(bad, EXPECTED, spec, cfg, pres, "CCF"))}
    assert ("DB_MISMATCH_PDF", "identificacion.numeroControl") in got
    assert ("DB_MISMATCH_PDF", "emisor.nit") in got
    assert ("DB_MISMATCH_PDF", "cuerpoDocumento.cantidad[1]") in got
    assert ("DB_MISMATCH_PDF", "cuerpoDocumento.codigo[2]") in got
    assert _codes(check_pdf_db(ok, None, spec, cfg, pres, "CCF")) == [("ERROR", "DB_MISSING", "")]


def test_pdf_catalog_multiword(spec):
    """El PDF imprime el nombre del catálogo (varias palabras) y la BD trae el código."""
    path = "DatosEntrada.InformacioGeneral.InformacionEmisor.recintoFiscal"
    pres = Presentation({"catalogos": {path: {"01": "Terrestre San Bartolo", "02": "Maritima de Acajutla"}}})
    cfg = BdConfig({("FAC_EXP", "emisor.recintoFiscal"): "DIRECTA"})
    info = lambda ln: PdfInfo(path="x_MINH.pdf", text=ln, kind="MINH", uuid=U, control="", title="FAC_EXP", lines=[ln])
    exp = {"emisor.recintoFiscal": "01"}
    ok = check_pdf_db(info("Recinto fiscal: Terrestre San Bartolo"), exp, spec, cfg, pres, "FAC_EXP")
    assert ok.ok and ok.checked == 1, _codes(ok)
    bad = check_pdf_db(info("Recinto fiscal: Maritima de Acajutla"), exp, spec, cfg, pres, "FAC_EXP")
    assert ("ERROR", "DB_MISMATCH_PDF", "emisor.recintoFiscal") in _codes(bad)
    sin_catalogo = check_pdf_db(info("Recinto fiscal: Terrestre San Bartolo"), exp, spec, cfg, Presentation(), "FAC_EXP")
    assert not sin_catalogo.ok  # sin el catálogo aprobado, código vs texto no se da por igual


def test_load_bd_config(tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "mapeo"
    ws.append(["tipo_doc", "ruta", "ruta_contraparte", "cambio_contrato", "categoria", "fuente_bd", "transformacion",
               "evidencia", "valor_observado", "notas", "aprobado"])
    ws.append(["CCF", "emisor.nit", "", "", "DIRECTA", "", "", "", "", "", "SI"])
    ws.append(["CCF", "emisor.nrc", "", "", "TRANSFORMACION", "", "", "", "", "", ""])  # sin aprobar
    ws.append(["CCF", "identificacion.horEmi", "", "", "SIN_FUENTE", "", "", "", "", "", "DB_MISSING"])
    m = tmp_path / "mapeo.xlsx"
    wb.save(m)
    c = tmp_path / "norm.json"
    c.write_text(json.dumps({"tolerancia_montos": 0.01}), encoding="utf-8")
    cfg = load_bd_config(str(m), str(c))
    assert cfg.category("CCF", "emisor.nit") == "DIRECTA"
    assert cfg.category("CCF", "emisor.nrc") == ""
    assert cfg.category("CCF", "identificacion.horEmi") == "SIN_FUENTE"
    assert str(cfg.tolerance) == "0.01"
    with pytest.raises(FileNotFoundError):
        load_bd_config(str(tmp_path / "no.xlsx"))


def test_time_tolerance_for_emission_hour():
    """horEmi: el generador usa SYSDATE - 1 h antes de la autorización; la BD solo guarda FEC_AUTORIZA."""
    cfg = BdConfig({}, 0.01, 60)
    assert same("10:36:18", "10:36:08", cfg)            # 10 s: dentro de la tolerancia
    assert not same("10:38:18", "10:36:08", cfg)        # 2 min: diferencia real
    assert same("00:00:20", "23:59:50", cfg)            # cruce de medianoche
    assert not same("10:36:18", "10:36:08", BdConfig({}, 0.01))   # sin tolerancia configurada: exacto
    assert not same("2026-11-13", "2026-11-14", cfg)    # solo aplica a horas HH:MM:SS


def test_load_time_tolerance(tmp_path):
    c = tmp_path / "norm.json"
    c.write_text(json.dumps({"tolerancia_montos": 0.01, "tolerancia_hora_segundos": 60}), encoding="utf-8")
    assert load_bd_config(None, str(c)).time_tolerance == 60


def test_query_covers_former_sin_fuente_fields():
    """Los 13 campos antes SIN_FUENTE salen del query (fuente: generador / CONTROL / FEC_AUTORIZA)."""
    sql = open(os.path.join(ROOT, "queries", "valores_esperados.sql"), encoding="utf-8").read()
    for ruta in ("identificacion.ambiente", "identificacion.horEmi", "identificacion.tipoMoneda", "resumen.descuNoSuj",
                 "resumen.descuExenta", "resumen.descuGravada", "resumen.porcentajeDescuento", "resumen.saldoFavor",
                 "resumen.totalLetras", "resumen.totalNoGravado", "cuerpoDocumento.uniMedida", "cuerpoDocumento.noGravado",
                 "cuerpoDocumento.psv"):
        assert f"AS '{ruta}'" in sql, ruta
    from contract_validator.db_fetch import check_read_only
    check_read_only(sql)   # sigue siendo una sola consulta de solo lectura


def test_real_time_tolerance_is_approved_value():
    """Valor aprobado por el usuario (2026-11-17): 120 s entre FEC_AUTORIZA - 1 h y horEmi."""
    cfg = load_bd_config(None, os.path.join(ROOT, "data", "specs", "normalizaciones_bd.json"))
    assert cfg.time_tolerance == 120
    assert same("10:38:00", "10:36:08", cfg) and not same("10:38:10", "10:36:08", cfg)


def test_query_covers_export_fields():
    sql = open(os.path.join(ROOT, "queries", "valores_esperados.sql"), encoding="utf-8").read()
    assert "AS 'emisor.tipoItemExpor'" in sql and "'TIPO_ITEM'" in sql
    assert "k.tipo_doc = 'FAC_EXP' THEN '37'" in sql
    assert "Para efectos de aduana: valor FOB: " in sql and "precioneto_moneda_sin_prorr" in sql
    cfg = load_bd_config()
    for ruta in ("emisor.tipoItemExpor", "receptor.tipoDocumento", "resumen.observaciones"):
        assert cfg.category("FAC_EXP", ruta) in ("CATALOGO", "TRANSFORMACION"), ruta


def test_pdf_regimen_uses_abbreviated_type_and_next_line(spec):
    emi = "DatosEntrada.InformacioGeneral.InformacionEmisor."
    texto = "Exportación Definitiva Proveniente de Franquicia Definitiva, Distribución Internacional"
    pres = Presentation({"catalogos": {emi + "regimen": {"EX-1.1048.031": texto}},
                         "catalogos_compuestos": {emi + "regimen": emi + "tipoRegimen"}})
    cfg = BdConfig({("FAC_EXP", "emisor.regimen"): "DIRECTA"})
    lines = ["Regimen de exportacion: Exportación Definitiva Proveniente de Franquicia Definitiva,", "Distribución Internacional"]
    info = PdfInfo(path="x_MINH.pdf", text=" ".join(lines), kind="MINH", uuid=U, control="", title="FAC_EXP", lines=lines)
    exp = {"emisor.tipoRegimen": "EX-1", "emisor.regimen": "1048.031"}
    assert check_pdf_db(info, exp, spec, cfg, pres, "FAC_EXP").ok
    assert not check_pdf_db(info, {**exp, "emisor.tipoRegimen": "EX-3"}, spec, cfg, pres, "FAC_EXP").ok


def test_pdf_must_print_each_apendice01_value_with_its_label(spec):
    pres = Presentation({"apendices_impresos": {"apendices": ["apendice01"]}})
    cfg = BdConfig({("CCF", "apendice.valor"): "TRANSFORMACION", ("CCF", "apendice.etiqueta"): "TRANSFORMACION"})
    exp = {"apendice.etiqueta[apendice01]": "Regional:|Bodega:|Pedido:", "apendice.valor[apendice01]": "22|01|011711"}
    lines = ["Regional: 22 Bodega: 99", "Otro dato 011711"]
    info = PdfInfo(path="x_MINH.pdf", text=" ".join(lines), kind="MINH", uuid=U, control="", title="CCF", lines=lines)
    res = check_pdf_db(info, exp, spec, cfg, pres, "CCF")
    assert [(f.code, f.path) for f in res.findings if "apendice" in f.path] == [
        ("DB_MISMATCH_PDF", "apendice[apendice01].Bodega"), ("DB_MISMATCH_PDF", "apendice[apendice01].Pedido")]
    assert "PDF='99'" in next(f.message for f in res.findings if f.path.endswith("Bodega"))
    assert "(no impreso)" in next(f.message for f in res.findings if f.path.endswith("Pedido"))
