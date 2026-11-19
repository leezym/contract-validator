"""Validación PDF vs XML. Par real: NCR_SAL_MINH_* (PDF oficial) contra tests/fixtures/tq-notas-credito.xml."""
import glob
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.build_contract import build as build_flat
from contract_validator.flatten import load_document
from contract_validator.pdf_check import (Presentation, check_pdf, load_presentation, validate_pdfs)
from contract_validator.pdf_text import find_tool, read_pdf
from contract_validator.spec_loader import load_spec

SPECS = os.path.join(ROOT, "data", "specs")
SOURCE = os.path.join(SPECS, "contrato_ejemplo.xlsx")
FIXTURE = os.path.join(HERE, "fixtures", "tq-notas-credito.xml")
NCR_MINH = os.path.join(SPECS, "NCR_SAL_MINH_531601105009878001895.pdf")
# Los PDF internos de TQ (*_SAL_<fecha>_*) se retiraron de data/specs (2026-11-17): solo quedan los MINH. Los casos de
# PDF interno y PDF vacío se prueban con archivos generados en tmp_path.
INTERNAL_NAME = "NCR_SAL_20261112195735_531601105001895009878.pdf"


def _blank_pdf(path):
    """PDF válido de una página sin texto (como el interno de NCR que venía vacío)."""
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << >> >>",
            b"<< /Length 0 >>\nstream\n\nendstream"]
    out, offs = bytearray(b"%PDF-1.4\n"), []
    for n, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % n + o + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % x for x in offs)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    open(path, "wb").write(bytes(out))
    return str(path)


@pytest.fixture()
def internal_pdf(tmp_path):
    """PDF interno: mismo contenido que el MINH de NCR, con nombre de PDF interno (*_SAL_<fecha>_*)."""
    import shutil
    dst = tmp_path / INTERNAL_NAME
    shutil.copy(NCR_MINH, dst)
    return str(dst)

pytestmark = [
    pytest.mark.skipif(find_tool() is None, reason="falta pdftotext (Xpdf/Poppler)"),
    pytest.mark.skipif(not os.path.exists(SOURCE) or not os.path.exists(NCR_MINH), reason="faltan los specs de data/specs"),
]


@pytest.fixture(scope="module")
def spec(tmp_path_factory):
    out = str(tmp_path_factory.mktemp("plano") / "contrato_plano.xlsx")
    build_flat(SOURCE, out)
    return load_spec(out)


@pytest.fixture(scope="module")
def cfg():
    return load_presentation(os.path.join(SPECS, "pdf_presentacion.json"))


@pytest.fixture(scope="module")
def minh():
    return read_pdf(NCR_MINH)


def _xml(tmp_path, *replacements):
    text = open(FIXTURE, encoding="utf-8").read()
    for old, new in replacements:
        assert old in text, old
        text = text.replace(old, new, 1)
    p = tmp_path / "doc.xml"
    p.write_text(text, encoding="utf-8")
    return load_document(str(p))


def codes(res):
    return {(f.code, f.path) for f in res.findings}


# ------------------------------------------------------------------ extracción
def test_extraction_of_all_pdfs():
    pdfs = sorted(glob.glob(os.path.join(SPECS, "*.pdf")))
    assert len(pdfs) == 4 and all("_MINH_" in p for p in pdfs)   # data/specs solo trae los PDF oficiales MINH
    expect = {"CCF_SAL_": "CCF", "EXP_SAL_": "FAC_EXP", "NCR_SAL_": "NCR", "REM_SAL_": "REM"}
    for p in pdfs:
        i = read_pdf(p)
        assert not i.error, (p, i.error)
        assert i.title == expect[os.path.basename(p)[:8]], p
        assert i.uuid and i.control.startswith("DTE-")
        assert i.kind == "MINH" and i.seal  # el oficial trae sello de recepción


def test_pdf_without_text_is_empty(tmp_path):
    assert read_pdf(_blank_pdf(tmp_path / INTERNAL_NAME)).error[0] == "PDF_EMPTY"


# ------------------------------------------------------------------ par real
APENDICE01 = {f"DatosEntrada.Notas.{x}" for x in ("Regional", "Bodega", "Concepto", "Pedido", "Plazo")}


def _sin_apendice01(errors):
    """Estos PDF no imprimen apendice01: esos 5 faltantes son hallazgos esperados."""
    return [f for f in errors if not (f.code == "PDF_VALUE_MISSING" and f.path in APENDICE01)]


def _apendice01_faltante(res):
    return {f.path for f in res.errors if f.code == "PDF_VALUE_MISSING"} == APENDICE01


def test_real_pair_has_no_errors(spec, cfg, minh):
    res = check_pdf(minh, "fixture", load_document(FIXTURE), spec, cfg)
    assert res.doc_type == "NCR" and res.checked > 50
    assert _sin_apendice01(res.errors) == [], [(f.code, f.path, f.message) for f in res.errors]
    assert _apendice01_faltante(res)
    assert {f.code for f in res.findings} <= {"PDF_NOT_PRINTED", "PDF_VALUE_NOT_IN_XML", "PDF_VALUE_MISSING"}


def test_real_pair_detects_changes(spec, cfg, minh, tmp_path):
    def run(*rep):
        return codes(check_pdf(minh, "x", _xml(tmp_path, *rep), spec, cfg))

    assert ("PDF_IDENTITY", "CabeceraEntrada.uuid") in run(("21E51C60-3D7A-7F84-E063-9B78A8C07B69", "21E51C60-3D7A-7F84-E063-9B78A8C07B00"))
    assert ("PDF_IDENTITY", "DatosEntrada.Encabezado.secuenciaDocumento") in run(("000000000200014", "000000000200015"))
    assert ("PDF_MISMATCH", "DatosEntrada.Items.cantidadArticulo") in run(("<cantidadArticulo>50<", "<cantidadArticulo>51<"))
    assert ("PDF_MISMATCH", "DatosEntrada.Items.monedaPrecioUnitario") in run(("<monedaPrecioUnitario>1.17<", "<monedaPrecioUnitario>1.18<"))
    assert ("PDF_MISMATCH", "DatosEntrada.Items.CargoDescuentos/valor") in run(("<valor>23.50<", "<valor>23.51<"))
    assert ("PDF_MISMATCH", "DatosEntrada.Encabezado.versionFormato") in run(("<versionFormato>1<", "<versionFormato>2<"))  # Previo vs Diferido
    assert ("PDF_VALUE_MISSING", "DatosEntrada.Notas.Fecha Vencimiento") in run(("001568||2026/11/12", "001568||2026/11/13"))
    assert any(c == "PDF_NOT_PRINTED" and p.endswith("nitCompania") for c, p in run(("<nitCompania>05010912191013</nitCompania><versionUBL>", "<nitCompania>05010912191099</nitCompania><versionUBL>")))
    assert any(c == "PDF_NOT_PRINTED" and p.endswith("totalFactura") for c, p in run(("<totalFactura>168.04<", "<totalFactura>168.05<")))
    assert any(c == "PDF_DOCTYPE" for c, _ in run(("<tipoFactura>05<", "<tipoFactura>03<")))  # el XML pasa a CCF y el PDF es NCR


def test_missing_item_row_is_detected(spec, cfg, minh, tmp_path):
    text = open(FIXTURE, encoding="utf-8").read()
    start = text.index("<Items><numero>3<")
    end = text.index("</Items>", start) + len("</Items>")
    p = tmp_path / "sin_item3.xml"
    p.write_text(text[:start] + text[end:], encoding="utf-8")
    found = codes(check_pdf(minh, "x", load_document(str(p)), spec, cfg))
    assert ("PDF_COUNT_MISMATCH", "DatosEntrada.Items") in found


# ------------------------------------------------------------------ presentación (lista explícita)
def test_presentation_list_is_explicit(cfg):
    assert cfg.canon("1,603.69") == cfg.canon("1603.69") and cfg.canon("285.6000") == cfg.canon("285.60") == "285.6"
    assert cfg.canon("35.00") == cfg.canon("35") and cfg.canon("0.00") == cfg.canon("0")
    assert cfg.canon("13/11/2026") == cfg.canon("2026-11-13") == cfg.canon("2026/11/13")
    assert cfg.canon("0501-091219-101-3") == "05010912191013" and cfg.canon("2248-5100") == "22485100"
    assert cfg.canon("001918") != cfg.canon("1918")                 # ceros a la izquierda NO se alteran
    assert cfg.canon("DÓLAR") != cfg.canon("DOLAR") and cfg.canon("Ab") != cfg.canon("AB")  # texto: exacto
    strict = Presentation()
    assert strict.canon("1,603.69") != strict.canon("1603.69") and strict.canon("13/11/2026") != strict.canon("2026-11-13")


def test_accent_and_case_rules_are_opt_in():
    loose = Presentation({"normalizaciones": {"texto_ignorar_acentos": True, "texto_ignorar_mayusculas": True}})
    assert loose.canon("DÓLAR") == loose.canon("dolar")


# ------------------------------------------------------------------ validate_pdfs / CLI
def test_only_minh_pdfs_are_validated(spec, cfg):
    docs = {FIXTURE: load_document(FIXTURE)}
    pdfs = sorted(glob.glob(os.path.join(SPECS, "*.pdf")))
    results = validate_pdfs(pdfs, docs, spec, cfg)
    assert all("MINH" in r.file for r in results)          # ningun PDF interno produce resultados
    assert not any(f.code in ("PDF_EMPTY", "PDF_INTERNAL_VS_MINH") for r in results for f in r.findings)  # ni el NCR interno vacio
    paired = [r for r in results if r.fmt == "PDF(MINH)<->XML"]
    assert len(paired) == 1 and not _sin_apendice01(paired[0].errors) and paired[0].doc_type == "NCR"
    no_pair = [r for r in results if r.findings and r.findings[0].code == "PDF_NO_PAIR"]
    assert len(no_pair) == 3 and all(f.severity == "WARNING" for r in no_pair for f in r.findings)  # CCF, EXP y REM MINH sin XML
    assert len(results) == 4


def test_explicit_pair_without_xml_is_an_error(spec, cfg):
    res = validate_pdfs([NCR_MINH], {}, spec, cfg, explicit=[("falta.xml", NCR_MINH)])
    assert res[0].findings[0].code == "PDF_NO_PAIR" and res[0].findings[0].severity == "ERROR"


def test_explicit_internal_pdf_is_ignored_with_warning(spec, cfg, internal_pdf):
    res = validate_pdfs([internal_pdf], {FIXTURE: load_document(FIXTURE)}, spec, cfg, explicit=[(FIXTURE, internal_pdf)])
    assert [f.code for r in res for f in r.findings] == ["PDF_IGNORED"] and res[0].ok


def test_kind_depends_only_on_file_name(internal_pdf):
    assert read_pdf(NCR_MINH).kind == "MINH"
    assert read_pdf(internal_pdf).kind == "INTERNO"   # mismo contenido, otro nombre


def test_cli_with_pdfs(tmp_path):
    empty_internal = _blank_pdf(tmp_path / INTERNAL_NAME)
    from contract_validator.cli import main
    plano = str(tmp_path / "plano.xlsx")
    build_flat(SOURCE, plano)
    out = str(tmp_path / "r.xlsx")
    assert main(["--spec", plano, NCR_MINH, FIXTURE, "--out", out]) == 1  # el XML del fixture tiene diferencias con el Excel
    from openpyxl import load_workbook
    rows = list(load_workbook(out)["Hallazgos"].iter_rows(values_only=True))
    assert any(r[2] and "PDF" in str(r[2]) for r in rows[1:])
    # --no-pdf: los PDF se ignoran por completo
    assert main(["--spec", plano, NCR_MINH, "--no-pdf"]) == 0
    # el PDF interno (aunque esté vacío) no se valida: solo cuentan los MINH
    assert main(["--spec", plano, empty_internal]) == 0


# ------------------------------------------------------------------ par real CCF y evento de retorno (XML entregados en data/specs)
CCF_XML = glob.glob(os.path.join(SPECS, "53_16_01_001_131126_020656_*.xml"))
CCF_MINH = os.path.join(SPECS, "CCF_SAL_MINH_531601001131126020656.pdf")
RET_XML = glob.glob(os.path.join(SPECS, "53_16_01_105_009867_000002_*.xml"))
needs_ccf = pytest.mark.skipif(not CCF_XML or not os.path.exists(CCF_MINH), reason="faltan el XML o el PDF MINH del CCF")


@needs_ccf
def test_real_ccf_pair_has_no_errors(spec, cfg):
    xml = load_document(CCF_XML[0])
    assert xml.doc_type == "CCF"
    res = validate_pdfs([CCF_MINH], {CCF_XML[0]: xml}, spec, cfg)
    assert len(res) == 1 and res[0].doc_type == "CCF" and res[0].fmt == "PDF(MINH)<->XML"
    assert not _sin_apendice01(res[0].errors), [(f.code, f.path, f.message) for f in res[0].errors]  # incluye el caso 'Lista: Fecha Nota de Remisión:'


@needs_ccf
def test_real_ccf_pair_detects_changes(spec, cfg, tmp_path):
    text = open(CCF_XML[0], encoding="utf-8").read()
    info = read_pdf(CCF_MINH)

    def run(old, new):
        assert old in text, old
        p = tmp_path / "ccf.xml"
        p.write_text(text.replace(old, new, 1), encoding="utf-8")
        return codes(check_pdf(info, str(p), load_document(str(p)), spec, cfg))

    assert ("PDF_MISMATCH", "DatosEntrada.Items.cantidadArticulo") in run("<cantidadArticulo>35<", "<cantidadArticulo>36<")
    assert any(c == "PDF_IDENTITY" for c, _ in run("DTE-03-M001P001-000000000002561", "DTE-03-M001P001-000000000002562"))
    assert any(c == "PDF_VALUE_MISSING" and "Notas" in p for c, p in run("<notasAdicionales>034344", "<notasAdicionales>099999") ) or \
        any(c == "PDF_VALUE_MISSING" for c, _ in run("|020656<", "|029999<"))


@pytest.mark.skipif(not RET_XML, reason="falta el XML del evento de retorno")
def test_return_event_type_is_detected_from_the_soap_operation():
    doc = load_document(RET_XML[0])
    assert doc.fmt == "XML" and doc.doc_type == "RETORNO"   # no trae tipoFactura: lo identifica <EventoRetornoEntrada>
