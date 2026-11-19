"""load_expected: lectura del resultado del query de valores esperados. Solo datos sintéticos."""
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator.expected import load_expected

U1 = "AAAAAAAA-0000-4000-8000-000000000001"
U2 = "bbbbbbbb-0000-4000-8000-000000000002"
ROWS = [
    ("CODIGO_GENERACION", "TIPO_DOC", "RUTA", "VALOR"),
    (U1, "CCF", "identificacion.tipoDte", "03"),
    (U1, "CCF", "resumen.totalGravada", "100.5"),
    (U1, "CCF", "cuerpoDocumento.codigo[1]", "0000001"),
    (U1, "CCF", "cuerpoDocumento.codigo[2]", "0000002"),
    (U1, "CCF", "apendice.valor[apendice02]", "000001|000002"),
    (U1, "CCF", "receptor.nrc", "[NULL]"),
    (U2, "FAC", "identificacion.tipoDte", "01"),
    (U2, "FAC", "receptor.nombre", "Cliente Ñandú, S.A."),
]
EXPECTED = {
    U1: {"identificacion.tipoDte": "03", "resumen.totalGravada": "100.5", "cuerpoDocumento.codigo[1]": "0000001",
         "cuerpoDocumento.codigo[2]": "0000002", "apendice.valor[apendice02]": "000001|000002"},
    U2.upper(): {"identificacion.tipoDte": "01", "receptor.nombre": "Cliente Ñandú, S.A."},
}


def _write(path, delim):
    with open(path, "w", encoding="utf-8", newline="") as fh:
        for r in ROWS:
            fh.write(delim.join(f'"{c}"' if delim == "," and "," in c else c for c in r) + "\n")


@pytest.mark.parametrize("ext,delim", [(".tsv", "\t"), (".csv", ","), (".csv", ";")])
def test_text_exports(tmp_path, ext, delim):
    p = tmp_path / f"esperados{ext}"
    _write(p, delim)
    assert load_expected(str(p)) == EXPECTED


def test_xlsx_export(tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    for r in ROWS:
        wb.active.append(list(r))
    p = tmp_path / "esperados.xlsx"
    wb.save(p)
    assert load_expected(str(p)) == EXPECTED


def test_iterables():
    header, data = ROWS[0], ROWS[1:]
    assert load_expected([dict(zip(header, r)) for r in data]) == EXPECTED
    assert load_expected(data) == EXPECTED  # tuplas de 4 columnas (cursor)
    assert load_expected([(r[0], r[2], r[3]) for r in data]) == EXPECTED  # tuplas de 3 columnas


def test_values_kept_as_text():
    got = load_expected([(U1, "resumen.totalPagar", "0.10"), (U1, "receptor.codActividad", "046484")])
    assert got[U1] == {"resumen.totalPagar": "0.10", "receptor.codActividad": "046484"}


def test_nulls_skipped():
    assert load_expected([(U1, "a", ""), (U1, "b", None), (U1, "c", "null"), (None, "d", "x")]) == {}


def test_missing_columns():
    with pytest.raises(ValueError, match="RUTA"):
        load_expected([{"CODIGO_GENERACION": U1, "VALOR": "x"}])


def test_conflicting_duplicate():
    assert load_expected([(U1, "a", "1"), (U1, "a", "1")]) == {U1: {"a": "1"}}
    with pytest.raises(ValueError, match="duplicado"):
        load_expected([(U1, "a", "1"), (U1, "a", "2")])
