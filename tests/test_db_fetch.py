"""db_fetch: ejecución del query de valores esperados con una conexión FALSA (sin BD, datos sintéticos)."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from contract_validator import db_fetch
from contract_validator.db_fetch import (DbFetchError, check_read_only, fetch, fetch_expected, load_db_config,
                                         make_dsn, read_sql, write_tsv)
from contract_validator.expected import load_expected

U1 = "AAAAAAAA-0000-4000-8000-000000000001"
U2 = "AAAAAAAA-0000-4000-8000-000000000002"
DATA = {
    U1: [(U1, "CCF", "identificacion.tipoDte", "03"), (U1, "CCF", "cuerpoDocumento.descripcion[1]", "PROD\tUNO\nLOTE: X")],
    U2: [(U2, "FAC", "identificacion.tipoDte", "01"), (U2, "FAC", "resumen.totalPagar", 12.5)],
}


class FakeCursor:
    def __init__(self, log):
        self.log, self.rows, self.description = log, [], None

    def execute(self, sql, binds=None):
        self.log.append((sql.split()[0].upper(), binds))
        if binds is not None:
            u = binds["codigo_generacion"]
            self.rows = DATA.get(u, []) if u else [r for v in DATA.values() for r in v]
            self.description = [("CODIGO_GENERACION",), ("TIPO_DOC",), ("RUTA",), ("VALOR",)]

    def fetchall(self):
        return self.rows

    def close(self):
        self.log.append(("CLOSE", None))


class FakeConn:
    def __init__(self):
        self.log = []

    def cursor(self):
        return FakeCursor(self.log)

    def rollback(self):
        self.log.append(("ROLLBACK", None))

    def close(self):
        self.log.append(("CONN_CLOSE", None))


SQL = "WITH k AS (SELECT 1 FROM dual WHERE :codigo_generacion IS NULL AND :desde IS NULL AND :hasta IS NULL) SELECT * FROM k FETCH FIRST NVL(:max_docs, 100) ROWS ONLY"


def test_real_query_is_read_only():
    from contract_validator.db_fetch import _strip
    sql = read_sql()  # lanza DbFetchError si no es una única consulta SELECT/WITH
    assert _strip(sql).strip().upper().startswith("WITH")
    for bind in (":codigo_generacion", ":desde", ":hasta", ":max_docs"):
        assert bind in sql


@pytest.mark.parametrize("bad", [
    "DELETE FROM doc", "UPDATE doc SET x=1", "SELECT 1 FROM dual; DROP TABLE doc", "BEGIN NULL; END;",
    "SELECT * FROM doc FOR UPDATE", "INSERT INTO t SELECT 1 FROM dual", "  -- comentario\n MERGE INTO t USING s ON (1=1)",
])
def test_rejects_non_select(bad):
    with pytest.raises(DbFetchError):
        check_read_only(bad)


def test_allows_keywords_inside_comments_and_strings():
    sql = "-- no hace DELETE\nSELECT 'UPDATE' txt /* DROP */ FROM dual;"
    assert check_read_only(sql).endswith("FROM dual")


def test_fetch_per_uuid_read_only_and_rollback():
    conn = FakeConn()
    rows = fetch_expected(conn, SQL, uuids=[U1.lower(), U2])
    assert conn.log[0] == ("SET", None)                                   # SET TRANSACTION READ ONLY primero
    binds = [b for k, b in conn.log if b is not None]
    assert [b["codigo_generacion"] for b in binds] == [U1, U2] and all(b["max_docs"] == 1 for b in binds)
    assert ("ROLLBACK", None) in conn.log and conn.log[-1] == ("CLOSE", None)
    assert rows[-1] == (U2, "FAC", "resumen.totalPagar", "12.5")          # todo como texto


def test_fetch_by_range():
    conn = FakeConn()
    rows = fetch_expected(conn, SQL, desde="2026-11-01", hasta="2026-11-15", max_docs=5)
    b = [b for _, b in conn.log if b is not None][0]
    assert b == {"codigo_generacion": None, "desde": "2026-11-01", "hasta": "2026-11-15", "max_docs": 5}
    assert len(rows) == 4


def test_tsv_roundtrip_with_tabs_and_newlines(tmp_path):
    rows = fetch_expected(FakeConn(), SQL, uuids=[U1, U2])
    p = tmp_path / "esperados.tsv"
    write_tsv(rows, str(p))
    got = load_expected(str(p))
    assert got[U1]["cuerpoDocumento.descripcion[1]"] == "PROD\tUNO\nLOTE: X"
    assert got[U2] == {"identificacion.tipoDte": "01", "resumen.totalPagar": "12.5"}


def test_config_env_overrides_file_and_password_only_from_env(tmp_path):
    no_env = str(tmp_path / "no.env")  # aislar del .env real del proyecto
    p = tmp_path / "oracle.local.json"
    p.write_text(json.dumps({"host": "h", "port": 1530, "sid": "s", "user": "u"}), encoding="utf-8")
    cfg = load_db_config(str(p), environ={"CV_ENV_FILE": no_env, "CV_ORACLE_USER": "otro", "CV_ORACLE_PASSWORD": "x"})
    assert cfg["user"] == "otro" and cfg["password"] == "x" and cfg["host"] == "h"
    assert "SID=s" in make_dsn(cfg) and "PORT=1530" in make_dsn(cfg)
    assert make_dsn({"host": "h", "port": 1521, "service": "svc"}) == "h:1521/svc"
    p.write_text(json.dumps({"host": "h", "sid": "s", "password": "no"}), encoding="utf-8")
    with pytest.raises(DbFetchError, match="CV_ORACLE_PASSWORD"):
        load_db_config(str(p), environ={"CV_ENV_FILE": no_env})


def test_project_env_file(tmp_path):
    from contract_validator.db_fetch import read_env_file
    env = tmp_path / ".env"
    env.write_text("# credenciales\n\nCV_ORACLE_USER=lector\nCV_ORACLE_PASSWORD=\"p=a#ss\"\nexport CV_ORACLE_SID='s2'\n"
                   "OTRA=ignorada\nlinea sin igual\n", encoding="utf-8")
    assert read_env_file(str(env)) == {"CV_ORACLE_USER": "lector", "CV_ORACLE_PASSWORD": "p=a#ss", "CV_ORACLE_SID": "s2"}
    p = tmp_path / "oracle.local.json"
    p.write_text(json.dumps({"host": "h", "port": 1530, "sid": "s"}), encoding="utf-8")
    cfg = load_db_config(str(p), environ={"CV_ENV_FILE": str(env)})
    assert (cfg["user"], cfg["password"], cfg["sid"], cfg["host"]) == ("lector", "p=a#ss", "s2", "h")  # .env > json
    cfg = load_db_config(str(p), environ={"CV_ENV_FILE": str(env), "CV_ORACLE_USER": "proceso", "CV_ORACLE_PASSWORD": ""})
    assert cfg["user"] == "proceso" and cfg["password"] == "p=a#ss"   # entorno > .env; vacío no pisa
    empty = tmp_path / "vacio.env"
    empty.write_text("CV_ORACLE_USER=\nCV_ORACLE_PASSWORD=\n", encoding="utf-8")
    cfg = load_db_config(str(p), environ={"CV_ENV_FILE": str(empty)})
    with pytest.raises(DbFetchError, match="Falta el usuario"):
        db_fetch.connect(cfg)


def test_repo_env_file_is_git_ignored():
    gi = open(os.path.join(ROOT, ".gitignore"), encoding="utf-8").read().splitlines()
    assert ".env" in gi and "config/*.local.json" in gi


def test_connect_requires_user_and_password():
    with pytest.raises(DbFetchError, match="usuario"):
        db_fetch.connect({"host": "h", "sid": "s", "password": "x"})
    with pytest.raises(DbFetchError, match="CV_ORACLE_PASSWORD"):
        db_fetch.connect({"host": "h", "sid": "s", "user": "u", "password": ""})


def test_fetch_shortcut_closes_connection(tmp_path):
    conn = FakeConn()
    sql = tmp_path / "q.sql"
    sql.write_text(SQL, encoding="utf-8")
    rows = fetch([U1], sql_path=str(sql), db_config=None, connector=lambda cfg: conn)
    assert len(rows) == 2 and conn.log[-1] == ("CONN_CLOSE", None)


def test_cli_expected_db(tmp_path, monkeypatch):
    """validate --expected-db consulta solo los UUID de los archivos de entrada (conexión falsa)."""
    from contract_validator import cli
    doc = tmp_path / "d.json"
    doc.write_text(json.dumps({"identificacion": {"tipoDte": "03", "codigoGeneracion": U1}}), encoding="utf-8")
    asked = {}

    def fake_fetch(uuids, db_config=None, **kw):
        asked["uuids"] = uuids
        return DATA[U1]
    monkeypatch.setattr(db_fetch, "fetch", fake_fetch)
    out = tmp_path / "esp.tsv"
    cli.main([str(doc), "--expected-db", "--no-cross", "--save-expected", str(out)])
    assert asked["uuids"] == [U1]
    assert load_expected(str(out))[U1]["identificacion.tipoDte"] == "03"
