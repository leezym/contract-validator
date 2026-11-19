"""Ejecuta queries/valores_esperados.sql contra Oracle (SOLO LECTURA) y entrega las filas para `load_expected`.

Conexión (python-oracledb, modo thin; no requiere cliente Oracle):
  * `config/oracle.local.json` (ignorado por git, NO lleva contraseña):
        {"host": "...", "port": 1530, "sid": "..."}   o   {"dsn": "host:puerto/servicio"}   y opcional "user"
  * variables CV_ORACLE_DSN, CV_ORACLE_HOST, CV_ORACLE_PORT, CV_ORACLE_SID, CV_ORACLE_SERVICE, CV_ORACLE_USER y
    **CV_ORACLE_PASSWORD** (la contraseña solo se lee de variables), tomadas de:
      1. el entorno del proceso (prioridad), o
      2. el archivo `.env` del proyecto (raíz del repo, ignorado por git; otra ruta con CV_ENV_FILE).
    Precedencia: entorno > .env > config/oracle.local.json.

Salvaguardas de solo lectura:
  1. el SQL debe ser una única sentencia SELECT/WITH (se rechaza DML/DDL/PLSQL);
  2. la sesión abre con `SET TRANSACTION READ ONLY` (Oracle rechaza cualquier escritura);
  3. al terminar se hace rollback y se cierra la conexión;
  4. tope de documentos por consulta (`max_docs`).

Compañías: cada documento se consulta en el esquema de su CIA (data/specs/companias.json). El query usa
`{ESQUEMA_VENTAS}` en lugar del esquema y el bind `:cia`; el esquema sale solo de la configuración y se valida como
identificador Oracle antes de reemplazarlo.
"""
from __future__ import annotations
import csv
import io
import json
import os
import re

DEFAULT_SQL = os.path.join("queries", "valores_esperados.sql")
DEFAULT_DB_CONFIG = os.path.join("config", "oracle.local.json")
COLUMNS = ("CODIGO_GENERACION", "TIPO_DOC", "RUTA", "VALOR")
FORBIDDEN = re.compile(r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|TRUNCATE|GRANT|REVOKE|RENAME|COMMIT|"
                       r"ROLLBACK|SAVEPOINT|LOCK|EXECUTE|EXEC|BEGIN|DECLARE|CALL)\b|\bFOR\s+UPDATE\b", re.I)
ENV = {"dsn": "CV_ORACLE_DSN", "host": "CV_ORACLE_HOST", "port": "CV_ORACLE_PORT", "sid": "CV_ORACLE_SID",
       "service": "CV_ORACLE_SERVICE", "user": "CV_ORACLE_USER"}
PASSWORD_ENV = "CV_ORACLE_PASSWORD"
SCHEMA_PLACEHOLDER = "{ESQUEMA_VENTAS}"


class DbFetchError(RuntimeError):
    pass


# ------------------------------------------------------------------ SQL
def _strip(sql: str) -> str:
    """Quita comentarios y literales de texto (para revisar solo la estructura de la sentencia)."""
    sql = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
    sql = re.sub(r"--[^\n]*", " ", sql)
    return re.sub(r"'(?:[^']|'')*'", "''", sql)


def check_read_only(sql: str) -> str:
    """Valida que sea una única consulta SELECT/WITH y la devuelve sin ';' final."""
    body = sql.strip().rstrip(";").strip()
    bare = _strip(body).strip()
    if not re.match(r"^(WITH|SELECT)\b", bare, re.I):
        raise DbFetchError("El query debe empezar con SELECT o WITH (solo lectura)")
    if ";" in bare:
        raise DbFetchError("El query debe ser una sola sentencia")
    m = FORBIDDEN.search(bare)
    if m:
        raise DbFetchError(f"El query contiene una instrucción no permitida: {m.group(0).upper()}")
    return body


def render_sql(sql: str, esquema: str | None) -> str:
    """Pone el esquema de la compañía donde el query dice {ESQUEMA_VENTAS}."""
    if SCHEMA_PLACEHOLDER not in sql:
        return sql
    from .companias import SCHEMA_RX
    if not esquema or not SCHEMA_RX.match(esquema):
        raise DbFetchError(f"Esquema de compañía inválido o ausente: {esquema!r}")
    return sql.replace(SCHEMA_PLACEHOLDER, esquema)


def read_sql(path: str | None = None) -> str:
    with open(path or DEFAULT_SQL, encoding="utf-8") as fh:
        return check_read_only(fh.read())


# ------------------------------------------------------------------ conexión
DEFAULT_ENV_FILE = ".env"


def read_env_file(path: str) -> dict:
    """Lee KEY=VALUE (líneas vacías y # comentarios se ignoran; comillas opcionales). Solo claves CV_*."""
    out = {}
    if not path or not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig") as fh:
        for ln in fh:
            ln = ln.strip()
            if not ln or ln.startswith("#") or "=" not in ln:
                continue
            k, v = ln.split("=", 1)
            k = k.strip().removeprefix("export ").strip()
            v = v.strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
                v = v[1:-1]
            if k.startswith("CV_"):
                out[k] = v
    return out


def _environ(environ=None) -> dict:
    """Entorno efectivo: archivo .env del proyecto + variables del proceso (estas tienen prioridad)."""
    real = os.environ if environ is None else environ
    env_file = real.get("CV_ENV_FILE") or DEFAULT_ENV_FILE
    merged = read_env_file(env_file)
    merged.update({k: v for k, v in real.items() if v})
    return merged


def load_db_config(path: str | None = None, environ=None) -> dict:
    environ = _environ(environ)
    cfg = {}
    p = path or DEFAULT_DB_CONFIG
    if os.path.exists(p):
        with open(p, encoding="utf-8-sig") as fh:
            cfg = json.load(fh)
        if any(k.lower() in ("password", "pass", "pwd", "clave") for k in cfg):
            raise DbFetchError(f"{p} no debe contener la contraseña: use {PASSWORD_ENV} en el archivo .env del proyecto")
    elif path:
        raise DbFetchError(f"No existe el archivo de conexión: {path}")
    for k, var in ENV.items():
        if environ.get(var):
            cfg[k] = environ[var]
    cfg["password"] = environ.get(PASSWORD_ENV, "")
    return cfg


def make_dsn(cfg: dict) -> str:
    if cfg.get("dsn"):
        return cfg["dsn"]
    if not cfg.get("host"):
        raise DbFetchError(f"Falta la conexión: defina {DEFAULT_DB_CONFIG} (host, port, sid|service) o {ENV['dsn']}")
    port = int(cfg.get("port") or 1521)
    if cfg.get("sid"):
        return f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={cfg['host']})(PORT={port}))(CONNECT_DATA=(SID={cfg['sid']})))"
    if cfg.get("service"):
        return f"{cfg['host']}:{port}/{cfg['service']}"
    raise DbFetchError("Falta 'sid' o 'service' en la configuración de conexión")


def connect(cfg: dict):
    if not cfg.get("user"):
        raise DbFetchError(f"Falta el usuario: {ENV['user']} en el archivo .env del proyecto (o 'user' en {DEFAULT_DB_CONFIG})")
    if not cfg.get("password"):
        raise DbFetchError(f"Falta la contraseña: {PASSWORD_ENV} en el archivo .env del proyecto")
    try:
        import oracledb
    except ImportError as e:
        raise DbFetchError("Falta el driver: pip install -r requirements.txt (oracledb)") from e
    try:
        return oracledb.connect(user=cfg["user"], password=cfg["password"], dsn=make_dsn(cfg), tcp_connect_timeout=15)
    except Exception as e:  # red/VPN, credenciales, servicio
        raise DbFetchError(f"No se pudo conectar a Oracle ({type(e).__name__}): {str(e).splitlines()[0]}. "
                           "Revise VPN, host/puerto/SID y credenciales") from None


# ------------------------------------------------------------------ ejecución
def fetch_expected(conn, sql: str, uuids=None, desde: str | None = None, hasta: str | None = None,
                   max_docs: int | None = None, cia: str | None = None, esquema: str | None = None) -> list[tuple]:
    """Ejecuta el query (una vez por UUID, o una vez por rango) y devuelve filas (uuid, tipo_doc, ruta, valor).
    cia/esquema: compañía de los documentos (el query debe consultarse una vez por compañía)."""
    sql = render_sql(check_read_only(sql), esquema)
    uses_cia = re.search(r":cia\b", _strip(sql), re.I) is not None
    if uses_cia and not cia:
        raise DbFetchError("El query requiere la compañía (:cia) del documento")
    rows: list[tuple] = []
    cur = conn.cursor()
    try:
        cur.execute("SET TRANSACTION READ ONLY")
        runs = [{"codigo_generacion": u.strip().upper(), "max_docs": 1} for u in uuids] if uuids else \
               [{"codigo_generacion": None, "max_docs": max_docs}]
        for b in runs:
            binds = {"codigo_generacion": b["codigo_generacion"], "desde": desde, "hasta": hasta, "max_docs": b["max_docs"]}
            if uses_cia:
                binds["cia"] = str(cia)
            cur.execute(sql, binds)
            head = [d[0].upper() for d in cur.description]
            ix = [head.index(c) for c in COLUMNS]
            for r in cur.fetchall():
                rows.append(tuple(None if r[i] is None else str(r[i]) for i in ix))
    finally:
        try:
            conn.rollback()
        finally:
            cur.close()
    return rows


def write_tsv(rows, path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, delimiter="\t", lineterminator="\n")
        w.writerow(COLUMNS)
        w.writerows(rows)


def to_tsv_text(rows) -> str:
    buf = io.StringIO()
    w = csv.writer(buf, delimiter="\t", lineterminator="\n")
    w.writerow(COLUMNS)
    w.writerows(rows)
    return buf.getvalue()


def fetch(uuids=None, desde=None, hasta=None, max_docs=None, sql_path=None, db_config=None, connector=connect,
          companias=None, cia_of=None, cias=None) -> list[tuple]:
    """Atajo: config + conexión + query, una vez por compañía. `connector` permite inyectar una conexión de prueba.

    companias: {cia: Compania} (default: data/specs/companias.json).
    cia_of: {uuid: cia} para consultar cada UUID en el esquema de su compañía (obligatorio con `uuids`).
    cias: compañías a consultar por rango de fechas (default: todas las configuradas).
    """
    from .companias import load_companias
    sql = read_sql(sql_path)
    companias = companias if companias is not None else load_companias()
    if uuids:
        cia_of = {u.strip().upper(): str(c) for u, c in (cia_of or {}).items()}
        sin = sorted({u.strip().upper() for u in uuids} - set(cia_of))
        if sin:
            raise DbFetchError(f"Falta la compañía (CIA) de: {', '.join(sin)}")
        groups: dict[str, list] = {}
        for u in uuids:
            groups.setdefault(cia_of[u.strip().upper()], []).append(u)
    else:
        groups = {str(c): None for c in (cias or companias)}
    unknown = sorted(set(groups) - set(companias))
    if unknown:
        raise DbFetchError(f"CIA sin esquema en data/specs/companias.json: {', '.join(unknown)}")
    conn = connector(load_db_config(db_config))
    try:
        rows = []
        for cia, us in groups.items():
            rows += fetch_expected(conn, sql, us, desde, hasta, max_docs, cia=cia, esquema=companias[cia].esquema)
        return rows
    finally:
        conn.close()
