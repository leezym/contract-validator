"""Uso:
  python -m contract_validator verificar <pdf/xml/json o carpeta>...   # TODO en uno: estructura + BD + veredicto por documento
  python -m contract_validator build --source contrato.xlsx --out contrato_plano.xlsx
  python -m contract_validator audit-mapping [--spec contrato_plano.xlsx] [--doc-type FAC] [--out reports/auditoria_mapeo.xlsx]
  python -m contract_validator [validate] [--pair XML JSON] [--pair-pdf XML PDF] --spec contrato_plano.xlsx archivo.json carpeta/ [--doc-type NDB] [--format text|json] [--out reporte.xlsx]
      [--expected esperados.tsv | --expected-db]   # contenido contra la BD (archivo exportado o consulta directa)
  python -m contract_validator fetch-expected (--cia 53 --codigo-generacion UUID... | --desde YYYY-MM-DD --hasta YYYY-MM-DD [--cia 41]) [--out reports/esperados.tsv]
"""
from __future__ import annotations
import argparse
import glob
import os
import sys
from .spec_loader import load_spec
from .flatten import load_document, doc_type_from_filename
from .comparator import compare
from .crosscheck import pair_documents, cross_check
from .pdf_check import load_presentation, validate_pdfs
from .models import Result, Finding
from . import report

DEFAULT_SPEC = os.path.join("data", "specs", "contrato_plano.xlsx")


# Archivos de configuración del propio validador: al recorrer una carpeta no se tratan como documentos
CONFIG_FILES = {"pdf_presentacion.json", "normalizaciones_bd.json", "ajustes_contrato.json", "oracle.example.json",
                "oracle.local.json", "companias.json"}


def collect(inputs):
    files = []
    for i in inputs:
        if os.path.isdir(i):
            for ext in ("xml", "json", "pdf"):
                files += sorted(f for f in glob.glob(os.path.join(i, "**", f"*.{ext}"), recursive=True)
                                if os.path.basename(f).lower() not in CONFIG_FILES and not f.lower().endswith(".local.json"))
        else:
            files += sorted(glob.glob(i)) or [i]
    return files


def cmd_build(argv) -> int:
    from .build_contract import build
    ap = argparse.ArgumentParser(prog="contract_validator build",
                                 description="Genera el contrato plano (hojas 'contrato' y 'lint') desde el libro original")
    ap.add_argument("--source", default=os.path.join("data", "specs", "contrato_ejemplo.xlsx"))
    ap.add_argument("--out", default=DEFAULT_SPEC)
    ap.add_argument("--ajustes", default=os.path.join("data", "specs", "ajustes_contrato.json"),
                    help="excepciones al libro aprobadas por el usuario (default: data/specs/ajustes_contrato.json)")
    ap.add_argument("--sin-ajustes", action="store_true", help="generar solo desde el libro, sin aplicar ajustes")
    a = ap.parse_args(argv)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    s = build(a.source, a.out, None if a.sin_ajustes else a.ajustes)
    print(f"Contrato plano: {a.out}\n  filas: {s['rows']}  lint: {s['lint_levels']}  ajustes: {', '.join(s['ajustes']) or 'ninguno'}")
    for (code, fmt), n in sorted(s["by_type"].items()):
        print(f"  {code:<13}{fmt:<5}{n}")
    return 0


def cmd_audit(argv) -> int:
    from . import audit_mapping
    ap = argparse.ArgumentParser(prog="contract_validator audit-mapping",
                                 description="Audita el mapeo JSON <-> XML_TQ del contrato (no requiere archivos reales)")
    ap.add_argument("--spec", default=DEFAULT_SPEC, help="contrato plano (hoja 'contrato')")
    ap.add_argument("--doc-type", help="limitar a un tipo: CCF, FAC, FAC_EXP, NCR, NDB, REM, INVALIDACION, RETORNO")
    ap.add_argument("--out", default=os.path.join("reports", "auditoria_mapeo.xlsx"))
    ap.add_argument("--fail-on-error", action="store_true", help="exit 1 si hay hallazgos de nivel ERROR")
    a = ap.parse_args(argv)
    rows = audit_mapping.audit(load_spec(a.spec), (a.doc_type or "").upper())
    audit_mapping.to_xlsx(rows, a.out)
    print(audit_mapping.to_text(rows))
    print(f"Reporte: {a.out}")
    return 1 if a.fail_on_error and any(r["nivel"] == "ERROR" for r in rows) else 0


def cmd_audit_plano(argv) -> int:
    from . import audit_plano
    ap = argparse.ArgumentParser(prog="contract_validator audit-plano",
                                 description="Verifica que el contrato plano sea fiel al libro original")
    ap.add_argument("--source", default=os.path.join("data", "specs", "contrato_ejemplo.xlsx"))
    ap.add_argument("--plano", default=DEFAULT_SPEC)
    ap.add_argument("--out", default=os.path.join("reports", "auditoria_plano.xlsx"))
    ap.add_argument("--fail-on-error", action="store_true", help="exit 1 si hay hallazgos de nivel ERROR")
    ap.add_argument("--ajustes", default=os.path.join("data", "specs", "ajustes_contrato.json"),
                    help="ajustes aprobados que el plano debe contener (default: data/specs/ajustes_contrato.json)")
    a = ap.parse_args(argv)
    rows, stats = audit_plano.audit(a.source, a.plano, a.ajustes)
    audit_plano.to_xlsx(rows, stats, a.out)
    print(audit_plano.to_text(rows, stats))
    print(f"Reporte: {a.out}")
    return 1 if a.fail_on_error and any(r["nivel"] == "ERROR" for r in rows) else 0


def cmd_fetch_expected(argv) -> int:
    from .db_fetch import DEFAULT_SQL, DbFetchError, fetch, write_tsv
    ap = argparse.ArgumentParser(prog="contract_validator fetch-expected",
                                 description="Ejecuta queries/valores_esperados.sql en Oracle (solo lectura) y guarda el resultado")
    ap.add_argument("--codigo-generacion", nargs="+", metavar="UUID", help="uno o más Códigos de Generación (de la compañía --cia)")
    ap.add_argument("--cia", nargs="+", help="compañía(s) a consultar (data/specs/companias.json); obligatoria con "
                                             "--codigo-generacion; con rango, default: todas las configuradas")
    ap.add_argument("--companias", help="compañías y esquemas (default: data/specs/companias.json)")
    ap.add_argument("--desde", help="fecha de autorización inicial YYYY-MM-DD (sin --codigo-generacion)")
    ap.add_argument("--hasta", help="fecha de autorización final YYYY-MM-DD")
    ap.add_argument("--max-docs", type=int, default=100, help="tope de documentos por rango (default 100)")
    ap.add_argument("--out", default=os.path.join("reports", "esperados.tsv"), help="archivo .tsv de salida (contiene datos reales)")
    ap.add_argument("--sql", default=DEFAULT_SQL)
    ap.add_argument("--db-config", help="conexión sin contraseña (default: config/oracle.local.json)")
    a = ap.parse_args(argv)
    if not a.codigo_generacion and not (a.desde or a.hasta):
        ap.error("indique --codigo-generacion o un rango --desde/--hasta")
    if a.codigo_generacion and (not a.cia or len(a.cia) != 1):
        ap.error("con --codigo-generacion indique la compañía de esos documentos: --cia <código> (p. ej. 53 o 41)")
    from .companias import CompaniaError, load_companias
    try:
        companias = load_companias(a.companias)
        cia_of = {u.upper(): a.cia[0] for u in a.codigo_generacion} if a.codigo_generacion else None
        rows = fetch(a.codigo_generacion, a.desde, a.hasta, a.max_docs, a.sql, a.db_config,
                     companias=companias, cia_of=cia_of, cias=a.cia)
    except (DbFetchError, CompaniaError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2
    write_tsv(rows, a.out)
    docs = len({r[0] for r in rows})
    print(f"{docs} documento(s), {len(rows)} valor(es) esperado(s) -> {a.out}")
    if a.codigo_generacion and docs < len(a.codigo_generacion):
        faltan = sorted({u.upper() for u in a.codigo_generacion} - {r[0] for r in rows})
        print(f"Sin resultado en BD: {', '.join(faltan)}", file=sys.stderr)
    return 0


def cmd_validate(argv, verificar: bool = False) -> int:
    ap = argparse.ArgumentParser(
        prog="contract_validator verificar" if verificar else "contract_validator",
        description=("Flujo completo: estructura contra el .xlsx, coherencia XML/JSON/PDF, consulta de la BD (solo lectura) "
                     "y comparación del contenido; termina con un veredicto por documento") if verificar else None)
    ap.add_argument("inputs", nargs="+", help="archivos .xml/.json, globs o carpetas")
    ap.add_argument("--spec", default=DEFAULT_SPEC, help="Excel con el contrato (hoja plana 'contrato')")
    ap.add_argument("--sheet", help="hoja del Excel (default: 'contrato' o la primera)")
    ap.add_argument("--doc-type", help="forzar tipo de documento: CCF, FAC, FAC_EXP, NCR, NDB, REM, INVALIDACION, RETORNO")
    ap.add_argument("--format", choices=["text", "json"], default="text")
    ap.add_argument("--out", help="ruta de reporte (.xlsx o .json)")
    ap.add_argument("--pair", nargs=2, action="append", metavar=("XML", "JSON"),
                    help="cruzar un XML contra el JSON del mismo documento (repetible). Sin esto se emparejan por UUID")
    ap.add_argument("--no-cross", action="store_true", help="no cruzar XML contra JSON")
    ap.add_argument("--pair-pdf", nargs=2, action="append", metavar=("XML", "PDF"),
                    help="validar un PDF contra su XML (repetible). Sin esto se emparejan por Código de Generación")
    ap.add_argument("--no-pdf", action="store_true", help="no procesar archivos .pdf")
    ap.add_argument("--pdf-config", help="normalizaciones de presentación (default: data/specs/pdf_presentacion.json)")
    ap.add_argument("--expected", help="resultado exportado de queries/valores_esperados.sql (.tsv/.csv/.xlsx): "
                                       "compara JSON/XML/PDF contra la BD (DB_MISMATCH_*, DB_MISSING)")
    ap.add_argument("--expected-db", action="store_true",
                    help="consulta la BD (solo lectura) para los Códigos de Generación de los archivos y compara contra ella")
    ap.add_argument("--db-config", help="conexión Oracle sin contraseña (default: config/oracle.local.json; "
                                        "contraseña en la variable CV_ORACLE_PASSWORD)")
    ap.add_argument("--save-expected", help="con --expected-db: guarda lo consultado en este .tsv (contiene datos reales)")
    ap.add_argument("--companias", help="compañías (CIA) y su esquema Oracle (default: data/specs/companias.json)")
    ap.add_argument("--mapeo-bd", help="mapeo aprobado BD -> contrato (default: data/specs/mapeo_bd.xlsx)")
    ap.add_argument("--bd-config", help="normalizaciones BD aprobadas (default: data/specs/normalizaciones_bd.json)")
    ap.add_argument("--no-rules", action="store_true", help="omitir la regla de número de control")
    ap.add_argument("--allow-extra", action="store_true", help="campos no definidos = WARNING en vez de ERROR")
    ap.add_argument("--ajustes", default=os.path.join("data", "specs", "ajustes_contrato.json"),
                    help="reglas generales de validación aprobadas (p. ej. apéndices obligatorios); default: data/specs/ajustes_contrato.json")
    a = ap.parse_args(argv)
    if verificar:
        if not a.expected:
            a.expected_db = True
        if not a.out:
            from datetime import datetime
            a.out = os.path.join("reports", f"verificacion_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

    spec = load_spec(a.spec, a.sheet)
    from .build_contract import load_reglas
    reglas = load_reglas(a.ajustes)
    results, docs, pdfs = [], {}, []
    for f in collect(a.inputs):
        if f.lower().endswith(".pdf"):
            pdfs.append(f)
            continue
        try:
            doc = load_document(f)
            dt = (a.doc_type or "").upper() or doc.doc_type or doc_type_from_filename(f)
            docs[f] = doc
            results.append(compare(f, doc, spec, strict_extra=not a.allow_extra, doc_type=dt, rules=not a.no_rules,
                                   reglas=reglas))
        except Exception as e:  # archivo mal formado
            results.append(Result(f, [Finding("ERROR", "PARSE_ERROR", "", str(e))]))
    if not a.no_cross:
        pairs, _ = pair_documents(docs)
        for x, j in (a.pair or []):
            x, j = next((k for k in docs if os.path.abspath(k) == os.path.abspath(x)), x), next((k for k in docs if os.path.abspath(k) == os.path.abspath(j)), j)
            if x not in docs or j not in docs:
                results.append(Result(f"{x} <-> {j}", [Finding("ERROR", "XREF_PAIR_NOT_LOADED", "", "Alguno de los archivos del par no se pudo leer o no está en la entrada")]))
            elif docs[x].fmt != "XML" or docs[j].fmt != "JSON":
                results.append(Result(f"{x} <-> {j}", [Finding("ERROR", "XREF_PAIR_FORMAT", "", "--pair espera XML y luego JSON")]))
            elif (x, j) not in pairs:
                pairs.append((x, j))
        dt = (a.doc_type or "").upper()
        for x, j in pairs:
            results.append(cross_check(x, docs[x], j, docs[j], spec, doc_type=dt, reglas=reglas))
    if pdfs and not a.no_pdf:
        results += validate_pdfs(pdfs, docs, spec, load_presentation(a.pdf_config), a.pair_pdf, (a.doc_type or "").upper())
    if a.expected and a.expected_db:
        ap.error("use --expected (archivo) o --expected-db (consulta directa), no ambos")
    from .crosscheck import document_id
    from .pdf_text import read_pdf
    db_pdfs = [] if a.no_pdf else pdfs
    pdf_infos = {p: read_pdf(p) for p in db_pdfs}
    pdf_ids = {p: i.uuid for p, i in pdf_infos.items() if i.kind == "MINH" and i.uuid}
    ignored_pdfs = {p: i.uuid or "" for p, i in pdf_infos.items() if i.kind != "MINH"}
    uuid_of = {f: document_id(d) for f, d in docs.items() if d.raw is not None and document_id(d)} | pdf_ids
    unicos = set(uuid_of.values())
    if len(unicos) == 1:  # un PDF MINH que no se pudo leer se asocia al único documento recibido
        uuid_of |= {p: next(iter(unicos)) for p, i in pdf_infos.items() if i.kind == "MINH" and not i.uuid}
    db_status: dict | str = ""
    compania_of: dict = {}
    if a.expected or a.expected_db:
        from .db_check import load_bd_config, validate_db
        from .expected import load_expected
        expected = None
        if a.expected_db:
            from .db_fetch import DbFetchError, fetch, write_tsv
            from .companias import CompaniaError, identify, load_companias
            try:
                # 1. compañía de cada documento -> esquema donde se consulta
                files_by_uuid: dict = {}
                for f, u in uuid_of.items():
                    files_by_uuid.setdefault(u, []).append(f)
                found, sin_cia = identify(files_by_uuid, docs, load_companias(a.companias))
                compania_of = {u: c.etiqueta for u, c in found.items()}
                db_status = dict(sin_cia)
                for u, motivo in sorted(sin_cia.items()):
                    print(f"AVISO: {u}: {motivo}", file=sys.stderr)
                for u, c in sorted(found.items()):
                    print(f"BD: {u} -> {c.etiqueta}", file=sys.stderr)
                # 2. consulta de solo lectura, una vez por compañía
                uuids = sorted(found)
                rows = fetch(uuids, db_config=a.db_config, cia_of={u: c.cia for u, c in found.items()},
                             companias=load_companias(a.companias)) if uuids else []
                print(f"BD: {len(uuids)} documento(s) consultado(s), {len(rows)} valor(es) esperado(s)", file=sys.stderr)
                if a.save_expected:
                    write_tsv(rows, a.save_expected)
                expected = load_expected(rows)
            except (DbFetchError, CompaniaError, OSError) as e:  # sin BD se valida igual la estructura; la información queda NO VERIFICADA
                db_status = f"no se pudo consultar la BD: {e}"
                print(f"AVISO: {db_status}", file=sys.stderr)
        else:
            expected = load_expected(a.expected)
        if expected is not None:
            sin_cia = db_status if isinstance(db_status, dict) else {}  # sin compañía identificada: no se consultaron
            results += validate_db(expected, {f: d for f, d in docs.items() if uuid_of.get(f) not in sin_cia},
                                   [p for p in db_pdfs if uuid_of.get(p) not in sin_cia], spec,
                                   load_bd_config(a.mapeo_bd, a.bd_config), load_presentation(a.pdf_config),
                                   (a.doc_type or "").upper())
    else:
        db_status = "no se solicitó comparar contra la BD (use 'verificar' o --expected-db)"
    from . import verdict
    verdicts = (verdict.build(results, uuid_of, db_status, ignored_pdfs, bool(a.no_pdf and pdfs), compania_of)
                if (verificar or a.expected or a.expected_db) else None)
    if a.format == "json":
        print(report.to_json(results, verdicts, spec))
    else:
        print(report.to_text(results))
        if verdicts is not None:
            print(verdict.to_text(verdicts))
    if a.out:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        if a.out.endswith(".xlsx"):
            report.to_xlsx(results, a.out, verdicts, spec)
        else:
            with open(a.out, "w", encoding="utf-8") as fh:
                fh.write(report.to_json(results, verdicts, spec))
        print(f"Reporte: {a.out}", file=sys.stderr)
    if verdicts is not None:
        return verdict.exit_code(verdicts)
    return 0 if all(r.ok for r in results) else 1


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "build":
        return cmd_build(argv[1:])
    if argv and argv[0] == "audit-mapping":
        return cmd_audit(argv[1:])
    if argv and argv[0] == "audit-plano":
        return cmd_audit_plano(argv[1:])
    if argv and argv[0] == "fetch-expected":
        return cmd_fetch_expected(argv[1:])
    if argv and argv[0] == "verificar":
        return cmd_validate(argv[1:], verificar=True)
    if argv and argv[0] == "validate":
        argv = argv[1:]
    return cmd_validate(argv)


if __name__ == "__main__":
    sys.exit(main())
