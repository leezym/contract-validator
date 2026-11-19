"""Veredicto final por documento (Código de Generación): ¿la estructura cumple el .xlsx? ¿la información es correcta
según la BD?

Agrupa los resultados de todas las validaciones de un mismo documento (JSON, XML, PDF MINH):
  * Estructura vs contrato (.xlsx): comparador de cada JSON/XML        -> CUMPLE | NO CUMPLE | SIN JSON/XML
  * Información vs BD: db_check (JSON y XML contra la BD)               -> CORRECTA | INCORRECTA | NO VERIFICADA
  * Coherencia entre archivos: cruce XML<->JSON                         -> OK | CON DIFERENCIAS | NO APLICA
  * Validación del PDF MINH: PDF vs XML y PDF vs BD                     -> CUMPLE | NO CUMPLE | NO VERIFICADO |
                                                                           NO RECIBIDO | IGNORADO | NO VALIDADO
  * RESULTADO: CUMPLE (sin incumplimientos), NO CUMPLE (algún incumplimiento) o INCOMPLETO (sin incumplimientos, pero
    la información o el PDF recibido no se pudieron verificar). Un PDF no recibido se informa, pero no cambia el resultado.
"""
from __future__ import annotations
import os
from collections import Counter

NO_DOC = "(sin Código de Generación)"


def _kind(r) -> str:
    fmt = r.fmt or ""
    if fmt.startswith("PDF(") or fmt.startswith("BD<->PDF"):
        return "pdf"
    if fmt.startswith("BD<->"):
        return "bd"
    if fmt == "XML<->JSON":
        return "coherencia"
    return "estructura"


def section(r) -> str:
    """Validación a la que pertenece un resultado (columna 'validacion' del reporte)."""
    fmt = r.fmt or ""
    if fmt.startswith("BD<->PDF"):
        return "PDF vs BD"
    if fmt.startswith("PDF("):
        return "PDF vs XML"
    return {"bd": "Información vs BD", "coherencia": "Coherencia XML/JSON"}.get(_kind(r), "Estructura vs contrato")


def _pdf_part(name, rs) -> str:
    if not rs:
        return f"{name}: no comparado"
    errs, warns = _codes(rs), _codes(rs, "WARNING")
    txt = f"{name}: {sum(r.checked for r in rs)} revisado(s), " + (errs or "sin errores")
    return txt + (f"; avisos: {warns}" if warns else "")


def _pdf_state(pdf, ignored, skipped) -> tuple[str, str]:
    if not pdf:
        if ignored:
            names = ", ".join(sorted(os.path.basename(p) for p in ignored))
            return "IGNORADO", f"solo se recibió PDF interno ({names}); se validan los PDF MINH"
        if skipped:
            return "NO VALIDADO", "se omitió la validación de PDF (--no-pdf)"
        return "NO RECIBIDO", "no se recibió PDF MINH de este documento"
    tool = next((f.message for r in pdf for f in r.findings if f.code == "PDF_TOOL_MISSING"), "")
    if tool:
        return "NO VERIFICADO", f"el PDF MINH se recibió, pero no se pudo leer: {tool}"
    vs_xml = [r for r in pdf if not r.fmt.startswith("BD<->")]
    vs_bd = [r for r in pdf if r.fmt.startswith("BD<->")]
    detail = _pdf_part("vs XML", vs_xml) + " | " + _pdf_part("vs BD", vs_bd)
    if any(f.severity == "ERROR" for r in pdf for f in r.findings):
        return "NO CUMPLE", detail
    no_pair = any(f.code == "PDF_NO_PAIR" for r in pdf for f in r.findings)
    if no_pair or not vs_xml or not vs_bd:
        return "NO VERIFICADO", detail
    return "CUMPLE", detail


def _codes(results, sev="ERROR") -> str:
    c = Counter(f.code for r in results for f in r.findings if f.severity == sev)
    return ", ".join(f"{k} x{n}" for k, n in c.most_common())


def _file_of(r) -> str:
    return r.file.split(" <-> ")[0]


def build(results, uuid_of: dict, db_status: str | dict = "", ignored_pdfs: dict | None = None, pdf_skipped: bool = False,
          compania_of: dict | None = None) -> list[dict]:
    """uuid_of: {ruta de archivo: Código de Generación}. db_status: '' si se consultó la BD; si no, el motivo.
    ignored_pdfs: {ruta: Código de Generación o ''} de PDF recibidos que no son MINH.
    db_status también puede ser {Código de Generación: motivo} (p. ej. compañía no identificada).
    compania_of: {Código de Generación: 'CIA nn (nombre, esquema X)'} consultada en la BD."""
    norm = lambda p: os.path.normcase(os.path.abspath(p))
    ids = {norm(k): v for k, v in uuid_of.items()}
    groups: dict[str, list] = {}
    for r in results:
        if r.fmt == "PDF(INTERNO)":
            continue
        u = ids.get(norm(_file_of(r))) or f"{NO_DOC} {os.path.basename(_file_of(r))}"
        groups.setdefault(u, []).append(r)
    ignored: dict[str, list] = {}
    for p, u in (ignored_pdfs or {}).items():
        key = u or f"{NO_DOC} {os.path.basename(p)}"
        ignored.setdefault(key, []).append(p)
        groups.setdefault(key, [])
    out = []
    for u, rs in groups.items():
        by = {k: [r for r in rs if _kind(r) == k] for k in ("estructura", "bd", "coherencia", "pdf")}
        files = sorted({_file_of(r) for r in rs} | set(ignored.get(u, [])))
        exts = sorted({os.path.splitext(f)[1].lstrip(".").upper() for f in files})
        tipo = next((r.doc_type for r in rs if r.doc_type), "")

        est = by["estructura"]
        est_v = "SIN JSON/XML" if not est else ("CUMPLE" if all(r.ok for r in est) else "NO CUMPLE")
        est_d = _codes(est) or ("sin errores" if est else "solo se recibió PDF")

        bd = by["bd"]
        bd_err = [f for r in bd for f in r.findings if f.severity == "ERROR"]
        compared = sum(r.checked for r in bd)
        status = db_status.get(u, "") if isinstance(db_status, dict) else db_status  # por documento (p. ej. sin CIA)
        if status:
            bd_v, bd_d = "NO VERIFICADA", status
        elif not bd:
            bd_v, bd_d = "NO VERIFICADA", "no hubo comparación contra la BD para este documento"
        elif any(f.code.startswith("DB_MISMATCH") for f in bd_err):
            n = sum(1 for f in bd_err if f.code.startswith("DB_MISMATCH"))
            bd_v, bd_d = "INCORRECTA", f"{n} diferencia(s) en {compared} valor(es) comparado(s): {_codes(bd)}"
        elif bd_err:
            bd_v, bd_d = "NO VERIFICADA", "; ".join(sorted({f.message for f in bd_err}))[:300]
        elif compared == 0:  # nada se pudo comparar (p. ej. tipo fuera de alcance)
            bd_v = "NO VERIFICADA"
            bd_d = "; ".join(sorted({f.message for r in bd for f in r.findings})) or "sin valores comparables en la BD"
        else:
            avisos = sum(1 for r in bd for f in r.findings if f.severity == "WARNING")
            bd_v = "CORRECTA"
            bd_d = f"{compared} valor(es) coinciden con la BD" + (f"; {avisos} aviso(s) de campos no verificables" if avisos else "")

        if (compania_of or {}).get(u):
            bd_d = f"{bd_d} [{compania_of[u]}]"

        co = by["coherencia"]
        co_v = "NO APLICA" if not co else ("OK" if all(r.ok for r in co) else "CON DIFERENCIAS")
        co_d = _codes(co) or ("sin diferencias" if co else "no hay otro archivo del mismo documento para cruzar")

        pdf_v, pdf_d = _pdf_state(by["pdf"], ignored.get(u, []), pdf_skipped)

        if est_v == "NO CUMPLE" or bd_v == "INCORRECTA" or co_v == "CON DIFERENCIAS" or pdf_v == "NO CUMPLE":
            res = "NO CUMPLE"
        elif bd_v == "NO VERIFICADA" or est_v == "SIN JSON/XML" or pdf_v == "NO VERIFICADO":
            res = "INCOMPLETO"
        else:
            res = "CUMPLE"
        out.append({"codigo_generacion": u, "tipo_doc": tipo, "archivos": ", ".join(exts), "rutas": files,
                    "estructura": est_v, "estructura_detalle": est_d, "informacion_bd": bd_v, "informacion_bd_detalle": bd_d,
                    "coherencia": co_v, "coherencia_detalle": co_d, "validacion_pdf": pdf_v, "validacion_pdf_detalle": pdf_d,
                    "pdf_ignorados": sorted(ignored.get(u, [])), "resultado": res})
    order = {"NO CUMPLE": 0, "INCOMPLETO": 1, "CUMPLE": 2}
    return sorted(out, key=lambda v: (order[v["resultado"]], v["codigo_generacion"]))


def exit_code(verdicts) -> int:
    """0 = todo CUMPLE; 1 = algún NO CUMPLE; 2 = sin incumplimientos pero algo quedó sin verificar."""
    if any(v["resultado"] == "NO CUMPLE" for v in verdicts):
        return 1
    return 2 if any(v["resultado"] == "INCOMPLETO" for v in verdicts) else 0


def to_text(verdicts) -> str:
    if not verdicts:
        return "VEREDICTO: no se recibieron documentos"
    c = Counter(v["resultado"] for v in verdicts)
    lines = ["", "=" * 100, "VEREDICTO POR DOCUMENTO", "=" * 100]
    for v in verdicts:
        lines += [
            f"Documento {v['codigo_generacion']}  ({v['tipo_doc'] or '?'})  archivos: {v['archivos']}",
            f"  Estructura vs contrato (.xlsx) : {v['estructura']:<14} {v['estructura_detalle']}",
            f"  Información vs base de datos   : {v['informacion_bd']:<14} {v['informacion_bd_detalle']}",
            f"  Coherencia XML / JSON          : {v['coherencia']:<14} {v['coherencia_detalle']}",
            f"  Validación del PDF             : {v['validacion_pdf']:<14} {v['validacion_pdf_detalle']}",
            f"  RESULTADO                      : {v['resultado']}",
            "-" * 100,
        ]
    lines.append(f"TOTAL: {len(verdicts)} documento(s) — CUMPLE: {c['CUMPLE']}, NO CUMPLE: {c['NO CUMPLE']}, "
                 f"INCOMPLETO: {c['INCOMPLETO']}")
    return "\n".join(lines)


def add_sheet(wb, verdicts) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    ws = wb.create_sheet("Veredicto", 0)
    head = ["codigo_generacion", "tipo_doc", "archivos", "estructura_xlsx", "detalle_estructura", "informacion_bd",
            "detalle_bd", "coherencia", "detalle_coherencia", "validacion_pdf", "detalle_pdf", "resultado"]
    ws.append(head)
    color = {"CUMPLE": "C6EFCE", "CORRECTA": "C6EFCE", "OK": "C6EFCE", "NO CUMPLE": "FFC7CE", "INCORRECTA": "FFC7CE",
             "CON DIFERENCIAS": "FFC7CE", "NO VERIFICADA": "FFEB9C", "INCOMPLETO": "FFEB9C", "SIN JSON/XML": "FFEB9C",
             "NO VERIFICADO": "FFEB9C", "NO RECIBIDO": "FFEB9C", "IGNORADO": "FFEB9C", "NO VALIDADO": "FFEB9C"}
    for v in verdicts:
        ws.append([v["codigo_generacion"], v["tipo_doc"], v["archivos"], v["estructura"], v["estructura_detalle"],
                   v["informacion_bd"], v["informacion_bd_detalle"], v["coherencia"], v["coherencia_detalle"],
                   v["validacion_pdf"], v["validacion_pdf_detalle"], v["resultado"]])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in (5, 7, 9, 11))
            if cell.value in color and cell.column in (4, 6, 8, 10, 12):
                cell.fill = PatternFill("solid", fgColor=color[cell.value])
                cell.font = Font(bold=True)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, w in zip("ABCDEFGHIJKL", (40, 9, 14, 14, 50, 15, 60, 16, 45, 15, 60, 13)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B2"
