from __future__ import annotations
import json
import os
from dataclasses import asdict


def to_text(results) -> str:
    lines = []
    for r in results:
        warns = len(r.findings) - len(r.errors)
        lines.append(f"[{'OK' if r.ok else 'FAIL'}] {r.file}  [{r.doc_type or '?'}/{r.fmt or '?'}] (campos del contrato: {r.checked}, errores: {len(r.errors)}, avisos: {warns})")
        for f in r.findings:
            lines.append(f"   {f.severity:<7} {f.code:<22} {f.path}: {f.message}")
    return "\n".join(lines)


def to_json(results, verdicts=None, spec=None) -> str:
    from .explain import explain, field_names
    from .verdict import section
    names = field_names(spec)
    detail = []
    for r in results:
        d = {**asdict(r), "ok": r.ok, "validacion": section(r)}
        d["findings"] = [{**asdict(f), "explicacion": explain(f, r, names)} for f in r.findings]
        detail.append(d)
    if verdicts is None:
        return json.dumps(detail, ensure_ascii=False, indent=2)
    return json.dumps({"veredicto": verdicts, "resultados": detail}, ensure_ascii=False, indent=2)


def to_xlsx(results, path: str, verdicts=None, spec=None):
    """spec: contrato cargado (para nombrar los campos con su Descripcion en la columna 'explicacion')."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from .explain import explain, field_names
    from .verdict import section
    names = field_names(spec)
    pdf_codes = {"NO RECIBIDO": "PDF_NOT_RECEIVED", "IGNORADO": "PDF_IGNORED", "NO VALIDADO": "PDF_NOT_VALIDATED"}
    sin_pdf = [v for v in verdicts or [] if v["validacion_pdf"] in pdf_codes]
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    if verdicts is not None:
        from .verdict import add_sheet
        add_sheet(wb, verdicts)
        wb.active = 0
    ws.append(["archivo", "tipo_doc", "formato", "validacion", "resultado", "campos_contrato", "errores", "avisos"])
    for r in results:
        ws.append([r.file, r.doc_type, r.fmt, section(r), "OK" if r.ok else "FAIL", r.checked, len(r.errors), len(r.findings) - len(r.errors)])
    for v in sin_pdf:
        ws.append([", ".join(os.path.basename(p) for p in v["pdf_ignorados"]) or v["codigo_generacion"], v["tipo_doc"], "PDF",
                   "Validación del PDF", v["validacion_pdf"], 0, 0, 0])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, w in zip("ABCDEFGH", (60, 9, 18, 24, 14, 15, 8, 8)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws = wb.create_sheet("Hallazgos")
    ws.append(["archivo", "tipo_doc", "formato", "validacion", "severidad", "codigo", "campo", "explicacion", "mensaje"])
    for r in results:
        if not r.findings:
            ws.append([r.file, r.doc_type, r.fmt, section(r), "OK", "", "", "Sin hallazgos: todo lo revisado coincide.", "Sin hallazgos"])
        for f in r.findings:
            ws.append([r.file, r.doc_type, r.fmt, section(r), f.severity, f.code, f.path, explain(f, r, names), f.message])
    for v in sin_pdf:
        texto = v["validacion_pdf_detalle"][0].upper() + v["validacion_pdf_detalle"][1:] + "."
        ws.append([", ".join(os.path.basename(p) for p in v["pdf_ignorados"]) or v["codigo_generacion"], v["tipo_doc"], "PDF",
                   "Validación del PDF", "WARNING", pdf_codes[v["validacion_pdf"]], "", texto, texto])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, w in zip("ABCDEFGHI", (45, 9, 16, 22, 9, 22, 40, 80, 70)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, min_col=8, max_col=9):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
