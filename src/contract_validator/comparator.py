"""Compara los campos aplanados de un documento contra el contrato del Excel."""
from __future__ import annotations
from .flatten import doc_type_from_filename
from .models import Finding, Result, NEW, REMOVED, RENAMED, SpecField

COMPAT = {("integer", "number")}  # un integer es válido donde se espera number
CONTAINERS = {"object", "array"}


def _src(s) -> str:
    """Origen de la regla: hoja y fila del libro, o el id del ajuste (fila 0)."""
    return s.sheet if not s.row else f"{s.sheet} fila {s.row}"


def _type_ok(expected: str, found: set[str], fmt: str = "") -> bool:
    if not expected or expected in found:
        return True
    if fmt == "XML" and expected == "string":  # en XML todo es texto: solo se exige que no sea contenedor
        return not (found & CONTAINERS)
    return all(f == expected or f == "null" or (f, expected) in COMPAT for f in found)


def _is_null_or_empty(path: str, doc) -> bool:
    """El campo existe pero solo como null, o es un array/objeto sin hijos (nada que exigir debajo)."""
    kinds = doc[path]
    if kinds == {"null"}:
        return True
    return bool(kinds & CONTAINERS) and not any(p.startswith(path + ".") for p in doc)


def _ancestors(path: str):
    parts = path.split(".")
    return [".".join(parts[:i]) for i in range(1, len(parts))]


def select_spec(spec: list[SpecField], doc_type: str, fmt: str) -> list[SpecField]:
    return [s for s in spec if s.doc_type in ("", doc_type) and s.fmt in ("", fmt)]


def compare(file: str, doc, spec: list[SpecField], strict_extra: bool = True,
            doc_type: str = "", fmt: str = "", rules: bool = True, reglas: dict | None = None) -> Result:
    """reglas: reglas generales de validación activas (p. ej. 'apendice_obligatorio' de ajustes_contrato.json)."""
    content_type = getattr(doc, "doc_type", "")
    doc_type = doc_type or content_type
    fmt = fmt or getattr(doc, "fmt", "")
    res = Result(file=file, doc_type=doc_type, fmt=fmt)

    def add(sev, code, path, msg):
        res.findings.append(Finding(sev, code, path, msg))

    for code, msg in getattr(doc, "warnings", []):
        add("WARNING", code, "", msg)

    hint = doc_type_from_filename(file)
    if hint and content_type and hint != content_type and doc_type == content_type:
        add("WARNING", "DOCTYPE_NAME_MISMATCH", "",
            f"El nombre del archivo sugiere {hint} pero el contenido es {content_type}; se valida como {content_type}")

    known_types = {s.doc_type for s in spec if s.doc_type}
    if known_types and doc_type not in known_types:
        add("ERROR", "DOCTYPE_UNKNOWN", "",
            f"No se pudo determinar el tipo de documento (detectado: '{doc_type or 'ninguno'}'; "
            f"contrato define: {sorted(known_types)}). Use --doc-type.")
        return res
    full_spec = spec
    json_removed = {s.path for s in full_spec if s.doc_type == doc_type and s.fmt == "JSON" and s.change == REMOVED}
    spec = select_spec(spec, doc_type, fmt)
    by_path = {s.path: s for s in spec if s.change != REMOVED}

    expected = {s.path for s in spec if s.change != REMOVED}
    removed = {s.path for s in spec if s.change == REMOVED}
    old_names = {s.old_path for s in spec if s.change == RENAMED}
    implicit = {a for e in expected for a in _ancestors(e)}  # contenedores implícitos (p.ej. XML DatosEntrada.Items)
    reported_containers: set[str] = set()

    for s in spec:
        res.checked += 1
        present = s.path in doc
        if s.change == REMOVED:
            if present:
                add("ERROR", "REMOVED_STILL_PRESENT", s.path, f"Campo eliminado en el contrato pero sigue presente ({_src(s)})")
            continue
        if s.change == RENAMED and s.old_path in doc:
            add("ERROR", "OLD_NAME_PRESENT", s.old_path, f"Nombre anterior aún presente; debe ser '{s.path}' ({_src(s)})")
        if not present:
            if not s.required:
                # XML: un tag del libro que no llega nunca pasa en silencio. Se excluyen los campos TQ agregados por
                # ajuste (no están en el libro) y los tags cuyo campo JSON el contrato eliminó.
                if (fmt == "XML" and s.doc_type and s.row and s.counterpart not in json_removed
                        and "posicional" not in (s.note or "")):
                    add("WARNING", "XML_OPTIONAL_MISSING", s.path,
                        f"Tag del contrato no recibido; el libro no lo marca como requerido ({_src(s)})")
                continue
            absent = next((a for a in _ancestors(s.path) + [s.path] if a not in doc), s.path)
            if absent != s.path:
                if absent in by_path and not by_path[absent].required:
                    continue  # el contenedor opcional no viene: no se exige su contenido
                if absent in by_path:
                    continue  # el contenedor requerido ya se reporta por su propia fila
                if absent in reported_containers:
                    continue
                reported_containers.add(absent)
                add("ERROR", "REQUIRED_MISSING", absent, f"Contenedor requerido no encontrado (contiene campos requeridos, p.ej. '{s.path}')")
                continue
            anc_null = any(a in doc and _is_null_or_empty(a, doc) for a in _ancestors(s.path))
            if anc_null:
                continue
            code = {NEW: "NEW_MISSING", RENAMED: "RENAMED_MISSING"}.get(s.change, "REQUIRED_MISSING")
            add("ERROR", code, s.path, f"Campo requerido no encontrado (cambio: {s.change}, {_src(s)})")
            continue
        kinds = doc[s.path]
        if "solo_null" in (s.note or "") and kinds != {"null"}:
            add("ERROR", "VALUE_NOT_NULL", s.path, f"El campo debe venir con valor null ({_src(s)})")
            continue
        if s.required and not s.nullable and kinds == {"null"} and s.type not in CONTAINERS:
            add("ERROR", "NULL_NOT_ALLOWED", s.path, f"Campo requerido con valor nulo ({_src(s)})")
        # Campo con catálogo en el libro (regla catalogos_como_string): texto, y también se acepta un número sin comillas.
        catalogo_ok = "catalogos_como_string" in (s.note or "") and kinds <= {"string", "integer", "number", "null"}
        if s.type and not catalogo_ok and not _type_ok(s.type, kinds, fmt):
            add("ERROR", "TYPE_MISMATCH", s.path, f"Tipo esperado '{s.type}', encontrado {sorted(kinds)} ({_src(s)})")
        if s.values:
            samples = getattr(doc, "samples", {}).get(s.path, set())
            bad = sorted(v for v in samples if v not in s.values)
            if bad:
                add("WARNING", "VALUE_NOT_IN_ENUM", s.path, f"Valores {bad} fuera de los permitidos {list(s.values)[:8]}{'…' if len(s.values) > 8 else ''}")

    for p in sorted(doc):
        if p in expected or p in removed or p in old_names or p in implicit:
            continue
        add("ERROR" if strict_extra else "WARNING", "UNEXPECTED_FIELD", p, "Campo presente en el archivo pero no definido en el contrato")
    if rules and getattr(doc, "raw", None) is not None:
        from .rules import check, check_apendice, check_apendice_attributes, check_xml_occurrences
        for sev, code, path, msg in check(doc, doc_type):
            add(sev, code, path, msg)
        # el XML ubica cada apéndice por la contraparte de apendice.campo (contrato con ajustes)
        for sev, code, path, msg in check_xml_occurrences(doc, doc_type, full_spec) + check_apendice_attributes(doc, doc_type, full_spec):
            add(sev, code, path, msg)
        key_path = next((s.counterpart for s in full_spec if s.doc_type == doc_type and s.fmt == "JSON"
                         and s.path == "apendice.campo" and s.counterpart), "DatosEntrada.Notas.etiqueta")
        for sev, code, path, msg in check_apendice(doc, doc_type, (reglas or {}).get("apendice_obligatorio"), key_path):
            add(sev, code, path, msg)
    return res
