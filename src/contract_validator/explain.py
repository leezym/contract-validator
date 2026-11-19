"""Explicación en lenguaje natural de cada hallazgo (columna 'explicacion' del reporte).

No agrega reglas: reformula el código y el mensaje técnico del hallazgo con el nombre del campo tomado del contrato
(columna Descripcion del libro), el formato del archivo y los valores que ya trae el mensaje.
"""
from __future__ import annotations
import re

DOC_NAMES = {
    "CCF": "Comprobante de Crédito Fiscal", "FAC": "Factura", "FAC_EXP": "Factura de Exportación",
    "NCR": "Nota de Crédito", "NDB": "Nota de Débito", "REM": "Nota de Remisión",
    "INVALIDACION": "evento de Invalidación", "RETORNO": "evento de Retorno",
}
ITEM_ARRAYS = ("cuerpoDocumento", "DatosEntrada.Items")


def field_names(spec) -> dict:
    """{(tipo_doc, ruta con '.'): descripción del libro} para JSON y XML."""
    out = {}
    for s in spec or []:
        if s.description:
            out.setdefault((s.doc_type, s.path), s.description.strip())
    return out


def _fmt_word(fmt: str) -> str:
    f = (fmt or "").upper()
    if "PDF" in f and "BD" in f:
        return "el PDF"
    if f.startswith("BD<->"):
        f = f[5:]
    if f.startswith("PDF"):
        return "el PDF"
    return {"JSON": "el JSON", "XML": "el XML"}.get(f, "el archivo")


def _doc(doc_type: str) -> str:
    return DOC_NAMES.get(doc_type, "documento")


def _field(path: str, doc_type: str, names: dict) -> str:
    """«Descripción» (ruta) [del ítem n]."""
    if not path:
        return "el documento"
    base, idx = path, ""
    m = re.match(r"^(.*?)\[([^\]]+)\]$", path)
    if m:
        base, idx = m.group(1), m.group(2)
    desc = names.get((doc_type, base)) or names.get((doc_type, base.replace("/", ".")))
    txt = f"«{desc}» ({base})" if desc else f"el campo {base}"
    if idx:
        if idx.isdigit():
            txt += f" del ítem {idx}" if base.startswith(ITEM_ARRAYS) else f" (posición {idx})"
        else:
            txt += f" ({idx})"
    return txt


def _vals(msg: str, a: str, b: str):
    """Extrae los valores de 'A=x B=y' (con o sin comillas)."""
    m = re.search(rf"{a}=('(?:[^']|'')*'|\S+)\s+{b}=('(?:[^']|'')*'|.+)$", msg)
    if not m:
        return None
    return tuple(v.strip().strip("'") for v in m.groups())


def _clean(v: str) -> str:
    return "vacío (null)" if v in ("null", "None", "") else f"«{v}»"


def explain(f, r, names: dict | None = None) -> str:
    """Texto en lenguaje natural para el hallazgo f del resultado r."""
    return _fix(_explain(f, r, names))


def _fix(t: str) -> str:
    """Contracciones del español (de el -> del, a el -> al)."""
    t = re.sub(r"\b([Dd])e el\b", r"\1el", t)
    return re.sub(r"\b([Aa]) el\b", r"\1l", t)


def _explain(f, r, names: dict | None = None) -> str:
    names = names or {}
    code, msg, path, dt = f.code, f.message or "", f.path or "", r.doc_type or ""
    fw = _fmt_word(r.fmt)
    field = _field(path, dt, names)
    doc = _doc(dt)

    attr = re.fullmatch(r"apendice\[([^\]]+)\]\.(\w+)", path)
    if attr and code in ("REQUIRED_MISSING", "NULL_NOT_ALLOWED"):
        estado = "no llegó" if code == "REQUIRED_MISSING" else "llegó vacío"
        return f"En el apéndice «{attr.group(1)}» de {fw}, el atributo «{attr.group(2)}» {estado}. Cada apéndice debe traer campo, etiqueta y valor."
    if path.startswith("apendice[") and path.endswith("]") and code != "APENDICE_MISSING" and code != "APENDICE_EXTRA":
        name = path[9:-1]
        m = re.search(r"Apéndice \w+, (\w+): (BD|JSON)=(.+?) (documento|XML)=(.+)$", msg)
        parte = {"etiqueta": "las etiquetas", "valor": "los datos unidos"}.get(m.group(1), "") if m else ""
        if code == "APENDICE_NOT_RECEIVED":
            return (f"El apéndice «{name}» se arma uniendo varios datos de la base de datos, pero no llegó en {fw}. "
                    "El libro no lo marca como requerido: confirme que su ausencia es correcta.")
        falta = re.search(r"BD=(.+) falta en el documento$", msg)
        if code.startswith("DB_MISMATCH") and falta:
            return f"El apéndice «{name}» no llegó en {fw}. Según la base de datos debe traer los datos unidos {falta.group(1)}."
        if code.startswith("DB_MISMATCH") and m:
            return f"En el apéndice «{name}» de {fw}, {parte} no coinciden con la base de datos: se esperaba {m.group(3)} y llegó {m.group(5)}."
        if code == "XREF_MISMATCH" and m:
            return f"En el apéndice «{name}», {parte} difieren entre archivos: el JSON trae {m.group(3)} y el XML {m.group(5)}."
        if code == "XREF_MISSING":
            return f"El apéndice «{name}» no llegó completo en ambos archivos: {msg.split(': ', 1)[-1]}."

    # ---------------- estructura contra el contrato (.xlsx)
    if code in ("REQUIRED_MISSING", "NEW_MISSING", "RENAMED_MISSING"):
        if msg.startswith("Contenedor"):
            return f"Falta el bloque {path} en {fw}; el contrato exige campos dentro de él para el {doc}."
        extra = {"NEW_MISSING": " Es un campo nuevo de la versión 4 del contrato.",
                 "RENAMED_MISSING": " El contrato lo renombró: debe venir con su nombre nuevo."}.get(code, "")
        src = re.search(r"\(cambio: [^,]+, (.+)\)$", msg)
        items = re.search(r"\(ítems ([\d, ]+) de (\d+)\)", msg)
        if items:
            field += f" en los ítems {items.group(1)} de {items.group(2)}"
        return f"Falta {field} en {fw}. El contrato lo exige para el {doc}.{extra}" + (f" Definido en: {src.group(1)}." if src else "")
    if code == "NULL_NOT_ALLOWED":
        return f"{field[0].upper() + field[1:]} viene vacío (null) en {fw}, pero el contrato exige que traiga un valor."
    if code == "VALUE_NOT_NULL":
        return f"{field[0].upper() + field[1:]} debe venir con valor null en {fw} para el {doc}, pero trae información."
    if code == "TYPE_MISMATCH":
        m = re.search(r"esperado '(\w+)', encontrado \[(.*?)\]", msg)
        return (f"{field[0].upper() + field[1:]} no tiene el tipo de dato que indica el contrato"
                + (f": debe ser {m.group(1)} y llegó como {m.group(2).replace(chr(39), '')}" if m else "") + f" en {fw}.")
    if code == "XML_OPTIONAL_MISSING":
        return (f"{field[0].upper() + field[1:]} está definido en el contrato pero no llegó en {fw}. El libro no lo marca "
                "como requerido, así que no es un error: confirme que su ausencia es correcta para este documento.")
    if code == "REMOVED_STILL_PRESENT":
        return f"{fw[0].upper() + fw[1:]} todavía trae {field}, pero el contrato lo eliminó en la versión 4: debe quitarse."
    if code == "OLD_NAME_PRESENT":
        new = re.search(r"debe ser '([^']+)'", msg)
        return f"{fw[0].upper() + fw[1:]} usa el nombre anterior {path}; según el contrato ahora debe llamarse {new.group(1) if new else 'con el nombre nuevo'}."
    if code == "UNEXPECTED_FIELD":
        return (f"{fw[0].upper() + fw[1:]} trae el campo {path}, que el contrato no define para el {doc}. "
                "Si debe aceptarse, hay que agregarlo al contrato (ajuste aprobado); si no, el generador debe dejar de enviarlo.")
    if code == "VALUE_NOT_IN_ENUM":
        return f"{field[0].upper() + field[1:]} trae un valor que no está en la lista de valores permitidos del contrato."
    if code == "FORMAT_INVALID":
        return f"El número de control no tiene el formato esperado DTE-##-<establecimiento y punto>-<15 dígitos>. {msg.split(' debe ser')[0]}."
    if code == "CONSISTENCY_ERROR":
        return f"El número de control indica un tipo de documento distinto del documento: {msg}."
    if code == "DOCTYPE_UNKNOWN":
        return "No se pudo saber qué tipo de documento es (el archivo no trae el tipo y el nombre no lo indica); indíquelo con --doc-type."
    if code == "DOCTYPE_NAME_MISMATCH":
        return f"El nombre del archivo sugiere otro tipo de documento, pero su contenido es un {doc}; se validó como {doc}."
    if code == "ENCODING":
        return "El XML no venía en UTF-8 y se leyó como Windows-1252: revise que tildes y eñes se vean bien."
    if code == "APENDICE_MISSING":
        m = re.search(r"'(\w+)' \(obligatorios: (.+)\)", msg)
        n = m.group(1) if m else path
        return (f"Falta el apéndice obligatorio «{n}» en {fw}. Los apéndices {m.group(2) if m else 'apendice01 a apendice04'} "
                f"deben venir siempre en el {doc}.")
    if code == "APENDICE_EXTRA":
        m = re.search(r"'(\w+)' con valor '(.*)': ", msg)
        n, v = (m.group(1), m.group(2)) if m else (path, "")
        return (f"{fw[0].upper() + fw[1:]} trae el apéndice adicional «{n}» (valor «{v}»), que no está entre los obligatorios. "
                "No es un error: revíselo manualmente para confirmar que corresponde a este documento.")
    if code == "PARSE_ERROR":
        return f"No se pudo leer el archivo (está dañado o no es un JSON/XML válido): {msg}"

    # ---------------- coherencia XML <-> JSON
    if code == "XREF_MISMATCH":
        v = _vals(msg, "JSON", "XML")
        if v:
            return f"El JSON y el XML del mismo documento no coinciden en {field}: el JSON dice {_clean(v[0])} y el XML dice {_clean(v[1])}."
    if code == "XREF_FORMAT_DIFF":
        v = _vals(msg, "JSON", "XML")
        if v:
            return f"{field[0].upper() + field[1:]} tiene el mismo valor en el JSON y en el XML, pero escrito distinto ({v[0]} vs {v[1]}). Es solo de formato."
    if code == "XREF_MISSING":
        m = re.search(r"existe en (JSON|XML) \((.*)\) y falta en (JSON|XML)", msg)
        if m:
            return f"{field[0].upper() + field[1:]} viene en el {m.group(1)} ({_clean(m.group(2).strip(chr(39)))}) pero no en el {m.group(3)}."
    if code == "XREF_COUNT_MISMATCH":
        m = re.search(r"JSON tiene (\d+) valor\(es\) y XML (\d+)", msg)
        if m:
            return f"El JSON y el XML no traen la misma cantidad de elementos en {field}: el JSON tiene {m.group(1)} y el XML {m.group(2)}."
    if code == "XREF_DOCTYPE":
        return f"El XML y el JSON emparejados son de tipos de documento distintos: {msg}."
    if code in ("XREF_PAIR_NOT_LOADED", "XREF_PAIR_FORMAT"):
        return f"No se pudo cruzar el par XML/JSON indicado: {msg}."

    # ---------------- PDF oficial vs XML
    if code == "PDF_IDENTITY":
        return f"El PDF no corresponde a este XML: {msg}."
    if code == "PDF_DOCTYPE":
        return f"El título impreso en el PDF no coincide con el tipo de documento del XML: {msg}."
    if code == "PDF_MISMATCH":
        m = re.match(r"'(.+?)': PDF=(.+?) XML=(.+)$", msg)
        if m:
            return f"En el PDF, «{m.group(1)}» muestra {_clean(m.group(2).strip(chr(39)))}, pero el XML tiene {_clean(m.group(3).strip(chr(39)))}."
        m = re.match(r"Ítem (\S+) \((.+?)\): (\w+) XML=(.+) no figura", msg)
        if m:
            return f"En el PDF, la fila del ítem {m.group(1)} (código {m.group(2)}) no muestra la {m.group(3)} {m.group(4)} que trae el XML."
    if code in ("PDF_FORMAT_DIFF", "PDF_TRUNCATED"):
        m = re.match(r"'(.+?)': (XML|BD)=(\S+) PDF=(\S+)", msg)
        if m:
            fuente = "el XML" if m.group(2) == "XML" else "la BD"
            como = ("es el mismo valor con otro formato (mayúsculas/minúsculas)" if code == "PDF_FORMAT_DIFF"
                    else "el PDF lo imprime cortado")
            return (f"En el PDF, «{m.group(1)}» muestra {m.group(4)} y {fuente} trae {m.group(3)}: {como}. "
                    "Es un aviso, no un error.")
    if code == "PDF_FORMAT_DIFF":
        m = re.match(r"Ítem (\S+) \((.+?)\): (\w+) (XML|BD)=(\S+) PDF=(\S+)", msg)
        if m:
            fuente = "el XML" if m.group(4) == "XML" else "la BD"
            return (f"En la fila del ítem {m.group(1)} (código {m.group(2)}) el PDF muestra la {m.group(3)} como "
                    f"{m.group(6)} y {fuente} trae {m.group(5)}: es el mismo valor con otro formato (menos decimales). "
                    "Es un aviso, no un error.")
    if code == "PDF_VALUE_MISSING":
        return f"Un dato que trae el XML no aparece impreso en el PDF: {msg}."
    if code == "PDF_COUNT_MISMATCH":
        m = re.search(r"PDF tiene (\d+) fila.* XML (\d+)", msg)
        if m:
            return f"El PDF muestra {m.group(1)} ítem(s) y el XML trae {m.group(2)}."
    if code == "PDF_NO_PAIR":
        return "No se recibió el XML de este PDF, así que no se pudo comparar lo impreso contra el XML."
    if code == "PDF_NOT_PRINTED":
        m = re.search(r"Valor del XML '(.*)' no aparece", msg)
        return (f"El valor {_clean(m.group(1)) if m else ''} de {field} está en el XML pero no aparece en el PDF. "
                "Si el PDF no lo imprime por diseño, puede marcarse como 'no_impreso'.")
    if code == "PDF_VALUE_NOT_IN_XML":
        m = re.search(r"importe (\S+)", msg)
        return f"El PDF muestra el importe {m.group(1) if m else ''}, que no aparece entre los valores del XML."
    if code == "PDF_IGNORED":
        return "Este es el PDF interno de TQ; solo se revisan los PDF oficiales de Hacienda (MINH)."
    if code == "PDF_EMPTY":
        return "El PDF oficial no tiene texto legible."
    if code == "PDF_TOOL_MISSING":
        return "No se pudo leer el PDF porque falta la herramienta pdftotext en el equipo."

    # ---------------- información contra la base de datos
    if code.startswith(("DB_MISMATCH", "XREF_MISMATCH")):
        more = re.search(r": (\d+) diferencia\(s\) más$", msg)
        if more:
            return f"Además, {field} tiene {more.group(1)} diferencia(s) más del mismo tipo (se muestran las primeras para no repetir)."
    if code.startswith("DB_MISMATCH"):
        k = re.search(r"\(elemento con (\w+)='([^']+)', clave según el contrato\): BD=('(?:[^']|'')*'|\S+) (.*)$", msg)
        if k:
            bd = _clean(k.group(3).strip(chr(39)))
            parent = re.search(r"<-> ([\w.]+)\.\w+ \(", msg)
            elem = f"<{parent.group(1).rsplit('.', 1)[-1]}>" if parent else "el elemento"
            if k.group(4).startswith("falta"):
                return (f"Según la base de datos, {field} debería ser {bd}, pero {fw} no lo trae: no hay ningún {elem} "
                        f"cuyo <{k.group(1)}> sea «{k.group(2)}», que es como el contrato identifica ese elemento.")
            got = _vals(msg, "BD", "documento")
            return f"Según la base de datos, {field} debería ser {bd}, pero {fw} trae {_clean(got[1]) if got else 'otro valor'}."
        v = _vals(msg, "BD", "documento")
        if v:
            return f"Según la base de datos, {field} debería ser {_clean(v[0])}, pero {fw} trae {_clean(v[1])}."
        m = re.search(r"BD=('(?:[^']|'')*'|\S+) falta en el documento", msg)
        if m:
            return f"Según la base de datos, {field} debería ser {_clean(m.group(1).strip(chr(39)))}, pero {fw} no lo trae."
        m = re.search(r"La BD tiene (\d+) ítem\(s\) y el documento (\d+)", msg)
        if m:
            return f"La base de datos registra {m.group(1)} ítem(s) para este documento, pero {fw} trae {m.group(2)}."
        m = re.match(r"'(.+?)': BD=(.+?) PDF=(.+)$", msg)
        if m:
            return f"En el PDF, «{m.group(1)}» muestra {_clean(m.group(3).strip(chr(39)))}, pero según la base de datos debería ser {_clean(m.group(2).strip(chr(39)))}."
        m = re.match(r"(Código de Generación|Número de Control): BD=(.+?) PDF=(.+)$", msg)
        if m:
            return f"El {m.group(1)} impreso en el PDF ({m.group(3).strip(chr(39))}) no es el de la base de datos ({m.group(2).strip(chr(39))})."
        m = re.match(r"Ítem (\d+) \((.+?)\): (\w+) BD=(.+) no figura", msg)
        if m:
            return f"En el PDF, la fila del ítem {m.group(1)} (código {m.group(2)}) no muestra el valor {m.group(4)} de {m.group(3)} que registra la base de datos."
        m = re.match(r"Ítem (\d+): el código (\S+) de la BD no aparece", msg)
        if m:
            return f"El ítem {m.group(1)} (código {m.group(2)}) que registra la base de datos no aparece en el PDF."
    if code == "DB_MISSING":
        if f.severity == "ERROR":
            return ("Este documento no se encontró en la base de datos consultada (por su Código de Generación), "
                    "así que su información no se pudo verificar. Confirme el ambiente (pruebas o producción).")
        if "fuera del alcance" in msg:
            return "Las invalidaciones y los retornos todavía no se comparan contra la base de datos; su información no se verificó."
        m = re.search(r"(\d+) campo\(s\) sin fuente en BD \(no verificables\): (.+)$", msg)
        if m:
            return (f"{m.group(1)} campo(s) de {fw} no se pueden verificar contra la base de datos porque no tienen un dato "
                    f"equivalente allí ({m.group(2)}). Es informativo, no un error del documento.")
        m = re.search(r"(\d+) campo\(s\) con valor en el documento y sin valor en BD: (.+)$", msg)
        if m:
            return (f"{fw[0].upper() + fw[1:]} trae {m.group(1)} campo(s) para los que la base de datos no tiene valor "
                    f"({m.group(2)}); no se pudieron verificar.")
        if msg.startswith("No se pudo ubicar"):
            return (f"Algunos valores no se compararon contra la base de datos porque el mapeo del contrato no permite "
                    f"ubicarlos en {fw}: {msg.split(': ', 1)[-1]}")
    return msg
