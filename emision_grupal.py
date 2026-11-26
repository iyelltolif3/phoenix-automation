"""
Phoenix v4 - Emisión Grupal de Viajes (Playwright)
===================================================
Extiende la emisión individual en Playwright para soportar carga masiva
de pasajeros mediante fichero Excel + validación estricta + notificación
automática al cliente cuando el formato es incorrecto.

Plantilla oficial Phoenix descargada desde #plantillaExcel (hoja 'Sheet1'):
- ID Type | Fiscal ID | First name | Surname | D.o.b. | Age | Gender | Email | Mobile

También mantiene compatibilidad de lectura para planillas de clientes en español
(Formato 1 y Formato 2 del manual) y las normaliza al formato oficial al subirlas.
"""

import io
import os
import re
import time
import logging
import unicodedata
import zipfile
from datetime import date, datetime
from xml.sax.saxutils import escape as xml_escape
import openpyxl

from phoenix_emision import PhoenixAutomation, DatosEmision, Pasajero, _calcular_edad
from utils import (
    cargar_config,
    determinar_producto,
    debe_enviar_precios,
    generar_resumen_emision,
    limpiar_rut,
    validar_rut_modulo11,
)
from email_sender import (
    enviar_voucher_al_asegurado,
    enviar_correo,
    enviar_aviso_formato_incorrecto,
)

logger = logging.getLogger(__name__)

COLUMNAS_OFICIAL_PHOENIX = [
    "ID Type",
    "Fiscal ID",
    "First name",
    "Surname",
    "D.o.b.",
    "Age",
    "Gender",
    "Email",
    "Mobile",
]

COLUMNAS_FORMATO_1 = [
    "Tipo de Documento",
    "Número de Documento",
    "DV",
    "Nombres",
    "Apellidos",
    "Fecha de Nacimiento",
    "Edad",
    "Teléfono Móvil",
    "Email",
]

COLUMNAS_FORMATO_2 = [
    "Tipo de Documento",
    "Número de Documento",
    "DV",
    "Nombres",
    "Apellidos",
    "Fecha de Nacimiento",
    "Edad",
    "Género",
    "Cod. País Móvil",
    "Teléfono Móvil",
    "Email",
]


def _norm_header(texto: str) -> str:
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def _parsear_fecha(val) -> datetime | None:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime(val.year, val.month, val.day)
    txt = str(val).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(txt, fmt)
        except ValueError:
            continue
    return None


def _cargar_workbook_seguro(ruta_excel: str) -> openpyxl.Workbook:
    """
    Carga un .xlsx incluso si fue generado por Phoenix con <x:fill /> vacío en styles.xml,
    lo cual provoca TypeError en openpyxl.
    """
    try:
        return openpyxl.load_workbook(ruta_excel, data_only=True)
    except TypeError:
        buf = io.BytesIO()
        with zipfile.ZipFile(ruta_excel, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    xml_txt = data.decode("utf-8", errors="replace")
                    xml_txt = re.sub(
                        r"<(x:)?fill\s*/>",
                        lambda m: f"<{m.group(1) or ''}fill><{m.group(1) or ''}patternFill patternType=\"none\" /></{m.group(1) or ''}fill>",
                        xml_txt,
                    )
                    data = xml_txt.encode("utf-8")
                zout.writestr(item, data)
        buf.seek(0)
        return openpyxl.load_workbook(buf, data_only=True)


def _codigo_tipo_doc_grupal(tipo_doc: str, config: dict | None = None) -> int:
    cfg = config or cargar_config()
    mapa = cfg.get("tipo_documento_grupal") or {"RUN": 1007, "PASAPORTE": 2}
    clave = tipo_doc.strip().upper()
    if clave in mapa:
        return int(mapa[clave])
    return 1007 if clave == "RUN" else 2


def _excel_serial_date(d: date) -> int:
    """Convierte una fecha Python a número serial de Excel (sistema 1900 con bug bisiesto de Lotus)."""
    return (d - date(1899, 12, 30)).days


def _escribir_xlsx_oficial_phoenix(ruta_salida: str, filas_datos: list[dict]):
    """
    Genera un .xlsx con la estructura OpenXML exacta de la plantilla oficial descargada desde
    Phoenix (#plantillaExcel): xl/worksheets/sheet.xml, styles.xml con dd/MM/yyyy (s=1) y enteros (s=3),
    y cadenas inline t="str" sin sharedStrings.xml. Validado en vivo contra Phoenix PRE.
    """
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)

    rows_xml = [
        '<x:row r="1">'
        + "".join(f'<x:c t="str"><x:v>{xml_escape(col)}</x:v></x:c>' for col in COLUMNAS_OFICIAL_PHOENIX)
        + "</x:row>"
    ]

    for idx, fila in enumerate(filas_datos, start=2):
        dob = fila["dob"]
        if isinstance(dob, datetime):
            dob = dob.date()
        dob_cell = (
            f'<x:c s="1"><x:v>{_excel_serial_date(dob)}</x:v></x:c>'
            if isinstance(dob, date)
            else f'<x:c t="str"><x:v>{xml_escape(str(dob))}</x:v></x:c>'
        )
        mobile_val = re.sub(r"\D", "", str(fila["mobile"] or "")) or "900000000"
        row_xml = (
            f'<x:row r="{idx}">'
            f'<x:c s="3"><x:v>{int(fila["id_type"])}</x:v></x:c>'
            f'<x:c t="str"><x:v>{xml_escape(str(fila["fiscal_id"]))}</x:v></x:c>'
            f'<x:c t="str"><x:v>{xml_escape(str(fila["first_name"]))}</x:v></x:c>'
            f'<x:c t="str"><x:v>{xml_escape(str(fila["surname"]))}</x:v></x:c>'
            f"{dob_cell}"
            f'<x:c s="3"><x:v>{int(fila["age"])}</x:v></x:c>'
            f'<x:c s="3"><x:v>{int(fila["gender"])}</x:v></x:c>'
            f'<x:c t="str"><x:v>{xml_escape(str(fila["email"]))}</x:v></x:c>'
            f'<x:c s="3"><x:v>{int(mobile_val)}</x:v></x:c>'
            f"</x:row>"
        )
        rows_xml.append(row_xml)

    sheet_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<x:sheetData>{"".join(rows_xml)}</x:sheetData>'
        "</x:worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<x:sheets><x:sheet name="Sheet1" sheetId="1" r:id="R3a666ff061f64190" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></x:sheets>'
        "</x:workbook>"
    )
    rels_root = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="/xl/workbook.xml" Id="R52f09e11e9fc4059" />'
        "</Relationships>"
    )
    wb_rels = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="/xl/worksheets/sheet.xml" Id="R3a666ff061f64190" />'
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="/xl/styles.xml" Id="rId3" />'
        "</Relationships>"
    )
    styles_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<x:numFmts>'
        '<x:numFmt numFmtId="0" formatCode="dd/MM/yyyy" />'
        '<x:numFmt numFmtId="100" formatCode="0.###" />'
        '<x:numFmt numFmtId="200" formatCode="0" />'
        '</x:numFmts>'
        '<x:fonts><x:font /></x:fonts>'
        '<x:fills><x:fill /></x:fills>'
        '<x:borders><x:border /></x:borders>'
        '<x:cellStyleXfs><x:xf /></x:cellStyleXfs>'
        '<x:cellXfs>'
        '<x:xf />'
        '<x:xf numFmtId="0" applyNumberFormat="1" />'
        '<x:xf numFmtId="100" applyNumberFormat="1" />'
        '<x:xf numFmtId="200" applyNumberFormat="1" />'
        '</x:cellXfs>'
        '</x:styleSheet>'
    )
    content_types = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />'
        '<Override PartName="/xl/worksheets/sheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />'
        '</Types>'
    )

    with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.encode("utf-8-sig"))
        zf.writestr("_rels/.rels", rels_root.encode("utf-8-sig"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8-sig"))
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels.encode("utf-8-sig"))
        zf.writestr("xl/styles.xml", styles_xml.encode("utf-8-sig"))
        zf.writestr("xl/worksheets/sheet.xml", sheet_xml.encode("utf-8-sig"))


def generar_plantilla_optima_excel(
    ruta_salida: str | None = None,
    formato: str | int = "oficial",
    con_ejemplo: bool = True,
) -> str:
    """Genera una plantilla Excel limpia con el formato oficial esperado por Phoenix ('Sheet1')."""
    if not ruta_salida:
        suf = "oficial" if str(formato).lower() == "oficial" else f"formato_{formato}"
        ruta_salida = os.path.join(os.path.dirname(__file__), "ficheros", f"plantilla_grupal_{suf}.xlsx")
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)

    if str(formato).lower() == "oficial":
        filas = []
        if con_ejemplo:
            filas.append({
                "id_type": 1007,
                "fiscal_id": "12345678-9",
                "first_name": "JUAN ANTONIO",
                "surname": "PEREZ ROJAS",
                "dob": date(1990, 12, 9),
                "age": 35,
                "gender": 5,
                "email": "ejemplo@correo.cl",
                "mobile": "981571641",
            })
        _escribir_xlsx_oficial_phoenix(ruta_salida, filas)
        return ruta_salida

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asegurados"
    columnas = COLUMNAS_FORMATO_1 if int(formato) == 1 else COLUMNAS_FORMATO_2
    ws.append(columnas)
    if con_ejemplo:
        if int(formato) == 1:
            ws.append([1007, "12345678", "9", "JUAN ANTONIO", "PEREZ ROJAS", "09-12-1990", 35, "981571641", "ejemplo@correo.cl"])
        else:
            ws.append([1007, "12345678", "9", "JUAN ANTONIO", "PEREZ ROJAS", "09-12-1990", 35, 5, 56, "981571641", "ejemplo@correo.cl"])

    for col_cells in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = max(max_len + 4, 14)

    wb.save(ruta_salida)
    return ruta_salida


def _detectar_formato_headers(headers_norm: list[str]) -> str | None:
    oficial_norm = [_norm_header(c) for c in COLUMNAS_OFICIAL_PHOENIX]
    fmt1_norm = [_norm_header(c) for c in COLUMNAS_FORMATO_1]
    fmt2_norm = [_norm_header(c) for c in COLUMNAS_FORMATO_2]

    if headers_norm[:len(oficial_norm)] == oficial_norm:
        return "oficial"
    if headers_norm[:len(fmt2_norm)] == fmt2_norm:
        return "2"
    if headers_norm[:len(fmt1_norm)] == fmt1_norm:
        return "1"
    if any("cmbtipocif" in h for h in headers_norm) or (
        any("tipo de documento" in h for h in headers_norm)
        and any("numero de documento" in h for h in headers_norm)
        and any("nombres" in h for h in headers_norm)
    ):
        return "cliente_ss"
    return None


def _normalizar_tipo_doc_cliente(val) -> str:
    txt = str(val or "").strip().upper()
    if any(k in txt for k in ("PASAPORTE", "(PS)", "[2]")) or txt in ("2", "PS"):
        return "PASAPORTE"
    if any(k in txt for k in ("RUN", "CEDULA", "CÉDULA", "(CI)", "[1]", "[4]", "1007")) or txt in ("1", "4", "1007"):
        return "RUN"
    return ""


def _normalizar_genero_cliente(val) -> int | None:
    txt = _norm_header(str(val or ""))
    if txt in ("2", "f", "femenino", "femenino [2]", "mujer"):
        return 2
    # Incluye tolerancia al typo real frecuente 'maculino'
    if txt in ("5", "m", "masculino", "maculino", "masculino [5]", "hombre"):
        return 5
    if txt in ("8", "no indica", "no indica [8]"):
        return 2
    return None


def _limpiar_telefono_cliente(val) -> str:
    digs = re.sub(r"\D", "", str(val or ""))
    if len(digs) > 9 and digs.startswith("56"):
        digs = digs[2:]
    return digs


def validar_excel_grupal(
    ruta_excel: str,
    formato: str | int = "auto",
    codigo_run: int = 1007,
    estricto_rut: bool = True,
) -> dict:
    """
    Valida estrictamente una planilla Excel enviada por un cliente antes de subirla a Phoenix.
    Acepta:
    - Plantilla oficial Phoenix ('ID Type | Fiscal ID | ...')
    - Plantilla clientes MAWDY ('Datos de los pasajeros' con cmbTipoCIF / Tipo de Documento *)
    - Formatos en español del manual (1 y 2)
    Retorna {"valido": bool, "formato_detectado": str, "errores": list[str], "advertencias": list[str], "pasajeros": list[Pasajero]}.
    """
    errores: list[str] = []
    advertencias: list[str] = []
    pasajeros: list[Pasajero] = []

    if not os.path.exists(ruta_excel):
        return {"valido": False, "formato_detectado": "", "errores": [f"No se encontró el archivo: {ruta_excel}"], "advertencias": [], "pasajeros": []}

    try:
        wb = _cargar_workbook_seguro(ruta_excel)
    except Exception as e:
        return {"valido": False, "formato_detectado": "", "errores": [f"El archivo no es un Excel (.xlsx) válido: {e}"], "advertencias": [], "pasajeros": []}

    ws = wb["Datos de los pasajeros"] if "Datos de los pasajeros" in wb.sheetnames else wb.active
    filas = list(ws.iter_rows(values_only=True))
    if len(filas) < 2:
        return {
            "valido": False,
            "formato_detectado": "",
            "errores": ["El Excel debe incluir encabezados en la fila 1 y al menos un pasajero desde la fila 2."],
            "advertencias": [],
            "pasajeros": [],
        }

    # Detectar si encabezados están en fila 1 o fila 2 (ej: plantilla cliente con IDs internos en fila 1)
    fila_header_idx = 0
    detectado = _detectar_formato_headers([_norm_header(str(c or "")) for c in filas[0]])
    if detectado == "cliente_ss" and len(filas) >= 3:
        h2 = [_norm_header(str(c or "")) for c in filas[1]]
        if any("tipo de documento" in h for h in h2):
            fila_header_idx = 1
    elif not detectado and len(filas) >= 3:
        d2 = _detectar_formato_headers([_norm_header(str(c or "")) for c in filas[1]])
        if d2:
            detectado = d2
            fila_header_idx = 1

    headers_raw = [str(c or "").strip() for c in filas[fila_header_idx]]
    headers_norm = [_norm_header(h) for h in headers_raw]

    if str(formato).lower() != "auto":
        objetivo = str(formato).lower()
        if detectado != objetivo:
            mapa_cols = {
                "oficial": COLUMNAS_OFICIAL_PHOENIX,
                "1": COLUMNAS_FORMATO_1,
                "2": COLUMNAS_FORMATO_2,
            }
            esperadas = mapa_cols.get(objetivo, COLUMNAS_OFICIAL_PHOENIX)
            errores.append(
                f"Encabezados o orden de columnas incorrectos (formato '{objetivo}'). "
                f"Se esperaba: {' | '.join(esperadas)}. Recibido: {' | '.join(headers_raw)}."
            )
            return {"valido": False, "formato_detectado": detectado or "", "errores": errores, "advertencias": [], "pasajeros": []}
    elif not detectado:
        errores.append(
            f"Encabezados o orden de columnas incorrectos. "
            f"Se esperaba plantilla oficial: {' | '.join(COLUMNAS_OFICIAL_PHOENIX)}. "
            f"Recibido: {' | '.join(headers_raw)}."
        )
        return {"valido": False, "formato_detectado": "", "errores": errores, "advertencias": [], "pasajeros": []}

    if detectado == "cliente_ss":
        esperadas_norm = [h.rstrip(" *") for h in headers_norm]
    else:
        cols_activas = (
            COLUMNAS_OFICIAL_PHOENIX if detectado == "oficial"
            else (COLUMNAS_FORMATO_1 if detectado == "1" else COLUMNAS_FORMATO_2)
        )
        esperadas_norm = [_norm_header(c) for c in cols_activas]

    for idx_fila, row in enumerate(filas[fila_header_idx + 1:], start=fila_header_idx + 2):
        if not any(row):
            continue
        vals = list(row[:len(esperadas_norm)])
        while len(vals) < len(esperadas_norm):
            vals.append("")
        data = dict(zip(esperadas_norm, vals))

        if detectado == "oficial":
            td_raw = str(data["id type"] or "").strip().upper()
            fiscal_raw = str(data["fiscal id"] or "").strip()
            nombres = str(data["first name"] or "").strip()
            apellidos = str(data["surname"] or "").strip()
            dob_raw = data["d.o.b."]
            gen_raw = str(data["gender"] or "").strip()
            email_raw = str(data["email"] or "").strip()
            tel_raw = _limpiar_telefono_cliente(data["mobile"])
            tipo_doc = "RUN" if td_raw in ("1007", "4", "RUN") else ("PASAPORTE" if td_raw in ("2", "PASAPORTE", "PS") else "")
        elif detectado == "cliente_ss":
            td_raw = str(data.get("tipo de documento") or "").strip()
            tipo_doc = _normalizar_tipo_doc_cliente(td_raw)
            fiscal_raw = str(data.get("numero de documento") or "").strip()
            nombres = str(data.get("nombres") or "").strip()
            apellidos = str(data.get("apellidos") or "").strip()
            dob_raw = data.get("fecha de nacimiento")
            gen_raw = str(data.get("genero") or "").strip()
            email_raw = str(data.get("email") or "").strip()
            tel_raw = _limpiar_telefono_cliente(data.get("telefono movil"))
        else:
            td_raw = str(data["tipo de documento"] or "").strip().upper()
            num_raw = str(data["numero de documento"] or "").strip()
            dv_raw = str(data["dv"] if data["dv"] is not None else "").strip().upper()
            fiscal_raw = f"{num_raw}-{dv_raw}" if dv_raw and dv_raw != "0" else num_raw
            nombres = str(data["nombres"] or "").strip()
            apellidos = str(data["apellidos"] or "").strip()
            dob_raw = data["fecha de nacimiento"]
            gen_raw = str(data.get("genero") or "2").strip()
            email_raw = str(data["email"] or "").strip()
            tel_raw = _limpiar_telefono_cliente(data["telefono movil"])
            tipo_doc = "RUN" if td_raw in ("1007", "4", "RUN") else ("PASAPORTE" if td_raw in ("2", "PASAPORTE", "PS") else "")

        errores_fila: list[str] = []

        # 1. Tipo de Documento
        if not tipo_doc:
            errores_fila.append(f"Fila {idx_fila}: ID Type / Tipo de Documento inválido ('{td_raw}'). Debe ser {codigo_run} (RUN) o 2 (PASAPORTE).")

        # 2. Documento y DV
        num_doc, dv_doc = "", ""
        if not fiscal_raw:
            errores_fila.append(f"Fila {idx_fila}: Fiscal ID / Número de Documento vacío.")
        elif "." in fiscal_raw or " " in fiscal_raw:
            errores_fila.append(f"Fila {idx_fila}: Documento ('{fiscal_raw}') no debe contener puntos ni espacios.")
        elif tipo_doc == "RUN":
            if "-" in fiscal_raw:
                num_doc, dv_doc = limpiar_rut(fiscal_raw)
            elif detectado not in ("oficial", "cliente_ss"):
                num_doc, dv_doc = num_raw, dv_raw
            else:
                num_doc, dv_doc = limpiar_rut(fiscal_raw) if len(fiscal_raw) >= 8 else (fiscal_raw, "")

            dv_doc = dv_doc.upper()
            if not re.fullmatch(r"\d{7,9}", num_doc):
                msg_rut = f"Fila {idx_fila}: RUN '{fiscal_raw}' debe tener entre 7 y 9 dígitos numéricos antes del DV."
                (errores_fila if estricto_rut else advertencias).append(msg_rut)
            elif not re.fullmatch(r"[0-9K]", dv_doc):
                msg_dv = f"Fila {idx_fila}: DV '{dv_doc}' inválido para RUN (debe ser 0-9 o K)."
                (errores_fila if estricto_rut else advertencias).append(msg_dv)
            elif not validar_rut_modulo11(num_doc, dv_doc):
                advertencias.append(f"Fila {idx_fila}: RUN '{num_doc}-{dv_doc}' no coincide con dígito verificador módulo 11.")
        else:
            num_doc = fiscal_raw
            if detectado in ("1", "2") and dv_raw not in ("0", ""):
                errores_fila.append(f"Fila {idx_fila}: Para PASAPORTE, la columna DV debe ser 0 (recibido '{dv_raw}').")

        # 3. Nombres y Apellidos
        if not nombres:
            errores_fila.append(f"Fila {idx_fila}: Falta First name / Nombres.")
        if not apellidos:
            errores_fila.append(f"Fila {idx_fila}: Falta Surname / Apellidos.")

        # 4. Fecha de Nacimiento
        dt_nac = _parsear_fecha(dob_raw)
        if not dt_nac:
            errores_fila.append(f"Fila {idx_fila}: D.o.b. / Fecha de Nacimiento inválida ('{dob_raw}'). Usar DD/MM/AAAA o DD-MM-AAAA.")
            fnac_str = ""
        else:
            fnac_str = dt_nac.strftime("%d/%m/%Y")

        # 5. Género (en Phoenix PRE #ayudaExcel: 2=Femenino, 5=Masculino)
        sexo_val = 2
        if detectado in ("oficial", "2", "cliente_ss"):
            g_norm = _normalizar_genero_cliente(gen_raw)
            if g_norm is None:
                errores_fila.append(f"Fila {idx_fila}: Gender / Género inválido ('{gen_raw}'). Usar 2 (Femenino) o 5 (Masculino).")
            else:
                sexo_val = g_norm
        if detectado == "2":
            cod_pais_raw = str(data["cod. pais movil"] or "").strip()
            if not re.fullmatch(r"\d{1,4}", cod_pais_raw):
                errores_fila.append(f"Fila {idx_fila}: Cod. País Móvil inválido ('{cod_pais_raw}'). Ejemplo: 56.")

        # 6. Teléfono Móvil
        if not re.fullmatch(r"\d{8,12}", tel_raw):
            errores_fila.append(f"Fila {idx_fila}: Mobile / Teléfono Móvil inválido ('{tel_raw}'). Debe contener solo dígitos (ej: 912345678).")

        # 7. Email
        if email_raw and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email_raw):
            errores_fila.append(f"Fila {idx_fila}: Email inválido ('{email_raw}').")

        errores.extend(errores_fila)
        if not errores_fila:
            partes_ap = apellidos.split(None, 1)
            ap_pat = partes_ap[0] if partes_ap else apellidos
            ap_mat = partes_ap[1] if len(partes_ap) > 1 else ""
            pasajeros.append(
                Pasajero(
                    tipo_documento=tipo_doc,
                    numero_documento=num_doc,
                    digito_verificador=dv_doc if tipo_doc == "RUN" else "",
                    nombre=nombres,
                    apellido_paterno=ap_pat,
                    apellido_materno=ap_mat,
                    fecha_nacimiento=fnac_str,
                    sexo=sexo_val,
                    email=email_raw,
                    telefono=tel_raw,
                )
            )

    if not pasajeros and not errores:
        errores.append("La planilla no contiene filas de pasajeros con datos.")

    return {
        "valido": len(errores) == 0,
        "formato_detectado": detectado or "",
        "errores": errores,
        "advertencias": advertencias,
        "pasajeros": pasajeros,
    }


class PhoenixEmisionGrupal(PhoenixAutomation):
    """Automatización de emisiones grupales con Playwright, validación y aviso al cliente."""

    def crear_fichero_carga(
        self,
        pasajeros: list[Pasajero],
        formato: str | int = "oficial",
        ruta_salida: str | None = None,
        correo_respaldo: str = "",
        codigo_run: int | None = None,
    ) -> str:
        """
        Crea el fichero Excel (.xlsx) con la estructura OpenXML exacta de Phoenix (#plantillaExcel):
        Hoja 'Sheet1', columnas: ID Type | Fiscal ID | First name | Surname | D.o.b. | Age | Gender | Email | Mobile.
        """
        if not ruta_salida:
            ruta_salida = os.path.join(os.path.dirname(__file__), "ficheros", "carga_temporal.xlsx")

        config = cargar_config()
        filas_datos: list[dict] = []

        for p in pasajeros:
            es_run = p.tipo_documento.upper() == "RUN"
            if es_run and codigo_run is not None:
                id_type = int(codigo_run)
            else:
                id_type = _codigo_tipo_doc_grupal(p.tipo_documento, config)

            num = str(p.numero_documento).strip()
            dv = str(p.digito_verificador or "").strip().upper()
            fiscal_id = f"{num}-{dv}" if (es_run and dv) else num

            apellidos = f"{p.apellido_paterno} {p.apellido_materno}".strip().upper()
            dt_nac = _parsear_fecha(p.fecha_nacimiento)
            dob_val = dt_nac.date() if dt_nac else p.fecha_nacimiento.strip()
            edad = _calcular_edad(p.fecha_nacimiento) or 0
            genero = int(p.sexo) if str(p.sexo) in ("2", "5") else 2
            telefono = _limpiar_telefono_cliente(p.telefono) or "900000000"
            email = (p.email or "").strip() or correo_respaldo.strip()

            filas_datos.append({
                "id_type": id_type,
                "fiscal_id": fiscal_id,
                "first_name": p.nombre.strip().upper(),
                "surname": apellidos,
                "dob": dob_val,
                "age": int(edad),
                "gender": genero,
                "email": email,
                "mobile": telefono,
            })

        _escribir_xlsx_oficial_phoenix(ruta_salida, filas_datos)
        logger.info(f"✅ Fichero oficial Phoenix creado: {ruta_salida} ({len(pasajeros)} pasajeros)")
        return ruta_salida

    def cargar_fichero(self, ruta_fichero: str) -> tuple[bool, str]:
        """
        Sube el fichero Excel en la pestaña T1 usando Playwright set_input_files(),
        esperando de forma determinista el resultado del callback de carga masiva.
        """
        logger.info(f"Cargando fichero en #windowCargaAsegurados_PW-1: {ruta_fichero}")
        self.mostrar_hud("Paso 5: Carga Masiva", f"Subiendo {os.path.basename(ruta_fichero)}...")

        self.frame.locator("#Wcr_carTabPage_T1").click()
        self._esperar_sin_loading()

        modal_carga = self.frame.locator("#windowCargaAsegurados_PW-1")
        for intento in range(2):
            self.frame.locator("#btnCargaPlantillaAsegurado").click()
            try:
                modal_carga.wait_for(state="visible", timeout=10000)
                break
            except Exception:
                if intento == 0:
                    logger.warning("⚠️ Reintentando abrir #windowCargaAsegurados_PW-1...")
                else:
                    raise

        self.frame.locator("#fuBulk_TextBox0_Input").set_input_files(os.path.abspath(ruta_fichero))

        deadline = time.monotonic() + (self.timeout_ms / 1000.0)
        modal_text = ""
        estado = ""
        while time.monotonic() < deadline:
            self.page.wait_for_timeout(500)
            self._esperar_sin_loading(3000)
            modal_text = modal_carga.inner_text().strip()
            low = modal_text.lower()
            if (
                "realizado correctamente" in low
                or "validacion del fichero es correcta" in low
                or "validación del fichero es correcta" in low
                or "file validation is correct" in low
                or "register import was successful" in low
            ):
                estado = "ok"
                break
            if any(
                kw in low
                for kw in (
                    "incorrecto",
                    "incorrect",
                    "error",
                    "revise la plantilla",
                    "check the template",
                    "no se pudo",
                    "inválid",
                    "invalid",
                )
            ):
                estado = "error"
                break

        if estado != "ok":
            detalle = modal_text or "Timeout esperando respuesta de carga masiva en Phoenix."
            logger.error(f"❌ Phoenix rechazó o no confirmó la carga del fichero: {detalle}")
            self._capturar_screenshot("error_formato_grupal")
            return False, detalle

        btn_cerrar = self.frame.locator("#windowCargaAsegurados_HCB-1")
        if btn_cerrar.count() > 0 and btn_cerrar.is_visible():
            btn_cerrar.click()
            try:
                modal_carga.wait_for(state="hidden", timeout=10000)
            except Exception:
                pass
        self._esperar_sin_loading()

        logger.info("✅ Fichero cargado correctamente")
        return True, modal_text

    def eliminar_fila_en_blanco(self):
        """Elimina la fila 1 si quedó en blanco tras la carga (#chkElimRiesgo_1 -> #btnEliminarRiesgos)."""
        doc1 = self.frame.locator("#txtIdFiscal_1")
        if doc1.count() > 0 and doc1.input_value().strip():
            return

        logger.info("Eliminando fila 1 en blanco (#chkElimRiesgo_1 -> #btnEliminarRiesgos)...")
        self.mostrar_hud("Limpiando Fila 1", "Eliminando fila en blanco inicial...")
        self._limpiar_alertas()
        chk1 = self.frame.locator("#chkElimRiesgo_1")
        if chk1.count() > 0 and not chk1.is_checked():
            chk1.check()
        self.frame.locator("#btnEliminarRiesgos").click()
        self._esperar_sin_loading()
        self.page.wait_for_timeout(600)

    def seleccionar_titular_grupal(self, indice_titular: int = 1):
        """Marca el radio circular de 'Contrata' (#rbTomador_{indice_titular})."""
        logger.info(f"Seleccionando titular grupal (#rbTomador_{indice_titular})...")
        rb = self.frame.locator(f"#rbTomador_{indice_titular}")
        if rb.count() > 0 and rb.is_enabled():
            rb.check()
            self.page.wait_for_timeout(500)

    def notificar_formato_incorrecto(
        self,
        datos: DatosEmision,
        errores: list[str],
        formato_fichero: str | int = "oficial",
        email_modo: str | None = None,
    ) -> dict:
        """Genera la plantilla óptima oficial y envía aviso al cliente solicitando corrección."""
        ruta_plantilla = generar_plantilla_optima_excel(formato="oficial", con_ejemplo=True)
        destinatario = datos.correo_solicitante or (datos.pasajeros[0].email if datos.pasajeros else "")
        return enviar_aviso_formato_incorrecto(
            correo_cliente=destinatario,
            errores=errores,
            nombre_cliente=datos.nombre_agencia,
            file_ref=datos.file1,
            ruta_plantilla_adjunta=ruta_plantilla,
            modo_override=email_modo,
        )

    def emitir_voucher_grupal(
        self,
        datos: DatosEmision,
        formato_fichero: str | int = "oficial",
        indice_titular: int = 1,
        dry_run: bool = False,
        enviar_correo_asegurado: bool | None = None,
        email_modo: str | None = None,
        ruta_excel_cliente: str | None = None,
        notificar_cliente_si_error: bool = True,
        confirmar_antes_aceptar: bool = False,
    ) -> dict:
        """
        Ejecuta emisión grupal completa con Playwright.
        Si se entrega `ruta_excel_cliente`, valida primero la planilla; si falla
        (o si Phoenix rechaza el formato al subir), avisa automáticamente al cliente.
        """
        email_cfg = self.config.get("email", {})
        if enviar_correo_asegurado is None:
            enviar_correo_asegurado = bool(email_cfg.get("enviar_al_asegurado", True))

        resultado = {
            "exito": False,
            "dry_run": dry_run,
            "numero_voucher": "",
            "precios": {},
            "archivo_voucher": "",
            "resumen": "",
            "correo_asegurado": {},
            "aviso_formato_incorrecto": {},
            "pax": len(datos.pasajeros),
            "error": "",
        }

        try:
            logger.info("=" * 60)
            logger.info(f"EMISIÓN GRUPAL PLAYWRIGHT {'[DRY-RUN] ' if dry_run else ''}: {datos.codigo_cliente}")
            logger.info("=" * 60)

            if ruta_excel_cliente:
                validacion = validar_excel_grupal(ruta_excel_cliente, formato="auto")
                if not validacion["valido"]:
                    resultado["error"] = "Formato de planilla grupal incorrecto"
                    if notificar_cliente_si_error:
                        resultado["aviso_formato_incorrecto"] = self.notificar_formato_incorrecto(
                            datos, validacion["errores"], formato_fichero=formato_fichero, email_modo=email_modo
                        )
                    return resultado
                datos.pasajeros = validacion["pasajeros"]
                resultado["pax"] = len(datos.pasajeros)

            ruta_fichero = self.crear_fichero_carga(
                datos.pasajeros,
                formato=formato_fichero,
                correo_respaldo=datos.correo_solicitante,
            )

            self.ir_a_emision_vouchers()
            self.nuevo_contrato()

            self.buscar_cliente(datos.codigo_cliente)
            self.seleccionar_cliente_final(datos.codigo_cliente, datos.cliente_final or datos.nombre_agencia)

            producto = datos.producto or determinar_producto(datos.nombre_agencia, datos.tipo_viaje) or ""
            self.seleccionar_producto(producto)

            self.completar_vigencia(datos.fecha_inicio, datos.fecha_fin, datos.zona)
            self.completar_datos_asistencia(
                datos.pais_origen, datos.destino, datos.lista_destino, datos.tipo_viaje_sv
            )

            carga_ok, msg_carga = self.cargar_fichero(ruta_fichero)
            if not carga_ok:
                resultado["error"] = f"Phoenix rechazó el formato del fichero: {msg_carga}"
                if notificar_cliente_si_error:
                    resultado["aviso_formato_incorrecto"] = self.notificar_formato_incorrecto(
                        datos,
                        [f"Rechazo de plataforma Phoenix: {msg_carga}"],
                        formato_fichero=formato_fichero,
                        email_modo=email_modo,
                    )
                return resultado

            self.eliminar_fila_en_blanco()
            self.seleccionar_titular_grupal(indice_titular)

            if datos.file1:
                self.completar_file(datos.file1)

            if dry_run:
                self.mostrar_hud("Dry-Run Grupal", "Revisando precios antes de cancelar...")
                resultado["precios"] = self.obtener_precios(datos.tipo_viaje)
                self._capturar_screenshot("dry_run_grupal_antes_aceptar")
                self.frame.locator("#Wcr_btnCancelar").click()
                resultado["exito"] = True
                resultado["resumen"] = "Dry-run grupal completado con Playwright sin emitir."
                return resultado

            if confirmar_antes_aceptar:
                self.mostrar_hud("Esperando Confirmación", f"{len(datos.pasajeros)} pax cargados. Confirma en consola.")
                resp = input("\n👀 Revisa los pasajeros cargados en Phoenix. ¿Confirmar emisión grupal (Aceptar)? (s/n) [s]: ").strip().lower()
                if resp == "n":
                    self.frame.locator("#Wcr_btnCancelar").click()
                    resultado["error"] = "Emisión grupal cancelada manualmente por el usuario antes de Aceptar."
                    return resultado

            exito, err_msg = self.terminar_emision()
            if not exito:
                resultado["error"] = err_msg or "Error al confirmar emisión grupal"
                return resultado

            titular_nombre = ""
            if datos.pasajeros:
                idx = min(max(indice_titular - 1, 0), len(datos.pasajeros) - 1)
                p_tit = datos.pasajeros[idx]
                titular_nombre = f"{p_tit.nombre} {p_tit.apellido_paterno}".strip()

            num_voucher, ruta_pdf = self.descargar_voucher(datos.file1, titular_nombre)
            precios = self.obtener_precios(datos.tipo_viaje)

            resultado["exito"] = True
            resultado["numero_voucher"] = num_voucher
            resultado["archivo_voucher"] = ruta_pdf or ""
            resultado["precios"] = precios

            if debe_enviar_precios(datos.codigo_cliente, datos.nombre_agencia):
                resultado["resumen"] = generar_resumen_emision(
                    datos.tipo_viaje,
                    clp=precios.get("clp", ""),
                    usd=precios.get("usd", ""),
                    tc=precios.get("tc", ""),
                    pax=len(datos.pasajeros),
                )
            else:
                resultado["resumen"] = "EstimadX, buenos días:\n\nEmisión realizada según lo solicitado. Adjunto voucher.\n\nSaludos."

            if enviar_correo_asegurado and ruta_pdf:
                resultado["correo_asegurado"] = enviar_voucher_al_asegurado(
                    pasajeros=datos.pasajeros,
                    numero_voucher=num_voucher,
                    ruta_pdf=ruta_pdf,
                    fecha_inicio=datos.fecha_inicio,
                    fecha_fin=datos.fecha_fin,
                    destino=datos.lista_destino or datos.destino,
                    nombre_agencia=datos.nombre_agencia,
                    correo_respaldo=datos.correo_solicitante,
                    modo_override=email_modo,
                )

            if email_cfg.get("enviar_al_ejecutivo") and datos.correo_solicitante and ruta_pdf:
                enviar_correo(
                    destinatarios=datos.correo_solicitante,
                    asunto=f"Voucher N° {num_voucher} - {datos.file1 or titular_nombre}",
                    cuerpo=resultado["resumen"],
                    ruta_pdf=ruta_pdf,
                    modo_override=email_modo,
                    etiqueta_archivo=f"ejecutivo_{num_voucher}",
                )

            logger.info(f"✅ EMISIÓN GRUPAL PLAYWRIGHT COMPLETADA (Voucher: {num_voucher})")

        except Exception as e:
            resultado["error"] = str(e)
            logger.error(f"❌ ERROR EN EMISIÓN GRUPAL PLAYWRIGHT: {e}")
            self._capturar_screenshot("error_grupal")

        return resultado
