"""
Phoenix v4 - Módulo MVP de Autoemisión desde Correo (.eml / Texto + Excel)
==========================================================================
Segmento MVP interactivo para:
1. Leer correos reales (.eml con adjuntos/tablas HTML o texto pegado).
2. Extraer automáticamente cliente, producto, vigencia, destino, File1 y pasajeros.
3. Auditar rigurosamente cada pasajero (RUT módulo 11, edades, fechas, teléfonos) antes de tocar Phoenix.
4. Ejecutar la emisión en pantalla (navegador visible + HUD paso a paso + slow_mo + confirmación opcional).
"""

import email
import os
import re
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser

from phoenix_emision import DatosEmision, Pasajero, PhoenixAutomation, _calcular_edad, _parsear_fecha_phoenix
from emision_grupal import PhoenixEmisionGrupal, validar_excel_grupal
from utils import (
    BASE_DIR,
    buscar_dealer_por_codigo,
    buscar_dealer_por_correo,
    buscar_dealer_por_nombre,
    cargar_config,
    determinar_producto,
    es_destino_internacional,
    formatear_file_nacional,
    inferir_region_destino,
    limpiar_rut,
    normalizar_texto,
    validar_rut_modulo11,
)

MESES_ES = {
    "enero": 1, "ene": 1,
    "febrero": 2, "feb": 2,
    "marzo": 3, "mar": 3,
    "abril": 4, "abr": 4,
    "mayo": 5, "may": 5,
    "junio": 6, "jun": 6,
    "julio": 7, "jul": 7,
    "agosto": 8, "ago": 8,
    "septiembre": 9, "setiembre": 9, "sep": 9, "sept": 9,
    "octubre": 10, "oct": 10,
    "noviembre": 11, "nov": 11,
    "diciembre": 12, "dic": 12,
}

DESTINOS_INTERNACIONALES_HINT = (
    "argentina", "bariloche", "buenos aires", "mendoza", "brasil", "rio", "florianopolis",
    "europa", "schengen", "espana", "españa", "francia", "italia", "peru", "perú",
    "colombia", "mexico", "méxico", "eeuu", "usa", "estados unidos", "miami", "orlando",
    "sudamerica", "sudamérica", "latinoamerica", "latinoamérica", "todo el mundo", "internacional",
    "caribe", "cuba", "punta cana", "cancun", "cancún", "uruguay", "paraguay", "bolivia",
)


class _HTMLTableParser(HTMLParser):
    """Extrae tablas HTML y texto limpio desde el cuerpo de un correo."""

    def __init__(self):
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._cur_table: list[list[str]] = []
        self._cur_row: list[str] = []
        self._cur_cell: list[str] = []
        self._text_parts: list[str] = []

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t == "table":
            self._in_table = True
            self._cur_table = []
        elif t == "tr" and self._in_table:
            self._in_row = True
            self._cur_row = []
        elif t in ("td", "th") and self._in_row:
            self._in_cell = True
            self._cur_cell = []
        elif t in ("br", "p", "div", "li", "tr"):
            self._text_parts.append("\n")

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in ("td", "th") and self._in_cell:
            cell_txt = re.sub(r"\s+", " ", "".join(self._cur_cell)).strip()
            self._cur_row.append(cell_txt)
            self._in_cell = False
            self._text_parts.append(" | ")
        elif t == "tr" and self._in_row:
            if any(self._cur_row):
                self._cur_table.append(self._cur_row)
            self._in_row = False
            self._text_parts.append("\n")
        elif t == "table" and self._in_table:
            if self._cur_table:
                self.tables.append(self._cur_table)
            self._in_table = False

    def handle_data(self, data):
        if self._in_cell:
            self._cur_cell.append(data)
        self._text_parts.append(data)

    def get_text(self) -> str:
        raw = "".join(self._text_parts)
        return re.sub(r"\n{3,}", "\n\n", raw).strip()


def extraer_desde_eml(ruta_eml: str) -> dict:
    """Lee un archivo .eml y extrae asunto, remitente original, texto, tablas HTML y adjuntos Excel."""
    if not os.path.exists(ruta_eml):
        raise FileNotFoundError(f"No se encontró el archivo .eml: {ruta_eml}")

    with open(ruta_eml, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    fecha_dt = None
    if msg["date"]:
        try:
            fecha_dt = parsedate_to_datetime(msg["date"])
        except Exception:
            fecha_dt = None

    text_plain = ""
    html_raw = ""
    adjuntos_excel: list[str] = []
    adj_dir = os.path.join(BASE_DIR, "ficheros", "mvp_adjuntos")
    os.makedirs(adj_dir, exist_ok=True)

    for part in msg.walk():
        ctype = part.get_content_type()
        disp = part.get_content_disposition()
        fname = part.get_filename()
        if fname and fname.lower().endswith((".xlsx", ".xlsm")):
            safe_name = re.sub(r'[<>:"/\\|?*]', "_", fname)
            dest = os.path.join(adj_dir, safe_name)
            payload = part.get_payload(decode=True)
            if payload:
                with open(dest, "wb") as out:
                    out.write(payload)
                adjuntos_excel.append(dest)
        elif ctype == "text/plain" and disp != "attachment" and not text_plain:
            text_plain = part.get_content() or ""
        elif ctype == "text/html" and disp != "attachment" and not html_raw:
            html_raw = part.get_content() or ""

    tablas_html: list[list[list[str]]] = []
    text_html = ""
    if html_raw:
        parser = _HTMLTableParser()
        parser.feed(html_raw)
        tablas_html = parser.tables
        text_html = parser.get_text()

    cuerpo = text_plain.strip() or text_html

    # Si es un correo reenviado (RV:), detectar el remitente real interno (De: Nombre <correo>)
    remitente = str(msg["from"] or "")
    m_inner_from = re.search(r"(?im)^De:\s*.*?([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", cuerpo)
    if m_inner_from:
        remitente = m_inner_from.group(1)
    else:
        m_mail = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", remitente)
        if m_mail:
            remitente = m_mail.group(1)

    return {
        "subject": str(msg["subject"] or ""),
        "from": remitente,
        "date": fecha_dt,
        "body": cuerpo,
        "tablas_html": tablas_html,
        "adjuntos_excel": adjuntos_excel,
    }


_MESES_PAT = "|".join(sorted((re.escape(k) for k in MESES_ES), key=len, reverse=True))
_SEP = r"\s*(?:al|hasta|a|-|–|—)\s*(?:el\s+)?(?:d[ií]a\s+)?"
_FECHA_NUM = r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
_DIA_MES_TXT = rf"(\d{{1,2}})\s*(?:de\s*)?({_MESES_PAT})\b(?:\s*(?:de|del)?\s*(\d{{4}}))?"


def _fecha_valida(d: int, m: int, y: int) -> datetime | None:
    try:
        return datetime(y, m, d)
    except ValueError:
        return None


def _anio_para(d: int, m: int, ref: datetime) -> int:
    """Año de una fecha sin año: el de la fecha del correo, o el siguiente si ya quedó muy atrás."""
    y = ref.year
    cand = _fecha_valida(d, m, y)
    if cand and (ref - cand).days > 60:
        y += 1
    return y


def _fmt(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")


def _parsear_rango_fechas(texto: str, ref_dt: datetime | None = None) -> tuple[str, str]:
    """
    Detecta rangos de fechas en el texto. Soporta:
    '27/11/2026 al 29/11/2026', '27-11-26 - 29-11-26', '27/11 al 29/11',
    '30 noviembre al 4 diciembre', 'del 27 de noviembre hasta el 29 de noviembre de 2026',
    'del 27 al 29 de noviembre'. Devuelve ('', '') si no hay un rango válido.
    """
    ref = ref_dt.replace(tzinfo=None) if ref_dt else datetime.now()
    t = normalizar_texto(texto)

    # 1. Numérico con año
    m = re.search(rf"({_FECHA_NUM}){_SEP}({_FECHA_NUM})", t)
    if m:
        def _norm_d(sv: str) -> datetime | None:
            sv = sv.replace("-", "/")
            for fmt in ("%d/%m/%Y", "%d/%m/%y"):
                try:
                    return datetime.strptime(sv, fmt)
                except ValueError:
                    continue
            return None
        a, b = _norm_d(m.group(1)), _norm_d(m.group(2))
        if a and b:
            return _fmt(a), _fmt(b)

    # 2. Numérico sin año: 27/11 al 29/11
    m = re.search(rf"(?<![\d/-])(\d{{1,2}})[/-](\d{{1,2}}){_SEP}(\d{{1,2}})[/-](\d{{1,2}})(?![\d/-])", t)
    if m:
        d1, m1, d2, m2 = (int(m.group(i)) for i in range(1, 5))
        y1 = _anio_para(d1, m1, ref)
        y2 = y1 + 1 if (m2, d2) < (m1, d1) else y1
        a, b = _fecha_valida(d1, m1, y1), _fecha_valida(d2, m2, y2)
        if a and b:
            return _fmt(a), _fmt(b)

    # 3. Texto con dos meses: 30 noviembre al 4 diciembre [2026]
    m = re.search(rf"{_DIA_MES_TXT}{_SEP}{_DIA_MES_TXT}", t)
    if m:
        d1, m1 = int(m.group(1)), MESES_ES[m.group(2)]
        d2, m2 = int(m.group(4)), MESES_ES[m.group(5)]
        y1 = int(m.group(3)) if m.group(3) else _anio_para(d1, m1, ref)
        if m.group(6):
            y2 = int(m.group(6))
        else:
            y2 = y1 + 1 if (m2, d2) < (m1, d1) else y1
        a, b = _fecha_valida(d1, m1, y1), _fecha_valida(d2, m2, y2)
        if a and b:
            return _fmt(a), _fmt(b)

    # 4. Mismo mes: del 27 al 29 de noviembre [de 2026]
    m = re.search(rf"(?<![\d/-])(\d{{1,2}}){_SEP}(\d{{1,2}})\s*(?:de\s*)?({_MESES_PAT})\b(?:\s*(?:de|del)?\s*(\d{{4}}))?", t)
    if m:
        d1, d2, mes = int(m.group(1)), int(m.group(2)), MESES_ES[m.group(3)]
        y = int(m.group(4)) if m.group(4) else _anio_para(d2, mes, ref)
        a, b = _fecha_valida(d1, mes, y), _fecha_valida(d2, mes, y)
        if a and b:
            return _fmt(a), _fmt(b)

    return "", ""


def _label_regex(etiqueta: str) -> str:
    """Patrón de una etiqueta tolerante a tildes y mayúsculas (Codigo/Código, Cod./Cód.)."""
    base = re.escape(normalizar_texto(etiqueta))
    for vocal, clase in (("a", "[aáàä]"), ("e", "[eéèë]"), ("i", "[iíìï]"), ("o", "[oóòö]"), ("u", "[uúùü]"), ("n", "[nñ]")):
        base = base.replace(vocal, clase)
    return base


def _extraer_campo_tabla_o_texto(texto: str, tablas: list[list[list[str]]], etiquetas: list[str]) -> str:
    """Busca primero en tablas HTML clave-valor y luego en líneas del texto (sin depender de tildes)."""
    etq_norm = [normalizar_texto(e) for e in etiquetas]
    for tbl in tablas:
        for row in tbl:
            if len(row) >= 2:
                k = normalizar_texto(row[0]).strip().rstrip(":").strip()
                if any(e == k or e in k for e in etq_norm):
                    val = row[1].strip()
                    if val:
                        return val

    for etq in etiquetas:
        pat = _label_regex(etq)
        m = re.search(rf"(?im)^\s*{pat}\s*[:\t]+\s*([^\r\n]+)", texto)
        if m and m.group(1).strip():
            return m.group(1).strip()
        m2 = re.search(rf"(?im)^\s*{pat}\s*\r?\n\s*([^\r\n]+)", texto)
        if m2 and m2.group(1).strip():
            return m2.group(1).strip()
    return ""


def _aplanar_tablas(tablas: list[list[list[str]]]) -> str:
    """Texto plano de las tablas (una fila por línea), para buscar fechas que vengan dentro de ellas."""
    return "\n".join(" ".join(c for c in row if c) for tbl in tablas for row in tbl)


_CODIGO_RE = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]{2}(?:\d{3,}|\.[A-Za-z0-9]+))(?![A-Za-z0-9])")


def _prefijos_codigo_conocidos() -> set[str]:
    """Prefijos de código de cliente (SV, PC, WX, MV...) sacados del catálogo de dealers + los conocidos."""
    prefijos = {"WX", "SV", "PC", "MV"}
    try:
        for d in cargar_config().get("dealers", []):
            cod = str(d.get("codigo", ""))
            if re.match(r"^[A-Za-z]{2}\d", cod):
                prefijos.add(cod[:2].upper())
    except Exception:
        pass
    return prefijos


def _normalizar_codigo_cliente(valor: str) -> tuple[str, str]:
    """
    Separa el valor de un campo 'cliente' en (código, nombre).
    'WX01526 TURISMO RAYS' -> ('WX01526', 'TURISMO RAYS'); 'Colegio San Pedro' -> ('', 'Colegio San Pedro');
    un valor de una sola palabra (ej. 'PARANCIB@cverde.cl') se conserva como código.
    """
    v = (valor or "").strip()
    if not v:
        return "", ""
    m = _CODIGO_RE.search(v)
    if m:
        codigo = m.group(1).upper()
        nombre = re.sub(r"\s*[-–:]\s*$", "", v.replace(m.group(1), "", 1)).strip(" -–:")
        return codigo, nombre
    if len(v.split()) == 1:
        return v, ""
    return "", v


def _buscar_codigo_en_texto(texto: str) -> str:
    """Último recurso: busca en el texto libre un código con prefijo conocido (ej. WX01526)."""
    prefijos = _prefijos_codigo_conocidos()
    for m in _CODIGO_RE.finditer(texto):
        if m.group(1)[:2].upper() in prefijos:
            return m.group(1).upper()
    return ""


def _buscar_destino_en_texto(texto: str) -> str:
    """Último recurso para correos sin etiquetas: 'viaje a Bariloche del 27...' -> 'Bariloche'."""
    m = re.search(
        r"(?i)\b(?:viaje|viajan|viajar|viajaran|destino|rumbo)\s+(?:a|al|hacia|con destino a)\s+"
        r"([A-ZÁÉÍÓÚÑa-záéíóúñ][A-Za-zÁÉÍÓÚÑáéíóúñ ]{2,40}?)(?=\s+(?:del|desde|entre|el|para|con|durante)\b|[,.;\n]|\s*$)",
        texto,
    )
    return m.group(1).strip() if m else ""


def _pasajeros_desde_tablas(tablas: list[list[list[str]]]) -> list[Pasajero]:
    """Extrae pasajeros cuando vienen en una tabla inline dentro del correo (ej: NOMBRE | RUT | CELULAR | CORREO)."""
    pasajeros: list[Pasajero] = []
    for tbl in tablas:
        if len(tbl) < 2:
            continue
        headers = [c.strip().lower() for c in tbl[0]]
        if not any("rut" in h or "documento" in h for h in headers):
            continue
        if not any("nombre" in h for h in headers):
            continue

        for row in tbl[1:]:
            if len(row) < len(headers):
                row = row + [""] * (len(headers) - len(row))
            rec = dict(zip(headers, row))
            nombre_full = next((v for k, v in rec.items() if "nombre" in k), "").strip()
            rut_raw = next((v for k, v in rec.items() if "rut" in k or "documento" in k), "").strip()
            tel_raw = next((v for k, v in rec.items() if "cel" in k or "tel" in k or "movil" in k or "móvil" in k), "").strip()
            mail_raw = next((v for k, v in rec.items() if "correo" in k or "mail" in k), "").strip()
            fnac_raw = next((v for k, v in rec.items() if "nac" in k), "").strip()

            if not nombre_full and not rut_raw:
                continue
            num_doc, dv_doc = limpiar_rut(rut_raw) if rut_raw else ("", "")
            partes = nombre_full.split()
            if len(partes) >= 3:
                nom = " ".join(partes[:-2])
                ap_pat, ap_mat = partes[-2], partes[-1]
            elif len(partes) == 2:
                nom, ap_pat, ap_mat = partes[0], partes[1], ""
            else:
                nom, ap_pat, ap_mat = nombre_full, "", ""

            tel_digits = re.sub(r"\D", "", tel_raw)
            if len(tel_digits) > 9 and tel_digits.startswith("56"):
                tel_digits = tel_digits[2:]

            pasajeros.append(
                Pasajero(
                    tipo_documento="RUN",
                    numero_documento=num_doc,
                    digito_verificador=dv_doc,
                    nombre=nom,
                    apellido_paterno=ap_pat,
                    apellido_materno=ap_mat,
                    fecha_nacimiento=fnac_raw.replace("-", "/"),
                    sexo=5,
                    email=mail_raw,
                    telefono=tel_digits or "900000000",
                    es_titular=len(pasajeros) == 0,
                )
            )
    return pasajeros


def parsear_solicitud_correo(
    texto: str,
    asunto: str = "",
    remitente: str = "",
    fecha_correo: datetime | None = None,
    tablas_html: list[list[list[str]]] | None = None,
    rutas_excel: list[str] | None = None,
) -> dict:
    """Construye DatosEmision y metadatos de auditoría a partir de un correo real."""
    tablas = tablas_html or []
    excels = rutas_excel or []
    texto_completo = f"{asunto}\n{texto}"

    # 1. Cliente / Dealer
    valor_cliente = _extraer_campo_tabla_o_texto(
        texto,
        tablas,
        [
            "EMITIR BAJO CLIENTE", "EMITIR BAJO", "COD CLIENTE", "COD. CLIENTE", "CODIGO CLIENTE",
            "CODIGO DE CLIENTE", "CLIENTE",
        ],
    )
    codigo_cliente, nombre_en_cliente = _normalizar_codigo_cliente(valor_cliente)
    if not codigo_cliente:
        codigo_cliente = _buscar_codigo_en_texto(texto_completo + "\n" + _aplanar_tablas(tablas))

    producto = _extraer_campo_tabla_o_texto(texto, tablas, ["Producto", "Plan"])
    negocio = _extraer_campo_tabla_o_texto(
        texto, tablas, ["N° de negocio", "N° negocio", "Numero de negocio", "File", "File1", "Booking", "Expediente"]
    )
    origen = _extraer_campo_tabla_o_texto(texto, tablas, ["Origen"]) or "SANTIAGO"
    destino_raw = _extraer_campo_tabla_o_texto(texto, tablas, ["Destino"]) or _buscar_destino_en_texto(texto_completo)
    colegio = _extraer_campo_tabla_o_texto(texto, tablas, ["Colegio", "Empresa", "Sociedad"]) or nombre_en_cliente

    agencia = ""
    if remitente:
        matches = buscar_dealer_por_correo(remitente)
        if matches:
            agencia = matches[0]["agencia"]
            if not codigo_cliente:
                codigo_cliente = matches[0]["codigo"]

    if codigo_cliente and not agencia:
        dealer = buscar_dealer_por_codigo(codigo_cliente)
        if dealer:
            agencia = dealer["agencia"]

    if not codigo_cliente and colegio:
        matches_nom = buscar_dealer_por_nombre(colegio)
        if matches_nom:
            codigo_cliente = matches_nom[0]["codigo"]
            agencia = matches_nom[0]["agencia"]

    # 2. Vigencia (texto del correo + asunto + contenido de las tablas + campos etiquetados)
    fecha_inicio, fecha_fin = _parsear_rango_fechas(texto_completo, ref_dt=fecha_correo)
    if not (fecha_inicio and fecha_fin):
        extra = _extraer_campo_tabla_o_texto(
            texto, tablas, ["Vigencia", "Fechas", "Fecha de viaje", "Fechas de viaje", "Periodo", "Periodo de viaje"]
        )
        ini_txt = _extraer_campo_tabla_o_texto(texto, tablas, ["Fecha inicio", "Fecha de inicio", "Desde", "Salida"])
        fin_txt = _extraer_campo_tabla_o_texto(texto, tablas, ["Fecha fin", "Fecha de fin", "Fecha termino", "Hasta", "Regreso"])
        candidatos = [extra, f"{ini_txt} al {fin_txt}" if ini_txt and fin_txt else "", _aplanar_tablas(tablas)]
        for cand in candidatos:
            if cand:
                fecha_inicio, fecha_fin = _parsear_rango_fechas(cand, ref_dt=fecha_correo)
                if fecha_inicio and fecha_fin:
                    break

    # 3. Tipo de viaje / Destino Phoenix
    texto_destino = f"{destino_raw} {asunto}"
    es_internacional = es_destino_internacional(texto_destino)
    tipo_viaje = "internacional" if es_internacional else "nacional"
    # Internacional: región de Phoenix (LATAM/EUROPA/ESTADOS UNIDOS/TODO EL MUNDO) o '' si no se reconoce.
    destino_combo = inferir_region_destino(texto_destino) if es_internacional else "CL - NACIONAL"
    lista_destino = destino_raw.split("(")[0].strip().rstrip(".").upper() if destino_raw else ""

    if not producto and (agencia or codigo_cliente):
        producto = determinar_producto(agencia or codigo_cliente, tipo_viaje) or ""

    # 4. File1
    file1 = negocio.strip()
    if not file1 and tipo_viaje == "nacional" and lista_destino and fecha_inicio and fecha_fin:
        try:
            d1 = int(fecha_inicio[:2])
            d2 = int(fecha_fin[:2])
            m2 = int(fecha_fin[3:5])
            a2 = int(fecha_fin[6:])
            file1 = formatear_file_nacional(origen, lista_destino, d1, d2, m2, a2)
        except Exception:
            file1 = ""

    # 5. Pasajeros (desde Excel adjunto o tabla HTML del correo)
    pasajeros: list[Pasajero] = []
    validacion_excel = None
    excel_usado = excels[0] if excels else ""
    if excel_usado:
        validacion_excel = validar_excel_grupal(excel_usado, formato="auto")
        pasajeros = validacion_excel.get("pasajeros", [])
    if not pasajeros and tablas:
        pasajeros = _pasajeros_desde_tablas(tablas)

    if pasajeros and not any(p.es_titular for p in pasajeros):
        pasajeros[0].es_titular = True

    datos = DatosEmision(
        codigo_cliente=codigo_cliente.strip(),
        nombre_agencia=agencia or colegio or codigo_cliente.strip(),
        cliente_final=agencia or colegio or codigo_cliente.strip(),
        producto=producto.strip(),
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
        zona="ALL" if es_internacional else "",
        pais_origen="CHILE",
        destino=destino_combo,
        lista_destino=lista_destino,
        pasajeros=pasajeros,
        correo_solicitante=remitente,
        file1=file1,
        tipo_viaje=tipo_viaje,
    )

    return {
        "datos": datos,
        "excel_usado": excel_usado,
        "validacion_excel": validacion_excel,
        "colegio_o_empresa": colegio,
        "origen": origen,
    }


def auditar_solicitud_mvp(parsed: dict) -> dict:
    """
    Auditoría rigurosa previa a abrir Phoenix:
    - Valida campos obligatorios del contrato.
    - Valida RUT chileno con módulo 11 y detecta errores de tipeo en fechas/datos.
    """
    datos: DatosEmision = parsed["datos"]
    errores: list[str] = []
    advertencias: list[str] = []

    if not datos.codigo_cliente:
        errores.append("Falta código de cliente / agencia (ej: PC.LATITUD, SV24000000).")
    if not datos.fecha_inicio or not datos.fecha_fin:
        errores.append("No se pudo determinar la fecha de inicio y/o fin de vigencia.")
    if not datos.lista_destino and not datos.destino:
        errores.append("Falta el destino del viaje.")
    if datos.tipo_viaje == "internacional" and not datos.destino:
        errores.append(
            "Viaje internacional sin región de destino de Phoenix (LATAM, EUROPA, ESTADOS UNIDOS o TODO EL MUNDO): "
            "indícala con la opción 3 (Editar datos)."
        )
    if datos.fecha_inicio and datos.fecha_fin:
        f_ini = _parsear_fecha_phoenix(datos.fecha_inicio)
        f_fin = _parsear_fecha_phoenix(datos.fecha_fin)
        if not f_ini or not f_fin:
            errores.append(f"Fechas de vigencia con formato inválido: '{datos.fecha_inicio}' -> '{datos.fecha_fin}'.")
        elif f_fin < f_ini:
            errores.append(f"La vigencia termina ({datos.fecha_fin}) antes de empezar ({datos.fecha_inicio}).")
        elif f_ini.date() < datetime.now().date():
            advertencias.append(f"La vigencia empieza en el pasado ({datos.fecha_inicio}); Phoenix puede rechazarla.")
    if not datos.pasajeros:
        errores.append("No hay pasajeros cargados (adjunta Excel válido o tabla de pasajeros).")

    val_excel = parsed.get("validacion_excel")
    if val_excel and val_excel.get("errores"):
        errores.extend(val_excel["errores"])
    if val_excel and val_excel.get("advertencias"):
        advertencias.extend(val_excel["advertencias"])

    for idx, p in enumerate(datos.pasajeros, start=1):
        nombre_ref = f"{p.nombre} {p.apellido_paterno}".strip() or f"Pasajero #{idx}"
        if p.tipo_documento.upper() == "RUN":
            if not validar_rut_modulo11(p.numero_documento, p.digito_verificador):
                advertencias.append(
                    f"Pasajero #{idx} ({nombre_ref}): RUT '{p.numero_documento}-{p.digito_verificador}' "
                    f"no pasa validación módulo 11 o longitud (7-9 dígitos)."
                )
        if not p.fecha_nacimiento:
            errores.append(f"Pasajero #{idx} ({nombre_ref}): falta fecha de nacimiento.")
        else:
            edad = _calcular_edad(p.fecha_nacimiento)
            if edad is None or edad < 0 or edad > 100:
                advertencias.append(
                    f"Pasajero #{idx} ({nombre_ref}): fecha de nacimiento sospechosa '{p.fecha_nacimiento}' (edad={edad})."
                )
            elif edad == 0:
                advertencias.append(
                    f"Pasajero #{idx} ({nombre_ref}): edad calculada = 0 años ('{p.fecha_nacimiento}'); revisar posible error de año."
                )

    return {
        "ok": len(errores) == 0,
        "errores": errores,
        "advertencias": advertencias,
    }


def _imprimir_resumen_mvp(parsed: dict, audit: dict):
    d: DatosEmision = parsed["datos"]
    print("\n" + "=" * 68)
    print("  📋 RESUMEN PRE-EMISIÓN (MVP AUTOEMISIÓN)")
    print("=" * 68)
    print(f"  Cliente / Código : {d.codigo_cliente or '❌ FALTANTE'} ({d.nombre_agencia})")
    print(f"  Producto         : {d.producto or '[automático según cliente]'}")
    print(f"  Vigencia         : {d.fecha_inicio or '❌'} -> {d.fecha_fin or '❌'}")
    print(f"  Tipo / Destino   : {d.tipo_viaje.upper()} | {d.destino} | Lista: {d.lista_destino or '❌'}")
    print(f"  File1            : {d.file1 or '(vacío)'}")
    print(f"  Solicitante      : {d.correo_solicitante or '(sin correo)'}")
    print(f"  Pasajeros        : {len(d.pasajeros)} (Excel: {os.path.basename(parsed['excel_usado']) if parsed['excel_usado'] else 'tabla/manual'})")

    if d.pasajeros:
        print("\n  Primeros pasajeros:")
        for i, p in enumerate(d.pasajeros[:5], 1):
            tit = " [TITULAR]" if p.es_titular else ""
            doc = f"{p.numero_documento}-{p.digito_verificador}" if p.digito_verificador else p.numero_documento
            print(f"    {i}. {p.tipo_documento} {doc} | {p.nombre} {p.apellido_paterno} | Nac: {p.fecha_nacimiento}{tit}")
        if len(d.pasajeros) > 5:
            print(f"    ... y {len(d.pasajeros) - 5} pasajero(s) más.")

    if audit["errores"]:
        print("\n  ❌ ERRORES CRÍTICOS (bloquean emisión):")
        for e in audit["errores"]:
            print(f"     - {e}")
    if audit["advertencias"]:
        print("\n  ⚠️  ADVERTENCIAS DE AUDITORÍA (revisar antes de emitir):")
        for w in audit["advertencias"]:
            print(f"     - {w}")
    if audit["ok"] and not audit["advertencias"]:
        print("\n  ✅ Auditoría 100% limpia: todos los datos y RUTs están validados.")
    print("=" * 68)


def _editar_campos_manuales(parsed: dict):
    d: DatosEmision = parsed["datos"]
    print("\n✏️  EDICIÓN RÁPIDA (ENTER para mantener valor actual)")
    d.codigo_cliente = input(f"  Código cliente [{d.codigo_cliente}]: ").strip() or d.codigo_cliente
    d.producto = input(f"  Producto [{d.producto}]: ").strip() or d.producto
    d.fecha_inicio = input(f"  Fecha inicio (DD/MM/AAAA) [{d.fecha_inicio}]: ").strip() or d.fecha_inicio
    d.fecha_fin = input(f"  Fecha fin (DD/MM/AAAA) [{d.fecha_fin}]: ").strip() or d.fecha_fin
    d.tipo_viaje = (input(f"  Tipo viaje (nacional/internacional) [{d.tipo_viaje}]: ").strip().lower() or d.tipo_viaje)
    d.lista_destino = input(f"  Lista destino [{d.lista_destino}]: ").strip() or d.lista_destino
    d.destino = input(f"  Destino Phoenix (ej. LATAM, EUROPA, CL - NACIONAL) [{d.destino}]: ").strip() or d.destino
    d.file1 = input(f"  File1 [{d.file1}]: ").strip() or d.file1


def ejecutar_mvp_interactivo():
    """Menú segmentado del MVP de Autoemisión Visual desde Correo."""
    print("\n" + "╔" + "═" * 60 + "╗")
    print("║   🚀 MVP AUTOEMISIÓN VISUAL DESDE CORREO (EN PANTALLA)    ║")
    print("╚" + "═" * 60 + "╝")
    print("  1. Cargar solicitud desde archivo .eml (con adjunto/tabla)")
    print("  2. Pegar texto del correo + ruta de Excel opcional")
    print("  0. Volver al menú principal")

    op = input("\nSelecciona una opción: ").strip()
    if op == "0":
        return

    parsed = None
    if op == "1":
        ruta_eml = input("Ruta al archivo .eml: ").strip().strip('"')
        if not ruta_eml or not os.path.exists(ruta_eml):
            print(f"❌ No existe el archivo: {ruta_eml}")
            return
        info = extraer_desde_eml(ruta_eml)
        parsed = parsear_solicitud_correo(
            texto=info["body"],
            asunto=info["subject"],
            remitente=info["from"],
            fecha_correo=info["date"],
            tablas_html=info["tablas_html"],
            rutas_excel=info["adjuntos_excel"],
        )
    elif op == "2":
        print("Pega el texto del correo y termina con una línea que diga FIN:")
        lineas = []
        while True:
            ln = input()
            if ln.strip().upper() == "FIN":
                break
            lineas.append(ln)
        ruta_xlsx = input("Ruta del Excel de pasajeros (opcional, ENTER si no hay): ").strip().strip('"')
        remitente = input("Correo del solicitante (opcional): ").strip()
        parsed = parsear_solicitud_correo(
            texto="\n".join(lineas),
            remitente=remitente,
            rutas_excel=[ruta_xlsx] if ruta_xlsx else [],
        )
    else:
        print("❌ Opción inválida.")
        return

    while True:
        audit = auditar_solicitud_mvp(parsed)
        _imprimir_resumen_mvp(parsed, audit)
        print("\nAcciones:")
        print("  1. Ejecutar AUTOEMISIÓN VISIBLE en pantalla (con HUD y confirmación antes de Aceptar)")
        print("  2. Ejecutar DRY-RUN VISIBLE en pantalla (llena todo en Phoenix y presiona Cancelar)")
        print("  3. Editar datos detectados")
        print("  0. Cancelar y volver")

        accion = input("\nElige acción: ").strip()
        if accion == "0":
            return
        if accion == "3":
            _editar_campos_manuales(parsed)
            continue
        if accion not in ("1", "2"):
            continue
        if not audit["ok"]:
            print("❌ Corrige primero los errores críticos (opción 3) antes de abrir Phoenix.")
            continue

        es_dry = accion == "2"
        if audit["advertencias"] and not es_dry:
            seguro = input("⚠️ Hay advertencias de auditoría (RUT/fechas). ¿Continuar de todas formas? (s/n) [n]: ").strip().lower()
            if seguro != "s":
                continue

        browser = input("Navegador visible (edge/chrome/chromium/firefox) [edge]: ").strip().lower() or "edge"
        datos: DatosEmision = parsed["datos"]
        grupal = len(datos.pasajeros) > 1 or bool(parsed.get("excel_usado"))

        phoenix = (
            PhoenixEmisionGrupal(browser=browser, headless=False, slow_mo=250, visual_hud=True)
            if grupal
            else PhoenixAutomation(browser=browser, headless=False, slow_mo=250, visual_hud=True)
        )

        if not phoenix.conectar_navegador():
            return

        try:
            phoenix.navegar_a_phoenix()
            if grupal:
                res = phoenix.emitir_voucher_grupal(
                    datos,
                    dry_run=es_dry,
                    enviar_correo_asegurado=not es_dry,
                    confirmar_antes_aceptar=not es_dry,
                )
            else:
                res = phoenix.emitir_voucher(
                    datos,
                    dry_run=es_dry,
                    enviar_correo_asegurado=not es_dry,
                    confirmar_antes_aceptar=not es_dry,
                )

            print("\n" + "=" * 55)
            if res.get("exito"):
                print("  ✅ AUTOEMISIÓN FINALIZADA" + (" [DRY-RUN]" if es_dry else ""))
                if res.get("numero_voucher"):
                    print(f"  🎫 Voucher N°: {res['numero_voucher']}")
                if res.get("archivo_voucher"):
                    print(f"  📄 PDF guardado: {res['archivo_voucher']}")
                if res.get("precios"):
                    print(f"  💰 Precios: {res['precios']}")
            else:
                print(f"  ❌ FALLÓ LA EMISIÓN: {res.get('error')}")
            print("=" * 55)
        finally:
            input("\nPresiona ENTER para cerrar la ventana del navegador...")
            phoenix.cerrar()
        return


if __name__ == "__main__":
    ejecutar_mvp_interactivo()
