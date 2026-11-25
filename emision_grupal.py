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
import logging
import unicodedata
import zipfile
from datetime import date, datetime
import openpyxl

from phoenix_emision import PhoenixAutomation, DatosEmision, Pasajero, _calcular_edad
from utils import (
    cargar_config,
    determinar_producto,
    debe_enviar_precios,
    generar_resumen_emision,
    limpiar_rut,
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

    wb = openpyxl.Workbook()
    ws = wb.active

    if str(formato).lower() == "oficial":
        ws.title = "Sheet1"
        ws.append(COLUMNAS_OFICIAL_PHOENIX)
        if con_ejemplo:
            ws.append([
                1007,
                "12345678-9",
                "JUAN ANTONIO",
                "PEREZ ROJAS",
                date(1990, 12, 9),
                35,
                5,
                "ejemplo@correo.cl",
                "981571641",
            ])
            ws.cell(row=2, column=5).number_format = "dd/MM/yyyy"
    else:
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
    return None


def validar_excel_grupal(
    ruta_excel: str,
    formato: str | int = "auto",
    codigo_run: int = 1007,
) -> dict:
    """
    Valida estrictamente una planilla Excel enviada por un cliente antes de subirla a Phoenix.
    Acepta la plantilla oficial Phoenix ('ID Type | Fiscal ID | ...') y formatos en español (1 y 2).
    Retorna {"valido": bool, "formato_detectado": str, "errores": list[str], "pasajeros": list[Pasajero]}.
    """
    errores: list[str] = []
    pasajeros: list[Pasajero] = []

    if not os.path.exists(ruta_excel):
        return {"valido": False, "formato_detectado": "", "errores": [f"No se encontró el archivo: {ruta_excel}"], "pasajeros": []}

    try:
        wb = _cargar_workbook_seguro(ruta_excel)
    except Exception as e:
        return {"valido": False, "formato_detectado": "", "errores": [f"El archivo no es un Excel (.xlsx) válido: {e}"], "pasajeros": []}

    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    if len(filas) < 2:
        return {
            "valido": False,
            "formato_detectado": "",
            "errores": ["El Excel debe incluir encabezados en la fila 1 y al menos un pasajero desde la fila 2."],
            "pasajeros": [],
        }

    headers_raw = [str(c or "").strip() for c in filas[0]]
    headers_norm = [_norm_header(h) for h in headers_raw]
    detectado = _detectar_formato_headers(headers_norm)

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
            return {"valido": False, "formato_detectado": detectado or "", "errores": errores, "pasajeros": []}
    elif not detectado:
        errores.append(
            f"Encabezados o orden de columnas incorrectos. "
            f"Se esperaba plantilla oficial: {' | '.join(COLUMNAS_OFICIAL_PHOENIX)}. "
            f"Recibido: {' | '.join(headers_raw)}."
        )
        return {"valido": False, "formato_detectado": "", "errores": errores, "pasajeros": []}

    cols_activas = (
        COLUMNAS_OFICIAL_PHOENIX if detectado == "oficial"
        else (COLUMNAS_FORMATO_1 if detectado == "1" else COLUMNAS_FORMATO_2)
    )
    esperadas_norm = [_norm_header(c) for c in cols_activas]

    for idx_fila, row in enumerate(filas[1:], start=2):
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
            tel_raw = str(data["mobile"] or "").strip()
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
            tel_raw = str(data["telefono movil"] or "").strip()

        # 1. Tipo de Documento
        tipo_doc = ""
        if td_raw in ("1007", "4", "RUN"):
            tipo_doc = "RUN"
        elif td_raw in ("2", "PASAPORTE", "PS"):
            tipo_doc = "PASAPORTE"
        else:
            errores.append(f"Fila {idx_fila}: ID Type / Tipo de Documento inválido ('{td_raw}'). Debe ser {codigo_run} (RUN) o 2 (PASAPORTE).")

        # 2. Documento y DV
        num_doc, dv_doc = "", ""
        if not fiscal_raw:
            errores.append(f"Fila {idx_fila}: Fiscal ID / Número de Documento vacío.")
        elif "." in fiscal_raw or " " in fiscal_raw:
            errores.append(f"Fila {idx_fila}: Documento ('{fiscal_raw}') no debe contener puntos ni espacios.")
        elif tipo_doc == "RUN":
            if "-" in fiscal_raw:
                num_doc, dv_doc = limpiar_rut(fiscal_raw)
            elif detectado != "oficial":
                num_doc, dv_doc = num_raw, dv_raw
            else:
                num_doc, dv_doc = limpiar_rut(fiscal_raw) if len(fiscal_raw) >= 8 else (fiscal_raw, "")

            dv_doc = dv_doc.upper()
            if not re.fullmatch(r"\d{7,9}", num_doc):
                errores.append(f"Fila {idx_fila}: RUN '{fiscal_raw}' debe tener entre 7 y 9 dígitos numéricos antes del DV.")
            if not re.fullmatch(r"[0-9K]", dv_doc):
                errores.append(f"Fila {idx_fila}: DV '{dv_doc}' inválido para RUN (debe ser 0-9 o K).")
        else:
            num_doc = fiscal_raw
            if detectado != "oficial" and dv_raw not in ("0", ""):
                errores.append(f"Fila {idx_fila}: Para PASAPORTE, la columna DV debe ser 0 (recibido '{dv_raw}').")

        # 3. Nombres y Apellidos
        if not nombres:
            errores.append(f"Fila {idx_fila}: Falta First name / Nombres.")
        if not apellidos:
            errores.append(f"Fila {idx_fila}: Falta Surname / Apellidos.")

        # 4. Fecha de Nacimiento
        dt_nac = _parsear_fecha(dob_raw)
        if not dt_nac:
            errores.append(f"Fila {idx_fila}: D.o.b. / Fecha de Nacimiento inválida ('{dob_raw}'). Usar DD/MM/AAAA o DD-MM-AAAA.")
            fnac_str = ""
        else:
            fnac_str = dt_nac.strftime("%d/%m/%Y")

        # 5. Género (en Phoenix PRE #ayudaExcel: 2=Femenino, 5=Masculino)
        sexo_val = 2
        if detectado in ("oficial", "2"):
            if gen_raw not in ("2", "5"):
                errores.append(f"Fila {idx_fila}: Gender / Género inválido ('{gen_raw}'). Usar 2 (Femenino) o 5 (Masculino).")
            else:
                sexo_val = int(gen_raw)
        if detectado == "2":
            cod_pais_raw = str(data["cod. pais movil"] or "").strip()
            if not re.fullmatch(r"\d{1,4}", cod_pais_raw):
                errores.append(f"Fila {idx_fila}: Cod. País Móvil inválido ('{cod_pais_raw}'). Ejemplo: 56.")

        # 6. Teléfono Móvil
        if not re.fullmatch(r"\d{8,12}", tel_raw):
            errores.append(f"Fila {idx_fila}: Mobile / Teléfono Móvil inválido ('{tel_raw}'). Debe contener solo dígitos (ej: 912345678).")

        # 7. Email
        if email_raw and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email_raw):
            errores.append(f"Fila {idx_fila}: Email inválido ('{email_raw}').")

        if not errores:
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
        Crea el fichero Excel (.xlsx) en el formato oficial de Phoenix (#plantillaExcel):
        Hoja 'Sheet1', columnas: ID Type | Fiscal ID | First name | Surname | D.o.b. | Age | Gender | Email | Mobile.
        """
        if not ruta_salida:
            ruta_salida = os.path.join(os.path.dirname(__file__), "ficheros", "carga_temporal.xlsx")

        os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        config = cargar_config()

        ws.append(COLUMNAS_OFICIAL_PHOENIX)

        for idx_row, p in enumerate(pasajeros, start=2):
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
            telefono = (p.telefono or "").strip() or "900000000"
            email = (p.email or "").strip() or correo_respaldo.strip()

            ws.append([
                id_type,
                fiscal_id,
                p.nombre.strip().upper(),
                apellidos,
                dob_val,
                int(edad),
                genero,
                email,
                str(telefono),
            ])
            if isinstance(dob_val, date):
                ws.cell(row=idx_row, column=5).number_format = "dd/MM/yyyy"

        wb.save(ruta_salida)
        logger.info(f"✅ Fichero oficial Phoenix creado: {ruta_salida} ({len(pasajeros)} pasajeros)")
        return ruta_salida

    def cargar_fichero(self, ruta_fichero: str) -> tuple[bool, str]:
        """
        Sube el fichero Excel en la pestaña T1 usando Playwright set_input_files().
        Retorna (exito, mensaje_detalle).
        """
        logger.info(f"Cargando fichero en #windowCargaAsegurados_PW-1: {ruta_fichero}")

        self.frame.locator("#Wcr_carTabPage_T1").click()
        self.page.wait_for_timeout(500)

        self.frame.locator("#btnCargaPlantillaAsegurado").click()
        self.frame.locator("#windowCargaAsegurados_PW-1").wait_for(state="visible", timeout=self.timeout_ms)

        self.frame.locator("#fuBulk_TextBox0_Input").set_input_files(os.path.abspath(ruta_fichero))
        self.page.wait_for_timeout(4000)

        modal_text = self.frame.locator("#windowCargaAsegurados_PW-1").inner_text().strip()
        if "incorrecto" in modal_text.lower() or "error" in modal_text.lower():
            logger.error(f"❌ Phoenix rechazó el formato del fichero: {modal_text}")
            self._capturar_screenshot("error_formato_grupal")
            return False, modal_text

        btn_cerrar = self.frame.locator("#windowCargaAsegurados_HCB-1")
        if btn_cerrar.count() > 0:
            btn_cerrar.click()
            self.page.wait_for_timeout(800)

        logger.info("✅ Fichero cargado correctamente")
        return True, modal_text

    def eliminar_fila_en_blanco(self):
        """Elimina la fila 1 si quedó en blanco tras la carga (#chkElimRiesgo_1 -> #btnEliminarRiesgos)."""
        doc1 = self.frame.locator("#txtIdFiscal_1")
        if doc1.count() > 0 and doc1.input_value().strip():
            return

        logger.info("Eliminando fila 1 en blanco (#chkElimRiesgo_1 -> #btnEliminarRiesgos)...")
        chk1 = self.frame.locator("#chkElimRiesgo_1")
        if chk1.count() > 0 and not chk1.is_checked():
            chk1.check()
        self.frame.locator("#btnEliminarRiesgos").click()
        self.page.wait_for_timeout(1200)

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
                resultado["precios"] = self.obtener_precios(datos.tipo_viaje)
                self._capturar_screenshot("dry_run_grupal_antes_aceptar")
                self.frame.locator("#Wcr_btnCancelar").click()
                resultado["exito"] = True
                resultado["resumen"] = "Dry-run grupal completado con Playwright sin emitir."
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
