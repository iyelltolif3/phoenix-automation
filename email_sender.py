"""
Módulo de envío de vouchers por correo electrónico.
===================================================
Permite enviar (o generar borrador .eml / Outlook) el PDF del voucher emitido
directamente al asegurado y/o al ejecutivo solicitante.

Modos soportados (config["email"]["modo"]):
- "eml"           : Genera un archivo .eml listo con el PDF adjunto (no requiere credenciales).
- "smtp"          : Envío automático directo vía SMTP (Office 365 / Gmail / corporativo).
- "outlook_send"  : Envío automático vía Outlook clásico de Windows (COM).
- "outlook_draft" : Abre el borrador en Outlook clásico de Windows con el adjunto listo.
"""

import mimetypes
import os
import re
import smtplib
import subprocess
import logging
from datetime import datetime
from email.message import EmailMessage

from utils import cargar_config, resolver_ruta

logger = logging.getLogger(__name__)


def _sanitizar(texto: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "-", texto).strip()


def _validar_emails(destinatarios: list[str] | str) -> list[str]:
    if isinstance(destinatarios, str):
        candidatos = [x.strip() for x in re.split(r"[;,]", destinatarios) if x.strip()]
    else:
        candidatos = [str(x).strip() for x in destinatarios if x and str(x).strip()]

    validos = []
    for email in candidatos:
        if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) and email.lower() not in [v.lower() for v in validos]:
            validos.append(email)
    return validos


def construir_cuerpo_asegurado(
    nombre_asegurado: str,
    numero_voucher: str,
    fecha_inicio: str,
    fecha_fin: str,
    destino: str,
    nombre_agencia: str = "",
) -> str:
    """Genera el cuerpo del correo destinado al asegurado."""
    saludo = f"Estimado/a {nombre_asegurado}," if nombre_asegurado else "Estimado/a,"
    linea_agencia = f"\n· Agencia / Contratante: {nombre_agencia}" if nombre_agencia else ""
    return (
        f"{saludo}\n\n"
        f"Junto con saludar, adjuntamos su Voucher de Asistencia en Viaje.\n\n"
        f"Detalle de su asistencia:\n"
        f"· Voucher N°: {numero_voucher or 'Adjunto'}\n"
        f"· Vigencia: {fecha_inicio} al {fecha_fin}\n"
        f"· Destino: {destino or 'Según voucher'}"
        f"{linea_agencia}\n\n"
        f"Le recomendamos revisar sus datos personales y conservar este documento "
        f"durante todo su viaje.\n\n"
        f"Atentamente,\n"
        f"Equipo de Emisiones - MAWDY"
    )


def _construir_mensaje(
    remitente: str,
    destinatarios: list[str],
    asunto: str,
    cuerpo: str,
    ruta_pdf: str | None = None,
    cc: list[str] | None = None,
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = remitente
    msg["To"] = ", ".join(destinatarios)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = asunto
    msg.set_content(cuerpo, charset="utf-8")

    if ruta_pdf and os.path.exists(ruta_pdf):
        ctype, _ = mimetypes.guess_type(ruta_pdf)
        maintype, subtype = (ctype.split("/", 1) if ctype else ("application", "octet-stream"))
        with open(ruta_pdf, "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(ruta_pdf),
        )
    return msg


def _enviar_smtp(msg: EmailMessage, email_cfg: dict, smtp_password: str | None = None) -> dict:
    host = email_cfg.get("smtp_host", "smtp.office365.com")
    port = int(email_cfg.get("smtp_port", 587))
    use_tls = bool(email_cfg.get("use_tls", True))
    remitente = email_cfg.get("remitente", "")
    pass_env = email_cfg.get("password_env", "PHOENIX_SMTP_PASSWORD")
    password = smtp_password or os.environ.get(pass_env, "")

    if not remitente or not password:
        return {
            "enviado": False,
            "modo": "smtp",
            "error": (
                f"Faltan credenciales SMTP: configura 'remitente' en config.json y "
                f"la variable de entorno '{pass_env}'."
            ),
        }

    with smtplib.SMTP(host, port, timeout=30) as server:
        if use_tls:
            server.starttls()
        server.login(remitente, password)
        server.send_message(msg)

    return {"enviado": True, "modo": "smtp", "archivo_eml": "", "error": ""}


def _guardar_eml(msg: EmailMessage, email_cfg: dict, etiqueta: str) -> dict:
    out_dir = resolver_ruta(email_cfg.get("eml_output_dir"), "correos")
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = _sanitizar(f"{etiqueta}_{stamp}") + ".eml"
    ruta_eml = os.path.join(out_dir, nombre)
    with open(ruta_eml, "wb") as f:
        f.write(bytes(msg))
    logger.info(f"📧 Borrador .eml guardado: {ruta_eml}")
    return {"enviado": True, "modo": "eml", "archivo_eml": ruta_eml, "error": ""}


def _enviar_outlook_com(
    destinatarios: list[str],
    asunto: str,
    cuerpo: str,
    ruta_pdf: str | None,
    cc: list[str] | None = None,
    solo_borrador: bool = False,
) -> dict:
    """Usa Outlook clásico en Windows vía PowerShell COM sin requerir pywin32."""
    to_str = "; ".join(destinatarios).replace("'", "''")
    cc_str = "; ".join(cc or []).replace("'", "''")
    subj_str = asunto.replace("'", "''")
    body_str = cuerpo.replace("'", "''").replace("\r\n", "`n").replace("\n", "`r`n")
    adj_cmd = ""
    if ruta_pdf and os.path.exists(ruta_pdf):
        pdf_escaped = os.path.abspath(ruta_pdf).replace("'", "''")
        adj_cmd = f"$mail.Attachments.Add('{pdf_escaped}') | Out-Null;"

    accion = "$mail.Display()" if solo_borrador else "$mail.Send()"
    ps_script = (
        "$ErrorActionPreference = 'Stop'; "
        "$outlook = New-Object -ComObject Outlook.Application; "
        "$mail = $outlook.CreateItem(0); "
        f"$mail.To = '{to_str}'; "
        f"$mail.CC = '{cc_str}'; "
        f"$mail.Subject = '{subj_str}'; "
        f"$mail.Body = \"{body_str}\"; "
        f"{adj_cmd} "
        f"{accion}"
    )

    proc = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True,
        text=True,
        timeout=30,
    )
    if proc.returncode != 0:
        return {
            "enviado": False,
            "modo": "outlook_draft" if solo_borrador else "outlook_send",
            "error": proc.stderr.strip() or "No se pudo invocar Outlook COM",
        }

    return {
        "enviado": True,
        "modo": "outlook_draft" if solo_borrador else "outlook_send",
        "archivo_eml": "",
        "error": "",
    }


def enviar_correo(
    destinatarios: list[str] | str,
    asunto: str,
    cuerpo: str,
    ruta_pdf: str | None = None,
    cc: list[str] | str | None = None,
    modo_override: str | None = None,
    smtp_password: str | None = None,
    etiqueta_archivo: str = "voucher",
) -> dict:
    """Despacha un correo con el voucher adjunto según el modo configurado."""
    config = cargar_config()
    email_cfg = config.get("email", {})
    modo = (modo_override or email_cfg.get("modo", "eml")).lower()
    remitente = email_cfg.get("remitente", "emision.viajes@mawdy.com")

    lista_to = _validar_emails(destinatarios)
    lista_cc = _validar_emails(cc or [])

    if not lista_to:
        return {"enviado": False, "modo": modo, "destinatarios": [], "error": "No hay correos válidos de destinatario."}

    if not ruta_pdf or not os.path.exists(ruta_pdf):
        logger.warning(f"⚠️ No se encontró PDF para adjuntar: {ruta_pdf}")

    msg = _construir_mensaje(remitente, lista_to, asunto, cuerpo, ruta_pdf=ruta_pdf, cc=lista_cc)

    try:
        if modo == "smtp":
            res = _enviar_smtp(msg, email_cfg, smtp_password=smtp_password)
        elif modo == "outlook_send":
            res = _enviar_outlook_com(lista_to, asunto, cuerpo, ruta_pdf, cc=lista_cc, solo_borrador=False)
        elif modo == "outlook_draft":
            res = _enviar_outlook_com(lista_to, asunto, cuerpo, ruta_pdf, cc=lista_cc, solo_borrador=True)
        elif modo == "eml":
            res = _guardar_eml(msg, email_cfg, etiqueta=etiqueta_archivo)
        else:
            res = {"enviado": False, "modo": modo, "error": f"Modo de correo no soportado: {modo}"}

        res["destinatarios"] = lista_to
        if res.get("enviado"):
            logger.info(f"✅ Correo ({modo}) procesado para {lista_to}")
        else:
            logger.error(f"❌ Error enviando correo ({modo}): {res.get('error')}")
        return res
    except Exception as e:
        logger.error(f"❌ Excepción en envío de correo ({modo}): {e}")
        return {"enviado": False, "modo": modo, "destinatarios": lista_to, "error": str(e)}


def enviar_voucher_al_asegurado(
    pasajeros: list,
    numero_voucher: str,
    ruta_pdf: str | None,
    fecha_inicio: str,
    fecha_fin: str,
    destino: str,
    nombre_agencia: str = "",
    correo_respaldo: str = "",
    modo_override: str | None = None,
) -> dict:
    """
    Envía el voucher capturado a los correos de los asegurados (excluyendo correos vacíos).
    Si ningún pasajero tiene correo propio y se indica correo_respaldo, lo envía allí.
    """
    correos_pax = _validar_emails([getattr(p, "email", "") for p in pasajeros])
    if not correos_pax and correo_respaldo:
        correos_pax = _validar_emails(correo_respaldo)

    titular = next((p for p in pasajeros if getattr(p, "es_titular", False)), pasajeros[0] if pasajeros else None)
    nombre_tit = f"{getattr(titular, 'nombre', '')} {getattr(titular, 'apellido_paterno', '')}".strip()

    asunto = f"Voucher de Asistencia en Viaje N° {numero_voucher or 'Emitido'} - {nombre_tit or 'Pasajero'}"
    cuerpo = construir_cuerpo_asegurado(
        nombre_asegurado=nombre_tit,
        numero_voucher=numero_voucher,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
        destino=destino,
        nombre_agencia=nombre_agencia,
    )

    return enviar_correo(
        destinatarios=correos_pax,
        asunto=asunto,
        cuerpo=cuerpo,
        ruta_pdf=ruta_pdf,
        modo_override=modo_override,
        etiqueta_archivo=f"voucher_{numero_voucher or 'sin_num'}_{nombre_tit or 'asegurado'}",
    )


def enviar_aviso_formato_incorrecto(
    correo_cliente: str | list[str],
    errores: list[str],
    nombre_cliente: str = "",
    file_ref: str = "",
    ruta_plantilla_adjunta: str | None = None,
    modo_override: str | None = None,
) -> dict:
    """
    Notifica automáticamente al cliente/ejecutivo cuando la planilla grupal enviada
    no cumple con el formato requerido por Phoenix, detallando los errores detectados
    y adjuntando la plantilla Excel óptima.
    """
    saludo = f"Estimado/a {nombre_cliente}," if nombre_cliente else "Estimado/a,"
    ref_txt = f" (Ref: {file_ref})" if file_ref else ""
    lista_errores = "\n".join(f"  - {e}" for e in errores) if errores else "  - El archivo no coincide con la estructura esperada por Phoenix."

    asunto = f"Solicitud de corrección de planilla grupal{ref_txt} - Formato incorrecto"
    cuerpo = (
        f"{saludo}\n\n"
        f"Junto con saludar, informamos que no fue posible procesar la emisión grupal{ref_txt} "
        f"debido a que el archivo Excel enviado no cumple con el formato requerido por Phoenix.\n\n"
        f"Observaciones detectadas:\n"
        f"{lista_errores}\n\n"
        f"Por favor, reenviar la nómina utilizando la plantilla adjunta (hoja 'Sheet1') y respetando estas reglas:\n"
        f"1. Orden exacto de columnas (plantilla oficial Phoenix):\n"
        f"   ID Type | Fiscal ID | First name | Surname | D.o.b. | Age | Gender | Email | Mobile\n"
        f"2. ID Type: 1007 = RUN (Persona Natural) y 2 = Pasaporte.\n"
        f"3. Fiscal ID: sin puntos ni espacios (ej: 12345678-9 para RUN, AB123456 para Pasaporte).\n"
        f"4. D.o.b. (Fecha de Nacimiento): formato DD/MM/AAAA o DD-MM-AAAA.\n"
        f"5. Age: edad numérica del pasajero.\n"
        f"6. Gender: 2 = Femenino, 5 = Masculino.\n"
        f"7. Mobile: teléfono móvil numérico sin '+' ni espacios (ej: 912345678).\n\n"
        f"Adjuntamos la plantilla óptima para facilitar el llenado.\n\n"
        f"Atentamente,\n"
        f"Equipo de Emisiones - MAWDY"
    )

    return enviar_correo(
        destinatarios=correo_cliente,
        asunto=asunto,
        cuerpo=cuerpo,
        ruta_pdf=ruta_plantilla_adjunta,
        modo_override=modo_override,
        etiqueta_archivo=f"aviso_formato_incorrecto_{file_ref or 'grupal'}",
    )
