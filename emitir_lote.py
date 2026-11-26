"""
Phoenix v4 - Emisión por Lote desde Excel
=========================================
Lee un archivo Excel de solicitudes (`plantilla_solicitudes.xlsx`),
agrupa los pasajeros por `solicitud_id`, aplica las reglas de negocio
del manual y emite automáticamente cada solicitud en Phoenix v4.

Genera `resultados_emision.xlsx` con:
- Estado (OK / ERROR)
- Número de Voucher
- Ruta al PDF descargado
- Precios extraídos (CLP, USD, TC)
- Texto del correo de respuesta listo para copiar/enviar
"""

import argparse
import os
import sys
from datetime import datetime

import openpyxl

from phoenix_emision import PhoenixAutomation, DatosEmision, Pasajero
from emision_grupal import PhoenixEmisionGrupal, _cargar_workbook_seguro
from utils import buscar_dealer_por_codigo, buscar_dealer_por_correo, limpiar_rut


COLUMNAS_PLANTILLA = [
    "solicitud_id",       # Ej: SOL-001 (filas con el mismo ID van en el mismo voucher)
    "codigo_cliente",     # Ej: SV24000000 (opcional si se indica correo_dealer)
    "correo_dealer",      # Ej: jroman@esquerre.cl
    "producto",           # Opcional (vacío = automático o regla de negocio)
    "fecha_inicio",       # DD/MM/AAAA
    "fecha_fin",          # DD/MM/AAAA
    "tipo_viaje",         # nacional / internacional
    "zona",               # Opcional (ej: Nacional / ALL)
    "pais_origen",        # Ej: CHILE
    "destino",            # Ej: NACIONAL / EUROPA / SCHENGEN
    "lista_destino",      # Ciudad (nacional) o País (internacional)
    "file1",              # Número de negocio, booking, expediente o formato intrarregional
    "modo_grupal",        # SI / NO (SI usa carga por fichero Excel)
    "formato_fichero",    # 1 o 2 (solo si modo_grupal=SI)
    "es_titular",         # SI / NO
    "tipo_documento",     # RUN o PASAPORTE
    "numero_documento",   # Con o sin DV (ej: 12345678-9)
    "dv",                 # Opcional si viene con guión
    "nombre",             # Nombres del pasajero
    "apellido_paterno",   # Apellido paterno
    "apellido_materno",   # Apellido materno
    "fecha_nacimiento",   # DD/MM/AAAA
    "email",              # Email del pasajero (opcional)
    "telefono",           # Teléfono (opcional, defecto 900000000)
]


def crear_plantilla_excel(ruta: str):
    """Crea un archivo Excel plantilla con una fila de ejemplo."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solicitudes"
    ws.append(COLUMNAS_PLANTILLA)
    ws.append([
        "SOL-001",
        "SV24000000",
        "jroman@esquerre.cl",
        "",
        "20/12/2026",
        "27/12/2026",
        "nacional",
        "",
        "CHILE",
        "NACIONAL",
        "PUERTO VARAS",
        "SANTIAGO-PUERTO VARAS 20-27 DIC.26",
        "NO",
        1,
        "SI",
        "RUN",
        "12345678",
        "9",
        "JUAN",
        "PEREZ",
        "LOPEZ",
        "01/01/1990",
        "juan@ejemplo.cl",
        "912345678",
    ])
    wb.save(ruta)
    print(f"✅ Plantilla creada en: {ruta}")


def _formatear_fecha(val) -> str:
    if val is None:
        return ""
    if hasattr(val, "strftime"):
        return val.strftime("%d/%m/%Y")
    return str(val).strip().replace("-", "/")


def cargar_solicitudes(ruta_excel: str) -> list[dict]:
    """Lee el Excel y agrupa las filas por solicitud_id."""
    wb = _cargar_workbook_seguro(ruta_excel)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return []

    headers = [str(h or "").strip().lower() for h in rows[0]]
    grupos: dict[str, dict] = {}

    for row in rows[1:]:
        if not any(row):
            continue
        r = dict(zip(headers, row))
        sol_id = str(r.get("solicitud_id") or f"SOL-{len(grupos)+1}").strip()

        if sol_id not in grupos:
            codigo = str(r.get("codigo_cliente") or "").strip()
            correo = str(r.get("correo_dealer") or "").strip()
            agencia = ""

            if codigo:
                dealer = buscar_dealer_por_codigo(codigo)
                if dealer:
                    agencia = dealer["agencia"]
                    correo = correo or dealer["correo"]
            elif correo:
                matches = buscar_dealer_por_correo(correo)
                if matches:
                    codigo = matches[0]["codigo"]
                    agencia = matches[0]["agencia"]

            grupos[sol_id] = {
                "solicitud_id": sol_id,
                "modo_grupal": str(r.get("modo_grupal") or "NO").strip().upper() in ("SI", "SÍ", "YES", "TRUE", "1"),
                "formato_fichero": int(r.get("formato_fichero") or 1),
                "datos": DatosEmision(
                    codigo_cliente=codigo,
                    nombre_agencia=agencia,
                    cliente_final=agencia,
                    producto=str(r.get("producto") or "").strip(),
                    fecha_inicio=_formatear_fecha(r.get("fecha_inicio")),
                    fecha_fin=_formatear_fecha(r.get("fecha_fin")),
                    tipo_viaje=str(r.get("tipo_viaje") or "nacional").strip().lower(),
                    zona=str(r.get("zona") or "").strip(),
                    pais_origen=str(r.get("pais_origen") or "CHILE").strip(),
                    destino=str(r.get("destino") or "").strip(),
                    lista_destino=str(r.get("lista_destino") or "").strip(),
                    file1=str(r.get("file1") or "").strip(),
                    correo_solicitante=correo,
                    pasajeros=[],
                ),
            }

        tipo_doc = str(r.get("tipo_documento") or "RUN").strip().upper()
        num_doc = str(r.get("numero_documento") or "").strip()
        dv = str(r.get("dv") or "").strip()
        if tipo_doc == "RUN" and "-" in num_doc and not dv:
            num_doc, dv = limpiar_rut(num_doc)

        es_tit = str(r.get("es_titular") or "").strip().upper() in ("SI", "SÍ", "YES", "TRUE", "1")
        pasajero = Pasajero(
            tipo_documento=tipo_doc,
            numero_documento=num_doc,
            digito_verificador=dv,
            nombre=str(r.get("nombre") or "").strip(),
            apellido_paterno=str(r.get("apellido_paterno") or "").strip(),
            apellido_materno=str(r.get("apellido_materno") or "").strip(),
            fecha_nacimiento=_formatear_fecha(r.get("fecha_nacimiento")),
            email=str(r.get("email") or "").strip(),
            telefono=str(r.get("telefono") or "900000000").strip(),
            es_titular=es_tit,
        )
        grupos[sol_id]["datos"].pasajeros.append(pasajero)

    # Asegurar al menos un titular en cada grupo
    for g in grupos.values():
        paxs = g["datos"].pasajeros
        if paxs and not any(p.es_titular for p in paxs):
            paxs[0].es_titular = True

    return list(grupos.values())


def exportar_resultados(resultados: list[dict], ruta_salida: str):
    """Guarda el resumen de todas las emisiones procesadas en Excel."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append([
        "Solicitud ID",
        "Estado",
        "Dry Run",
        "N° Voucher",
        "CLP",
        "USD",
        "Tipo Cambio",
        "PDF Voucher",
        "Correo Asegurado",
        "Archivo EML / Error Correo",
        "Resumen Email",
        "Error",
    ])

    for r in resultados:
        precios = r.get("precios") or {}
        correo = r.get("correo_asegurado") or {}
        estado_correo = "ENVIADO" if correo.get("enviado") else ("NO_ENVIADO" if correo else "-")
        detalle_correo = correo.get("archivo_eml") or correo.get("error") or ", ".join(correo.get("destinatarios", []))
        ws.append([
            r.get("solicitud_id", ""),
            "OK" if r.get("exito") else "ERROR",
            "SI" if r.get("dry_run") else "NO",
            r.get("numero_voucher", ""),
            precios.get("clp", ""),
            precios.get("usd", ""),
            precios.get("tc", ""),
            r.get("archivo_voucher", ""),
            estado_correo,
            detalle_correo,
            r.get("resumen", ""),
            r.get("error", ""),
        ])

    wb.save(ruta_salida)
    print(f"\n📊 Reporte guardado en: {ruta_salida}")


def main():
    parser = argparse.ArgumentParser(description="Emisión automática por lote en Phoenix v4 (Playwright)")
    parser.add_argument("excel", nargs="?", default="plantilla_solicitudes.xlsx", help="Ruta al Excel de solicitudes")
    parser.add_argument("--browser", default=None, help="Navegador: edge, chrome, chromium o firefox")
    parser.add_argument("--dry-run", action="store_true", help="Llena los formularios sin emitir (presiona Cancelar)")
    parser.add_argument("--no-email", action="store_true", help="No enviar ni generar correo al asegurado")
    parser.add_argument("--email-modo", default=None, help="Modo de correo: eml, smtp, outlook_send, outlook_draft")
    parser.add_argument("--crear-plantilla", action="store_true", help="Solo crea plantilla_solicitudes.xlsx")
    args = parser.parse_args()

    base_dir = os.path.dirname(__file__)
    ruta_excel = args.excel if os.path.isabs(args.excel) else os.path.join(base_dir, args.excel)

    if args.crear_plantilla or not os.path.exists(ruta_excel):
        crear_plantilla_excel(ruta_excel)
        if args.crear_plantilla:
            return
        print("Edita la plantilla con tus solicitudes reales y vuelve a ejecutar emitir_lote.py")
        return

    solicitudes = cargar_solicitudes(ruta_excel)
    if not solicitudes:
        print("❌ No se encontraron solicitudes válidas en el archivo Excel.")
        return

    print(f"📋 Se cargaron {len(solicitudes)} solicitud(es) desde {ruta_excel}")

    phoenix = PhoenixEmisionGrupal(browser=args.browser)
    if not phoenix.conectar_navegador():
        return

    resultados = []
    try:
        phoenix.navegar_a_phoenix()

        for sol in solicitudes:
            datos: DatosEmision = sol["datos"]
            if sol["modo_grupal"] or len(datos.pasajeros) > 5:
                idx_tit = next((i + 1 for i, p in enumerate(datos.pasajeros) if p.es_titular), 1)
                res = phoenix.emitir_voucher_grupal(
                    datos,
                    formato_fichero=sol["formato_fichero"],
                    indice_titular=idx_tit,
                    dry_run=args.dry_run,
                    enviar_correo_asegurado=not args.no_email,
                    email_modo=args.email_modo,
                )
            else:
                res = phoenix.emitir_voucher(
                    datos,
                    dry_run=args.dry_run,
                    enviar_correo_asegurado=not args.no_email,
                    email_modo=args.email_modo,
                )

            res["solicitud_id"] = sol["solicitud_id"]
            resultados.append(res)

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        exportar_resultados(resultados, os.path.join(base_dir, f"resultados_emision_{stamp}.xlsx"))
    finally:
        phoenix.cerrar()


if __name__ == "__main__":
    main()
