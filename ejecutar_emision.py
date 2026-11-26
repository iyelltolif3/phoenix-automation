"""
Phoenix v4 - Script Interactivo de Emisión
===========================================
Interfaz interactiva para realizar emisiones individuales o grupales,
con búsqueda de dealers, reglas de negocio y modo de prueba (dry-run).

USO:
  python ejecutar_emision.py
"""

import json
import os
import subprocess
import sys

import openpyxl

from phoenix_emision import PhoenixAutomation, DatosEmision, Pasajero
from emision_grupal import PhoenixEmisionGrupal, _cargar_workbook_seguro, validar_excel_grupal
from mvp_autoemision import ejecutar_mvp_interactivo
from utils import (
    buscar_dealer_por_correo,
    buscar_dealer_por_nombre,
    es_precompra,
    debe_enviar_precios,
    determinar_producto,
    limpiar_rut,
    formatear_file_nacional,
)


def limpiar_pantalla():
    os.system("cls" if os.name == "nt" else "clear")


def mostrar_banner():
    print(
        """
╔══════════════════════════════════════════════════════╗
║       🛫 PHOENIX v4 - EMISIÓN AUTOMATIZADA 🛫       ║
║          Asistencia en Viajes - Mawdy               ║
╚══════════════════════════════════════════════════════╝
    """
    )


def buscar_dealer_interactivo() -> dict | None:
    """Busca un dealer interactivamente."""
    print("\n📋 BUSCAR DEALER")
    print("  1. Buscar por correo")
    print("  2. Buscar por nombre/agencia")
    print("  3. Ingresar código manualmente")

    opcion = input("\nOpción: ").strip()

    if opcion == "1":
        correo = input("Correo del dealer: ").strip()
        resultados = buscar_dealer_por_correo(correo)
    elif opcion == "2":
        nombre = input("Nombre o agencia: ").strip()
        resultados = buscar_dealer_por_nombre(nombre)
    elif opcion == "3":
        codigo = input("Código de agencia (ej: SV24000000 o API01): ").strip()
        agencia = input("Nombre de agencia (opcional): ").strip()
        return {"codigo": codigo, "agencia": agencia, "correo": "", "nombre": "", "voucher_individual": True}
    else:
        return None

    if not resultados:
        print("❌ No se encontraron resultados.")
        return None

    if len(resultados) == 1:
        d = resultados[0]
        print(f"\n✅ Encontrado: {d['nombre']} | {d['correo']} | {d['codigo']} | {d['agencia']}")
        return d

    print(f"\n📋 Se encontraron {len(resultados)} resultados:")
    for i, d in enumerate(resultados, 1):
        tipo = "PC" if es_precompra(d["codigo"]) else "SV"
        print(f"  {i}. [{tipo}] {d['nombre']} | {d['correo']} | {d['codigo']} | {d['agencia']}")

    seleccion = input(f"\nSelecciona (1-{len(resultados)}): ").strip()
    try:
        idx = int(seleccion) - 1
        return resultados[idx]
    except (ValueError, IndexError):
        print("❌ Selección inválida")
        return None


def ingresar_pasajero() -> Pasajero:
    """Ingresa datos de un pasajero interactivamente."""
    print("\n👤 DATOS DEL PASAJERO")

    tipo_doc = input("  Tipo documento (RUN/PASAPORTE) [RUN]: ").strip().upper() or "RUN"

    if tipo_doc == "RUN":
        rut_completo = input("  RUT (con guión, ej: 12345678-9): ").strip()
        num_doc, dv = limpiar_rut(rut_completo)
    else:
        num_doc = input("  Número de pasaporte: ").strip()
        dv = ""

    nombre = input("  Nombre: ").strip()
    ap_paterno = input("  Apellido paterno: ").strip()
    ap_materno = input("  Apellido materno: ").strip()
    fecha_nac = input("  Fecha nacimiento (DD/MM/AAAA): ").strip()
    email = input("  Email [vacío = usar correo solicitante]: ").strip()
    telefono = input("  Teléfono [vacío = 900000000]: ").strip() or "900000000"

    return Pasajero(
        tipo_documento=tipo_doc,
        numero_documento=num_doc,
        digito_verificador=dv,
        nombre=nombre,
        apellido_paterno=ap_paterno,
        apellido_materno=ap_materno,
        fecha_nacimiento=fecha_nac,
        email=email,
        telefono=telefono,
    )


def cargar_pasajeros_desde_excel(ruta_excel: str) -> list[Pasajero]:
    """Lee una lista de pasajeros desde un Excel oficial Phoenix o simple con encabezados."""
    val = validar_excel_grupal(ruta_excel, formato="auto")
    if val["valido"] and val["pasajeros"]:
        return val["pasajeros"]

    wb = _cargar_workbook_seguro(ruta_excel)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return []

    headers = [str(h or "").strip().lower() for h in rows[0]]
    pasajeros = []
    for row in rows[1:]:
        if not any(row):
            continue
        data = dict(zip(headers, row))
        tipo_doc = str(data.get("tipo_documento") or data.get("tipo") or "RUN").strip().upper()
        if tipo_doc in ("4", "RUN"):
            tipo_doc = "RUN"
        elif tipo_doc in ("2", "PASAPORTE", "PS"):
            tipo_doc = "PASAPORTE"

        num_raw = str(data.get("numero_documento") or data.get("documento") or data.get("rut") or "").strip()
        dv = str(data.get("dv") or "").strip()
        if tipo_doc == "RUN" and "-" in num_raw and not dv:
            num_raw, dv = limpiar_rut(num_raw)

        fnac_val = data.get("fecha_nacimiento") or ""
        if hasattr(fnac_val, "strftime"):
            fnac = fnac_val.strftime("%d/%m/%Y")
        else:
            fnac = str(fnac_val).strip().replace("-", "/")

        pasajeros.append(
            Pasajero(
                tipo_documento=tipo_doc,
                numero_documento=num_raw,
                digito_verificador=dv,
                nombre=str(data.get("nombre") or data.get("nombres") or "").strip(),
                apellido_paterno=str(data.get("apellido_paterno") or data.get("apellidos") or "").strip(),
                apellido_materno=str(data.get("apellido_materno") or "").strip(),
                fecha_nacimiento=fnac,
                email=str(data.get("email") or "").strip(),
                telefono=str(data.get("telefono") or data.get("movil") or "900000000").strip(),
            )
        )
    return pasajeros


def recopilar_datos_emision(grupal: bool = False) -> tuple[DatosEmision | None, bool, int, int]:
    """Recopila datos comunes para emisión individual o grupal."""
    dealer = buscar_dealer_interactivo()
    if not dealer:
        return None, False, 1, 1

    tipo = "PRECOMPRA" if es_precompra(dealer["codigo"]) else "AGENCIA (SV)"
    enviar_precios = debe_enviar_precios(dealer["codigo"], dealer["agencia"])
    print(f"\n  📌 Tipo: {tipo}")
    print(f"  📌 Enviar precios: {'Sí' if enviar_precios else 'No'}")

    print("\n📅 DATOS DE VIGENCIA")
    fecha_inicio = input("  Fecha inicio (DD/MM/AAAA): ").strip()
    fecha_fin = input("  Fecha fin (DD/MM/AAAA): ").strip()
    tipo_viaje = input("  Tipo viaje (nacional/internacional) [nacional]: ").strip().lower() or "nacional"
    zona = input("  Zona (ej: Nacional / Internacional / ALL) [vacío = auto]: ").strip()

    producto_especial = determinar_producto(dealer["agencia"], tipo_viaje)
    if producto_especial:
        print(f"  📌 Producto por regla de negocio: {producto_especial}")
    producto_custom = input(f"  Producto [{producto_especial or 'auto'}]: ").strip()
    producto_final = producto_custom or (producto_especial or "")

    print("\n🌍 DATOS DE ASISTENCIA")
    pais_origen = input("  País origen [CHILE]: ").strip() or "CHILE"
    if tipo_viaje == "internacional":
        destino = input("  Destino (ej: EUROPA / SCHENGEN / TODO EL MUNDO): ").strip()
        lista_destino = input("  País(es) en Lista de Destino: ").strip()
        ciudad_origen = ""
    else:
        destino = input("  Destino en combo (ej: CL - NACIONAL) [NACIONAL]: ").strip() or "NACIONAL"
        ciudad_origen = input("  Ciudad origen (para File1) [SANTIAGO]: ").strip() or "SANTIAGO"
        lista_destino = input("  Ciudad destino (para Lista de Destino): ").strip()

    pasajeros = []
    formato_fichero = 1
    indice_titular = 1

    if grupal:
        ruta_xlsx = input("\n📄 Ruta del Excel de pasajeros (o ENTER para ingresar manual): ").strip().strip('"')
        if ruta_xlsx and os.path.exists(ruta_xlsx):
            pasajeros = cargar_pasajeros_desde_excel(ruta_xlsx)
            print(f"  ✅ {len(pasajeros)} pasajeros cargados desde Excel")
        else:
            while True:
                pasajeros.append(ingresar_pasajero())
                if input("\n¿Agregar otro pasajero? (s/n) [n]: ").strip().lower() != "s":
                    break
        fmt_in = input("  Formato de fichero Phoenix (1 o 2) [1]: ").strip() or "1"
        formato_fichero = int(fmt_in) if fmt_in in ("1", "2") else 1
        tit_in = input(f"  Número de pasajero titular/guía (1-{max(len(pasajeros), 1)}) [1]: ").strip() or "1"
        indice_titular = max(1, int(tit_in))
        if pasajeros and indice_titular <= len(pasajeros):
            pasajeros[indice_titular - 1].es_titular = True
    else:
        while True:
            p = ingresar_pasajero()
            if len(pasajeros) == 0:
                p.es_titular = True
            pasajeros.append(p)
            if input("\n¿Agregar otro pasajero? (s/n) [n]: ").strip().lower() != "s":
                break

    print("\n📁 FILE / REFERENCIA")
    file1 = input("  File/Booking (vacío = autogenerar si nacional): ").strip()
    if not file1 and tipo_viaje == "nacional" and lista_destino:
        try:
            d_ini = int(fecha_inicio[:2])
            d_fin = int(fecha_fin[:2])
            mes_fin = int(fecha_fin[3:5])
            anio_fin = int(fecha_fin[6:])
            file1 = formatear_file_nacional(ciudad_origen, lista_destino, d_ini, d_fin, mes_fin, anio_fin)
            print(f"  File generado: {file1}")
        except Exception:
            file1 = ""

    dry_run = input("\n🧪 ¿Modo prueba DRY-RUN (llena todo pero NO emite)? (s/n) [n]: ").strip().lower() == "s"

    datos = DatosEmision(
        codigo_cliente=dealer["codigo"],
        nombre_agencia=dealer["agencia"],
        cliente_final=dealer.get("agencia", ""),
        producto=producto_final,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
        zona=zona,
        pais_origen=pais_origen,
        destino=destino,
        lista_destino=lista_destino,
        tipo_viaje=tipo_viaje,
        file1=file1,
        pasajeros=pasajeros,
        correo_solicitante=dealer.get("correo", ""),
    )
    return datos, dry_run, formato_fichero, indice_titular


def mostrar_resultado(resultado: dict):
    print("\n" + "=" * 50)
    if resultado.get("exito"):
        print("   ✅ PROCESO EXITOSO" + (" [DRY-RUN]" if resultado.get("dry_run") else ""))
        print("=" * 50)
        if resultado.get("numero_voucher"):
            print(f"  🎫 Voucher N°: {resultado['numero_voucher']}")
        if resultado.get("archivo_voucher"):
            print(f"  📄 PDF: {resultado['archivo_voucher']}")
        if resultado.get("correo_asegurado"):
            c = resultado["correo_asegurado"]
            if c.get("enviado"):
                print(f"  ✉️  Correo asegurado ({c.get('modo')}): {c.get('destinatarios')} {c.get('archivo_eml') or ''}")
            else:
                print(f"  ⚠️  Correo asegurado no enviado: {c.get('error')}")
        if resultado.get("precios"):
            print(f"  💰 Precios: {resultado['precios']}")
        if resultado.get("resumen"):
            print(f"\n  📧 Resumen para email:\n{resultado['resumen']}")
    else:
        print("   ❌ PROCESO FALLIDO")
        print("=" * 50)
        print(f"  Error: {resultado.get('error')}")


def main():
    while True:
        limpiar_pantalla()
        mostrar_banner()

        print("─── SECCIÓN MVP (AUTOEMISIÓN EN PANTALLA) ──────────────")
        print("  1. [MVP] Autoemisión visual desde correo (.eml / texto + Excel)")
        print("")
        print("─── EMISIÓN ESTÁNDAR / LOTE ────────────────────────────")
        print("  2. Emisión individual (asistida)")
        print("  3. Emisión grupal (con fichero Excel)")
        print("  4. Emisión por lote desde Excel (emitir_lote.py)")
        print("  5. Solo buscar dealer")
        print("  6. Test de utilidades")
        print("  0. Salir")

        opcion = input("\nSelecciona: ").strip()

        if opcion == "0":
            return
        if opcion == "1":
            ejecutar_mvp_interactivo()
            input("\nPresiona ENTER para volver al menú...")
            continue
        if opcion == "5":
            buscar_dealer_interactivo()
            input("\nPresiona ENTER para volver al menú...")
            continue
        if opcion == "6":
            subprocess.run([sys.executable, "utils.py"], cwd=os.path.dirname(__file__))
            input("\nPresiona ENTER para volver al menú...")
            continue
        if opcion == "4":
            ruta = input("Ruta al Excel de solicitudes [plantilla_solicitudes.xlsx]: ").strip().strip('"') or "plantilla_solicitudes.xlsx"
            subprocess.run([sys.executable, "emitir_lote.py", ruta], cwd=os.path.dirname(__file__))
            input("\nPresiona ENTER para volver al menú...")
            continue

        if opcion in ("2", "3"):
            grupal = opcion == "3"
            datos, dry_run, formato_fichero, indice_titular = recopilar_datos_emision(grupal=grupal)
            if not datos:
                input("\nPresiona ENTER para volver al menú...")
                continue

            if input("\n¿Iniciar navegador y ejecutar? (s/n) [s]: ").strip().lower() == "n":
                print("❌ Cancelado")
                continue

            browser = input("Navegador a usar (edge/chrome/chromium/firefox) [edge]: ").strip().lower() or "edge"
            email_modo = input("Envío de correo al asegurado (eml/smtp/outlook_send/outlook_draft/no) [eml]: ").strip().lower() or "eml"
            enviar_email = email_modo != "no"
            phoenix = (
                PhoenixEmisionGrupal(browser=browser, headless=False, slow_mo=150, visual_hud=True)
                if grupal
                else PhoenixAutomation(browser=browser, headless=False, slow_mo=150, visual_hud=True)
            )

            if not phoenix.conectar_navegador():
                input("\nPresiona ENTER para volver al menú...")
                continue

            try:
                phoenix.navegar_a_phoenix()
                if grupal:
                    resultado = phoenix.emitir_voucher_grupal(
                        datos,
                        formato_fichero=formato_fichero,
                        indice_titular=indice_titular,
                        dry_run=dry_run,
                        enviar_correo_asegurado=enviar_email,
                        email_modo=email_modo if enviar_email else None,
                    )
                else:
                    resultado = phoenix.emitir_voucher(
                        datos,
                        dry_run=dry_run,
                        enviar_correo_asegurado=enviar_email,
                        email_modo=email_modo if enviar_email else None,
                    )
                mostrar_resultado(resultado)
            finally:
                input("\nPresiona ENTER para cerrar el navegador de automatización...")
                phoenix.cerrar()


if __name__ == "__main__":
    main()
