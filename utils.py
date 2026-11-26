"""
Utilidades para la automatización de emisiones en Phoenix v4.
Funciones de búsqueda de dealers, formateo de datos y validación.
"""

import json
import os
import re
from datetime import datetime

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def resolver_ruta(ruta: str | None, default_relativa: str = "") -> str:
    """Resuelve rutas relativas al proyecto y expande '~' para portabilidad entre equipos."""
    objetivo = (ruta or default_relativa or "").strip()
    if not objetivo:
        return BASE_DIR
    expandida = os.path.expanduser(objetivo)
    if os.path.isabs(expandida):
        return os.path.normpath(expandida)
    return os.path.normpath(os.path.join(BASE_DIR, expandida))


def cargar_config():
    """Carga la configuración desde config.json."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def buscar_dealer_por_correo(correo: str) -> list[dict]:
    """
    Busca un dealer por su correo electrónico.
    Retorna una lista de coincidencias (puede tener múltiples códigos de agencia).
    """
    config = cargar_config()
    correo_lower = correo.strip().lower()
    return [
        d for d in config["dealers"]
        if d["correo"].lower() == correo_lower
    ]


def buscar_dealer_por_nombre(nombre: str) -> list[dict]:
    """Busca dealers cuyo nombre contenga el texto dado."""
    config = cargar_config()
    nombre_lower = nombre.strip().lower()
    return [
        d for d in config["dealers"]
        if nombre_lower in d["nombre"].lower() or nombre_lower in d["agencia"].lower()
    ]


def buscar_dealer_por_codigo(codigo: str) -> dict | None:
    """Busca un dealer por su código de agencia exacto."""
    config = cargar_config()
    for d in config["dealers"]:
        if d["codigo"].upper() == codigo.strip().upper():
            return d
    return None


def es_precompra(codigo_agencia: str) -> bool:
    """Determina si el código de agencia es Precompra (PC) o Agencia (SV)."""
    return codigo_agencia.strip().upper().startswith("PC")


def debe_enviar_precios(codigo_agencia: str, nombre_agencia: str = "") -> bool:
    """
    Determina si se deben enviar precios al ejecutivo.
    - SV: Siempre se envían precios.
    - PC: No se envían precios, EXCEPTO para TRAVELAWAY.
    """
    if not es_precompra(codigo_agencia):
        return True  # SV siempre envía precios
    # PC: solo Travelaway envía precios
    return "TRAVELAWAY" in nombre_agencia.upper()


def determinar_producto(nombre_agencia: str, tipo_viaje: str = "nacional") -> str | None:
    """
    Determina el producto basado en reglas de negocio.
    Retorna el nombre del producto si hay una regla especial, None si es estándar.
    """
    config = cargar_config()
    nombre_upper = nombre_agencia.upper()
    
    # Regla: MELI NEWEN, VALPOVIÑA, PATAGONIA PUNTO NORTE → 1966 INTRARREGIONAL
    agencias_intrarregional = config["reglas_negocio"]["producto_intrarregional"]["agencias"]
    for agencia in agencias_intrarregional:
        if agencia.upper() in nombre_upper:
            return config["reglas_negocio"]["producto_intrarregional"]["producto"]
    
    # Regla: MASAI TRAVEL
    if "MASAI" in nombre_upper:
        masai = config["reglas_negocio"]["masai_travel"]
        return masai["internacional"] if tipo_viaje.lower() == "internacional" else masai["nacional"]
    
    return None  # Producto estándar, seleccionado por Phoenix


def tipo_documento_codigo(tipo: str) -> int:
    """Convierte tipo de documento a código numérico: RUN=4, PASAPORTE=2."""
    config = cargar_config()
    return config["tipo_documento"].get(tipo.upper(), 4)


def formatear_file_nacional(origen: str, destino: str, dia_inicio: int, dia_fin: int, mes: str, anio: int) -> str:
    """
    Formatea el File1 para viajes nacionales (intrarregionales).
    Formato: ORIGEN-DESTINO DÍA_INICIO-DÍA_FIN MES.AÑO
    Ejemplo: SANTIAGO-PICHILEMU 15-29 DIC.24
    """
    meses = {
        1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
    }
    if isinstance(mes, int):
        mes_abrev = meses.get(mes, str(mes))
    elif mes.isdigit():
        mes_abrev = meses.get(int(mes), mes)
    else:
        mes_abrev = mes.upper()[:3]
    anio_corto = str(anio)[-2:]
    return f"{origen.upper()}-{destino.upper()} {dia_inicio}-{dia_fin} {mes_abrev}.{anio_corto}"


def formatear_nombre_voucher(numero_voucher: str, file_ref: str, titular: str) -> str:
    """
    Formatea el nombre del voucher según el estándar del manual.
    Formato: VOUCHER N°XXX - FILE - TIT
    """
    return f"VOUCHER N°{numero_voucher} - {file_ref} - {titular}"


def generar_resumen_emision(tipo_viaje: str, clp: str = "", usd: str = "", tc: str = "", pax: int = 1) -> str:
    """
    Genera el resumen de emisión para enviar al ejecutivo.
    """
    if tipo_viaje.lower() == "nacional":
        if usd:
            return (
                f"Emisión realizada según lo solicitado. Adjunto voucher.\n\n"
                f"Resumen de Emisión:\n"
                f"· USD: ${usd} + IVA\n"
                f"· CLP: ${clp} + IVA\n"
                f"· TC: ${tc}\n"
                f"· PAX: {pax}\n\n"
                f"Saludos"
            )
        else:
            return (
                f"Emisión realizada según lo solicitado. Adjunto voucher.\n\n"
                f"Resumen de Emisión:\n"
                f"· CLP: ${clp} + IVA\n"
                f"· PAX: {pax}\n\n"
                f"Saludos"
            )
    else:  # internacional
        return (
            f"Emisión realizada según lo solicitado. Adjunto voucher.\n\n"
            f"Resumen de Emisión:\n"
            f"· USD: ${usd}\n"
            f"· CLP: ${clp}\n"
            f"· TC: ${tc}\n"
            f"· PAX: {pax}\n\n"
            f"Saludos"
        )


def validar_rut(rut: str) -> bool:
    """Valida longitud básica de un RUT chileno (sin dígito verificador)."""
    rut_limpio = re.sub(r'[^0-9]', '', rut)
    return len(rut_limpio) >= 7 and len(rut_limpio) <= 9


def calcular_dv_rut(numero: str) -> str:
    """Calcula el dígito verificador módulo 11 de un RUT chileno (retorna '0'-'9' o 'K')."""
    digitos = re.sub(r"[^0-9]", "", str(numero or ""))
    if not digitos:
        return ""
    suma = 0
    factor = 2
    for ch in reversed(digitos):
        suma += int(ch) * factor
        factor = 2 if factor == 7 else factor + 1
    resto = 11 - (suma % 11)
    if resto == 11:
        return "0"
    if resto == 10:
        return "K"
    return str(resto)


def validar_rut_modulo11(numero: str, dv: str) -> bool:
    """Verifica que el número (7-9 dígitos) y el DV coincidan según módulo 11."""
    num_limpio = re.sub(r"[^0-9]", "", str(numero or ""))
    dv_limpio = str(dv or "").strip().upper()
    if len(num_limpio) < 7 or len(num_limpio) > 9 or not re.fullmatch(r"[0-9K]", dv_limpio):
        return False
    return calcular_dv_rut(num_limpio) == dv_limpio


def limpiar_rut(rut: str) -> tuple[str, str]:
    """
    Separa un RUT en número y dígito verificador.
    Entrada: '12345678-9' o '12.345.678-9'
    Retorna: ('12345678', '9')
    """
    rut_limpio = rut.strip().replace(".", "").replace(" ", "")
    if "-" in rut_limpio:
        partes = rut_limpio.split("-", 1)
        return partes[0], partes[1].upper()
    # Si no tiene guión, el último carácter es el DV
    return rut_limpio[:-1], rut_limpio[-1].upper()


if __name__ == "__main__":
    # Test de utilidades
    print("=== Test de utilidades ===")
    
    dealers = buscar_dealer_por_correo("jroman@esquerre.cl")
    print(f"\nDealer por correo 'jroman@esquerre.cl': {dealers}")
    
    dealers = buscar_dealer_por_nombre("masai")
    print(f"\nDealers con nombre 'masai': {len(dealers)} encontrados")
    for d in dealers:
        print(f"  - {d['nombre']} | {d['codigo']} | {d['agencia']}")
    
    print(f"\n¿SV24000000 es precompra? {es_precompra('SV24000000')}")
    print(f"¿PC06722 es precompra? {es_precompra('PC06722')}")
    
    print(f"\n¿Enviar precios SV24000000 (Esquerre)? {debe_enviar_precios('SV24000000', 'Esquerre Udec')}")
    print(f"¿Enviar precios PC06722 (Travelaway)? {debe_enviar_precios('PC06722', 'Travelaway')}")
    print(f"¿Enviar precios PC26000000 (Veci)? {debe_enviar_precios('PC26000000', 'Veci')}")
    
    prod = determinar_producto("Melinewen")
    print(f"\nProducto para Melinewen: {prod}")
    
    prod = determinar_producto("Masai travel", "internacional")
    print(f"Producto para Masai travel (internacional): {prod}")
    
    file_name = formatear_file_nacional("Santiago", "Pichilemu", 15, 29, 12, 2026)
    print(f"\nFile nacional: {file_name}")
    
    rut, dv = limpiar_rut("12.345.678-9")
    print(f"\nRUT: {rut}, DV: {dv}")
