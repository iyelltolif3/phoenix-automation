"""Pruebas del módulo MVP de autoemisión (sin abrir Phoenix). Ejecutar: python -m unittest discover -s tests"""
import os
import shutil
import tempfile
import unittest
from datetime import datetime
from email.message import EmailMessage

import _stubs  # noqa: F401  (debe ir antes de importar los módulos del proyecto)

import mvp_autoemision as m
import utils as u
from emision_grupal import PhoenixEmisionGrupal, validar_excel_grupal
from phoenix_emision import Pasajero

REF = datetime(2026, 11, 26)
TABLA_PAX = [[
    ["Nombre", "RUT", "Celular", "Correo", "Fecha nacimiento"],
    ["Juan Perez Soto", "11.111.111-1", "+56 9 8765 4321", "juan@x.cl", "01/01/1990"],
    ["Ana Diaz", "22.222.222-2", "987654321", "ana@x.cl", "15/03/1985"],
]]


def parsear(texto="", asunto="", remitente="", tablas=None, excels=None):
    return m.parsear_solicitud_correo(texto, asunto, remitente, REF, tablas, excels)


class TestFechas(unittest.TestCase):
    def verificar(self, texto, esperado):
        self.assertEqual(m._parsear_rango_fechas(texto, REF), esperado, texto)

    def test_formatos_numericos(self):
        ok = ("27/11/2026", "29/11/2026")
        self.verificar("Vigencia 27/11/2026 al 29/11/2026", ok)
        self.verificar("del 27-11-2026 al 29-11-2026", ok)
        self.verificar("27/11/26 - 29/11/26", ok)
        self.verificar("27/11 al 29/11", ok)

    def test_formatos_en_texto(self):
        ok = ("27/11/2026", "29/11/2026")
        self.verificar("del 27 de noviembre de 2026 al 29 de noviembre de 2026", ok)
        self.verificar("viaje del 27 al 29 de noviembre", ok)
        self.verificar("desde el 27 de noviembre hasta el 29 de noviembre", ok)
        self.verificar("27 noviembre - 29 noviembre", ok)
        self.verificar("10 al 15 de Noviembre", ("10/11/2026", "15/11/2026"))

    def test_cruce_de_mes_y_de_anio(self):
        self.verificar("30 noviembre al 4 diciembre", ("30/11/2026", "04/12/2026"))
        self.verificar("28 diciembre al 3 enero", ("28/12/2026", "03/01/2027"))

    def test_mes_ya_pasado_se_asume_proximo_anio(self):
        self.verificar("15 sept al 20 sept", ("15/09/2027", "20/09/2027"))

    def test_sin_fechas_o_invalidas(self):
        self.verificar("hola, necesito una emisión", ("", ""))
        self.verificar("31/02/2026 al 05/03/2026", ("", ""))


class TestDestino(unittest.TestCase):
    def test_no_confunde_palabras_con_destinos(self):
        for texto in ("Varios colegios", "Rio Bueno", "Causa social", "Pucón", "Santiago"):
            self.assertFalse(u.es_destino_internacional(texto), texto)

    def test_region_phoenix(self):
        casos = {
            "Bariloche, Argentina": "LATAM", "Punta Cana": "LATAM", "Miami": "ESTADOS UNIDOS",
            "Europa 2026": "EUROPA", "Todo el mundo": "TODO EL MUNDO", "Pucón": "",
        }
        for texto, region in casos.items():
            self.assertEqual(u.inferir_region_destino(texto), region, texto)

    def test_solicitud_internacional_trae_region(self):
        d = parsear("EMITIR BAJO CLIENTE: WX01526\nDestino: Bariloche\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"]
        self.assertEqual((d.tipo_viaje, d.destino, d.lista_destino), ("internacional", "LATAM", "BARILOCHE"))

    def test_solicitud_nacional(self):
        d = parsear("Código Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"]
        self.assertEqual((d.tipo_viaje, d.destino), ("nacional", "CL - NACIONAL"))
        self.assertEqual(d.file1, "SANTIAGO-PUCON 27-29 NOV.26")

    def test_asunto_con_varios_no_vuelve_internacional(self):
        d = parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", asunto="Varios pasajeros", tablas=TABLA_PAX)["datos"]
        self.assertEqual(d.tipo_viaje, "nacional")


class TestCliente(unittest.TestCase):
    def test_codigo_con_etiquetas_y_tildes(self):
        for etiqueta in ("EMITIR BAJO CLIENTE", "Código Cliente", "Codigo de cliente", "Cód. cliente", "Cliente"):
            d = parsear(f"{etiqueta}: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"]
            self.assertEqual(d.codigo_cliente, "WX01526", etiqueta)

    def test_codigo_con_nombre_pegado(self):
        d = parsear("Cliente: WX01526 TURISMO RAYS\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"]
        self.assertEqual(d.codigo_cliente, "WX01526")

    def test_texto_libre_sin_etiquetas(self):
        d = parsear(
            "Buenos días, necesito emitir bajo cliente WX01526 para viaje a Bariloche del 27 al 29 de noviembre",
            tablas=TABLA_PAX,
        )["datos"]
        self.assertEqual((d.codigo_cliente, d.fecha_inicio, d.fecha_fin, d.destino), ("WX01526", "27/11/2026", "29/11/2026", "LATAM"))

    def test_nombre_de_colegio_no_se_usa_como_codigo(self):
        r = parsear("Cliente: Colegio San Pedro\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)
        self.assertEqual(r["datos"].codigo_cliente, "")
        self.assertFalse(m.auditar_solicitud_mvp(r)["ok"])

    def test_numero_de_vuelo_no_es_cliente(self):
        d = parsear("Vuelo LA800 de Santiago a Bariloche\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"]
        self.assertEqual(d.codigo_cliente, "")

    def test_tabla_clave_valor(self):
        tablas = [[["Emitir bajo cliente", "WX01526"], ["Producto", "3392"], ["Destino", "Bariloche"],
                   ["Fechas", "27/11/2026 al 29/11/2026"]]] + TABLA_PAX
        d = parsear(tablas=tablas)["datos"]
        self.assertEqual((d.codigo_cliente, d.producto, d.fecha_inicio, d.fecha_fin, len(d.pasajeros)),
                         ("WX01526", "3392", "27/11/2026", "29/11/2026", 2))

    def test_fechas_en_campos_desde_hasta(self):
        tablas = [[["Cliente", "WX01526"], ["Destino", "Pucon"], ["Desde", "27/11/2026"], ["Hasta", "29/11/2026"]]] + TABLA_PAX
        d = parsear(tablas=tablas)["datos"]
        self.assertEqual((d.fecha_inicio, d.fecha_fin), ("27/11/2026", "29/11/2026"))

    def test_dealer_por_correo(self):
        d = parsear("Destino: Pucon\n27/11/2026 al 29/11/2026", remitente="pjara@esquerre.cl", tablas=TABLA_PAX)["datos"]
        self.assertEqual((d.codigo_cliente, d.nombre_agencia), ("SV24000000", "Esquerre Udec"))


class TestPasajerosYAuditoria(unittest.TestCase):
    def test_pasajeros_desde_tabla(self):
        ps = parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)["datos"].pasajeros
        self.assertEqual(len(ps), 2)
        self.assertTrue(ps[0].es_titular)
        self.assertEqual((ps[0].numero_documento, ps[0].digito_verificador, ps[0].telefono), ("11111111", "1", "987654321"))
        self.assertEqual((ps[0].nombre, ps[0].apellido_paterno, ps[0].apellido_materno), ("Juan", "Perez", "Soto"))

    def test_auditoria_ok(self):
        r = parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)
        a = m.auditar_solicitud_mvp(r)
        self.assertTrue(a["ok"], a)
        self.assertEqual(a["errores"], [])

    def test_auditoria_bloquea_lo_que_rompia_el_flujo(self):
        a = m.auditar_solicitud_mvp(parsear("hola"))
        self.assertFalse(a["ok"])
        texto = " ".join(a["errores"]).lower()
        for pista in ("cliente", "vigencia", "pasajeros"):
            self.assertIn(pista, texto)

    def test_auditoria_vigencia_invertida_y_region_faltante(self):
        r = parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=TABLA_PAX)
        r["datos"].fecha_inicio, r["datos"].fecha_fin = "29/11/2026", "27/11/2026"
        self.assertTrue(any("antes de empezar" in e for e in m.auditar_solicitud_mvp(r)["errores"]))
        r["datos"].tipo_viaje, r["datos"].destino = "internacional", ""
        self.assertTrue(any("región de destino" in e for e in m.auditar_solicitud_mvp(r)["errores"]))

    def test_auditoria_rut_invalido_es_advertencia(self):
        tabla = [[["Nombre", "RUT", "Fecha nacimiento"], ["Juan Perez Soto", "11.111.111-9", "01/01/1990"]]]
        a = m.auditar_solicitud_mvp(parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=tabla))
        self.assertTrue(a["ok"])
        self.assertTrue(any("módulo 11" in w for w in a["advertencias"]))

    def test_auditoria_sin_fecha_nacimiento_bloquea(self):
        tabla = [[["Nombre", "RUT"], ["Juan Perez Soto", "11.111.111-1"]]]
        a = m.auditar_solicitud_mvp(parsear("Cliente: WX01526\nDestino: Pucon\n27/11/2026 al 29/11/2026", tablas=tabla))
        self.assertFalse(a["ok"])


class TestEml(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _eml(self, nombre="solicitud.eml", asunto="RV: Emisión grupal", adjuntar_excel=False):
        msg = EmailMessage()
        msg["Subject"], msg["From"], msg["To"] = asunto, "jefe@mapfre.cl", "emision@mawdy.com"
        msg["Date"] = "Thu, 26 Nov 2026 09:00:00 -0300"
        msg.set_content(
            "De: Pamela Jara <pjara@esquerre.cl>\nEnviado: jueves, 26 de noviembre de 2026\n\n"
            "EMITIR BAJO CLIENTE: WX01526\nDestino: Bariloche\nVigencia: del 27 al 29 de noviembre\n"
        )
        if adjuntar_excel:
            ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ficheros", "plantilla_grupal_formato_1.xlsx")
            with open(ruta, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   filename="pasajeros.xlsx")
        ruta_eml = os.path.join(self.dir, nombre)
        with open(ruta_eml, "wb") as f:
            f.write(bytes(msg))
        return ruta_eml

    def test_reenviado_toma_remitente_original(self):
        info = m.extraer_desde_eml(self._eml())
        self.assertEqual(info["from"], "pjara@esquerre.cl")
        self.assertIn("WX01526", info["body"])

    def test_eml_con_excel_adjunto(self):
        info = m.extraer_desde_eml(self._eml(adjuntar_excel=True))
        try:
            self.assertEqual(len(info["adjuntos_excel"]), 1)
            r = m.parsear_solicitud_correo(info["body"], info["subject"], info["from"], info["date"], info["tablas_html"], info["adjuntos_excel"])
            d = r["datos"]
            self.assertEqual((d.codigo_cliente, d.fecha_inicio, d.fecha_fin, d.destino), ("WX01526", "27/11/2026", "29/11/2026", "LATAM"))
            self.assertGreaterEqual(len(d.pasajeros), 1)
            self.assertTrue(m.auditar_solicitud_mvp(r)["ok"], m.auditar_solicitud_mvp(r))
        finally:
            for ruta in info["adjuntos_excel"]:
                os.remove(ruta)

    def test_eml_inexistente(self):
        with self.assertRaises(FileNotFoundError):
            m.extraer_desde_eml(os.path.join(self.dir, "no_existe.eml"))


class TestExcelGrupal(unittest.TestCase):
    def test_ciclo_crear_y_validar_fichero_oficial(self):
        ps = [
            Pasajero(numero_documento="11111111", digito_verificador="1", nombre="Juan", apellido_paterno="Perez",
                     apellido_materno="Soto", fecha_nacimiento="01/01/1990", sexo=5, email="j@x.cl", telefono="987654321", es_titular=True),
            Pasajero(numero_documento="22222222", digito_verificador="2", nombre="Ana", apellido_paterno="Diaz",
                     fecha_nacimiento="15/03/1985", sexo=2, email="a@x.cl", telefono="912345678"),
        ]
        carpeta = tempfile.mkdtemp()
        try:
            ruta = os.path.join(carpeta, "carga.xlsx")
            PhoenixEmisionGrupal.__new__(PhoenixEmisionGrupal).crear_fichero_carga(ps, ruta_salida=ruta)
            r = validar_excel_grupal(ruta, formato="auto")
            self.assertTrue(r["valido"], r)
            self.assertEqual([(p.numero_documento, p.digito_verificador) for p in r["pasajeros"]],
                             [("11111111", "1"), ("22222222", "2")])
        finally:
            shutil.rmtree(carpeta, ignore_errors=True)

    def test_rut_modulo11(self):
        self.assertTrue(u.validar_rut_modulo11("11111111", "1"))
        self.assertFalse(u.validar_rut_modulo11("11111111", "9"))
        self.assertTrue(u.validar_rut_modulo11("12345678", "5"))


if __name__ == "__main__":
    unittest.main()
