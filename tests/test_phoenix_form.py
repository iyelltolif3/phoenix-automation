"""Pruebas de la lógica de formulario de phoenix_emision.py con un formulario simulado (sin navegador)."""
import unittest
from datetime import timedelta

import _stubs  # noqa: F401

import phoenix_emision as pe


class Opt:
    def __init__(self, val, txt):
        self.val, self.txt = val, txt

    def get_attribute(self, _):
        return self.val

    def inner_text(self):
        return self.txt


class Loc:
    def __init__(self, form, sel):
        self.f, self.sel = form, sel

    def _el(self):
        return self.f.els.get(self.sel)

    def count(self):
        return 1 if self._el() is not None else 0

    def is_visible(self):
        return self._el()["visible"]

    def is_enabled(self):
        return True

    def get_attribute(self, _):
        return None

    def input_value(self):
        return self._el()["value"]

    def fill(self, v):
        el = self._el()
        if el.get("maxlength0"):
            return  # Chromium no deja escribir en maxlength=0
        el["value"] = v
        self.f.on_change(self.sel)

    def press(self, _):
        self.f.on_change(self.sel)

    def evaluate(self, js, arg=None):
        if arg is not None:
            self._el()["value"] = arg

    def click(self, timeout=None):
        if not self._el()["visible"]:
            raise pe.PlaywrightTimeoutError("oculto")
        self.f.clicks.append(self.sel)

    def select_option(self, value=None, label=None):
        self._el()["value"] = value
        self.f.selected.append((self.sel, value))

    def locator(self, sel):
        assert sel == "option"
        return OptLoc(self._el()["options"])


class OptLoc:
    def __init__(self, opts):
        self.opts = opts

    def all(self):
        return [Opt(v, t) for v, t in self.opts]


class Form:
    def __init__(self, calcula_bien=True, dur_visible=True):
        self.ok = calcula_bien
        self.els = {
            "#Wcr_txtFhInicio": {"value": "26/11/2026", "visible": True},
            "#Wcr_txtFhFin": {"value": "", "visible": True},
            "#Wcr_txtDuracionDias": {"value": "", "visible": dur_visible},
        }
        self.clicks, self.selected = [], []

    def add_combo(self, sel, opciones, visible=True):
        self.els[sel] = {"value": "-1", "visible": visible, "options": [("-1", " ")] + opciones}

    def add_input(self, sel, **kw):
        self.els[sel] = {"value": "", "visible": True, **kw}

    def on_change(self, sel):
        ini = pe._parsear_fecha_phoenix(self.els["#Wcr_txtFhInicio"]["value"])
        if sel == "#Wcr_txtDuracionDias" and self.els[sel]["value"].isdigit():
            fin = (ini + timedelta(days=int(self.els[sel]["value"]) - 1)).strftime("%d/%m/%Y")
            self.els["#Wcr_txtFhFin"]["value"] = fin if self.ok else "NaN/NaN/NaN"
        if sel == "#Wcr_txtFhFin":
            fin = pe._parsear_fecha_phoenix(self.els[sel]["value"])
            self.els["#Wcr_txtDuracionDias"]["value"] = str((fin - ini).days + 1) if (self.ok and fin) else ""

    def locator(self, sel):
        return Loc(self, sel)

    @property
    def value(self):
        return {k: v["value"] for k, v in self.els.items()}


class _Page:
    def wait_for_timeout(self, ms):
        pass


def automatizacion(form):
    a = pe.PhoenixAutomation.__new__(pe.PhoenixAutomation)
    a.visual_hud, a.alertas, a.ultima_alerta, a.timeout_ms = False, [], None, 35000
    a.page = _Page()
    a._FORM = form
    a._esperar_sin_loading = lambda *x, **k: None
    a._capturar_screenshot = lambda *x, **k: None
    return a


pe.PhoenixAutomation.frame = property(lambda self: self._FORM)
pe.PhoenixAutomation._esperar_vigencia.__defaults__ = (50,)  # no esperar 6 s en las pruebas de fallo


class TestVigencia(unittest.TestCase):
    def test_duration_calcula_expiry(self):
        f = Form()
        automatizacion(f).completar_vigencia("26/11/2026", "02/12/2026")
        self.assertEqual((f.value["#Wcr_txtFhFin"], f.value["#Wcr_txtDuracionDias"]), ("02/12/2026", "7"))

    def test_respaldo_por_expiry_si_duration_no_esta_visible(self):
        f = Form(dur_visible=False)
        automatizacion(f).completar_vigencia("27/11/2026", "29/11/2026")
        self.assertEqual(f.value["#Wcr_txtFhFin"], "29/11/2026")
        self.assertEqual(f.value["#Wcr_txtDuracionDias"], "3")

    def test_expiry_nan_aborta_con_mensaje_claro(self):
        with self.assertRaisesRegex(RuntimeError, "no calculó una vigencia válida"):
            automatizacion(Form(calcula_bien=False)).completar_vigencia("27/11/2026", "29/11/2026")

    def test_fechas_invalidas_abortan_antes_de_tocar_phoenix(self):
        f = Form()
        a = automatizacion(f)
        with self.assertRaisesRegex(ValueError, "Vigencia inválida"):
            a.completar_vigencia("", "")
        with self.assertRaisesRegex(ValueError, "anterior al inicio"):
            a.completar_vigencia("29/11/2026", "27/11/2026")
        self.assertEqual(f.value["#Wcr_txtFhInicio"], "26/11/2026")  # no se modificó nada

    def test_formato_corto_se_normaliza(self):
        f = Form()
        automatizacion(f).completar_vigencia("27-11-26", "29-11-26")
        self.assertEqual((f.value["#Wcr_txtFhInicio"], f.value["#Wcr_txtFhFin"]), ("27/11/2026", "29/11/2026"))


DESTINOS_3392 = [("1008", "PV02 - PV EUROPA - EUROPA"), ("1009", "PV03 - PV LATAM - LATAM"),
                 ("1010", "PV04 - All world except USA and Canada - TODO EL MUNDO (EXC. EEUU Y CANADÁ)"),
                 ("1007", "PV01 - UNITED STATES AND CANADA - ESTADOS UNIDOS Y CANADA")]


class TestPestanaViaje(unittest.TestCase):
    def form_3392(self, tab_visible=True):
        f = Form()
        f.els["#Wcr_carTabPage_T0"] = {"value": "", "visible": tab_visible}
        f.add_combo("#CMBDESTINOSV", DESTINOS_3392)
        f.add_combo("#CMBTPVIAJESV", [("1", "1 - VACACIONES/HOLIDAYS"), ("2", "2 - NEGOCIO")])
        return f

    def test_destino_por_region(self):
        f = self.form_3392()
        automatizacion(f).completar_datos_asistencia("CHILE", "LATAM", "BARILOCHE")
        self.assertEqual(f.value["#CMBDESTINOSV"], "1009")
        self.assertEqual(f.value["#CMBTPVIAJESV"], "1")

    def test_no_adivina_destino_si_hay_varias_opciones(self):
        f = self.form_3392()
        with self.assertRaisesRegex(RuntimeError, "CMBDESTINOSV.*CL - NACIONAL.*Opciones"):
            automatizacion(f).completar_datos_asistencia("CHILE", "CL - NACIONAL")
        self.assertEqual(f.value["#CMBDESTINOSV"], "-1")  # no eligió Europa en silencio

    def test_sin_destino_tampoco_adivina(self):
        with self.assertRaisesRegex(RuntimeError, "no se indicó un valor"):
            automatizacion(self.form_3392()).completar_datos_asistencia("CHILE", "")

    def test_opcion_unica_se_usa(self):
        f = Form()
        f.els["#Wcr_carTabPage_T0"] = {"value": "", "visible": True}
        f.add_combo("#CMBDESTINOSV", [("5", "CL - NACIONAL")])
        automatizacion(f).completar_datos_asistencia("CHILE", "CL - NACIONAL")
        self.assertEqual(f.value["#CMBDESTINOSV"], "5")

    def test_tab_oculta_se_activa_por_dom_sin_esperar_35s(self):
        f = self.form_3392(tab_visible=False)
        automatizacion(f).completar_datos_asistencia("CHILE", "LATAM")
        self.assertEqual(f.value["#CMBDESTINOSV"], "1009")

    def test_controles_ausentes_se_omiten_pero_sin_ninguno_falla(self):
        f = self.form_3392()
        self.assertNotIn("#txtPaisOrigen", f.els)
        automatizacion(f).completar_datos_asistencia("CHILE", "LATAM", "BARILOCHE")  # no falla por txtPaisOrigen/txtListaDestinos
        vacio = Form()
        vacio.els["#Wcr_carTabPage_T0"] = {"value": "", "visible": True}
        with self.assertRaisesRegex(RuntimeError, "no muestra controles de viaje"):
            automatizacion(vacio).completar_datos_asistencia("CHILE", "LATAM")

    def test_tab_inexistente_falla_con_mensaje(self):
        with self.assertRaisesRegex(RuntimeError, "no existe"):
            automatizacion(Form()).completar_datos_asistencia("CHILE", "LATAM")


class TestPasajero(unittest.TestCase):
    def form_pasajero(self, con_dv=False, edad_maxlength0=True):
        f = Form()
        f.els["#Wcr_carTabPage_T1"] = {"value": "", "visible": True}
        f.add_combo("#cmbTipoCIF_1", [("2", "PASAPORTE (2)"), ("1007", "RUN (PERSONA NATURAL) (8)")])
        for sel in ("#txtIdFiscal_1", "#txtNmAsegurado_1", "#txtApeAsegurado_1", "#txtFhNacimiento_1",
                    "#txtDDD_movil_1", "#txtMovil_1", "#txtEmail_1"):
            f.add_input(sel)
        f.add_input("#TXTEDADSV_1", maxlength0=edad_maxlength0)
        if con_dv:
            f.add_input("#txtDV_1")
        return f

    PAX = pe.Pasajero(numero_documento="11111111", digito_verificador="k", nombre="Juan", apellido_paterno="Perez",
                      apellido_materno="Soto", fecha_nacimiento="01/01/1990", sexo=5, email="j@x.cl", telefono="987654321")

    def test_run_sin_campo_dv_lleva_dv_en_el_numero(self):
        f = self.form_pasajero()
        automatizacion(f).agregar_pasajero(self.PAX, 1)
        self.assertEqual(f.value["#txtIdFiscal_1"], "11111111-K")
        self.assertEqual(f.value["#cmbTipoCIF_1"], "1007")

    def test_run_con_campo_dv_separado(self):
        f = self.form_pasajero(con_dv=True)
        automatizacion(f).agregar_pasajero(self.PAX, 1)
        self.assertEqual((f.value["#txtIdFiscal_1"], f.value["#txtDV_1"]), ("11111111", "K"))

    def test_edad_se_escribe_aunque_el_input_tenga_maxlength_0(self):
        f = self.form_pasajero(edad_maxlength0=True)
        automatizacion(f).agregar_pasajero(self.PAX, 1)
        self.assertTrue(f.value["#TXTEDADSV_1"].isdigit() and int(f.value["#TXTEDADSV_1"]) >= 35)

    def test_datos_basicos_en_mayusculas(self):
        f = self.form_pasajero()
        automatizacion(f).agregar_pasajero(self.PAX, 1)
        self.assertEqual((f.value["#txtNmAsegurado_1"], f.value["#txtApeAsegurado_1"], f.value["#txtEmail_1"]),
                         ("JUAN", "PEREZ SOTO", "J@X.CL"))


class TestXslt(unittest.TestCase):
    def test_alerta_xslt_da_mensaje_explicito(self):
        f = Form()
        f.els["#Wcr_btnAceptar"] = {"value": "", "visible": True}
        f.els["#Wcr_PW-1"] = {"value": "", "visible": True}
        a = automatizacion(f)
        original = Loc.click
        Loc.click = lambda self, timeout=None: a.alertas.append("XSLTProcessor is not defined")
        try:
            ok, msg = a.terminar_emision()
        finally:
            Loc.click = original
        self.assertFalse(ok)
        self.assertIn("XSLT", msg)
        self.assertIn("158", msg)


if __name__ == "__main__":
    unittest.main()
