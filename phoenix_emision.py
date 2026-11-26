"""
Phoenix v4 - Automatización de Emisión de Viajes (Playwright)
==============================================================
Automatiza la emisión de vouchers de asistencia en viajes en Phoenix v4
usando Playwright con los selectores reales mapeados de la aplicación
ASP.NET / DevExpress (incluyendo soporte específico para _MAWDY [CHILE]).

Incluye:
- Perfil persistente en Edge / Chrome / Firefox (conserva sesión SSO de Microsoft).
- Captura determinista del voucher PDF con `expect_download()`.
- Envío automático (o generación .eml / Outlook) del voucher al asegurado y/o ejecutivo.
"""

import os
import re
import time
import logging
from datetime import datetime
from dataclasses import dataclass, field

from playwright.sync_api import (
    sync_playwright,
    Playwright,
    BrowserContext,
    Page,
    FrameLocator,
    Dialog,
    TimeoutError as PlaywrightTimeoutError,
)

from utils import (
    cargar_config,
    resolver_ruta,
    debe_enviar_precios,
    determinar_producto,
    tipo_documento_codigo,
    formatear_nombre_voucher,
    generar_resumen_emision,
)
from email_sender import enviar_voucher_al_asegurado, enviar_correo

# ─── Configuración de logging ───────────────────────────────────────────────
LOG_PATH = os.path.join(os.path.dirname(__file__), "phoenix_emision.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


# ─── Modelo de datos ────────────────────────────────────────────────────────
@dataclass
class Pasajero:
    """Datos de un pasajero/asegurado."""
    tipo_documento: str = "RUN"        # RUN o PASAPORTE
    numero_documento: str = ""         # Número sin DV
    digito_verificador: str = ""       # DV (vacío si pasaporte)
    nombre: str = ""
    apellido_paterno: str = ""
    apellido_materno: str = ""
    fecha_nacimiento: str = ""         # DD/MM/AAAA
    sexo: int = 2                      # 1 = Femenino, 2 = Masculino (en fichero grupal: 8)
    email: str = ""
    telefono: str = ""
    cod_pais_movil: int = 56
    es_titular: bool = False


@dataclass
class DatosEmision:
    """Datos completos para una emisión de voucher."""
    codigo_cliente: str = ""           # Código agencia (ej: MV00326, SV24000000)
    nombre_agencia: str = ""
    cliente_final: str = ""
    producto: str = ""

    fecha_inicio: str = ""             # DD/MM/AAAA
    fecha_fin: str = ""                # DD/MM/AAAA
    zona: str = ""

    pais_origen: str = "CHILE"
    destino: str = ""                  # Texto o código en CMBDESTINOSV
    lista_destino: str = ""            # País (int) o ciudad (nac) en #txtListaDestinos
    tipo_viaje_sv: str = ""            # Opcional para #CMBTPVIAJESV

    pasajeros: list = field(default_factory=list)
    correo_solicitante: str = ""
    file1: str = ""
    tipo_viaje: str = "nacional"       # nacional / internacional


def _calcular_dias_inclusivos(fecha_inicio: str, fecha_fin: str) -> int | None:
    """Calcula duración en días inclusivos entre DD/MM/AAAA y DD/MM/AAAA."""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            d1 = datetime.strptime(fecha_inicio.strip(), fmt)
            d2 = datetime.strptime(fecha_fin.strip(), fmt)
            return max((d2 - d1).days + 1, 1)
        except ValueError:
            continue
    return None


def _parsear_fecha_phoenix(valor: str) -> datetime | None:
    """Parsea fechas DD/MM/AAAA, DD-MM-AAAA o DD/MM/AA. Devuelve None si no es válida."""
    texto = (valor or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None


def _calcular_edad(fecha_nacimiento: str, fecha_ref: str = "") -> int | None:
    """Calcula edad en años a partir de DD/MM/AAAA."""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            fn = datetime.strptime(fecha_nacimiento.strip(), fmt)
            ref = datetime.strptime(fecha_ref.strip(), fmt) if fecha_ref else datetime.now()
            return ref.year - fn.year - ((ref.month, ref.day) < (fn.month, fn.day))
        except ValueError:
            continue
    return None


def _sanitizar_nombre_archivo(nombre: str) -> str:
    """Elimina caracteres inválidos de nombres de archivo en Windows."""
    return re.sub(r'[<>:"/\\|?*]', "-", nombre).strip()


class PhoenixAutomation:
    """Motor Playwright para automatizar emisiones en Phoenix v4."""

    IFRAME_SELECTOR = "#iframeContenidoMaestro"

    def __init__(
        self,
        browser: str | None = None,
        headless: bool | None = None,
        slow_mo: int | None = None,
        visual_hud: bool = False,
    ):
        self.config = cargar_config()
        self.browser_name = (browser or self.config.get("browser", "edge")).lower()
        self.headless = bool(self.config.get("headless", False) if headless is None else headless)
        self.timeout_ms = max(int(self.config.get("wait_timeout", 35)), 30) * 1000
        self.slow_mo = int(self.config.get("slow_mo_ms", 0) if slow_mo is None else slow_mo)
        self.visual_hud = visual_hud

        self.playwright: Playwright | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None
        self.ultima_alerta: str | None = None
        self.alertas: list[str] = []

    @property
    def frame(self) -> FrameLocator:
        """Retorna el FrameLocator del iframe principal de Phoenix."""
        return self.page.frame_locator(self.IFRAME_SELECTOR)

    def _limpiar_alertas(self):
        self.ultima_alerta = None
        self.alertas.clear()

    def _on_dialog(self, dialog: Dialog):
        """Acepta automáticamente alertas/confirmaciones JS y registra el mensaje."""
        msg = (dialog.message or "").strip()
        self.ultima_alerta = msg
        self.alertas.append(msg)
        logger.info(f"🔔 Alerta JS ({dialog.type}): {msg}")
        try:
            dialog.accept()
        except Exception:
            pass

    def mostrar_hud(self, titulo: str, detalle: str = ""):
        """Muestra un indicador flotante en pantalla durante la autoemisión visual."""
        if not self.visual_hud or not self.page:
            return
        try:
            self.page.evaluate(
                """([titulo, detalle]) => {
                    let box = document.getElementById('phoenix-mvp-hud');
                    if (!box) {
                        box = document.createElement('div');
                        box.id = 'phoenix-mvp-hud';
                        Object.assign(box.style, {
                            position: 'fixed',
                            bottom: '14px',
                            right: '14px',
                            zIndex: '2147483647',
                            background: 'rgba(15, 23, 42, 0.92)',
                            color: '#f8fafc',
                            padding: '10px 14px',
                            borderRadius: '10px',
                            borderLeft: '4px solid #ef4444',
                            fontFamily: 'Segoe UI, sans-serif',
                            fontSize: '12px',
                            maxWidth: '360px',
                            boxShadow: '0 8px 24px rgba(0,0,0,0.35)',
                            pointerEvents: 'none',
                        });
                        document.body.appendChild(box);
                    }
                    box.innerHTML = `<div style="font-weight:700;color:#fca5a5;margin-bottom:2px;">🤖 ${titulo}</div>` +
                        (detalle ? `<div style="color:#e2e8f0;">${detalle}</div>` : '');
                }""",
                [titulo, detalle],
            )
        except Exception:
            pass

    def _esperar_sin_loading(self, timeout_ms: int | None = None):
        """Espera a que desaparezcan los paneles de carga DevExpress dentro del iframe."""
        if not self.page:
            return
        t_ms = timeout_ms or self.timeout_ms
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=min(t_ms, 10000))
        except Exception:
            pass
        loading_sel = "[id$='_LP']:visible, .dxlpLoadingPanel:visible, .dxpcLoadingPanel:visible"
        try:
            lp = self.frame.locator(loading_sel)
            if lp.count() > 0:
                lp.first.wait_for(state="hidden", timeout=t_ms)
        except Exception:
            pass
        self.page.wait_for_timeout(250)

    def _detectar_errores_visibles_modal(self) -> list[str]:
        """Inspecciona mensajes de error o validación visibles dentro de #Wcr_PW-1."""
        errores: list[str] = []
        try:
            err_locs = self.frame.locator(
                "#Wcr_PW-1 .dxeOutOfRWarn, "
                "#Wcr_PW-1 .dxvsValidationSummary li, "
                "#Wcr_PW-1 [id*='lblError'], "
                "#Wcr_PW-1 [id*='MensajeError'], "
                "#Wcr_PW-1 span[style*='Red'], "
                "#Wcr_PW-1 span[style*='red']"
            )
            for i in range(min(err_locs.count(), 10)):
                el = err_locs.nth(i)
                if el.is_visible():
                    txt = el.inner_text().strip()
                    if txt and txt not in errores:
                        errores.append(txt)
        except Exception:
            pass
        return errores

    def conectar_navegador(self, browser: str | None = None) -> bool:
        """
        Inicia Playwright con perfil persistente para conservar la sesión SSO.
        Soporta: 'edge' (canal msedge), 'chrome' (canal chrome), 'chromium' y 'firefox'.
        """
        target = (browser or self.browser_name).lower()
        download_dir = resolver_ruta(self.config.get("download_dir"), "~/Downloads/vouchers")
        profile_base = resolver_ruta(self.config.get("browser_profile_dir"), "browser_profile")
        os.makedirs(download_dir, exist_ok=True)
        profile_dir = os.path.join(profile_base, f"pw_{target}")
        os.makedirs(profile_dir, exist_ok=True)

        logger.info(f"Iniciando Playwright ('{target}') con perfil persistente en: {profile_dir}")

        try:
            self.playwright = sync_playwright().start()

            common_args = {
                "user_data_dir": profile_dir,
                "headless": self.headless,
                "accept_downloads": True,
                "downloads_path": download_dir,
                "viewport": None,
                "slow_mo": self.slow_mo,
            }

            if target == "edge":
                self.context = self.playwright.chromium.launch_persistent_context(
                    channel="msedge", args=["--start-maximized"], **common_args
                )
            elif target == "chrome":
                self.context = self.playwright.chromium.launch_persistent_context(
                    channel="chrome", args=["--start-maximized"], **common_args
                )
            elif target == "chromium":
                self.context = self.playwright.chromium.launch_persistent_context(
                    args=["--start-maximized"], **common_args
                )
            elif target == "firefox":
                self.context = self.playwright.firefox.launch_persistent_context(**common_args)
            else:
                raise ValueError(f"Navegador no soportado: {target}")

            self.context.set_default_timeout(self.timeout_ms)
            self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
            self.page.on("dialog", self._on_dialog)

            logger.info(f"✅ Playwright '{target}' iniciado correctamente")
            return True

        except Exception as e:
            logger.error(f"❌ Error al iniciar Playwright '{target}': {e}")
            self.cerrar()
            return False

    def navegar_a_phoenix(self, url_override: str | None = None):
        """Navega a Phoenix y espera si requiere iniciar sesión en Microsoft SSO."""
        url = url_override or self.config["phoenix_url"]
        logger.info(f"Navegando a {url}")
        self.page.goto(url, wait_until="domcontentloaded")
        self.page.wait_for_timeout(1500)

        if "login.microsoftonline" in self.page.url or "login" in self.page.url.lower():
            logger.warning("⚠️ Se requiere iniciar sesión en Microsoft SSO.")
            input("Inicia sesión en la ventana del navegador y luego presiona ENTER aquí...")
            self.page.wait_for_url("**/default.aspx*", timeout=60000)

        logger.info(f"✅ Sesión activa en: {self.page.url}")

    def ir_a_emision_vouchers(self):
        """
        Navega a Pólizas/Vouchers -> Emisión y Consulta de Vouchers
        (o Certificates -> List of Certificates) y espera #btnNew dentro del iframe.
        """
        logger.info("Navegando a Emisión y Consulta de Vouchers / List of Certificates...")
        self.mostrar_hud("Navegando a Emisión", "Abriendo Emisión y Consulta de Vouchers...")

        for intento in range(2):
            submenu_href = self.page.locator("a[href*='listaContratos.aspx']")
            if submenu_href.count() > 0:
                submenu_href.first.evaluate("el => el.click()")
            else:
                menu = self.page.locator("a:has-text('Pólizas/Vouchers'), a:has-text('Certificates')").first
                menu.click()
                submenu = self.page.locator(
                    "a:has-text('Emisión y Consulta'), a:has-text('List of Certificates'), a[href*='listaContratos.aspx']"
                ).first
                submenu.click()
            try:
                self.frame.locator("#btnNew").wait_for(state="visible", timeout=self.timeout_ms)
                self._esperar_sin_loading()
                logger.info("✅ En listaContratos.aspx (#btnNew visible)")
                return
            except PlaywrightTimeoutError:
                if intento == 0:
                    logger.warning("⚠️ Timeout esperando #btnNew; reintentando navegación...")
                    self.page.wait_for_timeout(1500)
                else:
                    raise

    def seleccionar_sociedad_grilla(self, texto_sociedad: str = "CHILE"):
        """Filtra la grilla de listaContratos.aspx por sociedad (#cmbSociedad_I)."""
        if not texto_sociedad:
            return
        try:
            cmb = self.frame.locator("#cmbSociedad_I")
            if cmb.count() > 0 and cmb.is_visible():
                cmb.click()
                cmb.fill(texto_sociedad)
                self.page.wait_for_timeout(800)
                opt = self.frame.locator(f"#cmbSociedad_DDD_L_LBT td:has-text('{texto_sociedad}')").first
                if opt.count() > 0:
                    opt.click()
                    self.page.wait_for_timeout(1500)
                    logger.info(f"✅ Sociedad filtrada en grilla: {texto_sociedad}")
        except Exception as e:
            logger.warning(f"⚠️ No se pudo filtrar sociedad '{texto_sociedad}': {e}")

    def nuevo_contrato(self):
        """Hace clic en #btnNew y espera que se abra el modal #Wcr_PW-1."""
        logger.info("Abriendo modal de Nuevo contrato (#btnNew)...")
        self.mostrar_hud("Nuevo Contrato", "Abriendo formulario de emisión...")
        for intento in range(2):
            self.frame.locator("#btnNew").click()
            try:
                self.frame.locator("#Wcr_PW-1").wait_for(state="visible", timeout=self.timeout_ms)
                self.frame.locator("#Wcr_txtcodCliente").wait_for(state="visible", timeout=self.timeout_ms)
                self._esperar_sin_loading()
                logger.info("✅ Modal #Wcr_PW-1 abierto")
                return
            except PlaywrightTimeoutError:
                if intento == 0:
                    logger.warning("⚠️ #Wcr_PW-1 no abrió al primer clic; reintentando...")
                    self.page.wait_for_timeout(1000)
                else:
                    raise

    def buscar_cliente(self, codigo_cliente: str):
        """Paso 4: escribe hasta 6 caracteres en #Wcr_txtcodCliente y presiona Tab."""
        busqueda = codigo_cliente.strip()[:6]
        logger.info(f"Buscando cliente en #Wcr_txtcodCliente: '{busqueda}' (completo: {codigo_cliente})")
        self.mostrar_hud("Paso 1: Cliente", f"Buscando código '{busqueda}'...")

        campo = self.frame.locator("#Wcr_txtcodCliente")
        campo.fill(busqueda)
        campo.press("Tab")
        campo.evaluate("el => el.blur()")
        self._esperar_sin_loading()

        # Esperar opciones válidas en #Wcr_cmbClientes
        self.frame.locator("#Wcr_cmbClientes option").first.wait_for(state="attached", timeout=self.timeout_ms)
        self.page.wait_for_timeout(600)
        logger.info("✅ Cliente(s) cargado(s) en #Wcr_cmbClientes")

    def seleccionar_cliente_final(self, codigo_cliente: str = "", nombre_cliente: str = ""):
        """Paso 4b: si hay más de una razón social en #Wcr_cmbClientes, selecciona la coincidente."""
        cmb = self.frame.locator("#Wcr_cmbClientes")
        opciones = cmb.locator("option").all_text_contents()

        if len(opciones) <= 1:
            if opciones:
                logger.info(f"Cliente único auto-seleccionado: {opciones[0]}")
            return True

        objetivos = [x.strip().lower() for x in (codigo_cliente, nombre_cliente) if x and x.strip()]
        for texto in opciones:
            if any(obj in texto.lower() for obj in objetivos):
                cmb.select_option(label=texto)
                self._esperar_sin_loading()
                logger.info(f"✅ Cliente final seleccionado: {texto}")
                return True

        logger.warning(f"⚠️ Sin coincidencia exacta en #Wcr_cmbClientes; opciones: {opciones}")
        return False

    def seleccionar_producto(self, producto: str = ""):
        """Error 1 del manual: seleccionar producto (#Wcr_idRegProducto) ANTES de editar fechas."""
        self.mostrar_hud("Paso 2: Producto", f"Seleccionando '{producto or 'por defecto'}'...")
        cmb = self.frame.locator("#Wcr_idRegProducto")
        cmb.locator("option").first.wait_for(state="attached", timeout=self.timeout_ms)

        opciones = [
            (opt.get_attribute("value") or "", opt.inner_text().strip())
            for opt in cmb.locator("option").all()
        ]

        if producto:
            prod_lower = producto.strip().lower()
            for val, txt in opciones:
                if prod_lower in txt.lower() or prod_lower == val.lower():
                    cmb.select_option(value=val)
                    self._esperar_sin_loading()
                    logger.info(f"✅ Producto seleccionado: {txt}")
                    return
            logger.warning(f"⚠️ No se encontró producto '{producto}' en #Wcr_idRegProducto")
        else:
            actual = cmb.input_value()
            if not actual or actual == "-1":
                for val, txt in opciones:
                    if val and val != "-1" and txt:
                        cmb.select_option(value=val)
                        self._esperar_sin_loading()
                        logger.info(f"✅ Producto por defecto seleccionado: {txt}")
                        return

    def _leer_vigencia(self) -> tuple[str, str, str]:
        """Lee (Effect, Expiry, Duration) tal como los muestra el formulario."""
        def _val(sel: str) -> str:
            loc = self.frame.locator(sel)
            return loc.input_value().strip() if loc.count() > 0 else ""

        return _val("#Wcr_txtFhInicio"), _val("#Wcr_txtFhFin"), _val("#Wcr_txtDuracionDias")

    def _vigencia_coincide(self, esperado_ini: datetime, esperado_fin: datetime) -> bool:
        """True si Effect/Expiry del formulario son fechas reales (sin NaN) y coinciden con lo pedido."""
        ini, fin, dur = self._leer_vigencia()
        if "nan" in f"{ini} {fin} {dur}".lower():
            return False
        if self.frame.locator("#Wcr_txtDuracionDias").count() > 0 and not dur.isdigit():
            return False  # Duration vacía = Phoenix no terminó de calcular la vigencia
        f_ini, f_fin = _parsear_fecha_phoenix(ini), _parsear_fecha_phoenix(fin)
        return bool(f_ini and f_fin and f_ini.date() == esperado_ini.date() and f_fin.date() == esperado_fin.date())

    def _esperar_vigencia(self, esperado_ini: datetime, esperado_fin: datetime, espera_ms: int = 6000) -> bool:
        """Da tiempo al callback DevExpress de recalcular Expiry/Duration antes de dar la vigencia por mala."""
        limite = time.monotonic() + espera_ms / 1000.0
        while time.monotonic() < limite:
            if self._vigencia_coincide(esperado_ini, esperado_fin):
                return True
            self.page.wait_for_timeout(400)
        return self._vigencia_coincide(esperado_ini, esperado_fin)

    def completar_vigencia(self, fecha_inicio: str, fecha_fin: str, zona: str = ""):
        """
        Paso 5: completa Effect (#Wcr_txtFhInicio) y Duration (#Wcr_txtDuracionDias) / Expiry (#Wcr_txtFhFin).

        Valida las fechas ANTES de tocar Phoenix y verifica DESPUES que Expiry sea una fecha real
        (no vacía ni 'NaN') y coincida con la pedida; si no, falla con un mensaje claro en vez de
        seguir con una vigencia inválida (que deja ocultas pestañas como #Wcr_carTabPage_T0).
        """
        d_ini = _parsear_fecha_phoenix(fecha_inicio)
        d_fin = _parsear_fecha_phoenix(fecha_fin)
        if not d_ini or not d_fin:
            raise ValueError(
                f"Vigencia inválida: inicio='{fecha_inicio}' fin='{fecha_fin}' (se espera DD/MM/AAAA)."
            )
        if d_fin < d_ini:
            raise ValueError(f"Vigencia inválida: el fin ({fecha_fin}) es anterior al inicio ({fecha_inicio}).")
        str_ini, str_fin = d_ini.strftime("%d/%m/%Y"), d_fin.strftime("%d/%m/%Y")
        dias = (d_fin - d_ini).days + 1

        logger.info(f"Completando vigencia: {str_ini} -> {str_fin} ({dias} días)")
        self.mostrar_hud("Paso 3: Vigencia", f"{str_ini} al {str_fin}")
        self._limpiar_alertas()

        f_ini = self.frame.locator("#Wcr_txtFhInicio")
        f_ini.fill(str_ini)
        f_ini.press("Tab")
        self._esperar_sin_loading()

        f_fin = self.frame.locator("#Wcr_txtFhFin")
        dur_input = self.frame.locator("#Wcr_txtDuracionDias")
        dur_select = self.frame.locator("#Wcr_txtDuracion")

        def _por_duracion() -> bool:
            """Escribe Duration y Tab: Phoenix calcula Expiry = Effect + Duration - 1."""
            if dur_input.count() > 0 and dur_input.is_visible() and dur_input.is_enabled():
                dur_input.fill(str(dias))
                dur_input.press("Tab")
                dur_input.evaluate("el => el.blur()")
                self._esperar_sin_loading()
                logger.info(f"Duración establecida en {dias} días (#Wcr_txtDuracionDias)")
                return True
            if dur_select.count() > 0 and dur_select.is_visible():
                for opt in dur_select.locator("option").all():
                    if str(dias) in opt.inner_text():
                        dur_select.select_option(value=opt.get_attribute("value"))
                        self._esperar_sin_loading()
                        return True
            return False

        def _por_expiry() -> bool:
            es_readonly = f_fin.get_attribute("readonly") is not None or "inputDisabled" in (f_fin.get_attribute("class") or "")
            if f_fin.count() > 0 and not es_readonly and f_fin.is_enabled():
                f_fin.fill(str_fin)
                f_fin.press("Tab")
                f_fin.evaluate("el => el.blur()")
                self._esperar_sin_loading()
                logger.info(f"Expiry establecido en {str_fin} (#Wcr_txtFhFin)")
                return True
            return False

        # Duration primero (secuencia confirmada a mano en PRE); Expiry como respaldo, y viceversa.
        ok = False
        for metodo in (_por_duracion, _por_expiry):
            if metodo() and self._esperar_vigencia(d_ini, d_fin):
                ok = True
                break

        if self.alertas:
            logger.warning(f"⚠️ Alerta(s) durante vigencia: {' | '.join(self.alertas)}")

        if not ok:
            ini_f, fin_f, dur_f = self._leer_vigencia()
            self._capturar_screenshot("error_vigencia")
            raise RuntimeError(
                f"Phoenix no calculó una vigencia válida: pedido {str_ini}->{str_fin}, "
                f"formulario Effect='{ini_f}' Expiry='{fin_f}' Duration='{dur_f}'."
            )
        logger.info(f"✅ Vigencia validada en Phoenix: {'{}'.format(' | '.join(self._leer_vigencia()))}")

        if zona:
            cmb_zona = self.frame.locator("#cmbZona")
            if cmb_zona.count() > 0 and cmb_zona.is_visible():
                for opt in cmb_zona.locator("option").all():
                    if zona.strip().lower() in opt.inner_text().lower():
                        cmb_zona.select_option(value=opt.get_attribute("value"))
                        self._esperar_sin_loading()
                        logger.info(f"✅ Zona seleccionada: {opt.inner_text().strip()}")
                        break

    def _abrir_tab(self, tab_id: str, descripcion: str = ""):
        """
        Activa una pestaña del contrato (ej. #Wcr_carTabPage_T0) sin esperar 35 s si está oculta:
        intenta el clic normal 5 s y, si la pestaña existe pero está oculta/tapada, la activa por DOM.
        """
        tab = self.frame.locator(f"#{tab_id}")
        if tab.count() == 0:
            raise RuntimeError(f"La pestaña #{tab_id} {descripcion} no existe en el formulario de este producto.")
        try:
            tab.click(timeout=5000)
        except PlaywrightTimeoutError:
            ini, fin, dur = self._leer_vigencia()
            logger.warning(
                f"⚠️ #{tab_id} {descripcion} no es clicable (oculta). Vigencia actual: "
                f"Effect='{ini}' Expiry='{fin}' Duration='{dur}'. Activando por DOM..."
            )
            tab.evaluate("el => el.click()")
        self.page.wait_for_timeout(500)

    def _opciones_combo(self, selector: str) -> list[tuple[str, str]]:
        """Devuelve [(value, texto)] de las opciones reales de un <select>, sin la opción vacía (-1)."""
        cmb = self.frame.locator(selector)
        if cmb.count() == 0:
            return []
        return [
            (opt.get_attribute("value") or "", opt.inner_text().strip())
            for opt in cmb.locator("option").all()
            if (opt.get_attribute("value") or "") not in ("", "-1") and opt.inner_text().strip()
        ]

    def _seleccionar_en_combo(
        self,
        selector: str,
        texto_busqueda: str,
        fallback_primera_valida: bool = False,
        estricto: bool = False,
    ) -> bool:
        """
        Selecciona una opción por texto/valor. Si el control no existe en el formulario de este producto
        lo omite (devuelve False). Con `estricto=True` nunca adivina: el fallback solo se usa si hay UNA
        única opción válida; con varias y sin coincidencia lanza un error listando las opciones.
        """
        cmb = self.frame.locator(selector)
        if cmb.count() == 0:
            logger.info(f"ℹ️ {selector} no existe en el formulario de este producto; se omite")
            return False
        if not cmb.is_visible():
            logger.warning(f"⚠️ {selector} existe pero no está visible; se omite")
            return False

        opciones = self._opciones_combo(selector)
        if texto_busqueda:
            objetivo = texto_busqueda.strip().lower()
            for val, txt in opciones:
                if objetivo in txt.lower() or objetivo == val.lower():
                    cmb.select_option(value=val)
                    self.page.wait_for_timeout(400)
                    logger.info(f"✅ {selector} = {txt}")
                    return True

        actual = cmb.input_value()
        if actual not in ("", "-1"):
            return True

        if estricto and len(opciones) != 1:
            lista = " | ".join(t for _, t in opciones) or "(sin opciones)"
            motivo = f"'{texto_busqueda}' no coincide con ninguna opción" if texto_busqueda else "no se indicó un valor"
            raise RuntimeError(f"{selector}: {motivo}. Opciones de este producto: {lista}")

        if fallback_primera_valida or estricto:
            for val, txt in opciones:
                cmb.select_option(value=val)
                self.page.wait_for_timeout(400)
                logger.warning(f"⚠️ {selector}: sin coincidencia para '{texto_busqueda}', se usó '{txt}'")
                return True
        return False

    def completar_datos_asistencia(
        self,
        pais_origen: str,
        destino: str,
        lista_destino: str = "",
        tipo_viaje_sv: str = "",
    ):
        """
        Paso 6 (Tab T0, 'Vehicle Details' en algunos productos): Type of trip, Destination y, si el
        producto los tiene, País de origen y Lista de destino. Se adapta al formulario del producto:
        solo completa los controles que existen y no adivina el destino cuando hay varias opciones.
        """
        logger.info(f"Completando Tab T0: origen={pais_origen}, destino={destino}, lista={lista_destino}")
        self._abrir_tab("Wcr_carTabPage_T0", "(datos de viaje)")

        presentes = [
            sel
            for sel in ("#CMBTPVIAJESV", "#txtPaisOrigen", "#CMBDESTINOSV", "#txtListaDestinos")
            if self.frame.locator(sel).count() > 0
        ]
        logger.info(f"Controles de viaje presentes en este producto: {presentes or 'ninguno'}")
        if not presentes:
            raise RuntimeError("La pestaña T0 no muestra controles de viaje (Destination / Type of trip) para este producto.")

        self._seleccionar_en_combo("#CMBTPVIAJESV", tipo_viaje_sv, fallback_primera_valida=True)
        self._seleccionar_en_combo("#txtPaisOrigen", pais_origen, fallback_primera_valida=False)
        self._seleccionar_en_combo("#CMBDESTINOSV", destino, estricto=True)

        texto_lista = lista_destino or destino
        if texto_lista:
            txt_lista = self.frame.locator("#txtListaDestinos")
            if txt_lista.count() > 0 and txt_lista.is_visible():
                txt_lista.fill(texto_lista.upper())

    def agregar_pasajero(self, pasajero: Pasajero, indice: int = 1, correo_respaldo: str = ""):
        """Pasos 7-8 (Tab T1): agrega pasajero indexado _1, _2, ..."""
        logger.info(f"Agregando pasajero #{indice}: {pasajero.nombre} {pasajero.apellido_paterno}")

        if indice == 1:
            self.frame.locator("#Wcr_carTabPage_T1").click()
            self.page.wait_for_timeout(500)
            if pasajero.es_titular:
                chk = self.frame.locator("#chkTomadorRiesgo")
                if chk.count() > 0 and not chk.is_checked():
                    chk.check()
                    self.page.wait_for_timeout(400)
        else:
            self.frame.locator("#btnNuevoRiesgo").click()
            self.frame.locator(f"#txtIdFiscal_{indice}").wait_for(state="visible", timeout=self.timeout_ms)

        # Tipo de documento
        cmb_tipo = self.frame.locator(f"#cmbTipoCIF_{indice}")
        tipo_upper = pasajero.tipo_documento.strip().upper()
        seleccionado = False
        for opt in cmb_tipo.locator("option").all():
            if tipo_upper in opt.inner_text().upper():
                cmb_tipo.select_option(value=opt.get_attribute("value"))
                seleccionado = True
                break
        if not seleccionado:
            cmb_tipo.select_option(value=str(tipo_documento_codigo(tipo_upper)))

        # Documento y DV
        self.frame.locator(f"#txtIdFiscal_{indice}").fill(pasajero.numero_documento.strip())
        dv_loc = self.frame.locator(f"#txtDV_{indice}")
        if dv_loc.count() > 0 and dv_loc.is_visible():
            dv_loc.fill(pasajero.digito_verificador.strip() if tipo_upper == "RUN" else "")

        # Nombres y Apellidos
        self.frame.locator(f"#txtNmAsegurado_{indice}").fill(pasajero.nombre.strip().upper())
        apellidos = f"{pasajero.apellido_paterno} {pasajero.apellido_materno}".strip().upper()
        self.frame.locator(f"#txtApeAsegurado_{indice}").fill(apellidos)

        # Fecha de nacimiento + Edad explícita (#TXTEDADSV no se autocalcula en _MAWDY [CHILE])
        f_nac = self.frame.locator(f"#txtFhNacimiento_{indice}")
        f_nac.fill(pasajero.fecha_nacimiento.strip())
        f_nac.press("Tab")
        self.page.wait_for_timeout(300)

        edad_loc = self.frame.locator(f"#TXTEDADSV_{indice}")
        if edad_loc.count() > 0:
            val_edad = edad_loc.input_value().strip()
            if val_edad in ("", "0"):
                edad = _calcular_edad(pasajero.fecha_nacimiento)
                if edad is not None:
                    edad_loc.fill(str(edad))
                    edad_loc.press("Tab")
                    self.page.wait_for_timeout(400)

        # Sexo (#cmbSexo_{indice} en _MAWDY [CHILE]: 2=Femenino, 5=Masculino, 8=No Indica)
        sexo_loc = self.frame.locator(f"#cmbSexo_{indice}")
        if sexo_loc.count() > 0 and sexo_loc.is_visible():
            disponibles = [opt.get_attribute("value") or "" for opt in sexo_loc.locator("option").all()]
            val_sexo = str(pasajero.sexo)
            if val_sexo in disponibles and val_sexo not in ("", "-1"):
                sexo_loc.select_option(value=val_sexo)
            else:
                for cand in ("5", "2", "8", "1"):
                    if cand in disponibles:
                        sexo_loc.select_option(value=cand)
                        break

        # Prefijo, Teléfono y Email
        ddd_loc = self.frame.locator(f"#txtDDD_movil_{indice}")
        if ddd_loc.count() > 0 and ddd_loc.is_visible() and not ddd_loc.input_value().strip():
            ddd_loc.fill(str(pasajero.cod_pais_movil or 56))

        telefono = (pasajero.telefono or "").strip() or "900000000"
        self.frame.locator(f"#txtMovil_{indice}").fill(telefono)

        email = (pasajero.email or "").strip() or correo_respaldo.strip()
        if email:
            self.frame.locator(f"#txtEmail_{indice}").fill(email.upper())

        if indice > 1 and pasajero.es_titular:
            rb = self.frame.locator(f"#rbTomador_{indice}")
            if rb.count() > 0 and rb.is_enabled():
                rb.check()

    def completar_file(self, file_ref: str):
        """Paso 9 (Tab T2): llena File1 si existe en la configuración del producto/país."""
        if not file_ref:
            return
        self.frame.locator("#Wcr_carTabPage_T2").click()
        self.page.wait_for_timeout(400)
        campo_file = self.frame.locator(
            "#Wcr_carTabPage_C2 input[id*='File'], #Wcr_carTabPage_C2 input[id*='file'], "
            "#Wcr_carTabPage_C2 input[name*='File'], #Wcr_carTabPage_C2 input[name*='file']"
        )
        if campo_file.count() > 0:
            campo_file.first.fill(file_ref.strip().upper())
            logger.info(f"✅ File1 completado: {file_ref}")
        else:
            logger.info("ℹ️ Campo File1 no presente en este producto/entorno")

    def terminar_emision(self) -> tuple[bool, str]:
        """Paso 10: clic en #Wcr_btnAceptar, detecta alertas/errores de modal y espera confirmación real."""
        logger.info("Finalizando emisión (#Wcr_btnAceptar)...")
        self.mostrar_hud("Confirmando Emisión", "Enviando contrato a Phoenix...")
        self._limpiar_alertas()

        btn_aceptar = self.frame.locator("#Wcr_btnAceptar")
        btn_aceptar.click()

        modal = self.frame.locator("#Wcr_PW-1")
        deadline = time.monotonic() + (self.timeout_ms / 1000.0)
        ok_confirmado = False

        while time.monotonic() < deadline:
            self.page.wait_for_timeout(500)

            for al in list(self.alertas):
                al_low = al.lower()
                if any(
                    kw in al_low
                    for kw in (
                        "actualmente existe",
                        "error",
                        "obligatorio",
                        "debe ",
                        "inválid",
                        "invalid",
                        "incorrecto",
                        "no puede",
                        "revise",
                    )
                ):
                    self._capturar_screenshot("error_alerta_aceptar")
                    return False, " | ".join(self.alertas)
                if al_low.strip() == "ok" or "correctamente" in al_low:
                    ok_confirmado = True

            if modal.count() == 0 or not modal.is_visible():
                logger.info("✅ Modal #Wcr_PW-1 cerrado tras Aceptar")
                return True, ""

            errores_ui = self._detectar_errores_visibles_modal()
            if errores_ui:
                self._capturar_screenshot("error_validacion_modal")
                return False, " | ".join(errores_ui)

            # En PRE, tras la alerta 'OK', el modal a veces queda abierto con #Wcr_btnAceptar deshabilitado
            try:
                if ok_confirmado and not btn_aceptar.is_enabled():
                    logger.info("Confirmación 'OK' recibida y botón Aceptar deshabilitado; recargando grilla...")
                    self.ir_a_emision_vouchers()
                    return True, ""
            except Exception:
                pass

        if ok_confirmado:
            logger.info("Modal aún abierto tras confirmación 'OK'; recargando listaContratos.aspx...")
            self.ir_a_emision_vouchers()
            return True, ""

        self._capturar_screenshot("timeout_terminar_emision")
        return False, "Timeout esperando confirmación de emisión en #Wcr_PW-1 (revisar captura en screenshots/)."

    def obtener_precios(self, tipo_viaje: str = "internacional") -> dict:
        """
        Paso 13 del manual:
        - Viaje NACIONAL (pág. 15): toma valores netos (+ IVA) de la parte derecha (#Wcr_txtPrecioNeto y #Wcr_txt_PrecioCambioPN).
        - Viaje INTERNACIONAL (pág. 17): toma valores brutos (fila Total de Tab T4 #tablaPreciosCoberturas).
        """
        precios = {"clp": "", "usd": "", "tc": ""}
        es_nacional = tipo_viaje.strip().lower() == "nacional"

        try:
            campos_pref = ("#Wcr_txtPrecioNeto", "#Wcr_txtPrecioBruto") if es_nacional else ("#Wcr_txtPrecioBruto", "#Wcr_txtPrecioNeto")
            for cid in campos_pref:
                loc = self.frame.locator(cid)
                if loc.count() > 0 and loc.input_value().strip():
                    precios["clp"] = loc.input_value().strip()
                    break

            usd_loc = self.frame.locator("#Wcr_txt_PrecioCambioPN")
            if usd_loc.count() > 0 and usd_loc.input_value().strip():
                precios["usd"] = usd_loc.input_value().strip()

            tc_loc = self.frame.locator("#Wcr_cambioDivisa")
            if tc_loc.count() > 0 and tc_loc.input_value().strip():
                precios["tc"] = tc_loc.input_value().strip()

            tab_t4 = self.frame.locator("#Wcr_carTabPage_T4")
            if tab_t4.count() > 0:
                tab_t4.click()
                self.page.wait_for_timeout(600)
                texto_t4 = self.frame.locator("#Wcr_carTabPage_C4").inner_text()
                m_tc = re.search(r"tasa cambio[#:\s]*([\d.,]+)", texto_t4, re.IGNORECASE)
                if m_tc:
                    precios["tc"] = m_tc.group(1)

                if not es_nacional:
                    fila_total = self.frame.locator("#Wcr_carTabPage_tblPreciosCoberturas1 tr:has-text('Total')").first
                    if fila_total.count() > 0:
                        celdas = [c.strip() for c in fila_total.locator("td").all_inner_texts() if c.strip()]
                        if len(celdas) >= 2:
                            precios["clp"] = celdas[-2]
                            precios["usd"] = celdas[-1]
        except Exception as e:
            logger.warning(f"⚠️ Error al extraer precios: {e}")

        logger.info(f"Precios extraídos ({tipo_viaje}): {precios}")
        return precios

    def descargar_voucher(self, file_ref: str = "", titular: str = "") -> tuple[str, str | None]:
        """
        Pasos 11-12: reabre el voucher recién emitido, obtiene el número (#Wcr_txtNpoliza + #Wcr_txtSufijo)
        y captura el PDF usando Playwright `expect_download()` en Tab T6 o Tab T7 (#Wcr_carTabPage_btnReimprimeCertificado).
        """
        logger.info("Capturando número de voucher y descargando PDF...")
        numero_voucher = ""

        modal = self.frame.locator("#Wcr_PW-1")
        if modal.count() == 0 or not modal.is_visible():
            self.mostrar_hud("Descargando Voucher", "Reabriendo voucher emitido en grilla...")
            soc_filtro = self.config.get("sociedad_grilla_filtro", "CHILE")
            if soc_filtro:
                self.seleccionar_sociedad_grilla(soc_filtro)
            fila0 = self.frame.locator("tr[id*='DXDataRow0']").first
            fila0.wait_for(state="visible", timeout=self.timeout_ms)
            for intento in range(2):
                fila0.dblclick()
                try:
                    modal.wait_for(state="visible", timeout=10000)
                    break
                except PlaywrightTimeoutError:
                    if intento == 0:
                        logger.warning("⚠️ Primer doble clic no abrió #Wcr_PW-1; reintentando...")
                        self._esperar_sin_loading(5000)
                    else:
                        raise
            self._esperar_sin_loading()

        num_base = self.frame.locator("#Wcr_txtNpoliza").input_value().strip()
        suf_val = self.frame.locator("#Wcr_txtSufijo").input_value().strip() if self.frame.locator("#Wcr_txtSufijo").count() > 0 else ""
        numero_voucher = f"{num_base}{suf_val}" if (num_base and suf_val.isdigit() and len(num_base) < 5) else num_base

        download_dir = resolver_ruta(self.config.get("download_dir"), "~/Downloads/vouchers")
        os.makedirs(download_dir, exist_ok=True)
        nombre_base = formatear_nombre_voucher(numero_voucher or "SIN_NUM", file_ref or "SIN_FILE", titular or "TITULAR")
        ruta_destino = os.path.join(download_dir, _sanitizar_nombre_archivo(nombre_base) + ".pdf")

        # Intentar Tab T6 (solo si hay link PDF, excluyendo XML/XSD), si no Tab T7 Print Certificate
        boton_descarga = None
        self.frame.locator("#Wcr_carTabPage_T6").click()
        self.page.wait_for_timeout(600)
        links_t6 = self.frame.locator(
            "#Wcr_carTabPage_C6 tr:has-text('CREACION'):not(:has-text('XML')):not(:has-text('XSD')) a:has-text('Abrir'), "
            "#Wcr_carTabPage_C6 tr:has-text('CONTRATO'):not(:has-text('XML')):not(:has-text('XSD')) a:has-text('Open')"
        )
        if links_t6.count() > 0:
            boton_descarga = links_t6.first
        else:
            self.frame.locator("#Wcr_carTabPage_T7").click()
            self.page.wait_for_timeout(600)
            btn_print = self.frame.locator("#Wcr_carTabPage_btnReimprimeCertificado")
            if btn_print.count() > 0 and btn_print.is_enabled():
                boton_descarga = btn_print

        if not boton_descarga:
            logger.warning("⚠️ No se encontró botón de descarga de voucher en T6 ni T7")
            return numero_voucher, None

        try:
            with self.page.expect_download(timeout=max(self.timeout_ms, 45000)) as dl_info:
                boton_descarga.click()
            download = dl_info.value
            download.save_as(ruta_destino)
            logger.info(f"✅ Voucher PDF capturado y guardado en: {ruta_destino}")
            return numero_voucher, ruta_destino
        except PlaywrightTimeoutError:
            logger.error("❌ Tiempo de espera agotado esperando la descarga del PDF")
            self._capturar_screenshot("error_descarga_pdf")
            return numero_voucher, None

    def emitir_voucher(
        self,
        datos: DatosEmision,
        dry_run: bool = False,
        enviar_correo_asegurado: bool | None = None,
        email_modo: str | None = None,
        confirmar_antes_aceptar: bool = False,
    ) -> dict:
        """
        Ejecuta la emisión completa con Playwright, captura el PDF del voucher
        y opcionalmente lo envía por correo al asegurado (y/o ejecutivo).
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
            "error": "",
        }

        try:
            logger.info("=" * 60)
            logger.info(f"EMISIÓN PLAYWRIGHT {'[DRY-RUN] ' if dry_run else ''}: {datos.codigo_cliente}")
            logger.info("=" * 60)

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

            for i, pasajero in enumerate(datos.pasajeros, start=1):
                self.agregar_pasajero(pasajero, indice=i, correo_respaldo=datos.correo_solicitante)

            if datos.file1:
                self.completar_file(datos.file1)

            if dry_run:
                self.mostrar_hud("Dry-Run Completo", "Revisando precios antes de cancelar...")
                resultado["precios"] = self.obtener_precios(datos.tipo_viaje)
                self._capturar_screenshot("dry_run_antes_aceptar")
                self.frame.locator("#Wcr_btnCancelar").click()
                resultado["exito"] = True
                resultado["resumen"] = "Dry-run completado con Playwright sin emitir."
                return resultado

            if confirmar_antes_aceptar:
                self.mostrar_hud("Esperando Confirmación", "Revisa la pantalla y confirma en consola para emitir.")
                resp = input("\n👀 Revisa los datos cargados en Phoenix. ¿Confirmar emisión (Aceptar)? (s/n) [s]: ").strip().lower()
                if resp == "n":
                    self.frame.locator("#Wcr_btnCancelar").click()
                    resultado["error"] = "Emisión cancelada manualmente por el usuario antes de Aceptar."
                    return resultado

            exito, err_msg = self.terminar_emision()
            if not exito:
                resultado["error"] = err_msg or "Error al confirmar emisión"
                return resultado

            titular_nombre = ""
            if datos.pasajeros:
                p0 = next((p for p in datos.pasajeros if p.es_titular), datos.pasajeros[0])
                titular_nombre = f"{p0.nombre} {p0.apellido_paterno}".strip()

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

            # Captura y envío por correo al asegurado
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

            logger.info(f"✅ EMISIÓN PLAYWRIGHT COMPLETADA (Voucher: {num_voucher})")

        except Exception as e:
            resultado["error"] = str(e)
            logger.error(f"❌ ERROR EN EMISIÓN PLAYWRIGHT: {e}")
            self._capturar_screenshot("error_emision")

        return resultado

    def _capturar_screenshot(self, nombre: str):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta = os.path.join(os.path.dirname(__file__), "screenshots", f"{nombre}_{timestamp}.png")
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            if self.page:
                self.page.screenshot(path=ruta)
                logger.info(f"📸 Screenshot guardado: {ruta}")
        except Exception:
            pass

    def cerrar(self):
        """Cierra el contexto persistente y detiene Playwright."""
        try:
            if self.context:
                self.context.close()
            if self.playwright:
                self.playwright.stop()
            logger.info("Playwright cerrado (sesión guardada en perfil persistente)")
        except Exception:
            pass
