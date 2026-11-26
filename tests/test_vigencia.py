import sys, types
import os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
try:
    import playwright  # noqa
except ImportError:
    pw = types.ModuleType("playwright"); sync = types.ModuleType("playwright.sync_api")
    class _TE(Exception): pass
    for n in ("sync_playwright","Playwright","BrowserContext","Page","FrameLocator","Dialog"): setattr(sync, n, object)
    sync.TimeoutError = _TE
    sys.modules["playwright"] = pw; sys.modules["playwright.sync_api"] = sync
import phoenix_emision as pe
from datetime import datetime

class Loc:
    def __init__(s, form, sel): s.f, s.sel = form, sel
    def count(s): return 1 if s.sel in s.f.fields else 0
    def input_value(s): return s.f.fields[s.sel]
    def is_visible(s): return s.f.visible.get(s.sel, True)
    def is_enabled(s): return True
    def get_attribute(s, a): return None
    def fill(s, v): s.f.fill(s.sel, v)
    def press(s, k): s.f.tab(s.sel)
    def evaluate(s, js): pass
class Form:
    def __init__(s, calcula_bien=True, dur_visible=True):
        s.fields = {"#Wcr_txtFhInicio": "26/11/2026", "#Wcr_txtFhFin": "", "#Wcr_txtDuracionDias": ""}
        s.visible = {"#Wcr_txtDuracionDias": dur_visible}; s.ok = calcula_bien
    def fill(s, sel, v): s.fields[sel] = v
    def tab(s, sel):
        from datetime import timedelta
        ini = pe._parsear_fecha_phoenix(s.fields["#Wcr_txtFhInicio"])
        if sel == "#Wcr_txtDuracionDias" and s.fields[sel].isdigit():
            s.fields["#Wcr_txtFhFin"] = (ini + timedelta(days=int(s.fields[sel]) - 1)).strftime("%d/%m/%Y") if s.ok else "NaN/NaN/NaN"
        if sel == "#Wcr_txtFhFin":
            fin = pe._parsear_fecha_phoenix(s.fields[sel])
            s.fields["#Wcr_txtDuracionDias"] = str((fin - ini).days + 1) if (s.ok and fin) else ""
    def locator(s, sel): return Loc(s, sel)

def make(form):
    a = pe.PhoenixAutomation.__new__(pe.PhoenixAutomation)
    a.visual_hud = False; a.alertas = []; a.ultima_alerta = None; a.timeout_ms = 35000
    class P:
        def wait_for_timeout(s, ms): pass
    a.page = P(); a.__dict__["_FRAME"] = form
    pe.PhoenixAutomation.frame = property(lambda self: self.__dict__["_FRAME"])
    a._esperar_sin_loading = lambda *x, **k: None
    a._capturar_screenshot = lambda *x, **k: None
    return a

def run(label, fn, expect_err=None):
    try:
        fn(); print(f"[{'FAIL' if expect_err else 'PASS'}] {label}")
    except Exception as e:
        print(f"[{'PASS' if expect_err and expect_err in str(e) else 'FAIL'}] {label} -> {type(e).__name__}: {e}")

# 1) caso del usuario: Effect 26/11, 7 días -> Expiry 02/12/2026
f = Form(); a = make(f); a.completar_vigencia("26/11/2026", "02/12/2026")
print("[PASS] 7 días -> Expiry", f.fields["#Wcr_txtFhFin"], "Duration", f.fields["#Wcr_txtDuracionDias"]) if f.fields["#Wcr_txtFhFin"] == "02/12/2026" and f.fields["#Wcr_txtDuracionDias"] == "7" else print("[FAIL] 7 días", f.fields)
# 2) Duration oculto -> respaldo por Expiry
f = Form(dur_visible=False); a = make(f); a.completar_vigencia("27/11/2026", "29/11/2026")
print("[PASS] respaldo por Expiry", f.fields) if f.fields["#Wcr_txtFhFin"] == "29/11/2026" else print("[FAIL] respaldo", f.fields)
# 3) Phoenix devuelve NaN -> debe abortar con error claro
f = Form(calcula_bien=False); a = make(f)
pe.PhoenixAutomation._esperar_vigencia.__defaults__ = (50,)
run("Expiry=NaN aborta", lambda: a.completar_vigencia("27/11/2026", "29/11/2026"), expect_err="no calculó una vigencia válida")
# 4) fechas vacías / invertidas / formato corto
f = Form(); a = make(f)
run("fecha vacía aborta", lambda: a.completar_vigencia("", ""), expect_err="Vigencia inválida")
run("fin < inicio aborta", lambda: a.completar_vigencia("29/11/2026", "27/11/2026"), expect_err="anterior al inicio")
f = Form(); a = make(f); a.completar_vigencia("27-11-26", "29-11-26")
print("[PASS] formato DD-MM-AA normalizado", f.fields) if f.fields["#Wcr_txtFhFin"] == "29/11/2026" else print("[FAIL] formato corto", f.fields)
