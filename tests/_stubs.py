"""Permite importar los módulos del proyecto sin tener Playwright instalado (solo para pruebas)."""
import os
import sys
import types

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

try:
    import playwright  # noqa: F401
except ImportError:
    pw = types.ModuleType("playwright")
    sync = types.ModuleType("playwright.sync_api")
    for nombre in ("sync_playwright", "Playwright", "BrowserContext", "Page", "FrameLocator", "Dialog"):
        setattr(sync, nombre, object)
    sync.TimeoutError = type("TimeoutError", (Exception,), {})
    sys.modules["playwright"] = pw
    sys.modules["playwright.sync_api"] = sync
