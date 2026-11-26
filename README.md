# Phoenix v4 - Automatización de Emisiones de Viajes (Playwright)

Automatización completa para emisiones individuales y grupales en Phoenix v4 (ASP.NET + DevExpress) usando **Playwright**, con captura determinista del PDF del voucher y envío por correo al asegurado.

## Instalación en un Nuevo Equipo (Windows)

```powershell
# 1. Clonar el repositorio
git clone <URL_DEL_REPOSITORIO>
cd phoenix_automation

# 2. Crear y activar entorno virtual (recomendado)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. Instalar dependencias y navegador (si no usas Edge/Chrome instalado)
pip install -r requirements.txt
playwright install chromium
```

> **Nota sobre sesión SSO**: La primera vez que ejecutes el script, Edge/Chrome abrirá la página de Phoenix y te pedirá iniciar sesión con Microsoft SSO en la ventana del navegador; luego presiona `ENTER` en la consola. La sesión quedará guardada localmente en `browser_profile/` (excluida de Git por seguridad).

## Guía Rápida de Uso

### Opción A: Menú Interactivo (incluye sección `[MVP] Autoemisión visual desde correo`)
```powershell
python ejecutar_emision.py
```

El apartado **`[MVP] Autoemisión visual desde correo`** (`mvp_autoemision.py`) permite:
- Cargar un archivo `.eml` real (con adjuntos `.xlsx` o tablas HTML) o pegar el texto del correo.
- Auditar rigurosamente todos los datos y RUTs (módulo 11, edades/fechas sospechosas, teléfonos) antes de tocar Phoenix.
- Ejecutar la emisión en pantalla con navegador visible, indicador flotante paso a paso (HUD) y confirmación opcional antes de `Aceptar` (o en modo `--dry-run`).

### Opción B: Emisión por Lote desde Excel (recomendado)
```powershell
# 1. Probar sin emitir (Dry-Run)
python emitir_lote.py plantilla_solicitudes.xlsx --dry-run

# 2. Emitir lote real generando archivos .eml listos con el PDF adjunto
python emitir_lote.py plantilla_solicitudes.xlsx

# 3. Emitir lote y enviar correo automático vía SMTP (Office 365)
$env:PHOENIX_SMTP_PASSWORD = "tu_clave_o_app_password"
python emitir_lote.py plantilla_solicitudes.xlsx --email-modo smtp
```

## Rutas Portables (`config.json`)

- `browser_profile_dir`: `"browser_profile"` (relativa al proyecto, se crea automáticamente en cada equipo).
- `download_dir`: `"~/Downloads/vouchers"` (resuelve automáticamente la carpeta `Downloads` del usuario activo).
- `email.eml_output_dir`: `"~/Downloads/vouchers/correos"`.
- `browser`: `"edge"` (usa Microsoft Edge corporativo instalado en Windows; también soporta `"chrome"`, `"chromium"` y `"firefox"`).

## Modos de Envío de Correo al Asegurado (`config.json` -> `email.modo`)

| Modo | Comportamiento |
|------|----------------|
| `eml` (por defecto) | Genera un `.eml` con el PDF del voucher adjunto en `~/Downloads/vouchers/correos/` |
| `smtp` | Envío automático vía `smtp.office365.com:587` leyendo la clave desde la variable de entorno `PHOENIX_SMTP_PASSWORD` |
| `outlook_send` | Envía automáticamente usando Outlook clásico instalado en Windows (COM) |
| `outlook_draft` | Abre el borrador redactado con el PDF adjunto en Outlook clásico |

## Plantilla Grupal Oficial Phoenix

La plantilla oficial descargada desde `#plantillaExcel` en Phoenix PRE utiliza la hoja `Sheet1` con estas 9 columnas:

`ID Type | Fiscal ID | First name | Surname | D.o.b. | Age | Gender | Email | Mobile`

- `ID Type`: `1007` = RUN (Persona Natural), `2` = Pasaporte.
- `Gender`: `2` = Femenino, `5` = Masculino.

## Archivos del Proyecto

| Archivo | Descripción |
|---------|-------------|
| `config.json` | URLs (PRE y Chile), navegador, SMTP/email, reglas de negocio y catálogo de dealers |
| `utils.py` | Resolución de rutas portables, búsqueda de dealers, validación de RUT y reglas de negocio |
| `email_sender.py` | Construcción y envío/borrador de correos al asegurado y al ejecutivo con PDF adjunto |
| `phoenix_emision.py` | Motor **Playwright** con selectores reales de `_MAWDY [CHILE]`, `expect_download()` y manejo de alertas |
| `emision_grupal.py` | Emisión grupal en Playwright, lectura segura de `.xlsx` Phoenix, validación y carga por `#btnCargaPlantillaAsegurado` |
| `ejecutar_emision.py` | Asistente interactivo por consola |
| `emitir_lote.py` | Procesamiento masivo desde `plantilla_solicitudes.xlsx` + reporte `resultados_emision_*.xlsx` |
| `plantilla_solicitudes.xlsx` | Plantilla Excel de ejemplo |
| `requirements.txt` | Dependencias Python del proyecto |
