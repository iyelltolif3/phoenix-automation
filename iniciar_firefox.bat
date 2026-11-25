@echo off
echo ============================================
echo  INICIANDO FIREFOX CON DEPURACION REMOTA
echo ============================================
echo.
echo Esto abrira Firefox en modo de depuracion
echo para que el script de automatizacion pueda
echo conectarse a el.
echo.
echo 1. Inicia sesion en Phoenix normalmente
echo 2. Luego ejecuta: python ejecutar_emision.py
echo.

start "" "%LOCALAPPDATA%\Mozilla Firefox\firefox.exe" -marionette -start-debugger-server 2828

echo Firefox iniciado. Ahora inicia sesion en Phoenix:
echo https://phoenix-pre-net-eus-int.mia-assistance.com/default.aspx
echo.
pause
