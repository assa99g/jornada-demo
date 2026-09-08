@echo off
setlocal
if not exist "%~dp0runtime\pythonw.exe" (
  echo Falta la carpeta runtime. Extraiga todo el archivo ZIP antes de iniciar.
  pause
  exit /b 1
)
start "Jornada" "%~dp0runtime\pythonw.exe" "%~dp0iniciar.py"
endlocal
