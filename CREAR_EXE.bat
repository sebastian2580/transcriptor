@echo off
chcp 65001 >nul
title Creando TranscriptorFlotante
python --version >nul 2>&1
if errorlevel 1 (
  echo Python no esta instalado en ESTE equipo. Solo se necesita aqui, una vez,
  echo para crear el programa. Descargalo de https://www.python.org/downloads/
  echo y marca "Add Python to PATH".
  pause
  exit /b
)
if not exist "similitud.xlsx" (
  echo Falta el archivo similitud.xlsx en esta carpeta. Es la base de datos.
  pause
  exit /b
)
if not exist "base_datos.txt" (
  echo Falta el archivo base_datos.txt en esta carpeta.
  pause
  exit /b
)

set OPENBLAS_NUM_THREADS=1
set OMP_NUM_THREADS=1
set MKL_NUM_THREADS=1
set NUMEXPR_NUM_THREADS=1

echo [1/4] Preparando un entorno limpio...
if not exist ".venv_build\Scripts\python.exe" python -m venv .venv_build
set PY=.venv_build\Scripts\python.exe
%PY% -m pip install --upgrade pip
%PY% -m pip install SpeechRecognition pyperclip pyautogui pyaudio pyinstaller openpyxl
if errorlevel 1 (
  %PY% -m pip install SpeechRecognition pyperclip pyautogui pyinstaller openpyxl
  %PY% -m pip install pipwin
  %PY% -m pipwin install pyaudio
)
%PY% -m pip install vosk
%PY% -m pip install faster-whisper
if errorlevel 1 (
  echo AVISO: no se pudo instalar faster-whisper. El programa usara Vosk.
)

echo.
echo [2/4] Descargando los modelos de voz (solo la primera vez; Whisper pesa ~480 MB)...
%PY% descargar_modelo.py %1

set WOPT=
%PY% -c "import faster_whisper" >nul 2>&1
if not errorlevel 1 set WOPT=--collect-all faster_whisper --collect-all ctranslate2 --collect-all av --collect-all tokenizers --collect-all huggingface_hub
set VOPT=--collect-all vosk --hidden-import vosk

echo.
echo [3/4] Creando el programa (puede tardar varios minutos)...
%PY% -m PyInstaller --noconfirm --clean --onedir --noconsole --icon icono.ico --collect-all speech_recognition --hidden-import pyaudio --hidden-import openpyxl --hidden-import buscador_voz --add-data "similitud.xlsx;." --add-data "base_datos.txt;." --add-data "alias.txt;." %WOPT% %VOPT% --exclude-module scipy --exclude-module pandas --exclude-module matplotlib --name TranscriptorFlotante transcriptor_flotante.py

if not exist "dist\TranscriptorFlotante\TranscriptorFlotante.exe" (
  echo.
  echo Algo fallo. Reinicia el equipo, cierra otros programas, y mueve esta
  echo carpeta a C:\Transcriptor antes de volver a intentar.
  pause
  exit /b
)

echo.
echo [4/4] Copiando modelos y archivos editables junto al programa...
if exist "modelo_whisper" xcopy /E /I /Y /Q "modelo_whisper" "dist\TranscriptorFlotante\modelo_whisper" >nul
if exist "modelo_voz" xcopy /E /I /Y /Q "modelo_voz" "dist\TranscriptorFlotante\modelo_voz" >nul
copy /Y similitud.xlsx "dist\TranscriptorFlotante\" >nul
copy /Y alias.txt "dist\TranscriptorFlotante\" >nul
copy /Y vocabulario.txt "dist\TranscriptorFlotante\" >nul
powershell -NoProfile -Command "Compress-Archive -Path 'dist\TranscriptorFlotante\*' -DestinationPath 'dist\TranscriptorFlotante.zip' -Force"

echo.
echo LISTO.
echo  - Carpeta:  dist\TranscriptorFlotante   (abre TranscriptorFlotante.exe)
echo  - Para llevarlo a otro equipo copia  dist\TranscriptorFlotante.zip , descomprimelo
echo    alli y abre TranscriptorFlotante.exe. NO necesita Python ni internet.
echo  - Al abrirlo debe decir:  [VOZ] Motor WHISPER listo
pause
exit /b
