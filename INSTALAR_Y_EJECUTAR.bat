@echo off
chcp 65001 >nul
title Transcriptor flotante
python --version >nul 2>&1
if errorlevel 1 (
  echo Python no esta instalado. Descargalo de https://www.python.org/downloads/
  echo Marca la casilla "Add Python to PATH" al instalar.
  pause
  exit /b
)
if not exist ".instalado_v9" (
  echo Instalando dependencias, espera un momento...
  python -m pip install --upgrade pip
  python -m pip install SpeechRecognition pyperclip pyautogui setuptools openpyxl
  python -m pip install pyaudio
  if errorlevel 1 (
    echo Instalando PyAudio con pipwin...
    python -m pip install pipwin
    python -m pipwin install pyaudio
  )
  python -m pip install vosk faster-whisper
  python descargar_modelo.py
  echo ok> .instalado_v9
)
start "" pythonw transcriptor_flotante.py
