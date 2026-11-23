"""
Transcriptor flotante de voz a texto
------------------------------------
- Botón circular flotante, siempre visible encima de otras ventanas.
- Clic  -> empieza a escuchar / clic otra vez -> detiene.
- Arrastra el botón para moverlo. Clic derecho -> menú (idioma, escribir
  automáticamente, ver transcripción, salir).
- El texto reconocido se muestra en una ventana, se copia al portapapeles y,
  si "Escribir automáticamente" está activo, se pega donde esté el cursor.

Requisitos:
    pip install SpeechRecognition pyaudio pyperclip pyautogui
(Sin internet: Whisper o Vosk; Google solo como último recurso.)
"""

import os
for _v in ("OPENBLAS_NUM_THREADS", "OMP_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
    os.environ.setdefault(_v, "2")          # evita errores de memoria al cargar librerías de cálculo

import csv
import json
import math
import re
import unicodedata
import queue
import sys
import traceback
import threading
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext

import speech_recognition as sr
from array import array

from buscador_voz import (Buscador, SENSIBILIDAD, decidir_coincidencia, leer_alias,
                          numeros_a_cifras)
from collections import defaultdict, deque

try:
    import pyaudio
except ImportError:          # se avisa al iniciar si falta
    pyaudio = None

try:
    import pyperclip
    import pyautogui
    PASTE_OK = True
except ImportError:
    PASTE_OK = False


IS_WIN = sys.platform == "win32"

if IS_WIN:
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    _PUL = ctypes.POINTER(ctypes.c_ulong)

    class _KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                    ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("dwExtraInfo", _PUL)]

    class _MOUSEINPUT(ctypes.Structure):
        _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                    ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                    ("time", wintypes.DWORD), ("dwExtraInfo", _PUL)]

    class _INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", _KEYBDINPUT), ("mi", _MOUSEINPUT)]

    class _INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]

    def win_type_unicode(texto):
        """Escribe el texto como si fuera el teclado (acentos, ñ, emojis incluidos)."""
        data = texto.encode("utf-16-le")
        units = [int.from_bytes(data[i:i + 2], "little") for i in range(0, len(data), 2)]
        events = []
        for u in units:
            for flags in (0x0004, 0x0004 | 0x0002):  # UNICODE down / UNICODE up
                inp = _INPUT(type=1)
                inp.u.ki = _KEYBDINPUT(0, u, flags, 0, None)
                events.append(inp)
        arr = (_INPUT * len(events))(*events)
        user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))

def _log_path():
    base = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                           else os.path.abspath(__file__))
    return os.path.join(base, "transcriptor_log.txt")


def log(msg):
    try:
        with open(_log_path(), "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S  ") + str(msg) + "\n")
    except Exception:
        pass


def list_inputs(pa):
    """[(indice, 'Nombre [driver]', nombre_base)] de los dispositivos que pueden grabar."""
    out = []
    for i in range(pa.get_device_count()):
        info = pa.get_device_info_by_index(i)
        if info.get("maxInputChannels", 0) > 0:
            base = info.get("name", f"Dispositivo {i}")
            try:
                host = pa.get_host_api_info_by_index(info["hostApi"])["name"]
            except Exception:
                host = ""
            out.append((i, f"{base} [{host}]" if host else base, base))
    return out


def input_devices():
    """Lista [(indice, nombre)] de dispositivos que sí pueden grabar."""
    try:
        import pyaudio as _pa
        pa = _pa.PyAudio()
        try:
            return [(i, n) for i, n, _b in list_inputs(pa)]
        finally:
            pa.terminate()
    except Exception:
        return []


def limpiar_texto(texto):
    """SOLO LETRAS MAYÚSCULAS: sin tildes ni diéresis, sin signos de puntuación.
    Se conserva la Ñ (es una letra distinta) y un punto o coma entre dos dígitos
    (ej. 1.500 o 3,5) para no dañar cifras."""
    texto = texto.upper().replace("Ñ", "\x03")
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")   # quita tildes
    texto = re.sub(r"(?<=\d)([.,])(?=\d)",
                   lambda m: "\x01" if m.group(1) == "." else "\x02", texto)
    texto = re.sub(r"[^\w\s\x01\x02\x03]|_", " ", texto)
    texto = texto.replace("\x01", ".").replace("\x02", ",").replace("\x03", "Ñ")
    return re.sub(r"\s+", " ", texto).strip()


def _app_dir():
    return os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                           else os.path.abspath(__file__))


def leer_productos(ruta):
    """Lee la columna A de un .xlsx, o una línea por producto de un .txt/.csv."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            vals = [str(r[0]).strip() for r in ws.iter_rows(min_col=1, max_col=1, values_only=True)
                    if r and r[0] not in (None, "")]
        finally:
            wb.close()
        return vals
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, encoding=enc) as f:
                return [ln.strip() for ln in f if ln.strip()]
        except UnicodeDecodeError:
            continue
    return []


def ruta_alias():
    """alias.txt junto al programa (o el incluido en el .exe)."""
    for c in (_app_dir(), getattr(sys, "_MEIPASS", _app_dir())):
        r = os.path.join(c, "alias.txt")
        if os.path.isfile(r):
            return r
    return None


def candidatos_base():
    """Dónde buscar la base: primero junto al programa (para poder actualizarla
    sin recrear el .exe) y luego la que va incluida dentro del .exe."""
    carpetas = [_app_dir(), getattr(sys, "_MEIPASS", _app_dir())]
    rutas = []
    for c in carpetas:
        for nombre in ("similitud.xlsx", "base_datos.txt"):
            r = os.path.join(c, nombre)
            if r not in rutas:
                rutas.append(r)
    return rutas


# =====================================================================
#  VOZ SIN INTERNET (Vosk) - utilidades
# =====================================================================
def a_16k(raw, rate):
    """Audio PCM 16 bits mono -> 16 000 Hz (lo que espera el modelo de voz)."""
    if rate == 16000:
        return raw
    try:
        import audioop
        return audioop.ratecv(raw, 2, 1, rate, 16000, None)[0]
    except Exception:
        a = array("h")
        a.frombytes(raw)
        n = len(a)
        m = int(n * 16000 / rate)
        out = array("h", [0]) * m
        paso = rate / 16000.0
        for k in range(m):
            pos = k * paso
            i0 = int(pos)
            i1 = i0 + 1 if i0 + 1 < n else i0
            f = pos - i0
            out[k] = int(a[i0] * (1 - f) + a[i1] * f)
        return out.tobytes()


def buscar_modelo():
    """Carpeta 'modelo_voz' junto al programa o incluida dentro del .exe."""
    for base in (_app_dir(), getattr(sys, "_MEIPASS", _app_dir())):
        r = os.path.join(base, "modelo_voz")
        if os.path.isdir(os.path.join(r, "am")) and os.path.isdir(os.path.join(r, "conf")):
            return r
    return None


def buscar_modelo_whisper():
    """Carpeta 'modelo_whisper' (faster-whisper) junto al programa."""
    for base in (_app_dir(), getattr(sys, "_MEIPASS", _app_dir())):
        r = os.path.join(base, "modelo_whisper")
        if os.path.isfile(os.path.join(r, "model.bin")) and os.path.isfile(os.path.join(r, "config.json")):
            return r
    return None


def leer_vocabulario():
    """vocabulario.txt (opcional): palabras/marcas que quieres que oiga bien, una por línea."""
    for c in (_app_dir(), getattr(sys, "_MEIPASS", _app_dir())):
        r = os.path.join(c, "vocabulario.txt")
        if os.path.isfile(r):
            try:
                with open(r, encoding="utf-8-sig") as f:
                    return [ln.split("#")[0].strip() for ln in f if ln.split("#")[0].strip()]
            except OSError:
                pass
    return []


def construir_prompt(buscador, extra=(), n=70):
    """Texto guía para Whisper: las palabras más usadas de TU base (marcas y abreviaturas,
    escritas como están) para que las reconozca y las escriba igual. Lo último pesa más,
    por eso tu vocabulario.txt va al final."""
    cuenta = sorted(((len(buscador.post[t]), t) for t in buscador.vocab
                     if t.isalpha() and len(t) >= 3), reverse=True)
    pal = [t.capitalize() for _, t in cuenta[:n]][::-1] + [e for e in extra]
    return ", ".join(pal)[:800]


_ALUCINACIONES = ("subtitulos", "amara.org", "gracias por ver", "suscribete", "suscribanse",
                  "no olvides", "www.", "nos vemos en el proximo", "gracias por su atencion")


def es_alucinacion(texto):
    t = unicodedata.normalize("NFD", texto.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return any(x in t for x in _ALUCINACIONES)


def registrar(motor, segundos, oido, res, escrito):
    """registro_voz.csv: qué oyó el programa y qué escribió (sirve para afinar)."""
    try:
        ruta = os.path.join(_app_dir(), "registro_voz.csv")
        if os.path.exists(ruta) and os.path.getsize(ruta) > 5_000_000:
            return
        nuevo = not os.path.exists(ruta)
        with open(ruta, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            if nuevo:
                w.writerow(["hora", "motor", "seg", "oido", "modo", "similitud", "escrito", "alternativas"])
            w.writerow([time.strftime("%Y-%m-%d %H:%M:%S"), motor, f"{segundos:.2f}", oido,
                        (res or {}).get("modo", "-"), (res or {}).get("score", ""), escrito,
                        " | ".join((res or {}).get("alts", [])[:3])])
    except Exception:
        pass


# Clases de ventana de clientes de escritorio remoto: ahí se pega con el portapapeles
REMOTE_CLASSES = {"TscShellContainerClass"}   # Conexión a Escritorio remoto (mstsc)


W, H = 112, 164     # tamaño de la ventana del botón (robot de 92x148)
RING_COLOR = "#fca5a5"

# Imágenes embebidas (PNG en base64): robot del botón flotante e ícono de la app
ROBOT_BTN_B64 = "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"
APP_ICON_B64 = "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"
TRANSPARENT = "#ff00ff"

IDIOMAS = {
    "Español (Colombia)": "es-CO",
    "Español (España)": "es-ES",
    "Español (México)": "es-MX",
    "English (US)": "en-US",
}


CHUNK = 1024


class Listener(threading.Thread):
    """Captura el micrófono con PyAudio, detecta cuándo hablas y corta las frases
    por silencio. Cada frase se manda a transcribir en otro hilo."""

    def __init__(self, app, device, gen):
        super().__init__(daemon=True)
        self.app, self.device, self.gen = app, device, gen
        self.halt_event = threading.Event()

    def stop(self):
        self.halt_event.set()

    def _put(self, kind, data):
        self.app.queue.put((kind, data))

    # ---- abrir el micrófono probando varias configuraciones y dispositivos ----
    def _open_device(self, pa, idx, nombre):
        info = pa.get_device_info_by_index(idx)
        max_ch = int(info.get("maxInputChannels", 1)) or 1
        default_rate = int(info.get("defaultSampleRate", 44100))
        rates = []
        for r in (default_rate, 16000, 44100, 48000, 32000, 22050, 8000):
            if r not in rates:
                rates.append(r)
        chans = [1] if max_ch < 2 else [1, 2]
        last = None
        for ch in chans:
            for rate in rates:
                try:
                    stream = pa.open(format=pyaudio.paInt16, channels=ch, rate=rate,
                                     input=True, input_device_index=idx,
                                     frames_per_buffer=CHUNK)
                    log(f"Micrófono abierto: {nombre} | {rate} Hz | {ch} canal(es)")
                    return stream, rate, ch
                except Exception as ex:
                    last = ex
                    log(f"No abrió '{nombre}' a {rate} Hz / {ch} canal(es): {ex}")
        raise RuntimeError(str(last))

    def _open(self, pa):
        inputs = list_inputs(pa)
        names = {i: (n, b) for i, n, b in inputs}
        explicit = self.device is not None
        idx = self.device
        if idx is None:
            try:
                idx = pa.get_default_input_device_info()["index"]
            except Exception:
                idx = None
        if idx is None and not inputs:
            raise RuntimeError("No hay ningún micrófono disponible en este equipo.")

        # Orden de intento: el elegido (o predeterminado) y, en modo automático,
        # los demás: primero los que se llaman igual (otro driver), luego el resto.
        orden = [idx] if idx is not None else []
        if not explicit:
            base0 = names.get(idx, ("", ""))[1] if idx is not None else ""
            mismos = [i for i, n, b in inputs if i != idx and b == base0 and base0]
            resto = [i for i, n, b in inputs if i != idx and i not in mismos]
            orden += mismos + resto

        primer_error = None
        for k, d in enumerate(orden):
            nombre = names.get(d, (f"#{d}", ""))[0]
            try:
                res = self._open_device(pa, d, nombre)
                if k > 0:
                    log(f"Se usó un micrófono alternativo: {nombre}")
                    self._put("status", f"Usando: {nombre}")
                return res
            except Exception as ex:
                if primer_error is None:
                    primer_error = f"No se pudo abrir '{nombre}': {ex}"
        raise RuntimeError(primer_error or "No se pudo abrir ningún micrófono.")

    @staticmethod
    def _rms(data):
        a = array("h")
        a.frombytes(data)
        a = a[::4]                       # submuestreo: basta para medir volumen
        if not len(a):
            return 0.0
        return (sum(x * x for x in a) / len(a)) ** 0.5

    @staticmethod
    def _read(stream, ch):
        data = stream.read(CHUNK, exception_on_overflow=False)
        if ch == 2:                      # estéreo -> mono (canal izquierdo)
            a = array("h")
            a.frombytes(data)
            data = a[0::2].tobytes()
        return data

    # ---- bucle principal ----
    def run(self):
        pa = stream = None
        try:
            if pyaudio is None:
                raise RuntimeError("PyAudio no está disponible en este programa.")
            pa = pyaudio.PyAudio()
            stream, rate, ch = self._open(pa)
            self._put("status", "Escuchando…")
            self._capture(stream, rate, ch)
        except Exception as ex:
            log("Error del micrófono:\n" + traceback.format_exc())
            msg = f"Error al escuchar: {ex}"
            if "-9999" in str(ex) or "host error" in str(ex).lower():
                msg += ("\n\nPosible causa: Windows no deja usar el micrófono. En ESTE equipo "
                        "abre Configuración > Privacidad > Micrófono y activa el acceso, "
                        "incluido 'Permitir que las aplicaciones de escritorio accedan a "
                        "tu micrófono'.")
            self._put("error", msg)
            self._put("stopped", self.gen)
        finally:
            try:
                if stream is not None:
                    try:
                        stream.stop_stream()
                    except Exception:
                        pass
                    stream.close()
            except Exception:
                pass
            try:
                if pa is not None:
                    pa.terminate()
            except Exception:
                pass

    def _capture(self, stream, rate, ch):
        chunk_s = CHUNK / float(rate)
        pre = deque(maxlen=max(1, int(0.3 / chunk_s)))     # 0.3 s antes de hablar
        end_silence = int(getattr(self.app, "fin_frase", 0.9) / chunk_s)   # pausa que cierra la frase
        max_frames = int(15 / chunk_s)                     # frase máx. 15 s
        min_speech = max(1, int(0.25 / chunk_s))           # mínimo de voz para enviar

        # calibración del ruido de fondo (0.5 s)
        vals = []
        for _ in range(max(1, int(0.5 / chunk_s))):
            if self.halt_event.is_set():
                return
            vals.append(self._rms(self._read(stream, ch)))
        noise = sum(vals) / len(vals)
        log(f"Ruido de fondo medido: {noise:.0f}")

        frames, speaking, silence, voiced = [], False, 0, 0
        while not self.halt_event.is_set():
            data = self._read(stream, ch)
            level = self._rms(data)
            threshold = max(150.0, noise * 3.0)
            if not speaking:
                if level > threshold:
                    speaking, silence, voiced = True, 0, 1
                    frames = list(pre) + [data]
                    self._put("status", "Te escucho…")
                else:
                    noise = noise * 0.95 + level * 0.05
                    pre.append(data)
            else:
                frames.append(data)
                if level > threshold:
                    voiced += 1
                    silence = 0
                else:
                    silence += 1
                if silence >= end_silence or len(frames) >= max_frames:
                    if voiced >= min_speech:
                        self.app.phrase_q.put((b"".join(frames), rate, self.gen))
                    speaking, frames = False, []
                    pre.clear()
                    self._put("status", "Escuchando…")

        # al detener, no perder la frase que estaba a medias
        if speaking and voiced >= min_speech:
            self.app.phrase_q.put((b"".join(frames), rate, self.gen))


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=TRANSPARENT)
        try:
            self.root.attributes("-transparentcolor", TRANSPARENT)
        except tk.TclError:
            pass

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{W}x{H}+{sw - W - 40}+{sh - H - 80}")

        self.canvas = tk.Canvas(self.root, width=W, height=H,
                                bg=TRANSPARENT, highlightthickness=0)
        self.canvas.pack()
        self.img_robot = tk.PhotoImage(data=ROBOT_BTN_B64)
        # halo rojo que pulsa detrás del robot mientras escucha (oculto en reposo)
        self.ring = self.canvas.create_oval(2, 8, W - 2, H - 2, outline="", fill="")
        self.icon = self.canvas.create_image(W // 2, H // 2 + 2, image=self.img_robot)

        # Estado
        self.listening = False
        self.listener = None
        self.gen = 0           # número de sesión de escucha
        self.lang_code = "es-CO"
        self.phrase_q = queue.Queue()
        self.queue = queue.Queue()
        self.pulse = 0
        self.language = tk.StringVar(value="es-CO")
        self.auto_type = tk.BooleanVar(value=True)
        self.solo_mayus = tk.BooleanVar(value=True)
        self.last_fg = None
        self.mic_index = tk.IntVar(value=-1)
        self.drag = {"x": 0, "y": 0, "moved": False}

        self.recognizer = sr.Recognizer()
        self.language.trace_add("write", lambda *a: setattr(self, "lang_code", self.language.get()))
        # motor de voz: sin internet (Vosk) o Google
        self.recognizer.operation_timeout = 6          # no esperar 20 s si no hay internet
        self.motor = tk.StringVar(value="auto")
        self.motor_sel = "auto"
        self.motor.trace_add("write", lambda *a: setattr(self, "motor_sel", self.motor.get()))
        self.vosk_model = None
        self.whisper = None
        self.whisper_estado = "no"        # no | cargando | listo | error
        self.prompt_voz = ""
        self.fin_frase = 0.9
        threading.Thread(target=self._cargar_modelo, daemon=True).start()
        threading.Thread(target=self._cargar_whisper, daemon=True).start()
        # búsqueda por similitud
        self.buscador = None
        self.usar_similitud = tk.BooleanVar(value=True)
        self.sensibilidad = tk.StringVar(value="normal")
        self.sim_on, self.sim_params = True, SENSIBILIDAD["normal"]
        self.usar_similitud.trace_add("write", lambda *a: setattr(self, "sim_on", self.usar_similitud.get()))
        self.sensibilidad.trace_add("write", self._sens_changed)
        threading.Thread(target=self._cargar_base, daemon=True).start()
        threading.Thread(target=self._phrase_worker, daemon=True).start()

        try:
            self.app_icon = tk.PhotoImage(data=APP_ICON_B64)
            self.root.iconphoto(True, self.app_icon)
        except tk.TclError:
            pass
        self._build_text_window()
        self._build_menu()
        self._bind_events()
        self._no_activate()

        try:
            nombres = [n for _, n in input_devices()]
            log("Microfonos detectados: " + " | ".join(nombres))
            if not nombres:
                self.report("No se detectó ningún micrófono en este equipo.")
        except Exception as ex:
            log("No se pudo listar microfonos:\n" + traceback.format_exc())
            self.report(f"Problema con el audio (PyAudio): {ex}")

        self.root.after(100, self._process_queue)
        self.root.after(80, self._animate)

    # ---------- Interfaz ----------
    def _build_text_window(self):
        self.win = tk.Toplevel(self.root)
        self.win.title("Transcripción · v9")
        self.win.geometry("420x260")
        self.win.attributes("-topmost", True)
        self.win.protocol("WM_DELETE_WINDOW", self.win.withdraw)

        self.text = scrolledtext.ScrolledText(self.win, wrap="word",
                                              font=("Segoe UI", 11))
        self.text.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        bar = tk.Frame(self.win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(bar, text="Copiar todo", command=self.copy_all).pack(side="left")
        tk.Button(bar, text="Limpiar", command=lambda: self.text.delete("1.0", "end")
                  ).pack(side="left", padx=6)
        self.status = tk.Label(bar, text="Listo", fg="gray")
        self.status.pack(side="right")

    def _build_menu(self):
        self.menu = tk.Menu(self.root, tearoff=0)
        lang = tk.Menu(self.menu, tearoff=0)
        for nombre, code in IDIOMAS.items():
            lang.add_radiobutton(label=nombre, value=code, variable=self.language)
        self.menu.add_cascade(label="Idioma", menu=lang)
        self.mic_menu = tk.Menu(self.menu, tearoff=0, postcommand=self._refresh_mics)
        self.menu.add_cascade(label="Micrófono", menu=self.mic_menu)
        self.menu.add_checkbutton(label="Solo MAYÚSCULAS (sin tildes ni signos)",
                                  variable=self.solo_mayus)
        self.menu.add_checkbutton(label="Escribir en cualquier aplicación (donde esté el cursor)",
                                  variable=self.auto_type)
        motores = tk.Menu(self.menu, tearoff=0)
        for clave, etiqueta in (("auto", "Automático (el más preciso disponible)"),
                                ("whisper", "Whisper (preciso, sin internet)"),
                                ("offline", "Vosk (rápido, menos preciso, sin internet)"),
                                ("google", "Google (necesita internet)")):
            motores.add_radiobutton(label=etiqueta, value=clave, variable=self.motor)
        self.menu.add_cascade(label="Motor de voz", menu=motores)
        self.menu.add_checkbutton(label="Búsqueda inteligente por similitud (autorrellenar)",
                                  variable=self.usar_similitud)
        sens = tk.Menu(self.menu, tearoff=0)
        for clave, etiqueta in (("estricta", "Estricta (solo casi idéntico)"),
                                ("normal", "Normal (recomendada)"),
                                ("flexible", "Flexible (siempre el más parecido)")):
            sens.add_radiobutton(label=etiqueta, value=clave, variable=self.sensibilidad)
        self.menu.add_cascade(label="Sensibilidad de la búsqueda", menu=sens)
        self.menu.add_command(label="Recargar base de datos", command=lambda: self._recargar_base())
        self.menu.add_command(label="Cargar otra base de datos…", command=self._elegir_base)
        self.menu.add_command(label="Mostrar transcripción", command=self.win.deiconify)
        self.menu.add_separator()
        self.menu.add_command(label="Salir", command=self.quit)

    def _refresh_mics(self):
        self.mic_menu.delete(0, "end")
        self.mic_menu.add_radiobutton(label="Automático", value=-1,
                                      variable=self.mic_index, command=self._mic_changed)
        devs = input_devices()
        if not devs:
            self.mic_menu.add_command(label="(no se detectó ningún micrófono)",
                                      state="disabled")
        for idx, nombre in devs:
            self.mic_menu.add_radiobutton(label=nombre[:60], value=idx,
                                          variable=self.mic_index,
                                          command=self._mic_changed)

    def _mic_changed(self):
        if self.listening:
            self.stop()
        self.set_status("Micrófono cambiado")

    def _bind_events(self):
        c = self.canvas
        c.bind("<ButtonPress-1>", self._press)
        c.bind("<B1-Motion>", self._motion)
        c.bind("<ButtonRelease-1>", self._release)
        c.bind("<Button-3>", lambda e: self.menu.tk_popup(e.x_root, e.y_root))
        c.bind("<Button-2>", lambda e: self.menu.tk_popup(e.x_root, e.y_root))  # macOS

    def _no_activate(self):
        """En Windows el botón no roba el foco y recordamos la última ventana activa."""
        if not IS_WIN:
            return
        try:
            self.root.update_idletasks()
            GWL_EXSTYLE, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW = -20, 0x08000000, 0x80
            hwnd = user32.GetAncestor(self.root.winfo_id(), 2) or self.root.winfo_id()
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                                  style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW)
        except Exception:
            pass
        self._track_foreground()

    def _own_hwnds(self):
        out = set()
        for w in (self.root, self.win):
            try:
                out.add(user32.GetAncestor(w.winfo_id(), 2))
                out.add(w.winfo_id())
            except Exception:
                pass
        return out

    def _track_foreground(self):
        try:
            hwnd = user32.GetForegroundWindow()
            if hwnd and hwnd not in self._own_hwnds():
                self.last_fg = hwnd
        except Exception:
            pass
        self.root.after(150, self._track_foreground)

    def _fg_class(self):
        try:
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(user32.GetForegroundWindow(), buf, 256)
            return buf.value
        except Exception:
            return ""

    def _restore_focus(self):
        """Si el foco quedó en nuestra app, lo devuelve a la ventana donde escribías."""
        try:
            fg = user32.GetForegroundWindow()
            if self.last_fg and fg in self._own_hwnds():
                user32.SetForegroundWindow(self.last_fg)
                time.sleep(0.12)
        except Exception:
            pass

    # ---------- Arrastrar / clic ----------
    def _press(self, e):
        self.drag.update(x=e.x, y=e.y, moved=False)

    def _motion(self, e):
        dx, dy = e.x - self.drag["x"], e.y - self.drag["y"]
        if abs(dx) + abs(dy) > 4:
            self.drag["moved"] = True
        if self.drag["moved"]:
            x = self.root.winfo_x() + dx
            y = self.root.winfo_y() + dy
            self.root.geometry(f"+{x}+{y}")

    def _release(self, _e):
        if not self.drag["moved"]:
            self.toggle()

    # ---------- Reconocimiento ----------
    def toggle(self):
        if self.listening:
            self.stop()
        else:
            self.start()

    def start(self):
        if self.listening:
            return
        if pyaudio is None:
            self.report("PyAudio no está instalado en este programa.")
            return
        old = self.listener
        if old is not None and old.is_alive():       # todavía cerrándose
            old.stop()
            old.join(1.5)
        try:
            self.fin_frase = {"whisper": 0.8, "offline": 0.7}.get(self._motor_efectivo(), 0.9)
        except Exception:
            self.fin_frase = 0.9
        self.gen += 1
        idx = self.mic_index.get()
        self.listening = True
        self.set_status("Iniciando micrófono…")
        self.listener = Listener(self, None if idx < 0 else idx, self.gen)
        self.listener.start()

    def stop(self):
        self.listening = False
        if self.listener is not None:
            self.listener.stop()
        self.canvas.itemconfig(self.ring, fill="")
        self.canvas.coords(self.icon, W // 2, H // 2 + 2)
        self.set_status("Detenido")

    # ---------- base de datos / similitud ----------
    def _sens_changed(self, *_):
        self.sim_params = SENSIBILIDAD.get(self.sensibilidad.get(), SENSIBILIDAD["normal"])

    def _cargar_base(self, ruta=None):
        """Carga la base y construye el índice (en un hilo, sin congelar la ventana)."""
        rutas = [ruta] if ruta else candidatos_base()
        ultimo = None
        for r in rutas:
            if not os.path.isfile(r):
                continue
            try:
                prods = leer_productos(r)
                if not prods:
                    raise RuntimeError("el archivo no tiene productos en la columna A")
                self.buscador = Buscador(prods, leer_alias(ruta_alias()) if ruta_alias() else None)
                self.prompt_voz = construir_prompt(self.buscador, leer_vocabulario())
                log(f"Base cargada: {len(self.buscador.items)} productos de {r}")
                self.queue.put(("base", (len(self.buscador.items), r)))
                return
            except Exception as ex:
                ultimo = ex
                log(f"No se pudo leer la base {r}:\n" + traceback.format_exc())
        detalle = f" Último error: {ultimo}" if ultimo else ""
        self.queue.put(("error", "No se encontró o no se pudo leer la base de datos "
                                 "(similitud.xlsx). La búsqueda por similitud queda desactivada." + detalle))

    def _recargar_base(self, ruta=None):
        self.set_status("Cargando base de datos…")
        threading.Thread(target=self._cargar_base, args=(ruta,), daemon=True).start()

    def _elegir_base(self):
        ruta = filedialog.askopenfilename(
            parent=self.win, title="Elegir base de datos (columna A = productos)",
            filetypes=[("Excel o texto", "*.xlsx *.xlsm *.txt *.csv"), ("Todos", "*.*")])
        if ruta:
            self._recargar_base(ruta)

    # ---------- motor de voz ----------
    def _cargar_modelo(self):
        """Carga el modelo de voz sin internet (si existe la carpeta modelo_voz)."""
        try:
            ruta = buscar_modelo()
            if not ruta:
                self.queue.put(("voz", "No se encontró el modelo sin internet (carpeta modelo_voz). "
                                       "Se usará Google, que necesita internet."))
                return
            import vosk
            try:
                vosk.SetLogLevel(-1)
            except Exception:
                pass
            self.vosk_model = vosk.Model(ruta)
            log(f"Modelo de voz cargado: {ruta}")
            self.queue.put(("voz", "Motor SIN INTERNET listo (transcribe al instante)."))
        except Exception as ex:
            log("No se pudo cargar el modelo de voz:\n" + traceback.format_exc())
            self.queue.put(("voz", f"No se pudo cargar el modelo sin internet: {ex}. "
                                   "Se usará Google, que necesita internet."))

    def _cargar_whisper(self):
        """Whisper (faster-whisper): mucho más preciso con marcas y nombres raros."""
        ruta = buscar_modelo_whisper()
        if not ruta:
            return
        self.whisper_estado = "cargando"
        try:
            from faster_whisper import WhisperModel
            import numpy as np
            hilos = max(2, min(4, os.cpu_count() or 2))
            self.whisper = WhisperModel(ruta, device="cpu", compute_type="int8",
                                        cpu_threads=hilos, local_files_only=True)
            t0 = time.time()
            segs, _ = self.whisper.transcribe(np.zeros(8000, dtype=np.float32), language="es",
                                              beam_size=1, without_timestamps=True)
            list(segs)                                   # calentamiento: la 1.ª vez tarda más
            self.whisper_estado = "listo"
            log(f"Whisper cargado: {ruta} | calentamiento {time.time() - t0:.1f}s | hilos {hilos}")
            self.queue.put(("voz", "Motor WHISPER listo (el más preciso)."))
        except Exception as ex:
            self.whisper_estado = "error"
            self.whisper = None
            log("No se pudo cargar Whisper:\n" + traceback.format_exc())
            self.queue.put(("voz", f"No se pudo cargar Whisper: {ex}. Se usará el siguiente motor."))

    def _motor_efectivo(self):
        if self.motor_sel == "google":
            return "google"
        if self.motor_sel == "whisper":
            if self.whisper_estado == "cargando":
                raise RuntimeError("Whisper todavía se está cargando (unos segundos). Intenta de nuevo.")
            if self.whisper is None:
                raise RuntimeError("Whisper no está disponible (falta la carpeta modelo_whisper). "
                                   "Elige otro motor o ejecuta descargar_modelo.py.")
            return "whisper"
        if self.motor_sel == "offline":
            if self.vosk_model is None:
                raise RuntimeError("El motor sin internet no está disponible (falta o no cargó "
                                   "la carpeta modelo_voz). Elige 'Google'.")
            return "offline"
        if self.whisper is not None:
            return "whisper"
        return "offline" if self.vosk_model is not None else "google"

    def _transcribir_whisper(self, raw, rate):
        import numpy as np
        a = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
        if rate != 16000 and len(a) > 1:
            ancho = max(1, int(round(rate / 16000.0)))      # filtro suave antes de bajar la frecuencia
            if ancho > 1:
                a = np.convolve(a, np.ones(ancho, dtype=np.float32) / ancho, mode="same")
            n = int(len(a) * 16000 / rate)
            a = np.interp(np.linspace(0, len(a) - 1, n), np.arange(len(a)), a).astype(np.float32)
        pico = float(np.abs(a).max()) if len(a) else 0.0
        if 0.002 < pico < 0.5:                               # voz baja: subir volumen
            a = a * (0.6 / pico)
        segs, _info = self.whisper.transcribe(
            a, language="es", task="transcribe", beam_size=3, best_of=1, temperature=0.0,
            initial_prompt=self.prompt_voz or None, condition_on_previous_text=False,
            without_timestamps=True, vad_filter=False, no_speech_threshold=0.6,
            log_prob_threshold=-1.0, compression_ratio_threshold=2.4)
        partes = []
        for sg in segs:
            if sg.no_speech_prob > 0.6 and sg.avg_logprob < -0.8:
                continue
            if sg.avg_logprob < -1.5:
                continue
            partes.append(sg.text)
        texto = " ".join(partes).strip()
        if not texto or es_alucinacion(texto):
            raise sr.UnknownValueError()
        return texto

    def _transcribir(self, raw, rate):
        """Convierte el audio de una frase en texto con el motor elegido."""
        motor = self._motor_efectivo()
        self._ultimo_motor = motor
        if motor == "whisper":
            return self._transcribir_whisper(raw, rate)
        if motor == "offline":
            import vosk
            rec = vosk.KaldiRecognizer(self.vosk_model, 16000)
            rec.AcceptWaveform(a_16k(raw, rate))
            texto = json.loads(rec.FinalResult()).get("text", "").strip()
            if not texto:
                raise sr.UnknownValueError()
            return texto
        audio = sr.AudioData(raw, rate, 2)
        return self.recognizer.recognize_google(audio, language=self.lang_code)

    @staticmethod
    def _pista_red(ex):
        t = str(ex).lower()
        if "10060" in t or "timed out" in t or "timeout" in t or "10061" in t or "11001" in t:
            return ("\n\nEste equipo no logra conectarse al servicio de voz de Google. "
                    "Solución: usa el motor SIN internet (carpeta modelo_voz; "
                    "se instala con CREAR_EXE.bat o descargar_modelo.py).")
        return ""

    def _phrase_worker(self):
        """Transcribe las frases en orden, en un hilo aparte."""
        while True:
            raw, rate, gen = self.phrase_q.get()
            if gen != self.gen:
                continue                              # frase de una sesión anterior
            self.queue.put(("status", "Transcribiendo…"))
            try:
                t0 = time.time()
                texto = self._transcribir(raw, rate)
                dt = time.time() - t0
                res = None
                if self.sim_on:
                    res = decidir_coincidencia(self.buscador, numeros_a_cifras(texto), self.sim_params)
                self.queue.put(("text", (texto, res, getattr(self, "_ultimo_motor", ""), dt)))
            except sr.UnknownValueError:
                self.queue.put(("status", "Escuchando… (no se entendió)"))
            except sr.RequestError as ex:
                log(f"RequestError: {ex}")
                self.queue.put(("error", f"Sin conexión al servicio de voz: {ex}" + self._pista_red(ex)))
            except Exception as ex:
                log("Error al transcribir:\n" + traceback.format_exc())
                self.queue.put(("error", f"Error al transcribir: {ex}" + self._pista_red(ex)))

    def _process_queue(self):
        try:
            while True:
                kind, data = self.queue.get_nowait()
                if kind == "text":
                    self.add_text(*data)
                    self.set_status("Escuchando…" if self.listening else "Listo")
                elif kind == "voz":
                    self.text.insert("end", f"[VOZ] {data}\n")
                    self.text.see("end")
                elif kind == "base":
                    n, origen = data
                    self.text.insert("end", f"[BASE DE DATOS] {n} productos cargados "
                                            f"({os.path.basename(origen)})\n")
                    self.text.see("end")
                    self.set_status(f"Base lista: {n} productos")
                elif kind == "status":
                    self.set_status(data)
                elif kind == "error":
                    self.report(data)
                elif kind == "stopped":
                    if data is None or data == self.gen:
                        self.stop()
        except queue.Empty:
            pass
        self.root.after(100, self._process_queue)

    # ---------- Salida de texto ----------
    def add_text(self, texto, res=None, motor="", dt=0.0):
        texto = texto.strip()
        oido = limpiar_texto(texto)
        if res and res["modo"] == "auto":
            # Autorrelleno: se escribe el nombre EXACTO de la base de datos
            final = res["texto"]
            self.text.insert("end", f"\n[BASE {int(res['score'] * 100)}%] {final}\n   (oí: {oido})\n")
        else:
            # Sin producto único: se escribe lo que dijiste con las palabras como están en la base
            # (ej. "cuaderno kinota" -> "CUAD QNOTA"); lo que no reconoce queda como se oyó.
            # (solo si hay algún producto plausible; con charla cualquiera se escribe tal cual se oyó)
            base = (res.get("corregido") if res and res["modo"] == "ambiguo" else "") or texto
            if self.solo_mayus.get():
                base = limpiar_texto(base)
            elif base:
                base = base[0].upper() + base[1:]
            if not base:
                return
            final = base
            nota = f"\n   (oí: {oido})" if oido and oido != limpiar_texto(final) else ""
            if res and res["modo"] == "ambiguo":
                self.text.insert("end", f"\n[VARIAS COINCIDENCIAS] {final}{nota}\n   ¿Será?  "
                                        + "  |  ".join(a.strip() for a in res["alts"][:4]) + "\n")
            elif res:
                self.text.insert("end", f"\n[SIN COINCIDENCIA] {final}{nota}\n")
            else:
                self.text.insert("end", final + " ")
        self.text.see("end")
        registrar(motor, dt, oido, res, final)
        if PASTE_OK:
            try:
                pyperclip.copy(final + " ")
            except Exception:
                pass
        if self.auto_type.get():
            self.type_anywhere(final + " ")

    def type_anywhere(self, texto):
        try:
            if IS_WIN:
                self._restore_focus()
                cls = self._fg_class()
                if cls in REMOTE_CLASSES and PASTE_OK:
                    # Escritorio remoto: copiar y pegar (usa el portapapeles compartido)
                    log(f"Pegando en escritorio remoto ({cls})")
                    pyperclip.copy(texto)
                    time.sleep(0.35)          # da tiempo a que el portapapeles se sincronice
                    pyautogui.hotkey("ctrl", "v")
                else:
                    win_type_unicode(texto)
            elif PASTE_OK:
                pyperclip.copy(texto)
                mod = "command" if sys.platform == "darwin" else "ctrl"
                pyautogui.hotkey(mod, "v")
        except Exception as ex:
            self.set_status(f"No se pudo escribir: {ex}")

    def copy_all(self):
        contenido = self.text.get("1.0", "end").strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(contenido)
        self.set_status("Copiado al portapapeles")

    def report(self, msg):
        """Muestra un error en la ventana de transcripción y lo guarda en el log."""
        log(msg)
        self.set_status(msg)
        self.text.insert("end", f"\n[!] {msg}\n")
        self.text.see("end")
        self.win.deiconify()

    def set_status(self, msg):
        self.status.config(text=msg)

    # ---------- Animación ----------
    def _animate(self):
        if self.listening:
            self.pulse = (self.pulse + 1) % 20
            grow = abs(10 - self.pulse) / 10  # 0..1
            pad = 2 + int(8 * (1 - grow))
            self.canvas.coords(self.ring, pad, pad + 6, W - pad, H - pad)
            self.canvas.itemconfig(self.ring, fill=RING_COLOR)
            # el robot "salta" suavemente mientras escucha
            self.canvas.coords(self.icon, W // 2, H // 2 + 2 - int(5 * grow))
        self.root.after(80, self._animate)

    def quit(self):
        self.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
