"""Descarga los modelos de voz SIN INTERNET (una sola vez) junto a este archivo.

  modelo_whisper  Whisper (faster-whisper)  -> el MAS PRECISO.  small ~480 MB
  modelo_voz      Vosk                      -> respaldo rapido, ~40 MB

Uso:
  python descargar_modelo.py            descarga Whisper 'small' y Vosk
  python descargar_modelo.py base       Whisper 'base' (~145 MB, mas rapido, menos preciso)
  python descargar_modelo.py medium     Whisper 'medium' (~1.5 GB, mas preciso, mas lento)
  python descargar_modelo.py vosk       solo Vosk
"""
import os
import shutil
import sys
import tempfile
import urllib.request
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
DEST_W = os.path.join(AQUI, "modelo_whisper")
DEST_V = os.path.join(AQUI, "modelo_voz")
VOSK_URL = "https://alphacephei.com/vosk/models/vosk-model-small-es-0.42.zip"
HF = "https://huggingface.co/Systran/faster-whisper-{t}/resolve/main/{f}"
ARCHIVOS_W = ["config.json", "preprocessor_config.json", "tokenizer.json", "vocabulary.txt", "model.bin"]


def _bajar(url, destino, opcional=False):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r, open(destino, "wb") as f:
            total = int(r.headers.get("Content-Length") or 0)
            hecho = 0
            while True:
                bloque = r.read(512 * 1024)
                if not bloque:
                    break
                f.write(bloque)
                hecho += len(bloque)
                if total:
                    print(f"\r  {os.path.basename(destino)}: {hecho * 100 // total}% "
                          f"({hecho // 1048576} de {total // 1048576} MB)", end="", flush=True)
        print()
    except Exception:
        if os.path.exists(destino):
            os.remove(destino)
        if not opcional:
            raise


def whisper_ok(ruta):
    return os.path.isfile(os.path.join(ruta, "model.bin")) and os.path.isfile(os.path.join(ruta, "config.json"))


def instalar_whisper(tam="small"):
    if whisper_ok(DEST_W):
        print("Whisper ya esta instalado en", DEST_W)
        return True
    print(f"Descargando Whisper '{tam}' (una sola vez)...")
    tmp = tempfile.mkdtemp(prefix="whisper_")
    try:
        try:                                        # camino 1: la propia libreria
            from faster_whisper.utils import download_model
            download_model(tam, output_dir=tmp)
        except Exception as ex:
            print("  (descarga automatica fallo:", ex, ") probando descarga directa...")
            for f in ARCHIVOS_W:
                _bajar(HF.format(t=tam, f=f), os.path.join(tmp, f), opcional=f in ("preprocessor_config.json",))
        if not whisper_ok(tmp):
            raise RuntimeError("faltan archivos del modelo (model.bin / config.json)")
        if os.path.isdir(DEST_W):
            shutil.rmtree(DEST_W)
        os.makedirs(DEST_W)
        for f in os.listdir(tmp):
            p = os.path.join(tmp, f)
            if os.path.isfile(p):
                shutil.move(p, os.path.join(DEST_W, f))
        print("Listo. Whisper instalado en:", DEST_W)
        return True
    except Exception as ex:
        print("\nNo se pudo descargar Whisper:", ex)
        print("Hazlo a mano (en un equipo con internet): entra a")
        print(f"  https://huggingface.co/Systran/faster-whisper-{tam}/tree/main")
        print("descarga config.json, tokenizer.json, vocabulary.txt y model.bin (y preprocessor_config.json)")
        print("y ponlos en una carpeta llamada  modelo_whisper  junto a este archivo.")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def vosk_ok(ruta):
    return os.path.isdir(os.path.join(ruta, "am")) and os.path.isdir(os.path.join(ruta, "conf"))


def instalar_vosk(origen=None):
    if vosk_ok(DEST_V):
        print("Vosk ya esta instalado en", DEST_V)
        return True
    tmp = tempfile.mkdtemp(prefix="vosk_")
    try:
        zpath = os.path.join(tmp, "modelo.zip")
        print("Descargando Vosk (~40 MB)...")
        if origen and os.path.isfile(origen):
            shutil.copyfile(origen, zpath)
        else:
            _bajar(origen or VOSK_URL, zpath)
        salida = os.path.join(tmp, "x")
        with zipfile.ZipFile(zpath) as zf:
            zf.extractall(salida)
        raiz = next((d for d, _, _ in os.walk(salida) if vosk_ok(d)), None)
        if not raiz:
            raise RuntimeError("el archivo no contiene un modelo valido")
        if os.path.isdir(DEST_V):
            shutil.rmtree(DEST_V)
        shutil.move(raiz, DEST_V)
        print("Listo. Vosk instalado en:", DEST_V)
        return True
    except Exception as ex:
        print("\nNo se pudo descargar Vosk:", ex)
        print("A mano: descarga", VOSK_URL, "descomprime y renombra la carpeta a  modelo_voz")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    arg = sys.argv[1].lower() if len(sys.argv) > 1 else "small"
    if arg.endswith(".zip"):
        return 0 if instalar_vosk(sys.argv[1]) else 1
    if arg == "vosk":
        return 0 if instalar_vosk() else 1
    if arg not in ("base", "small", "medium"):
        print(__doc__)
        return 1
    w = instalar_whisper(arg)
    v = instalar_vosk()
    return 0 if (w or v) else 1


if __name__ == "__main__":
    sys.exit(main())
