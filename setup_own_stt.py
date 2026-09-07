#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Своё офлайн-распознавание речи (Vosk) вместо недоступного эндпоинта платформы."""
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent
MODELS = BASE / "models"
MDIR = MODELS / "vosk-ru"
URL = "https://alphacephei.com/vosk/models/vosk-model-small-ru-0.22.zip"
ZP = MODELS / "vosk-ru.zip"

# ── 1) ставим vosk в текущее окружение ──
print("📦 Ставлю vosk...")
subprocess.run([sys.executable, "-m", "pip", "install", "vosk"], check=False)

# ── 2) качаем и распаковываем русскую модель ──
if not MDIR.exists():
    MODELS.mkdir(parents=True, exist_ok=True)
    print("⬇️  Скачиваю русскую модель Vosk (~45 МБ)...")
    urllib.request.urlretrieve(URL, ZP)
    print("📂 Распаковываю...")
    with zipfile.ZipFile(ZP) as z:
        z.extractall(MODELS)
    cand = [p for p in MODELS.iterdir() if p.is_dir() and p.name.startswith("vosk-model")]
    if cand:
        if MDIR.exists():
            shutil.rmtree(MDIR)
        cand[0].rename(MDIR)
    ZP.unlink(missing_ok=True)
print("🧠 Модель:", MDIR, "→", "OK" if MDIR.exists() else "НЕ НАЙДЕНА")

# ── 3) переключаем transcribe() в engine.py на Vosk ──
E = BASE / "services" / "aichat" / "engine.py"
src = E.read_text(encoding="utf-8")

NEW = '''_VOSK_MODEL = None


def _vosk_model():
    global _VOSK_MODEL
    if _VOSK_MODEL is not None:
        return _VOSK_MODEL
    from pathlib import Path as _P
    mdir = _P(__file__).resolve().parents[2] / "models" / "vosk-ru"
    if not mdir.exists():
        return None
    try:
        from vosk import Model
        _VOSK_MODEL = Model(str(mdir))
        print("[aichat] vosk model loaded:", mdir)
        return _VOSK_MODEL
    except Exception as e:
        print("[aichat] vosk model load error:", e)
        return None


def transcribe(data: bytes, mime: str = "audio/webm") -> str:
    """Локальное офлайн-распознавание речи (Vosk) — как диктовка на телефоне."""
    import io
    import json
    import wave
    model = _vosk_model()
    if model is None:
        print("[aichat] vosk model not found at models/vosk-ru")
        return ""
    try:
        wf = wave.open(io.BytesIO(data), "rb")
    except Exception:
        # не WAV (например webm из браузера) — vosk не читает
        return ""
    try:
        from vosk import KaldiRecognizer
        rec = KaldiRecognizer(model, wf.getframerate())
        parts = []
        while True:
            chunk = wf.readframes(4000)
            if not chunk:
                break
            if rec.AcceptWaveform(chunk):
                parts.append(json.loads(rec.Result()).get("text", ""))
        parts.append(json.loads(rec.FinalResult()).get("text", ""))
        return " ".join(p for p in parts if p).strip()
    except Exception as e:
        print("[aichat] transcribe error:", e)
        return ""
    finally:
        try:
            wf.close()
        except Exception:
            pass
'''

idx = src.find("def transcribe(")
if idx != -1:
    src = src[:idx].rstrip() + "\n\n\n"
else:
    src = src.rstrip() + "\n\n\n"
src += NEW
E.write_text(src, encoding="utf-8")
print("✅ engine.py: transcribe() теперь на локальном Vosk")
print("\nГотово. Рестарт сервера:")
print("   taskkill /IM python.exe /F")
print("   python -m uvicorn app:app --host 0.0.0.0 --port 8000")