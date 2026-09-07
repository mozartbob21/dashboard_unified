#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Максимальная точность диктовки: faster-whisper (medium) + нормализация + Vosk как fallback."""
import re
import subprocess
import sys
from pathlib import Path

# ── 1) ставим faster-whisper ──
print("📦 Ставлю faster-whisper...")
subprocess.run([sys.executable, "-m", "pip", "install", "faster-whisper"], check=False)
chk = subprocess.run([sys.executable, "-c", "import faster_whisper, ctranslate2; print('ok')"],
                     capture_output=True, text=True)
print("faster-whisper:", "✅ OK" if chk.returncode == 0 else "❌ не встал -> останется Vosk")

# ── 2) переписываем распознавание в engine.py ──
E = Path("services") / "aichat" / "engine.py"
src = E.read_text(encoding="utf-8")

NEW_BLOCK = '''_VOSK_MODEL = None
_WHISPER = None


def _whisper_model():
    """Whisper (medium int8) — максимальная точность диктовки."""
    global _WHISPER
    if _WHISPER is not None:
        return _WHISPER or None
    import os
    try:
        from faster_whisper import WhisperModel
        size = os.getenv("STT_WHISPER_SIZE", "medium")
        _WHISPER = WhisperModel(size, device="cpu", compute_type="int8")
        print("[aichat] whisper model loaded:", size)
    except Exception as e:
        print("[aichat] whisper unavailable:", e)
        _WHISPER = False
    return _WHISPER or None


def _vosk_model():
    global _VOSK_MODEL
    if _VOSK_MODEL is not None:
        return _VOSK_MODEL
    from pathlib import Path as _P
    base = _P(__file__).resolve().parents[2] / "models"
    mdir = None
    if base.exists():
        dirs = [p for p in base.iterdir() if p.is_dir()]

        def score(p):
            n = p.name.lower()
            if "0.42" in n or "big" in n or "large" in n:
                return 0
            if n == "vosk-ru":
                return 1
            if n.startswith("vosk-"):
                return 2
            return 3

        for c in sorted(dirs, key=score):
            if (c / "conf").exists() and (c / "am").exists():
                mdir = c
                break
    if mdir is None:
        return None
    try:
        from vosk import Model
        _VOSK_MODEL = Model(str(mdir))
        print("[aichat] vosk model loaded:", mdir)
        return _VOSK_MODEL
    except Exception as e:
        print("[aichat] vosk load error:", e)
        return None


def _pcm_to_16k_mono(data: bytes, ch: int, width: int, rate: int) -> bytes:
    import array
    if width == 1:
        samples = [((b - 128) << 8) for b in data]
    elif width == 2:
        a = array.array("h")
        a.frombytes(data)
        samples = list(a)
    else:
        samples = [int.from_bytes(data[i + width - 2:i + width], "little", signed=True)
                   for i in range(0, max(0, len(data) - width + 1), width)]
    if ch > 1:
        samples = samples[::ch]
    if rate and rate != 16000:
        ratio = 16000.0 / rate
        n = int(len(samples) * ratio)
        out = []
        for i in range(n):
            s = i / ratio
            i0 = int(s)
            i1 = min(i0 + 1, len(samples) - 1)
            f = s - i0
            out.append(int(samples[i0] * (1 - f) + samples[i1] * f))
        samples = out
    a = array.array("h", samples)
    return a.tobytes()


def _normalize_pcm(pcm: bytes) -> bytes:
    """Убирает постоянную составляющую и поднимает громкость до целевого пика."""
    import array
    a = array.array("h")
    a.frombytes(pcm)
    if not a:
        return pcm
    probe = a[::53] or array.array("h", [0])
    mean = sum(probe) // max(1, len(probe))
    peak = max(abs(v - mean) for v in a[::7]) or 1
    gain = min(8.0, 24000.0 / peak)
    out = array.array("h", [max(-32000, min(32000, int((v - mean) * gain))) for v in a])
    return out.tobytes()


def transcribe(data: bytes, mime: str = "audio/webm") -> str:
    """Диктовка: сначала Whisper (точно), при неудаче — Vosk (офлайн)."""
    import io
    m = _whisper_model()
    if m is not None:
        try:
            try:
                segments, _info = m.transcribe(io.BytesIO(data), language="ru", vad_filter=True)
            except Exception:
                segments, _info = m.transcribe(io.BytesIO(data), language="ru")
            text = " ".join(s.text for s in segments).strip()
            print("[aichat] whisper transcribed:", text[:120])
            if text:
                return text
        except Exception as e:
            print("[aichat] whisper error:", e)
    model = _vosk_model()
    if model is None:
        return ""
    import json
    import wave
    try:
        wf = wave.open(io.BytesIO(data), "rb")
    except Exception:
        return ""
    ch, width, rate = wf.getnchannels(), wf.getsampwidth(), wf.getframerate()
    raw = wf.readframes(wf.getnframes())
    wf.close()
    pcm = _normalize_pcm(_pcm_to_16k_mono(raw, ch, width, rate))
    try:
        from vosk import KaldiRecognizer
        rec = KaldiRecognizer(model, 16000)
        parts = []
        for i in range(0, len(pcm), 8000):
            if rec.AcceptWaveform(pcm[i:i + 8000]):
                parts.append(json.loads(rec.Result()).get("text", ""))
        parts.append(json.loads(rec.FinalResult()).get("text", ""))
        text = " ".join(p for p in parts if p).strip()
        print("[aichat] vosk transcribed:", text[:120])
        return text
    except Exception as e:
        print("[aichat] vosk error:", e)
        return ""
'''

# вырезаем старые реализации и вставляем новый блок
src = re.sub(r'\n?_VOSK_MODEL\s*=\s*None\n', '\n', src)
src = re.sub(r'\n?_WHISPER\s*=\s*None\n', '\n', src)
src = re.sub(r'def _vosk_model\(\):.*?(?=\ndef |\Z)', '', src, flags=re.S)
src = re.sub(r'def _whisper_model\(\):.*?(?=\ndef |\Z)', '', src, flags=re.S)
src = re.sub(r'def _pcm_to_16k_mono\(.*?(?=\ndef |\Z)', '', src, flags=re.S)
src = re.sub(r'def _normalize_pcm\(.*?(?=\ndef |\Z)', '', src, flags=re.S)
src = re.sub(r'def transcribe\(.*?(?=\ndef |\Z)', '', src, flags=re.S)
src = src.rstrip() + "\n\n\n" + NEW_BLOCK
E.write_text(src, encoding="utf-8")
print("✅ engine.py: распознавание = Whisper(medium) → Vosk(fallback) + нормализация")
print("\nРестарт:")
print("   taskkill /IM python.exe /F")
print("   python -m uvicorn app:app --host 0.0.0.0 --port 8000")