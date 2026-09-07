"""Движок AI-чата: Нейрона ИИ (работает на локальном контуре Подмосковья)."""
from services.summarizer.engine import _qwen_chat

GREETING = ("Здравствуйте! Я — Нейрона ИИ, корпоративный помощник платформы ЖКХ "
            "Московской области. Могу разобрать приложенный файл, подготовить "
            "служебную записку, сверку или отчёт, подсказать по обращениям граждан. "
            "С чего начнём?")

SYSTEM_PROMPT = (
    "Ты — Нейрона ИИ, корпоративный ИИ-ассистент платформы ЖКХ Московской области. "
    "ЭТО ТВОЁ ЕДИНСТВЕННОЕ ИМЯ. Если спрашивают, кто ты, как тебя зовут, "
    "какая ты модель или технология — отвечай: «Я — Нейрона ИИ». "
    "НИКОГДА не упоминай Qwen, Qwen3, Tongyi, Alibaba, OpenAI, GPT, LLaMA, "
    "«языковая модель», «нейросеть такого-то вендора» и любые технические названия моделей.\n"
    "Помогаешь сотрудникам с документами, отчётами, анализом обращений, "
    "служебными записками и рабочими задачами.\n"
    "Правила: отвечай на русском; структурируй ответ (списки, подзаголовки); "
    "если приложены файлы — опирайся на их содержимое; "
    "если данных не хватает — честно скажи; без воды и лишних вступлений."
)

def ask(history, max_tokens=2500):
    """history: список {"role": "user"|"assistant", "content"}."""
    msgs = [{"role": "system", "content": SYSTEM_PROMPT}]
    for m in (history or [])[-12:]:
        msgs.append({"role": m["role"], "content": m["content"]})
    return _qwen_chat(msgs, max_tokens=max_tokens)


_VOSK_MODEL = None
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


def _platform_creds():
    """Base URL и ключ того же контура, что использует чат."""
    import os
    try:
        from services.summarizer import engine as _se
    except Exception:
        _se = None
    base = (getattr(_se, "BASE_URL", None) or getattr(_se, "API_BASE", None)
            or os.getenv("STT_BASE") or os.getenv("QWEN_BASE_URL")
            or "https://aiplatform.mosreg.ru/api/user-models/v1")
    key = (getattr(_se, "API_KEY", None) or getattr(_se, "QWEN_API_KEY", None)
           or os.getenv("QWEN_API_KEY") or os.getenv("AI_API_KEY") or os.getenv("API_KEY") or "")
    return base.rstrip("/"), key


def _asr_qwen(data: bytes) -> str:
    """Распознавание речи моделью Qwen3-ASR (qwen-asr) через chat/completions."""
    import base64
    import httpx
    base, key = _platform_creds()
    b64 = base64.b64encode(data).decode()
    headers = {"Authorization": "Bearer " + key} if key else {}
    variants = [
        {"model": "qwen-asr", "messages": [{"role": "user", "content": [
            {"type": "input_audio", "input_audio": {"data": b64, "format": "wav"}},
            {"type": "text", "text": "Расшифруй речь на русском языке, верни только текст."}]}]},
        {"model": "qwen-asr", "messages": [{"role": "user", "content": [
            {"type": "audio_url", "audio_url": {"url": "data:audio/wav;base64," + b64}},
            {"type": "text", "text": "Расшифруй речь на русском языке, верни только текст."}]}]},
    ]
    for payload in variants:
        try:
            with httpx.Client(timeout=90) as c:
                r = c.post(base + "/chat/completions", json=payload, headers=headers)
                if r.status_code != 200:
                    print("[aichat] qwen-asr http", r.status_code, r.text[:120])
                    continue
                j = r.json()
                txt = (j.get("choices") or [{}])[0].get("message", {}).get("content", "")
                if isinstance(txt, list):
                    txt = "".join(p.get("text", "") for p in txt if isinstance(p, dict))
                txt = (txt or "").strip()
                if txt:
                    return txt
        except Exception as e:
            print("[aichat] qwen-asr variant error:", e)
    return ""


def transcribe(data: bytes, mime: str = "audio/webm") -> str:
    """Цепочка: Qwen3-ASR (контур) -> Whisper (локально) -> Vosk (офлайн)."""
    import io
    try:
        t = _asr_qwen(data)
        if t:
            print("[aichat] qwen-asr transcribed:", t[:120])
            return t
    except Exception as e:
        print("[aichat] qwen-asr error:", e)
    m = _whisper_model()
    if m is not None:
        try:
            try:
                segments, _i = m.transcribe(io.BytesIO(data), language="ru", vad_filter=True)
            except Exception:
                segments, _i = m.transcribe(io.BytesIO(data), language="ru")
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
