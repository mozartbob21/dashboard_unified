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


def _vosk_model():
    """Загружает русскую модель Vosk из models/vosk-ru (офлайн, локально)."""
    global _VOSK_MODEL
    if _VOSK_MODEL is not None:
        return _VOSK_MODEL
    from pathlib import Path as _P
    mdir = _P(__file__).resolve().parents[2] / "models" / "vosk-ru"
    if not mdir.exists():
        print("[aichat] vosk model dir not found:", mdir)
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
    """Локальное офлайн-распознавание речи (Vosk) — без интернета и HTTPS."""
    import io
    import json
    import wave
    model = _vosk_model()
    if model is None:
        print("[aichat] vosk model not loaded")
        return ""
    try:
        wf = wave.open(io.BytesIO(data), "rb")
    except Exception as e:
        print("[aichat] not a WAV file:", e)
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
        text = " ".join(p for p in parts if p).strip()
        print("[aichat] transcribed:", text[:80])
        return text
    except Exception as e:
        print("[aichat] transcribe error:", e)
        return ""
    finally:
        try:
            wf.close()
        except Exception:
            pass
