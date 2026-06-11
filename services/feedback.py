import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel, Field

from services.auth.security import get_user_from_token

router = APIRouter()

# === НАСТРОЙКИ SMTP ЯНДЕКС ===
# Лучше задавать через переменные окружения, значения ниже — запасные.
YANDEX_SMTP_HOST = os.getenv("FEEDBACK_SMTP_HOST", "smtp.yandex.ru")
YANDEX_SMTP_PORT = int(os.getenv("FEEDBACK_SMTP_PORT", "465"))  # SSL
YANDEX_SMTP_USER = os.getenv("FEEDBACK_SMTP_USER", "ivanotvetov@yandex.ru")
YANDEX_SMTP_PASSWORD = os.getenv("FEEDBACK_SMTP_PASSWORD", "ruenkbzhdbikcoyh")
FEEDBACK_RECIPIENT = os.getenv("FEEDBACK_RECIPIENT", "nejrona_support@rambler.ru")

TOPIC_LABELS = {
    "bug": "🐞 Ошибка",
    "idea": "💡 Идея",
    "question": "❓ Вопрос",
    "other": "📝 Другое",
}


class FeedbackRequest(BaseModel):
    topic: str = Field(..., max_length=50)
    subject: str = Field("Без темы", max_length=200)
    message: str = Field("", max_length=5000)
    contact: Optional[str] = Field(None, max_length=200)
    page_url: Optional[str] = Field(None, max_length=500)
    user_agent: Optional[str] = Field(None, max_length=500)


@router.post("/api/feedback/send")
async def send_feedback(request: Request):
    try:
        raw = await request.json()
    except Exception:
        raw = {}

    if not isinstance(raw, dict):
        raw = {}

    # Поддержка разных имён полей из разных версий формы
    payload = FeedbackRequest(
        topic=str(raw.get("topic") or raw.get("type") or "other"),
        subject=str(raw.get("subject") or raw.get("title") or raw.get("name") or "Без темы").strip() or "Без темы",
        message=str(raw.get("message") or raw.get("text") or raw.get("body") or "").strip(),
        contact=str(raw.get("contact") or raw.get("email") or raw.get("phone") or "") or None,
        page_url=str(raw.get("page_url") or raw.get("url") or "") or None,
        user_agent=str(raw.get("user_agent") or "") or None,
    )

    if len(payload.message) < 2:
        raise HTTPException(status_code=400, detail="Сообщение не может быть пустым")
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}
    username = user.get("username", "Аноним")

    topic_label = TOPIC_LABELS.get(payload.topic, payload.topic)
    subject_full = f"[Нейрона ИИ] {topic_label}: {payload.subject}"

    body_text = f"""Обратная связь из Нейрона ИИ

Тип: {topic_label}
Тема: {payload.subject}
Пользователь: {username}
Контакт: {payload.contact or '—'}
Страница: {payload.page_url or '—'}
Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}

Сообщение:
{payload.message}

---
User-Agent: {payload.user_agent or '—'}
"""

    body_html = f"""
    <html><body style="font-family: Arial, sans-serif; color:#1e293b;">
        <h2 style="color:#667eea;">Обратная связь из Нейрона ИИ</h2>
        <table style="border-collapse:collapse; width:100%; max-width:600px;">
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Тип</b></td><td style="padding:8px;">{topic_label}</td></tr>
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Тема</b></td><td style="padding:8px;">{payload.subject}</td></tr>
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Пользователь</b></td><td style="padding:8px;">{username}</td></tr>
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Контакт</b></td><td style="padding:8px;">{payload.contact or '—'}</td></tr>
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Страница</b></td><td style="padding:8px;">{payload.page_url or '—'}</td></tr>
            <tr><td style="padding:8px; background:#f1f5f9;"><b>Дата</b></td><td style="padding:8px;">{datetime.now().strftime('%d.%m.%Y %H:%M:%S')}</td></tr>
        </table>
        <h3 style="margin-top:24px;">Сообщение:</h3>
        <div style="padding:14px; background:#f8fafc; border-left:4px solid #667eea; white-space:pre-wrap;">{payload.message}</div>
        <p style="font-size:11px; color:#94a3b8; margin-top:24px;">User-Agent: {payload.user_agent or '—'}</p>
    </body></html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject_full
    msg["From"] = YANDEX_SMTP_USER
    msg["To"] = FEEDBACK_RECIPIENT
    msg["Date"] = formatdate(localtime=True)
    if payload.contact and "@" in payload.contact:
        msg["Reply-To"] = payload.contact

    msg.attach(MIMEText(body_text, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(YANDEX_SMTP_HOST, YANDEX_SMTP_PORT, timeout=20) as server:
            server.login(YANDEX_SMTP_USER, YANDEX_SMTP_PASSWORD)
            server.sendmail(YANDEX_SMTP_USER, [FEEDBACK_RECIPIENT], msg.as_string())
    except smtplib.SMTPAuthenticationError:
        raise HTTPException(
            status_code=500,
            detail="Ошибка авторизации SMTP. Проверьте логин и пароль приложения Яндекс.",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка отправки: {str(e)}")

    return {"ok": True, "status": "ok", "message": "Сообщение отправлено"}
