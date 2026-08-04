"""Отправка писем через Яндекс.Почту (SMTP)."""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def _get_config():
    user = (os.getenv("YANDEX_SMTP_USER") or "").strip()
    # Если логин без домена — дополняем до полного адреса
    if user and "@" not in user:
        user += "@yandex.ru"
    return {
        "host": os.getenv("YANDEX_SMTP_HOST", "smtp.yandex.ru"),
        "port": int(os.getenv("YANDEX_SMTP_PORT", "465")),
        "user": user,
        "password": (os.getenv("YANDEX_SMTP_PASSWORD") or "").strip(),
        "from_name": (os.getenv("YANDEX_SMTP_FROM_NAME") or "Neurona II").strip(),
        "from_plain": os.getenv("YANDEX_SMTP_FROM_PLAIN", "0") == "1",
    }


def send_verification_code(to_email: str, code: str) -> tuple:
    cfg = _get_config()

    if not cfg["user"] or not cfg["password"]:
        return False, f"SMTP не настроен (user='{cfg['user']}', pass={'OK' if cfg['password'] else 'ПУСТО'})"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Код подтверждения — Нейрона ИИ"
    # From должен ТОЧНО принадлежать аккаунту
    if cfg["from_plain"]:
        msg["From"] = cfg["user"]                      # только адрес, без имени
    else:
        msg["From"] = f"{cfg['from_name']} <{cfg['user']}>"
    msg["To"] = to_email

    print(f"[mailer] Login: {cfg['user']} | From: {msg['From']} | To: {to_email}")

    text = f"Ваш код подтверждения: {code}\nКод действителен 10 минут."
    html = f"""
    <div style="font-family:system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;max-width:480px;margin:0 auto;
                border:1px solid #e5e7eb;border-radius:16px;overflow:hidden;">
      <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:18px 24px;color:#fff;">
        <div style="font-weight:800;font-size:18px;">🤖 Нейрона ИИ</div>
        <div style="font-size:12px;opacity:.85;">Подмосковная Нейрона Роботовна</div>
      </div>
      <div style="padding:24px;">
        <p style="margin:0 0 12px;color:#374151;font-size:14px;">Ваш код подтверждения регистрации:</p>
        <div style="font-size:32px;font-weight:800;letter-spacing:.35em;color:#4f46e5;
                    background:#f3f4f6;border-radius:12px;padding:14px 18px;text-align:center;">{code}</div>
        <p style="margin:14px 0 0;color:#6b7280;font-size:12px;">
          Код действителен 10 минут. Если вы не запрашивали регистрацию — просто проигнорируйте письмо.
        </p>
      </div>
    </div>
    """
    msg.attach(MIMEText(text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=15) as server:
            server.login(cfg["user"], cfg["password"])
            server.sendmail(cfg["user"], [to_email], msg.as_string())
        print("[mailer] Письмо отправлено успешно")
        return True, ""
    except smtplib.SMTPSenderRefused as e:
        print(f"[mailer] Sender refused: {e.smtp_code} {e.smtp_error}")
        return False, (f"Яндекс отклонил адрес отправителя ({e.smtp_code}). "
                       f"Совет: поставьте YANDEX_SMTP_FROM_PLAIN=1 в .env и проверьте, "
                       f"что YANDEX_SMTP_USER совпадает с ящиком, на котором создан пароль приложения.")
    except smtplib.SMTPAuthenticationError as e:
        print(f"[mailer] Auth error: {e.smtp_code} {e.smtp_error}")
        return False, f"Ошибка авторизации SMTP ({e.smtp_code})."
    except Exception as e:
        print(f"[mailer] Ошибка отправки: {type(e).__name__}: {e}")
        return False, str(e)
        