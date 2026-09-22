"""Explicit AI destinations: approved GosChat or loopback-only local inference."""
import ipaddress
from urllib.parse import urlsplit


class PrivacyError(RuntimeError):
    pass


def ai_endpoint(base):
    value = str(base or '').strip().rstrip('/')
    parsed = urlsplit(value)
    if parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise PrivacyError('Недопустимый адрес ИИ-сервиса.')
    host = (parsed.hostname or '').lower()
    if host == 'aiplatform.mosreg.ru' and parsed.scheme == 'https' and parsed.port in (None, 443):
        return value
    try:
        local = ipaddress.ip_address(host).is_loopback
    except ValueError:
        local = host == 'localhost'
    if local and parsed.scheme in ('http', 'https'):
        return value
    raise PrivacyError('Разрешены только ГосЧат и локальная модель на этом сервере. Другие внешние ИИ отключены.')


def account_identity(user):
    return ('kc:' + user['kc_sub']) if user.get('kc_sub') else ('local:' + str(user['id']))


def ai_tls_context():
    """Use OS trust (including office Windows CA roots), never disable verification."""
    import os
    import ssl
    return ssl.create_default_context(cafile=os.getenv('AI_CA_BUNDLE') or None)
