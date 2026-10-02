"""Safe AI failure codes: no prompts, credentials or response bodies in errors."""
import ssl

import httpx

from core.privacy import PrivacyError


class AIServiceError(RuntimeError):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


MESSAGES = {
    'GIGACHAT_CREDENTIALS_MISSING': 'На сервере не задан ключ GigaChat. Проверьте SUMMARIZER_GIGACHAT_CREDENTIALS или GIGACHAT_CREDENTIALS.',
    'GIGACHAT_SCOPE_INVALID': 'Проверьте GIGACHAT_SCOPE на сервере: допустимы GIGACHAT_API_PERS, GIGACHAT_API_B2B или GIGACHAT_API_CORP.',
    'AI_KEY_MISSING': 'На сервере не задан ключ доступа к ИИ. Администратору нужно проверить QWEN_API_KEY и перезапустить приложение.',
    'AI_AUTH_FAILED': 'ИИ отклонил ключ доступа. Администратору нужно проверить ключ на сервере.',
    'AI_ACCESS_DENIED': 'Настроенной учётной записи не разрешён доступ к выбранной модели ИИ.',
    'AI_MODEL_NOT_FOUND': 'Настроенная модель или адрес ИИ не найдены. Администратору нужно проверить настройки подключения.',
    'AI_CONTEXT_LIMIT': 'Запрос превышает объём, который принимает модель. Начните новый диалог или уменьшите объём приложенных данных.',
    'AI_REQUEST_REJECTED': 'Сервис ИИ отклонил запрос. Администратору нужно проверить совместимость модели и настройки подключения.',
    'AI_RATE_LIMIT': 'Сервис ИИ временно ограничил число запросов. Повторите немного позже.',
    'AI_TIMEOUT': 'Сервис ИИ не успел ответить. Повторите запрос позже.',
    'AI_TLS_ERROR': 'Сервер не смог проверить сертификат ИИ-сервиса. Администратору нужно проверить доверенные сертификаты.',
    'AI_CONNECTION_ERROR': 'С компьютера-сервера не удалось подключиться к ИИ. Проверьте доступ к настроенному ИИ-сервису.',
    'AI_ENDPOINT_BLOCKED': 'Адрес ИИ не соответствует настройкам разрешённых сервисов. Администратору нужно проверить адрес подключения.',
    'AI_SERVICE_ERROR': 'Сервис ИИ вернул ошибку. Повторите запрос позже.',
    'AI_EMPTY_RESPONSE': 'Модель не вернула текст ответа. Администратору нужно проверить настройки модели.',
    'AI_INVALID_RESPONSE': 'Сервис ИИ вернул ответ в неподдерживаемом формате.',
    'AI_INTERNAL_ERROR': 'При обработке запроса к ИИ произошла ошибка приложения. Передайте код ошибки администратору.',
}


def describe_ai_error(exc):
    status = None
    if isinstance(exc, AIServiceError):
        code = exc.code if exc.code in MESSAGES else 'AI_INTERNAL_ERROR'
    elif isinstance(exc, PrivacyError):
        code = 'AI_ENDPOINT_BLOCKED'
    elif isinstance(exc, httpx.HTTPStatusError):
        status = exc.response.status_code
        code = {401: 'AI_AUTH_FAILED', 403: 'AI_ACCESS_DENIED', 404: 'AI_MODEL_NOT_FOUND',
                413: 'AI_CONTEXT_LIMIT', 422: 'AI_REQUEST_REJECTED', 429: 'AI_RATE_LIMIT'}.get(status)
        if status == 400:
            # Inspect only to select a fixed code; never return or log this body.
            body = exc.response.text[:4000].casefold()
            code = ('AI_CONTEXT_LIMIT' if any(t in body for t in ('context_length', 'context length', 'maximum context', 'too many tokens'))
                    else 'AI_REQUEST_REJECTED')
        code = code or ('AI_SERVICE_ERROR' if status >= 500 else 'AI_REQUEST_REJECTED')
    elif isinstance(exc, httpx.TimeoutException):
        code = 'AI_TIMEOUT'
    elif isinstance(exc, (ssl.SSLError, ssl.CertificateError)):
        code = 'AI_TLS_ERROR'
    elif isinstance(exc, httpx.RequestError):
        code = 'AI_TLS_ERROR' if any(t in str(exc).casefold() for t in ('certificate_verify_failed', 'certificate verify failed')) else 'AI_CONNECTION_ERROR'
    else:
        code = 'AI_INTERNAL_ERROR'
    return {'code': code, 'message': MESSAGES[code], 'http_status': status}
