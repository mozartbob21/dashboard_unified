"""Browser tile configuration. Only the selected provider's public key is exposed."""
import json
import os

_ORIGINS = {
    "osm": "https://tile.openstreetmap.org",
    "yandex": "https://tiles.api-maps.yandex.ru",
}


def browser_config():
    provider = os.getenv("MAP_TILE_PROVIDER", "auto").strip().lower() or "auto"
    key = os.getenv("YANDEX_TILES_API_KEY", "").strip()
    notice = ""
    if provider == "auto":
        provider = "yandex" if key else "osm"
    elif provider == "yandex" and not key:
        provider = "osm"
        notice = "Ключ Яндекс Карт не настроен. Временно показана OpenStreetMap. Администратору нужно добавить YANDEX_TILES_API_KEY в .env сервера."
    elif provider not in _ORIGINS:
        provider = "osm"
        notice = "Неизвестный источник карты. Показана OpenStreetMap. MAP_TILE_PROVIDER должен быть auto, yandex или osm."
    result = {"provider": provider, "notice": notice}
    if provider == "yandex":
        result["api_key"] = key
    return result


def image_origin():
    return _ORIGINS[browser_config()["provider"]]


def script_config():
    """JSON safe inside an inert HTML script element, including unusual key values."""
    value = json.dumps(browser_config(), ensure_ascii=True)
    return value.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
