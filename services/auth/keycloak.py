"""Keycloak OIDC: валидация токенов через JWKS."""
import os

KC_SERVER = os.getenv("KC_SERVER_URL", "http://localhost:8080")
KC_REALM = os.getenv("KC_REALM", "neurona")
KC_CLIENT_ID = os.getenv("KC_CLIENT_ID", "neurona-web")
KC_CLIENT_SECRET = os.getenv("KC_CLIENT_SECRET", "")
KC_JWKS_URL = f"{KC_SERVER}/realms/{KC_REALM}/protocol/openid-connect/certs"

_cached_jwks = None


def _get_jwks(force_refresh=False):
    """Кэш JWKS с возможностью принудительного обновления."""
    global _cached_jwks
    if _cached_jwks is None or force_refresh:
        import httpx
        print(f"[keycloak] fetching JWKS from {KC_JWKS_URL}", flush=True)
        _cached_jwks = httpx.get(KC_JWKS_URL, timeout=10).json()
    return _cached_jwks


def get_user_from_token_keycloak(token: str):
    """Валидация Keycloak access_token -> user-словарь или None.

    Проверяем:
      1. Подпись ключом из JWKS
      2. Issuer == наш Keycloak + realm
      3. Токен не истёк (exp)

    НЕ проверяем audience (Keycloak ставит 'account' по умолчанию).
    """
    if not token:
        return None
    try:
        from jose import jwt
        from datetime import datetime, timezone

        header = jwt.get_unverified_header(token)
        token_kid = header.get("kid")

        # Первая попытка с кэшированным JWKS
        jwks = _get_jwks(force_refresh=False)
        key = next((k for k in jwks.get("keys", []) if k.get("kid") == token_kid), None)

        # Если kid не найден — сбрасываем кэш и пробуем снова
        if not key:
            print(f"[keycloak] kid {token_kid} not in cache, refreshing JWKS", flush=True)
            jwks = _get_jwks(force_refresh=True)
            key = next((k for k in jwks.get("keys", []) if k.get("kid") == token_kid), None)

        if not key:
            available_kids = [k.get("kid") for k in jwks.get("keys", [])]
            print(f"[keycloak] no matching kid: token={token_kid}, jwks={available_kids}", flush=True)
            return None

        payload = jwt.decode(
            token,
            key,
            algorithms=["RS256"],
            options={"verify_aud": False},
        )

        expected_iss = f"{KC_SERVER}/realms/{KC_REALM}"
        if payload.get("iss") != expected_iss:
            print(f"[keycloak] bad issuer: {payload.get('iss')}", flush=True)
            return None

        exp = payload.get("exp")
        if exp and datetime.fromtimestamp(exp, tz=timezone.utc) < datetime.now(tz=timezone.utc):
            print("[keycloak] token expired", flush=True)
            return None

        roles = payload.get("realm_access", {}).get("roles", []) or []
        return {
            "username": payload.get("preferred_username", ""),
            "email": payload.get("email", ""),
            "roles": roles,
            "role": "admin" if "admin" in roles else (roles[0] if roles else "viewer"),
            "kc_sub": payload.get("sub", ""),
        }
    except Exception as e:
        print(f"[keycloak] validate error: {type(e).__name__}: {e}", flush=True)
        return None


async def get_user_from_token_keycloak_async(token: str):
    """Async-обёртка (JWKS-запрос в threadpool)."""
    import asyncio
    return await asyncio.to_thread(get_user_from_token_keycloak, token)
