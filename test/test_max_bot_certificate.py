"""Offline checks for the explicitly approved, MAX-only public root certificate."""
import hashlib
import os
import re
import ssl
from pathlib import Path

import httpx
import pytest

from services.max_bot import client, config


FINGERPRINT = "d26d2d0231b7c39f92cc738512ba54103519e4405d68b5bd703e9788ca8ecf31"
RELATIVE_CERTIFICATE = Path("deploy/certs/russian-trusted-root-ca.pem")
SYNTHETIC_TOKEN = "SYNTHETIC_MAX_TOKEN_NOT_REAL"


def bundled_certificate():
    return config.ROOT / RELATIVE_CERTIFICATE


def test_bundle_contains_only_one_approved_public_certificate():
    pem = bundled_certificate().read_text(encoding="ascii").strip()
    assert re.fullmatch(r"-----BEGIN CERTIFICATE-----\s+[A-Za-z0-9+/=\s]+-----END CERTIFICATE-----", pem)
    assert pem.count("-----BEGIN CERTIFICATE-----") == 1
    assert "PRIVATE KEY" not in pem
    der = ssl.PEM_cert_to_DER_cert(pem)
    assert hashlib.sha256(der).hexdigest() == FINGERPRINT


@pytest.mark.parametrize("mapping", [{}, {"MAX_ZIP_BOT_CA_FILE": ""}, {"MAX_ZIP_BOT_CA_FILE": "  "}])
def test_unset_ca_override_uses_bundled_certificate_without_enabling_bot(mapping):
    settings = config.Settings.from_env(mapping)
    assert settings.ca_file == bundled_certificate()
    assert settings.ca_file.is_file()
    assert settings.enabled is False
    assert settings.bot_token == ""


def test_direct_settings_also_selects_bundled_certificate():
    assert config.Settings().ca_file == bundled_certificate()


def test_relative_override_is_resolved_against_project_not_shell_directory(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    settings = config.Settings.from_env({"MAX_ZIP_BOT_CA_FILE": str(RELATIVE_CERTIFICATE)})
    assert settings.ca_file == bundled_certificate()
    assert settings.ca_file.is_absolute() and settings.ca_file.is_file()
    assert settings.ca_file.parent != tmp_path


def test_absolute_existing_override_is_preserved_and_not_printed(tmp_path, capsys, caplog):
    operator_bundle = tmp_path / "SYNTHETIC_PRIVATE_OPERATOR_PATH.pem"
    operator_bundle.write_bytes(bundled_certificate().read_bytes())
    settings = config.Settings.from_env({"MAX_ZIP_BOT_CA_FILE": str(operator_bundle)})
    assert settings.ca_file == operator_bundle
    captured = capsys.readouterr()
    assert str(operator_bundle) not in captured.out + captured.err + caplog.text


def test_client_trust_is_scoped_and_still_verifies_certificate_and_hostname(monkeypatch):
    original_httpx_client = httpx.Client
    options = {}
    calls = []
    before_environment = dict(os.environ)
    before_default_paths = ssl.get_default_verify_paths()
    def capture_client(**kwargs):
        options.update(kwargs)
        return original_httpx_client(**kwargs)
    def respond(request):
        calls.append(request)
        return httpx.Response(200, json={"is_bot": True, "user_id": 55})
    monkeypatch.setattr(client.httpx, "Client", capture_client)
    settings = config.Settings.from_env({})
    with client.MaxClient(SYNTHETIC_TOKEN, ca_bundle=settings.ca_file,
                          transport=httpx.MockTransport(respond)) as api:
        assert api.get_me() == {"is_bot": True, "user_id": 55}
    context = options["verify"]
    assert isinstance(context, ssl.SSLContext)
    assert context.verify_mode == ssl.CERT_REQUIRED
    assert context.check_hostname is True
    fingerprints = {hashlib.sha256(der).hexdigest() for der in context.get_ca_certs(binary_form=True)}
    assert FINGERPRINT in fingerprints
    assert options["trust_env"] is False and options["follow_redirects"] is False
    assert len(calls) == 1 and calls[0].url.host == "platform-api2.max.ru"
    # No global trust variables or Python defaults are changed by this client.
    assert dict(os.environ) == before_environment
    assert ssl.get_default_verify_paths() == before_default_paths


def test_invalid_override_does_not_fall_back_to_unverified_tls_or_print_path(tmp_path, capsys, caplog):
    operator_path = tmp_path / "SYNTHETIC_PRIVATE_OPERATOR_PATH.pem"
    operator_path.write_text("not a certificate", encoding="ascii")
    with pytest.raises(client.MaxAPIError) as caught:
        client.MaxClient(SYNTHETIC_TOKEN, ca_bundle=operator_path)
    assert caught.value.code == "TLS_CONFIG_INVALID"
    captured = capsys.readouterr()
    output = str(caught.value) + captured.out + captured.err + caplog.text
    assert str(operator_path) not in output and SYNTHETIC_TOKEN not in output
