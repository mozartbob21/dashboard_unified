"""Real TLS handshakes through MemoryBIO: no sockets, network or real secrets."""
from datetime import datetime, timedelta, timezone
import ssl

import certifi
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
import pytest

from services.max_bot.client import MaxClient


def certificate_pair(directory, label, hostname):
    now = datetime.now(timezone.utc)
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, label)])
    root = (x509.CertificateBuilder().subject_name(name).issuer_name(name)
        .public_key(key.public_key()).serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1)).not_valid_after(now + timedelta(days=2))
        .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
        .add_extension(x509.KeyUsage(True, False, False, False, False, True, True, False, False), critical=True)
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
        .sign(key, hashes.SHA256()))
    leaf_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    leaf = (x509.CertificateBuilder().subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, hostname)]))
        .issuer_name(name).public_key(leaf_key.public_key()).serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(days=1)).not_valid_after(now + timedelta(days=2))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(x509.SubjectAlternativeName([x509.DNSName(hostname)]), critical=False)
        .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(key.public_key()), critical=False)
        .sign(key, hashes.SHA256()))
    ca_path, cert_path, key_path = [directory / (label + extension) for extension in ('.ca.pem', '.cert.pem', '.key.pem')]
    ca_path.write_bytes(root.public_bytes(serialization.Encoding.PEM))
    cert_path.write_bytes(leaf.public_bytes(serialization.Encoding.PEM))
    key_path.write_bytes(leaf_key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
    server = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    server.load_cert_chain(cert_path, key_path)
    return ca_path, server


def handshake(client_context, server_context, hostname):
    client_in, client_out, server_in, server_out = [ssl.MemoryBIO() for _ in range(4)]
    client = client_context.wrap_bio(client_in, client_out, server_hostname=hostname)
    server = server_context.wrap_bio(server_in, server_out, server_side=True)
    complete = [False, False]
    for _ in range(20):
        for index, connection in enumerate((client, server)):
            if not complete[index]:
                try:
                    connection.do_handshake()
                    complete[index] = True
                except ssl.SSLWantReadError:
                    pass
        server_in.write(client_out.read())
        client_in.write(server_out.read())
        if all(complete):
            return
    raise AssertionError('TLS handshake did not finish')


def test_regression_api_trust_must_not_remove_upload_storage_trust(tmp_path, monkeypatch):
    public_ca, storage = certificate_pair(tmp_path, 'synthetic-public-ca', 'storage.example')
    extra_ca, api = certificate_pair(tmp_path, 'synthetic-approved-ca', 'api.example')
    monkeypatch.setattr(certifi, 'where', lambda: str(public_ca))
    # Reproduce the former implementation: API works, file upload fails.
    old = ssl.create_default_context(cafile=str(extra_ca))
    handshake(old, api, 'api.example')
    with pytest.raises(ssl.SSLCertVerificationError):
        handshake(old, storage, 'storage.example')
    with MaxClient('synthetic-token', ca_bundle=extra_ca) as connection:
        handshake(connection._verify, api, 'api.example')
        handshake(connection._verify, storage, 'storage.example')
        assert connection._verify.check_hostname
        assert connection._verify.verify_mode == ssl.CERT_REQUIRED
        with pytest.raises(ssl.SSLCertVerificationError):
            handshake(connection._verify, storage, 'wrong.example')
        _, foreign = certificate_pair(tmp_path, 'synthetic-unapproved-ca', 'storage.example')
        with pytest.raises(ssl.SSLCertVerificationError):
            handshake(connection._verify, foreign, 'storage.example')


def test_approved_ca_extends_certifi_even_with_environment_override(monkeypatch):
    from services.max_bot.config import Settings
    default = ssl.create_default_context(cafile=certifi.where())
    expected = set(default.get_ca_certs(binary_form=True))
    monkeypatch.setenv('SSL_CERT_FILE', '/not-used/synthetic-ca.pem')
    monkeypatch.setenv('SSL_CERT_DIR', '/not-used/synthetic-ca-directory')
    with MaxClient('synthetic-token', ca_bundle=Settings().ca_file) as connection:
        actual = set(connection._verify.get_ca_certs(binary_form=True))
        assert expected <= actual
        assert len(actual) > 1
