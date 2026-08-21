"""Самоподписанный сертификат для HTTPS-диктовки (порт 8443)."""
import datetime
from pathlib import Path
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

OUT = Path("ssl"); OUT.mkdir(exist_ok=True)
key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "neurona-local")])
san = x509.SubjectAlternativeName([
    x509.DNSName("localhost"),
    x509.IPAddress(__import__("ipaddress").ip_address("127.0.0.1")),
    x509.IPAddress(__import__("ipaddress").ip_address("10.113.6.74")),
])
cert = (x509.CertificateBuilder()
        .subject_name(name).issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.utcnow())
        .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=825))
        .add_extension(san, critical=False)
        .sign(key, hashes.SHA256()))
(OUT / "key.pem").write_bytes(key.private_bytes(
    serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
    serialization.NoEncryption()))
(OUT / "cert.pem").write_bytes(cert.public_bytes(serialization.Encoding.PEM))
print("✅ ssl/key.pem + ssl/cert.pem готовы")
print("Запуск HTTPS: python -m uvicorn app:app --host 0.0.0.0 --port 8443 "
      "--ssl-keyfile ssl/key.pem --ssl-certfile ssl/cert.pem")
