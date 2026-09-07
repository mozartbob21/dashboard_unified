#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерирует cert.pem/key.pem для HTTPS (SAN: localhost + IP сервера)."""
import subprocess, sys, datetime, ipaddress
try:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "cryptography"], check=True)
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

IP = sys.argv[1] if len(sys.argv) > 1 else "10.113.7.78"
key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "unified-dashboard")])
san = x509.SubjectAlternativeName([
    x509.DNSName("localhost"),
    x509.IPAddress(ipaddress.ip_address("127.0.0.1")),
    x509.IPAddress(ipaddress.ip_address(IP)),
])
cert = (x509.CertificateBuilder()
        .subject_name(name).issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.utcnow() - datetime.timedelta(days=1))
        .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=1095))
        .add_extension(san, critical=False)
        .sign(key, hashes.SHA256()))
open("key.pem", "wb").write(key.private_bytes(
    serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL, serialization.NoEncryption()))
open("cert.pem", "wb").write(cert.public_bytes(serialization.Encoding.PEM))
print(f"✅ cert.pem/key.pem готовы для IP {IP} (3 года)")