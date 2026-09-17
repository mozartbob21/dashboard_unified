"""Read-only network checks. Run on the actual Neurona server; no login is sent."""
import argparse
import concurrent.futures
import json
import os
import platform
import socket
import ssl
import subprocess
from datetime import datetime
from urllib.parse import urlsplit

HOST = 'zkh-kontur.mosreg.ru'


def tcp(host, port):
    result = {'host': host, 'port': port}
    try:
        with socket.create_connection((host, port), timeout=5):
            result['connected'] = True
    except OSError as error:
        result.update(connected=False, error_type=type(error).__name__, errno=error.errno, error=str(error))
    return result


def tls(address, ca_bundle=None):
    result = {'address': address, 'server_name': HOST, 'certificate_verification': True}
    try:
        context = ssl.create_default_context(cafile=ca_bundle)
        with socket.create_connection((address, 443), timeout=5) as raw:
            with context.wrap_socket(raw, server_hostname=HOST) as connection:
                result.update(connected=True, protocol=connection.version())
    except (OSError, ValueError) as error:
        result.update(connected=False, error_type=type(error).__name__, error=str(error))
    return result


def diagnose(internal_ip='10.10.34.2', dns_server='10.10.51.2'):
    report = {'checked_at': datetime.now().astimezone().isoformat(timespec='seconds'),
              'platform': platform.system(), 'portal': HOST,
              'note': 'Run on the Neurona server. Internal IP/DNS supplied for comparison; no credentials used.'}
    try:
        addresses = sorted({item[4][0] for item in socket.getaddrinfo(HOST, 443, type=socket.SOCK_STREAM)})
        report['system_dns'] = addresses
    except OSError as error:
        addresses = []
        report['system_dns_error'] = {'type': type(error).__name__, 'error': str(error)}
    try:
        query = subprocess.run(['nslookup', HOST, dns_server], capture_output=True, text=True,
                               errors='replace', timeout=10)
        report['specified_dns'] = {'server': dns_server, 'return_code': query.returncode,
                                   'output': (query.stdout + query.stderr).strip()[:4000]}
    except (OSError, subprocess.TimeoutExpired) as error:
        report['specified_dns'] = {'server': dns_server, 'error_type': type(error).__name__}
    targets = list(dict.fromkeys([(internal_ip, 443), (dns_server, 53), *[(ip, 443) for ip in addresses]]))
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        report['tcp'] = list(pool.map(lambda target: tcp(*target), targets))
    bundle = os.getenv('EDDS_ARM_CA_BUNDLE') or os.getenv('REQUESTS_CA_BUNDLE') or os.getenv('CURL_CA_BUNDLE')
    report['custom_ca_bundle_configured'] = bool(bundle)
    tls_targets = [value['host'] for value in report['tcp'] if value['connected'] and value['port'] == 443]
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        report['tls_direct'] = list(pool.map(lambda address: tls(address, bundle), tls_targets))
    report['proxy_environment'] = {}
    for name in ('HTTPS_PROXY', 'https_proxy', 'ALL_PROXY', 'all_proxy'):
        raw = os.getenv(name)
        if raw:
            try:
                parsed = urlsplit(raw)
                report['proxy_environment'][name] = {'scheme': parsed.scheme, 'host': parsed.hostname,
                                                      'port': parsed.port, 'has_credentials': bool(parsed.username)}
            except ValueError:
                report['proxy_environment'][name] = {'invalid_url': True}
    report['no_proxy'] = os.getenv('NO_PROXY') or os.getenv('no_proxy') or ''
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Network diagnosis for ARM EDDS; never sends credentials.')
    parser.add_argument('--internal-ip', default='10.10.34.2')
    parser.add_argument('--dns-server', default='10.10.51.2')
    args = parser.parse_args()
    print(json.dumps(diagnose(args.internal_ip, args.dns_server), ensure_ascii=False, indent=2))
