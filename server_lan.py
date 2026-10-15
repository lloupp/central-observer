#!/usr/bin/env python3
"""Servidor LAN de uma unica instancia, protegido com Basic Auth e HTTPS.

Para uso institucional, nao iniciar sem TLS valido e acesso restrito via firewall.
Nao disponibiliza contas individuais; uso inicial/piloto controlado.
"""
from __future__ import annotations

import base64
import binascii
import hmac
import os
import ssl
import sys
import urllib.parse
from http.server import ThreadingHTTPServer
from pathlib import Path

import app


HOST = os.environ.get('OBS_LAN_HOST', '0.0.0.0')
PORT = int(os.environ.get('OBS_PORT', '8765'))
USERNAME = os.environ.get('OBS_LAN_USER', 'central')
PASSWORD = os.environ.get('OBS_LAN_PASSWORD', '')
TLS_CERT = os.environ.get('OBS_TLS_CERT', '')
TLS_KEY = os.environ.get('OBS_TLS_KEY', '')
TLS = bool(TLS_CERT and TLS_KEY)
SCHEME = 'https' if TLS else 'http'
PUBLIC_ORIGIN = os.environ.get('OBS_PUBLIC_ORIGIN', '').rstrip('/')


def validate_configuration():
    if len(PASSWORD) < 16 or PASSWORD.strip() != PASSWORD:
        raise ValueError('Defina OBS_LAN_PASSWORD com no minimo 16 caracteres sem espacos nas extremidades.')
    if not USERNAME or ':' in USERNAME:
        raise ValueError('OBS_LAN_USER invalido.')
    if bool(TLS_CERT) != bool(TLS_KEY):
        raise ValueError('Informe OBS_TLS_CERT e OBS_TLS_KEY juntos.')
    if not TLS and HOST not in ('127.0.0.1', 'localhost', '::1') and os.environ.get('OBS_ALLOW_INSECURE_HTTP') != '1':
        raise ValueError('Acesso LAN exige HTTPS. Configure OBS_TLS_CERT / OBS_TLS_KEY ou, somente em teste isolado sem dados reais, OBS_ALLOW_INSECURE_HTTP=1.')
    if TLS:
        for name in (TLS_CERT, TLS_KEY):
            if not Path(name).is_file():
                raise ValueError('Certificado/chave TLS nao encontrado: ' + name)
    if PUBLIC_ORIGIN:
        parsed = urllib.parse.urlsplit(PUBLIC_ORIGIN)
        if parsed.scheme != SCHEME or not parsed.hostname or parsed.path not in ('', '/') or parsed.query or parsed.fragment:
            raise ValueError('OBS_PUBLIC_ORIGIN deve ser uma origem com esquema, host e porta, sem caminho.')
    # A implementacao atual usa SQLite/WAL. O arquivo ativo deve ficar em disco local.
    if str(app.DB_PATH).startswith('\\\\') or str(app.DB_PATH).startswith('//'):
        raise ValueError('OBS_DB nao pode apontar para pasta UNC ou compartilhamento de rede (SQLite/WAL).')


class LanHandler(app.Handler):
    server_version = 'CentralObserverLAN/1.0'

    def end_headers(self):
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
        if TLS:
            self.send_header('Strict-Transport-Security', 'max-age=86400')
        super().end_headers()

    def authenticated(self):
        raw = self.headers.get('Authorization', '')
        if not raw.startswith('Basic ') or len(raw) > 512:
            return False
        try:
            supplied = base64.b64decode(raw[6:], validate=True).decode('utf-8')
            username, password = supplied.split(':', 1)
        except (ValueError, UnicodeError, binascii.Error):
            return False
        return (hmac.compare_digest(username.encode(), USERNAME.encode())
                and hmac.compare_digest(password.encode(), PASSWORD.encode()))

    def require_authentication(self):
        self.send_response(401)
        self.send_header('WWW-Authenticate', 'Basic realm="Central Observer", charset="UTF-8"')
        self.send_header('Content-Length', '0')
        self.end_headers()

    def safe_origin(self):
        # Verifica a origem exata nas alteracoes para reduzir risco de CSRF.
        # Nao confia em X-Forwarded-* fornecidos por clientes.
        host = self.headers.get('Host', '')
        if not host or '/' in host or '@' in host or '\\' in host:
            return False
        expected = PUBLIC_ORIGIN or (SCHEME + '://' + host)
        origin = self.headers.get('Origin', '')
        fetch_site = self.headers.get('Sec-Fetch-Site', '')
        if fetch_site in ('cross-site', 'same-site'):
            return False
        if origin and origin != expected:
            return False
        if self.command in ('POST', 'PATCH'):
            # Requisicoes fetch do navegador enviam Origin; bloquear alteracoes sem ele.
            return origin == expected
        return True

    def run_action(self):
        if not self.authenticated():
            return self.require_authentication()
        return super().run_action()


class LanServer(ThreadingHTTPServer):
    daemon_threads = True


def main():
    try:
        validate_configuration()
    except ValueError as ex:
        print('\nCONFIGURACAO RECUSADA: ' + str(ex), file=sys.stderr)
        return 2
    app.init_db()
    server = LanServer((HOST, PORT), LanHandler)
    if TLS:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        ctx.load_cert_chain(TLS_CERT, TLS_KEY)
        server.socket = ctx.wrap_socket(server.socket, server_side=True)
    print(f'\nCentral Observer — modo LAN ({SCHEME})')
    print(f'Endereco: {SCHEME}://<IP-OU-NOME-DO-SERVIDOR>:{PORT}')
    print(f'Banco ativo (disco do servidor): {app.DB_PATH}')
    print('Acesso limitado pela senha e pelas regras de rede/firewall.\n')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
