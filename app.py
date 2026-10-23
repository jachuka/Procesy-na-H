from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, parse_qs
import json
import argparse

from calculator import read_workbook

STATIC = Path(__file__).parent / 'static'


class Handler(BaseHTTPRequestHandler):
    def respond(self, status, body, content_type):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        paths = {'/': ('index.html', 'text/html; charset=utf-8'), '/app.js': ('app.js', 'text/javascript; charset=utf-8'), '/style.css': ('style.css', 'text/css; charset=utf-8')}
        item = paths.get(urlsplit(self.path).path)
        if not item:
            return self.respond(404, b'Not found', 'text/plain')
        self.respond(200, (STATIC / item[0]).read_bytes(), item[1])

    def do_POST(self):
        if urlsplit(self.path).path != '/api/calculate':
            return self.respond(404, b'Not found', 'text/plain')
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= 10 * 1024 * 1024:
                raise ValueError('Wybierz plik XLSX o rozmiarze do 10 MB.')
            data = self.rfile.read(size)
            sheet = parse_qs(urlsplit(self.path).query).get('sheet', [None])[0]
            result = read_workbook(data, sheet)
            self.respond(200, json.dumps(result).encode(), 'application/json; charset=utf-8')
        except ValueError as error:
            self.respond(400, json.dumps({'error': str(error)}).encode(), 'application/json; charset=utf-8')
        except Exception:
            self.respond(400, json.dumps({'error': 'Nie można odczytać pliku. Wgraj poprawny, niezaszyfrowany plik XLSX.'}).encode(), 'application/json; charset=utf-8')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8000)
    parser.add_argument('--host', default='127.0.0.1')
    args = parser.parse_args()
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()
