"""Run with python app.py, then open http://127.0.0.1:8765."""
import argparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import mimetypes
from pathlib import Path
import secrets
import traceback
from urllib.parse import quote, unquote, urlsplit
import webbrowser

from converter.docx import MAX_UPLOAD
from converter.engine import convert, word_available
from converter.errors import ConversionError
from converter.fonts import inventory, public_inventory, ROOT

FONT_ROUTES = {}
for slug, family in (("bell", "Bell MT"), ("lm", "LM Roman 10")):
    FONT_ROUTES[f"/api/fonts/{slug}"] = (family, "regular")
    for variant in ("regular", "bold", "italic", "bolditalic"):
        FONT_ROUTES[f"/api/fonts/{slug}/{variant}"] = (family, variant)


class Handler(BaseHTTPRequestHandler):
    def _allowed_host(self):
        port = self.server.server_port
        return self.headers.get("Host") in {f"127.0.0.1:{port}", f"localhost:{port}"}

    def _respond(self, status, content, mime="application/json; charset=utf-8", headers=None):
        if isinstance(content, dict):
            content = json.dumps(content, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; worker-src 'self'; style-src 'self'; font-src 'self' blob:; img-src 'self' data: blob:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'")
        for name, value in (headers or {}).items():
            self.send_header(name, value)
        self.end_headers()
        try:
            self.wfile.write(content)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def do_GET(self):
        if not self._allowed_host():
            return self._respond(403, {"error": "Host tidak diizinkan."})
        path = urlsplit(self.path).path
        if path == "/api/status":
            return self._respond(200, {"engine": "Microsoft Word", "ready": word_available(), "fonts": public_inventory(), "max_size_mb": 30, "token": self.server.token})
        if path.startswith('/api/download/'):
            result = self.server.last_result
            if not result or not secrets.compare_digest(path.rsplit('/', 1)[-1], result[0]):
                return self._respond(404, {"error": "Hasil sudah diganti. Konversi ulang dokumen untuk mengunduhnya."})
            return self._respond(200, result[2], 'application/pdf', {'Content-Disposition': "attachment; filename=\"document.pdf\"; filename*=UTF-8''" + quote(result[1], safe="")})
        if path in FONT_ROUTES:
            name, variant = FONT_ROUTES[path]
            font = inventory()[name]["variants"].get(variant)
            if not font or not font["embeddable"]:
                return self._respond(404, {"error": "Font tidak tersedia."})
            file = Path(font["path"])
            return self._respond(200, file.read_bytes(), "font/otf" if file.suffix == ".otf" else "font/ttf")
        static = {"/": "index.html", "/styles.css": "styles.css", "/app.js": "app.js", "/preview.js": "preview.js", "/favicon.svg": "favicon.svg", "/vendor/pdfjs/pdf.mjs": "vendor/pdfjs/pdf.mjs", "/vendor/pdfjs/pdf.worker.mjs": "vendor/pdfjs/pdf.worker.mjs"}
        if path not in static:
            return self._respond(404, {"error": "Halaman tidak ditemukan."})
        file = ROOT / "web" / static[path]
        mime = "text/javascript" if file.suffix in {'.js', '.mjs'} else (mimetypes.guess_type(file)[0] or "application/octet-stream")
        return self._respond(200, file.read_bytes(), mime + "; charset=utf-8")

    def do_POST(self):
        if not self._allowed_host():
            return self._respond(403, {"error": "Host tidak diizinkan."})
        if urlsplit(self.path).path != "/api/convert":
            return self._respond(404, {"error": "Endpoint tidak ditemukan."})
        origins = {f"http://127.0.0.1:{self.server.server_port}", f"http://localhost:{self.server.server_port}"}
        if self.headers.get("Origin") not in origins | {None} or not secrets.compare_digest(self.headers.get("X-Converter-Token", ""), self.server.token):
            return self._respond(403, {"error": "Sesi tidak valid. Muat ulang halaman lalu coba lagi."})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_UPLOAD:
                raise ConversionError("Ukuran DOCX harus antara 1 byte dan 30 MB.", 413)
            if self.headers.get("Transfer-Encoding"):
                raise ConversionError("Unggahan chunked tidak didukung.", 400)
            filename = unquote(self.headers.get("X-Filename", "document.docx"))
            if not filename.lower().endswith(".docx"):
                raise ConversionError("Pilih file dengan ekstensi .docx.", 400)
            self.connection.settimeout(30)
            data = self.rfile.read(length)
            self.connection.settimeout(None)
            if len(data) != length:
                raise ConversionError("Unggahan tidak lengkap.", 400)
            pdf, report = convert(data, unquote(self.headers.get("X-Font", "original")))
            safe_name = Path(filename.replace("\\", "/")).stem[:120] + ".pdf"
            download_id = secrets.token_hex(24)
            # Keep only the latest PDF in RAM for ordinary HTTP downloads.
            # Uploaded DOCX and Word's temporary files have already been deleted.
            self.server.last_result = (download_id, safe_name, pdf)
            self._respond(200, pdf, "application/pdf", {
                "Content-Disposition": "attachment; filename=\"document.pdf\"; filename*=UTF-8''" + quote(safe_name, safe=""),
                "X-Conversion-Report": quote(json.dumps(report), safe=""),
                "X-Download-Url": '/api/download/' + download_id,
            })
        except ConversionError as exc:
            self._respond(exc.status, {"error": str(exc)})
        except (ValueError, TimeoutError):
            self._respond(400, {"error": "Unggahan tidak valid atau koneksi terputus."})
        except Exception:
            traceback.print_exc()
            self._respond(500, {"error": "Terjadi kesalahan saat memproses dokumen. Periksa terminal aplikasi."})


def make_server(port=8765):
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.token = secrets.token_urlsafe(32)
    server.last_result = None
    return server


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DOCX to PDF with Bell MT and LM Roman 10")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--open", action="store_true", help="Open the browser automatically")
    args = parser.parse_args()
    with make_server(args.port) as server:
        url = f"http://127.0.0.1:{server.server_port}"
        print(f"DOCX to PDF berjalan di {url}\nTekan Ctrl+C untuk berhenti.", flush=True)
        if args.open:
            webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
