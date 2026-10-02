"""Vercel serverless function: POST a masterlist .xlsx, get back a .zip with the
AR and/or STUB files.

Request:  POST /api/generate?name=<file name>&formats=AR,STUB
          body = the raw .xlsx file
Response: 200 application/zip, or 400/500 JSON {"error": "..."}
"""
import io
import json
import sys
import zipfile
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import paffp_formatter as engine  # noqa: E402

MAX_BYTES = 4 * 1024 * 1024  # Vercel request bodies are limited to 4.5 MB


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            q = parse_qs(urlparse(self.path).query)
            name = Path(q.get("name", ["masterlist.xlsx"])[0]).name
            wanted = {s.strip().upper() for s in q.get("formats", ["AR,STUB"])[0].split(",") if s.strip()}
            formats = [f for f in engine.FORMATS if f["name"] in wanted]
            if not formats:
                return self._error(400, "Choose AR and/or STUB.")
            size = int(self.headers.get("Content-Length") or 0)
            if size <= 0:
                return self._error(400, "No file received.")
            if size > MAX_BYTES:
                return self._error(400, "File is larger than 4 MB. Split the masterlist and try again.")
            data = self.rfile.read(size)

            notes = []
            zbuf = io.BytesIO()
            summary = {}
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for out_name, xlsx, pages, groups in engine.generate(
                        io.BytesIO(data), name, formats, log=notes.append):
                    z.writestr(out_name, xlsx)
                    summary[out_name] = pages
            body = zbuf.getvalue()
        except ValueError as e:
            return self._error(400, f"{name}: {e}")
        except Exception as e:  # unreadable file etc.
            return self._error(400, f"{name}: could not read this file as a masterlist ({e.__class__.__name__}).")

        zip_name = f"{engine.clean_stem(name)} - PAFFP.zip"
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(zip_name)}")
        self.send_header("X-Summary", quote(json.dumps({"notes": notes, "pages": summary})))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self._json(200, {"ok": True, "formats": [f["name"] for f in engine.FORMATS]})

    def _error(self, code, msg):
        self._json(code, {"error": msg})

    def _json(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)
