"""Vercel serverless function for the "New format" tab.

Step 1  POST /api/newformat?step=read&name=<file name>   body = the source workbook (.xlsx or .pdf)
        -> JSON {"municipalities": [{"municipality", "region", "province", "records"}]}
        (only sheets named "STUB <MUNICIPALITY>")
Step 2  POST /api/newformat?step=render    JSON {"kind": "AR"|"STUB"|"SYSTEM",
                                                  "municipality": {...from step 1...}, "amount": 2325,
                                                  "province": "IFUGAO"}   (empty = Province on the sheet)
        -> the file (AR/STUB as PDF, SYSTEM as .xlsx)
One file per request keeps every response under Vercel's 4.5 MB limit.
"""
import io
import json
import sys
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import newformat  # noqa: E402

MAX_BYTES = 4 * 1024 * 1024
KINDS = {k for k, _, _ in newformat.OUTPUTS}
TYPES = {"pdf": "application/pdf",
         "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        step = parse_qs(urlparse(self.path).query).get("step", [""])[0]
        size = int(self.headers.get("Content-Length") or 0)
        if size <= 0:
            return self._json(400, {"error": "No data received."})
        if size > MAX_BYTES:
            return self._json(400, {"error": "File is larger than 4 MB. Split the workbook and try again."})
        body = self.rfile.read(size)
        try:
            if step == "read":
                name = parse_qs(urlparse(self.path).query).get("name", [""])[0]
                munis = newformat.read_source(io.BytesIO(body), Path(name).name or None)
                return self._json(200, {"municipalities": munis, "font": newformat.FONT_NAME})
            if step == "render":
                req = json.loads(body)
                kind = str(req.get("kind", "")).upper()
                muni = req.get("municipality") or {}
                if kind not in KINDS or not muni.get("records"):
                    return self._json(400, {"error": "Nothing to render."})
                amount = str(req.get("amount", newformat.DEFAULT_AMOUNT) or "").replace(",", "").strip()
                amount = (float(amount) if "." in amount else int(amount)) if amount else None
                name, data = newformat.render(kind, muni, amount, str(req.get("province") or "")[:60])
                self.send_response(200)
                self.send_header("Content-Type", TYPES[name.rsplit(".", 1)[-1]])
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(name)}")
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(data)
                return
            return self._json(400, {"error": "Unknown step."})
        except ValueError as e:
            return self._json(400, {"error": str(e)})
        except Exception as e:
            return self._json(400, {"error": f"Could not process the file ({e.__class__.__name__}: {e})"})

    def _json(self, code, obj):
        body = json.dumps(obj, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)
