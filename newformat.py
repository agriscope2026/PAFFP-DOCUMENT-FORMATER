"""
NEW FORMAT: source workbook -> AR (PDF), STUB with QR codes (PDF), SYSTEM UPLOAD (Excel)
======================================================================================
Source: a workbook like "sample to format.xlsx". Only sheets named
"STUB <MUNICIPALITY>" are used (e.g. "STUB HUNGDUAN"); all other sheets are ignored.
Each such sheet gives three files:
    <MUNICIPALITY> - AR FORMAT.pdf     layout of "AR FORMAT NEW.pdf" (14 per page, signatories)
    <MUNICIPALITY> - STUB FORMAT.pdf   layout of "STUB FORMAT NEW.xlsx" (12 per page),
                                       QR code of the RSBSA NO. in the last column (QR CODE)
    <MUNICIPALITY> - SYSTEM UPLOAD.xlsx  "SYSTEM UPLOAD.xlsx" filled in
Every barangay starts on a new page; NO. runs 1, 2, 3 ... through the municipality.

Fonts: Cambria when it is installed (desktop), otherwise the bundled Caladea
(fonts/, SIL Open Font License), which is designed to match Cambria.
"""
import io
import math
import re
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pymupdf
import qrcode

import paffp_formatter as engine

BASE_DIR = Path(__file__).resolve().parent
STUB_TEMPLATE = BASE_DIR / "STUB FORMAT NEW.xlsx"
SYSTEM_TEMPLATE = BASE_DIR / "SYSTEM UPLOAD.xlsx"
SHEET_PATTERN = re.compile(r"^\s*STUB\s+(.+?)\s*$", re.I)   # "STUB HUNGDUAN" -> HUNGDUAN

PAGE_W, PAGE_H = 936.0, 612.0          # 8.5 x 13 in (long bond), landscape
LINE_W = 0.72
GREY = (0.647, 0.647, 0.647)
DEFAULT_AMOUNT = engine.DEFAULT_AMOUNT

# ---- AR layout, measured from "AR FORMAT NEW.pdf" (points from the top-left)
AR = {
    "cols": [17.3, 48.5, 155.1, 255.2, 323.1, 394.4, 445.1, 538.7, 655.6, 706.2, 798.7, 917.3],
    "fields": ["number", "rsbsa", "surname", "first", "middle", "ext", "barangay", "municipality",
               "amount", "idtype", None],
    "heads": ["NO", "RSBSA NO.", "SURNAME", "FIRST NAME", "MIDDLE NAME", "EXT. NAME", "BARANGAY",
              "MUNICIPALITY", "AMOUNT", ["ID TYPE/ID NO. OF", "AUTHORIZED", "CLAIMANTS"], "SIGNATURE"],
    "right": {"number", "amount"},
    "size": 9.1, "title_y": [35.8, 49.7, 62.9], "region_col": 3, "province_col": 5,
    "head_top": 67.9, "head_bottom": 103.7, "row_h": 32.18, "rows": 14,
    "pad_left": 2.0, "pad_right": 4.8, "page_label_y": 595.5, "bottom_limit": 583.7,
}
# signatory block of the AR (last page), relative to the end of the last row
AR_FOOTER = {
    "size": 9.8,
    "rows": [26.2, 13.9, 35.8, 23.1, 30.0],        # grey row, labels, (space), names, titles
    "full_lines": [0, 1, 4, 5],                     # horizontal lines across the whole table
    "short_lines": [2, 3],                          # lines that stop at the DATE block
    "short_end": 706.2,
    "verticals": [17.3, 155.1, 323.1, 538.7, 706.2, 917.3],
    # (row index, x0, x1, text, style, align) - style: "", "b" bold, "i" italic
    "texts": [
        (1, 17.3, 155.1, "Certified true and correct:", "", "left"),
        (1, 155.1, 323.1, "Recommending Approval:", "", "left"),
        (1, 323.1, 538.7, "Approved for Payment:", "", "center"),
        (1, 538.7, 706.2, "Payment Received by Payee:", "", "left"),
        (3, 155.1, 323.1, "DANILO P. DAGUIO, Ph.D.,CESO IV", "b", "left"),
        (3, 323.1, 538.7, "ATTY. JENNILYN M. DAWAYAN, CESO III", "b", "center"),
        (4, 17.3, 155.1, "RSBSA/Personnel who handles the list of beneficiaries", "i", "left"),
        (4, 155.1, 323.1, "Regional Technical Director", "i", "left"),
        (4, 323.1, 538.7, "Regional Executive Director/Head of Agency", "i", "center"),
        (4, 538.7, 706.2, "Special Disbursing Officer", "i", "center"),
        (4, 706.2, 917.3, "DATE (MM/DD/YYYY)", "", "center"),
    ],
}
# ---- STUB layout: columns, headings, titles and left/right margins come from
# STUB FORMAT NEW.xlsx (the column headed "QR CODE" gets the QR codes); sizes
# measured from a stub printed from that template (PAFF HV HUNGDUAN STUB.pdf).
# The columns are scaled to fit between the margins; font and row height stay.
STUB = {
    "pt_per_char": 5.594, "size": 11.0, "top": 28.35, "row_h": 39.15, "rows": 12,
    "pad_left": 2.55, "pad_right": 5.25, "page_label_y": 595.0,
    "qr_pad": 6.0,               # clear space (pt) around the QR code inside its cell: keeps
                                 # the 4-module quiet zone when the rows are cut apart
}


# ================================================================ fonts
def _font_files():
    win = Path("C:/Windows/Fonts")
    cambria = {"": win / "cambria.ttc", "b": win / "cambriab.ttf", "i": win / "cambriai.ttf"}
    if all(p.exists() for p in cambria.values()):
        return cambria, "Cambria"
    fonts = BASE_DIR / "fonts"
    return {"": fonts / "Caladea-Regular.ttf", "b": fonts / "Caladea-Bold.ttf",
            "i": fonts / "Caladea-Italic.ttf"}, "Caladea"


FONT_FILES, FONT_NAME = _font_files()


class Writer:
    """Draws text and lines on the pages of one document with shared fonts."""

    def __init__(self):
        self.doc = pymupdf.open()
        self.fonts = {k: pymupdf.Font(fontbuffer=Path(p).read_bytes()) for k, p in FONT_FILES.items()}
        self.tw = {}

    def new_page(self):
        page = self.doc.new_page(width=PAGE_W, height=PAGE_H)
        self.tw[page.number] = pymupdf.TextWriter(page.rect)
        return page

    def text(self, page, x0, x1, yc, text, size, style="", align="left", pad_l=2.0, pad_r=2.0):
        """One line of text in the box x0..x1, vertically centred on yc; shrinks to fit."""
        text = "" if text is None else str(text)
        if not text:
            return
        font = self.fonts[style]
        avail = max(1.0, (x1 - x0) - pad_l - pad_r)
        w = font.text_length(text, size)
        if w > avail:
            size = size * avail / w
            w = font.text_length(text, size)
        if align == "right":
            x = x1 - pad_r - w
        elif align == "center":
            x = (x0 + x1 - w) / 2
        else:
            x = x0 + pad_l
        baseline = yc + (font.ascender + font.descender) / 2 * size
        self.tw[page.number].append((x, baseline), text, font=font, fontsize=size)

    def lines(self, page, x0, x1, yc, text, size, style="", align="left", pad=2.0, leading=1.2):
        """Word-wrapped text (a list = fixed lines) centred vertically on yc."""
        font = self.fonts[style]
        if isinstance(text, (list, tuple)):
            rows = list(text)
        else:
            rows, cur = [], ""
            for word in str(text).split():
                trial = f"{cur} {word}".strip()
                if cur and font.text_length(trial, size) > (x1 - x0) - 2 * pad:
                    rows.append(cur)
                    cur = word
                else:
                    cur = trial
            rows.append(cur)
        step = size * leading
        first = yc - step * (len(rows) - 1) / 2
        for i, r in enumerate(rows):
            self.text(page, x0, x1, first + i * step, r, size, style, align, pad, pad)

    # Cambria uses one glyph for "-" and the non-breaking hyphen (and for the two
    # kinds of space); the PDF text map then says U+2011, so a copied RSBSA NO.
    # would not match. Point those glyphs back to the plain characters.
    TOUNICODE_FIX = {"2011": "002D", "2010": "002D", "00A0": "0020"}

    def _fix_tounicode(self):
        """Only the character side of each mapping is changed - the glyph numbers
        on the left can look the same (in Caladea, glyph 00A0 is "b")."""
        seen = set()
        for page in self.doc:
            for font in page.get_fonts():
                kind, val = self.doc.xref_get_key(font[0], "ToUnicode")
                if kind != "xref" or val in seen:
                    continue
                seen.add(val)
                tu = int(val.split()[0])
                data = self.doc.xref_stream(tu).decode("latin-1")
                out = []
                for line in data.split("\n"):
                    tokens = re.findall(r"<([0-9A-Fa-f]+)>", line)
                    if len(tokens) in (2, 3) and tokens[-1].upper() in self.TOUNICODE_FIX:
                        new = self.TOUNICODE_FIX[tokens[-1].upper()]
                        head, _, _ = line.rpartition(f"<{tokens[-1]}>")
                        line = f"{head}<{new}>"
                    out.append(line)
                fixed = "\n".join(out)
                if fixed != data:
                    self.doc.update_stream(tu, fixed.encode("latin-1"))

    def finish(self):
        for pno, tw in self.tw.items():
            tw.write_text(self.doc[pno])
        try:
            self.doc.subset_fonts()
        except Exception:
            pass                      # fontTools missing: fonts stay complete
        self._fix_tounicode()
        out = self.doc.tobytes(garbage=3, deflate=True)
        self.doc.close()
        return out


def hline(page, x0, x1, y):
    page.draw_line((x0, y), (x1, y), color=(0, 0, 0), width=LINE_W)


def vline(page, x, y0, y1):
    page.draw_line((x, y0), (x, y1), color=(0, 0, 0), width=LINE_W)


def fill(page, x0, y0, x1, y1, color=GREY):
    page.draw_rect(pymupdf.Rect(x0, y0, x1, y1), color=None, fill=color, width=0)


# ================================================================ source workbook
def _fmt_date(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%b-%d-%Y")
    return v


def read_source(source):
    """Municipality sheets of the source workbook.
    Returns [{"municipality", "sheet", "region", "province", "records"}]."""
    wb = openpyxl.load_workbook(source, data_only=True)
    out = []
    for ws in wb.worksheets:
        m = SHEET_PATTERN.match(ws.title)
        if not m or ws.sheet_state != "visible":
            continue
        hr, cols = engine.find_heading_row(ws)
        if not hr:
            continue
        region = province = None
        for r in range(1, hr):
            for c in ws[r]:
                t = str(c.value or "").strip()
                if re.match(r"region\b", t, re.I):
                    region = t
                elif re.match(r"province\b", t, re.I):
                    province = t
        records = []
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            rec = {}
            for f, c in cols.items():
                v = row[c - 1] if c - 1 < len(row) else None
                v = re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v
                rec[f] = None if v in ("", None) else _fmt_date(v) if f == "birthdate" else v
            if not any(rec.get(f) for f in ("rsbsa", "surname", "first")):
                continue
            records.append(rec)
        if records:
            out.append({"municipality": m.group(1).strip().upper(), "sheet": ws.title,
                        "region": region or "Region: CAR", "province": province or "Province :",
                        "records": records})
    if not out:
        raise ValueError('no sheets named "STUB <MUNICIPALITY>" with beneficiaries were found')
    return out


def pages_by_barangay(records, per_page):
    groups = engine.group_by_barangay(records)
    pages = []
    for recs in groups.values():
        for i in range(0, len(recs), per_page):
            pages.append(recs[i:i + per_page])
    return pages


def numbered(records):
    """Copies of the records with NO. 1..n and text values."""
    out = []
    for i, r in enumerate(records, 1):
        r = dict(r)
        r["number"] = i
        out.append(r)
    return out


# ================================================================ AR (PDF)
def render_ar(muni, amount=DEFAULT_AMOUNT):
    L = AR
    w = Writer()
    recs = numbered(muni["records"])
    for r in recs:
        if r.get("amount") in (None, "") and amount not in (None, ""):
            r["amount"] = amount
    pages = pages_by_barangay(recs, L["rows"])
    x = L["cols"]
    size = L["size"]

    def header(page):
        w.text(page, x[0], x[-1], L["title_y"][0], "DEPARTMENT OF AGRICULTURE", size, "b", "center")
        w.text(page, x[0], x[-1], L["title_y"][1],
               "2026 PRESIDENTIAL ASSISTANCE FOR FARMERS AND FISHERFOLK PROGRAM (PAFFP) PAYROLL",
               size, "b", "center")
        rc, pc = L["region_col"], L["province_col"]
        w.text(page, x[rc], x[rc + 2], L["title_y"][2], muni["region"], size, "b", "left", pad_r=0)
        w.text(page, x[pc], x[pc + 2], L["title_y"][2], muni["province"], size, "b", "left", pad_r=0)
        top, bot = L["head_top"], L["head_bottom"]
        fill(page, x[0], top, x[-1], bot)
        for i, head in enumerate(L["heads"]):
            w.lines(page, x[i], x[i + 1], (top + bot) / 2, head if isinstance(head, list) else [head],
                    size, "b", "center", pad=1.0, leading=1.25)
        hline(page, x[0], x[-1], top)
        hline(page, x[0], x[-1], bot)
        return bot

    def grid(page, top, bottom):
        for xi in x:
            vline(page, xi, top, bottom)

    out_pages = []
    for pg in pages:
        page = w.new_page()
        y = header(page)
        top = L["head_top"]
        for rec in pg:
            yc = y + L["row_h"] / 2
            for i, f in enumerate(L["fields"]):
                if not f:
                    continue
                v = rec.get(f)
                if f == "amount" and isinstance(v, (int, float)):
                    v = f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}"
                w.text(page, x[i], x[i + 1], yc, v, size, "",
                       "right" if f in L["right"] else "left", L["pad_left"], L["pad_right"])
            y += L["row_h"]
            hline(page, x[0], x[-1], y)
        grid(page, top, y)
        out_pages.append((page, y))

    # signatory block after the last row (new page with headings if it does not fit)
    F = AR_FOOTER
    page, y = out_pages[-1]
    if y + sum(F["rows"]) > L["bottom_limit"]:
        page = w.new_page()
        y = header(page)
        grid(page, L["head_top"], y)
    edges = [y]
    for h in F["rows"]:
        edges.append(edges[-1] + h)
    fill(page, x[0], edges[0], x[-1], edges[1])
    for xi in x:                                   # grey row keeps all column lines
        vline(page, xi, edges[0], edges[1])
    for xi in F["verticals"]:
        vline(page, xi, edges[1], edges[-1])
    for i in F["full_lines"]:
        hline(page, x[0], x[-1], edges[i])
    for i in F["short_lines"]:
        hline(page, x[0], F["short_end"], edges[i])
    for row, x0, x1, text, style, align in F["texts"]:
        yc = (edges[row] + edges[row + 1]) / 2
        w.lines(page, x0, x1, yc, text, F["size"], style, align, pad=2.0, leading=1.2)

    _page_labels(w, L["page_label_y"], size)
    return w.finish()


def _page_labels(w, y, size):
    n = len(w.doc)
    for page in w.doc:
        w.text(page, 0, PAGE_W, y, f"Page {page.number + 1} of {n}", size, "", "center")


# ================================================================ STUB (PDF with QR codes)
def _stub_template():
    ws = openpyxl.load_workbook(STUB_TEMPLATE).active
    hr, cols = engine.find_heading_row(ws, max_scan=ws.max_row)
    if not hr:
        raise ValueError(f"no column headings found in {STUB_TEMPLATE.name}")
    heads = [(c.column, c.value) for c in ws[hr] if c.value not in (None, "")]
    last = max(c for c, _ in heads)
    widths = []
    for c in range(1, last + 1):
        dim = ws.column_dimensions.get(openpyxl.utils.get_column_letter(c))
        widths.append((dim.width if dim and dim.width else 8.43) * STUB["pt_per_char"])
    left, right = ws.page_margins.left * 72, ws.page_margins.right * 72
    avail = PAGE_W - left - right
    if sum(widths) > avail:                       # fit the table between the margins
        widths = [wd * avail / sum(widths) for wd in widths]
    x0 = left + (avail - sum(widths)) / 2
    xs = [x0]
    for wdt in widths:
        xs.append(xs[-1] + wdt)
    heights = [ws.row_dimensions[r].height or 15 for r in range(1, hr + 1)]
    titles = []                                   # (row, col, text, bold) above the headings
    for r in range(1, hr):
        for c in ws[r]:
            if c.value not in (None, ""):
                titles.append((r, c.column, str(c.value), bool(c.font and c.font.b)))
    merged = {(m.min_row, m.min_col): m.max_col for m in ws.merged_cells.ranges}
    head_text = {c: str(v) for c, v in heads}
    fields = {c: f for f, c in cols.items()}
    qr_col = next((c for c, v in heads if "qr" in engine.norm(v)), None)
    return {"xs": xs, "heights": heights, "titles": titles, "merged": merged, "heads": head_text,
            "fields": fields, "hr": hr, "last": last, "qr_col": qr_col}


def render_stub(muni):
    T = _stub_template()
    S = STUB
    w = Writer()
    recs = numbered(muni["records"])
    pages = pages_by_barangay(recs, S["rows"])
    xs, size, last = T["xs"], S["size"], T["last"]
    qr_col = T["qr_col"]

    def header(page):
        y = S["top"]
        tops = []
        for h in T["heights"]:
            tops.append(y)
            y += h
        for r, c, text, bold in T["titles"]:
            low = text.lower()
            if low.startswith("region"):
                text = muni["region"]
            elif low.startswith("province"):
                text = muni["province"]
            yc = tops[r - 1] + T["heights"][r - 1] / 2
            end = T["merged"].get((r, c))
            if end:                                               # merged title row: centred
                w.text(page, xs[c - 1], xs[end], yc, text, size, "b" if bold else "", "center")
            else:
                w.text(page, xs[c - 1], xs[min(c + 1, last)], yc, text, size, "b" if bold else "",
                       "left", S["pad_left"], 0)
        top, bot = tops[-1], y
        fill(page, xs[0], top, xs[-1], bot)
        for c in range(1, last + 1):
            w.lines(page, xs[c - 1], xs[c], (top + bot) / 2, T["heads"].get(c, ""), size, "b", "center", pad=1.5)
        hline(page, xs[0], xs[-1], top)
        hline(page, xs[0], xs[-1], bot)
        return top, bot

    for pg in pages:
        page = w.new_page()
        top, y = header(page)
        for rec in pg:
            yc = y + S["row_h"] / 2
            for c, f in T["fields"].items():
                if f == "number":
                    w.text(page, xs[c - 1], xs[c], yc, rec["number"], size, "", "right",
                           S["pad_left"], S["pad_right"])
                else:
                    w.text(page, xs[c - 1], xs[c], yc, rec.get(f), size, "", "left", S["pad_left"], 2.0)
            if qr_col and rec.get("rsbsa"):
                _stub_qr(page, xs[qr_col - 1], xs[qr_col], y, y + S["row_h"], rec["rsbsa"])
            y += S["row_h"]
            hline(page, xs[0], xs[-1], y)
        for xi in xs:
            vline(page, xi, top, y)

    _page_labels(w, S["page_label_y"], size)
    return w.finish()


def _qr_png(text):
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=4, border=0)
    qr.add_data(str(text), optimize=0)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").get_image().convert("1")
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def _stub_qr(page, x0, x1, y0, y1, rsbsa):
    """QR code of the RSBSA NO., as large as the QR CODE cell allows, centred."""
    pad = STUB["qr_pad"]
    q = min(x1 - x0, y1 - y0) - 2 * pad
    qx, qy = (x0 + x1 - q) / 2, (y0 + y1 - q) / 2
    page.insert_image(pymupdf.Rect(qx, qy, qx + q, qy + q), stream=_qr_png(rsbsa))


# ================================================================ SYSTEM UPLOAD (Excel)
def build_system(muni):
    """SYSTEM UPLOAD.xlsx filled with the municipality's beneficiaries (sample row removed).
    Values are written as text, like the template's sample row; EXT. NAME is added to
    the first name because the upload layout has no extension column."""
    wb = openpyxl.load_workbook(SYSTEM_TEMPLATE)
    ws = wb.active
    hr, cols = engine.find_heading_row(ws, max_scan=ws.max_row)
    if not hr:
        raise ValueError(f"no column headings found in {SYSTEM_TEMPLATE.name}")
    style_row = [copy(c._style) for c in ws[hr + 1]] if ws.max_row > hr else None
    if ws.max_row > hr:
        ws.delete_rows(hr + 1, ws.max_row - hr)
    for i, rec in enumerate(muni["records"], 1):
        r = hr + i
        first = rec.get("first")
        ext = str(rec.get("ext") or "").strip()
        if ext and first and not engine.norm(first).endswith(engine.norm(ext)):
            first = f"{first} {ext}"
        values = dict(rec, number=i, first=first)
        for f, c in cols.items():
            v = values.get(f)
            cell = ws.cell(r, c, None if v in (None, "") else str(v))
            if style_row and c - 1 < len(style_row):
                cell._style = copy(style_row[c - 1])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ================================================================ one-call helper
OUTPUTS = [("AR", "AR FORMAT", "pdf"), ("STUB", "STUB FORMAT", "pdf"), ("SYSTEM", "SYSTEM UPLOAD", "xlsx")]


def province_line(province):
    """'ifugao' -> 'Province : IFUGAO' (also accepts 'Province : IFUGAO')."""
    name = re.sub(r"^\s*province\s*:?\s*", "", str(province or ""), flags=re.I).strip()
    return f"Province : {name.upper()}" if name else None


def render(kind, muni, amount=DEFAULT_AMOUNT, province=None):
    """Returns (file name, bytes) for one output of one municipality.
    province: what to print on the Province line of the AR/STUB heading
    (e.g. "IFUGAO"); empty = the Province written on the STUB sheet."""
    label, ext = next((l, e) for k, l, e in OUTPUTS if k == kind)
    line = province_line(province)
    if line:
        muni = dict(muni, province=line)
    data = (render_ar(muni, amount) if kind == "AR" else render_stub(muni) if kind == "STUB"
            else build_system(muni))
    return f"{engine.safe_name(muni['municipality'])} - {label}.{ext}", data


def process(source, out_dir, kinds=("AR", "STUB", "SYSTEM"), amount=DEFAULT_AMOUNT, log=print,
            province=None):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    munis = read_source(source)
    log(f"{Path(str(source)).name}: {len(munis)} municipalit{'y' if len(munis) == 1 else 'ies'} "
        f"({', '.join(m['municipality'] for m in munis)}) - fonts: {FONT_NAME}")
    written = []
    for m in munis:
        log(f"   {m['municipality']}: {len(m['records'])} beneficiaries, "
            f"{province_line(province) or m['province']}, "
            f"{len(engine.group_by_barangay(m['records']))} barangays")
        for kind in kinds:
            name, data = render(kind, m, amount, province)
            path = out_dir / name
            try:
                path.write_bytes(data)
            except PermissionError:
                raise PermissionError(f"cannot write {name} - close it and try again")
            log(f"      -> {name}")
            written.append(path)
    return written
