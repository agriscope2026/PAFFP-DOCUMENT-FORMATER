"""
PAFFP FORMATTER  (AR + STUB)
============================
Turns beneficiary masterlists into the AR (payroll) format and the STUB
format at the same time.

HOW TO USE
  - Easiest: double-click "PAFFP FORMATTER.bat" and pick the masterlist files.
  - Command line:  python paffp_formatter.py "masterlist1.xlsx" "masterlist2.xlsx"
    (with no files given, every .xlsx in the MASTERLISTS folder is formatted)
  Output goes to the OUTPUT folder:
      <masterlist name> - AR FORMAT.xlsx
      <masterlist name> - STUB FORMAT.xlsx

TEMPLATES ("AR FORMAT.xlsx", "STUB FORMAT.xlsx")
  - Rows 1..N   : page header block (title, Region/Province, column headings).
                  The column-heading row is the one containing "RSBSA".
  - Row N+1     : one EMPTY formatted row, used as the format of every beneficiary row.
  - Rows after  : spacer + signatories, copied once after the last beneficiary.
  Each template column is filled by matching its heading (NO, RSBSA NO.,
  SURNAME, FIRST NAME, MIDDLE NAME, BIRTHDATE, BARANGAY, MUNICIPALITY,
  AMOUNT, ID TYPE...). Edit text such as signatory names or Region/Province
  directly in the templates; the formatter picks the changes up automatically.

RULES APPLIED
  - Each page repeats the header block and holds exactly rows_per_page rows
    (AR 15, STUB 20), except the last page of each barangay (the remainder).
  - Each barangay starts on a new page.
  - The NO. column continues across barangays (1, 2, 3, ... to the end).
  - The signatories follow the last beneficiary, on the same page.
  - STUB only: a small QR code of the RSBSA number is placed in the NO. column.
"""
import io
import re
import sys
import warnings
import zlib
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
import qrcode
from PIL.PngImagePlugin import PngInfo
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------- settings
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "MASTERLISTS"
OUTPUT_DIR = BASE_DIR / "OUTPUT"

# One entry per output document.
#   rows_per_page: beneficiaries per page (only a barangay's last page may have fewer)
#   qr_column:     column (1 = A) that gets a QR code of the RSBSA number, or None
#   qr_size_pt:    printed size of the QR code (points; 72 pt = 1 inch)
FORMATS = [
    {"name": "AR", "template": BASE_DIR / "AR FORMAT.xlsx", "rows_per_page": 15,
     "qr_column": None},
    {"name": "STUB", "template": BASE_DIR / "STUB FORMAT.xlsx", "rows_per_page": 20,
     "qr_column": 1, "qr_size_pt": 22},
]

SORT_BY_SURNAME = False     # True = sort beneficiaries alphabetically inside each barangay
PAPER_SIZE = 14             # 14 = Folio / Long bond 8.5x13 in, 5 = Legal, 9 = A4, 1 = Letter
ORIENTATION = "landscape"
QR_LEFT_PT = 4              # gap between the cell's left border and the QR code

# Heading text -> field. Matching ignores case, spaces and punctuation, so
# "RSBSA NO." / "RSBSA No" / "rsbsa_no" all work. Used for both the masterlist
# and the template headings.
FIELD_ALIASES = {
    "number": ["no", "nos", "number", "count"],
    "rsbsa": ["rsbsano", "rsbsa", "rsbsanumber", "referenceno"],
    "surname": ["surname", "lastname", "familyname"],
    "first": ["firstname", "givenname", "first"],
    "middle": ["middlename", "middle", "mi"],
    "birthdate": ["birthdate", "dateofbirth", "birthday", "bday", "dob"],
    "barangay": ["barangay", "brgy"],
    "municipality": ["municipality", "citymunicipality", "town", "city"],
    "amount": ["amount"],
    "idtype": ["idtypeidnoofauthorizedclaimants", "idtype", "idno"],
}
REQUIRED = ("rsbsa", "surname", "first", "barangay")
EMU_PER_PT = 12700


def norm(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def match_field(heading, exact_only=False):
    h = norm(heading)
    if not h:
        return None
    for field, aliases in FIELD_ALIASES.items():
        if h in aliases:
            return field
    if exact_only:
        return None
    for field, aliases in FIELD_ALIASES.items():
        if field != "number" and any(h.startswith(a) for a in aliases if len(a) > 2):
            return field
    return None


# ---------------------------------------------------------------- template
class Template:
    def __init__(self, path):
        if not Path(path).exists():
            raise FileNotFoundError(f"template not found: {path}")
        self.ws = openpyxl.load_workbook(path).active
        ws = self.ws
        self.heading_row = next(
            (r for r in range(1, ws.max_row + 1)
             if any("rsbsa" in norm(c.value) for c in ws[r])), None)
        if self.heading_row is None:
            raise ValueError(f"no column-heading row (with 'RSBSA') in {Path(path).name}")
        self.data_row = self.heading_row + 1
        self.header_rows = list(range(1, self.heading_row + 1))
        self.footer_rows = list(range(self.data_row + 1, ws.max_row + 1))
        heads = [c for c in ws[self.heading_row] if c.value not in (None, "")]
        self.last_col = max(c.column for c in heads)
        self.max_col = max(ws.max_column, self.last_col)
        # which field goes into which column
        self.columns = {}
        for c in heads:
            field = match_field(c.value)
            if field and field not in self.columns.values():
                self.columns[c.column] = field

    def copy_row(self, src_row, dst_ws, dst_row, with_values=True, merges=None):
        """Copy one template row's formatting (and values) onto dst_row."""
        for c in range(1, self.max_col + 1):
            s, d = self.ws.cell(src_row, c), dst_ws.cell(dst_row, c)
            if with_values and s.value is not None:
                d.value = s.value
            if s.has_style:
                d.font, d.border, d.fill = copy(s.font), copy(s.border), copy(s.fill)
                d.alignment, d.protection = copy(s.alignment), copy(s.protection)
                d.number_format = s.number_format
        dst_ws.row_dimensions[dst_row].height = self.ws.row_dimensions[src_row].height
        if merges is not None:
            for m in self.ws.merged_cells.ranges:
                if m.min_row == src_row:
                    merges.append((dst_row, m.min_col, dst_row + m.max_row - m.min_row, m.max_col))


# ---------------------------------------------------------------- masterlist
def read_masterlist(path):
    """Return list of dicts (one per beneficiary) from the first sheet that has headings."""
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        for hr in range(1, min(ws.max_row, 30) + 1):
            heads = [(c.value, c.column) for c in ws[hr] if c.value is not None]
            if not any("rsbsa" in norm(v) for v, _ in heads):
                continue
            cols = {}
            for exact in (True, False):          # exact heading matches win over prefixes
                for v, col in heads:
                    field = match_field(v, exact_only=exact)
                    if field and field not in cols:
                        cols[field] = col
            missing = [f for f in REQUIRED if f not in cols]
            if missing:
                raise ValueError(f"missing column(s) {', '.join(missing)} in sheet '{ws.title}'")
            records = []
            for row in ws.iter_rows(min_row=hr + 1, values_only=True):
                rec = {f: (row[c - 1] if c - 1 < len(row) else None) for f, c in cols.items()}
                rec = {f: (v.strip() if isinstance(v, str) else v) for f, v in rec.items()}
                if not rec.get("rsbsa") and not rec.get("surname"):
                    continue  # blank / filler row
                if norm(rec.get("surname")) in ("surname", "total", "grandtotal"):
                    continue  # repeated heading or total line
                records.append(rec)
            return records, sorted(cols)
    raise ValueError("no heading row containing 'RSBSA' was found")


def group_by_barangay(records):
    """Group records by barangay, keeping the order barangays first appear in."""
    groups = {}
    for rec in records:
        groups.setdefault(str(rec.get("barangay") or "(NO BARANGAY)").upper(), []).append(rec)
    if SORT_BY_SURNAME:
        for recs in groups.values():
            recs.sort(key=lambda r: (norm(r.get("surname")), norm(r.get("first")), norm(r.get("middle"))))
    return groups


def paginate(groups, rows_per_page):
    """Split every barangay into pages of rows_per_page. Returns list of pages."""
    pages = []
    for recs in groups.values():
        for i in range(0, len(recs), rows_per_page):
            pages.append(recs[i:i + rows_per_page])
    return pages


# ---------------------------------------------------------------- QR code
def qr_png(text, used_crcs=None):
    """PNG of a QR code for text. Excel treats pictures whose files share a CRC32
    as the same picture (it then prints the wrong QR code), so when used_crcs is
    given, a small text tag is added to the file until its CRC32 is unique."""
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=4, border=1)
    qr.add_data(str(text))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").get_image()
    for attempt in range(1000):
        info = PngInfo()
        info.add_text("rsbsa", f"{text}#{attempt}" if attempt else str(text))
        buf = io.BytesIO()
        img.save(buf, format="PNG", pnginfo=info)
        crc = zlib.crc32(buf.getvalue())
        if used_crcs is None or crc not in used_crcs:
            break
    if used_crcs is not None:
        used_crcs.add(crc)
    buf.seek(0)
    return buf


def add_qr(ws, row, col, text, row_height_pt, size_pt, used_crcs):
    img = XLImage(qr_png(text, used_crcs))
    size = int(size_pt * EMU_PER_PT)
    top = max(0, int((row_height_pt - size_pt) / 2 * EMU_PER_PT))
    marker = AnchorMarker(col=col - 1, colOff=int(QR_LEFT_PT * EMU_PER_PT), row=row - 1, rowOff=top)
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(size, size))
    ws.add_image(img)


# ---------------------------------------------------------------- build
def cell_value(field, val):
    if val is None or val == "":
        return None
    if field == "birthdate" and isinstance(val, str):
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d/%m/%Y", "%b-%d-%Y", "%B %d, %Y"):
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                pass
    return val


def build(groups, tpl, fmt, out_path):
    pages = paginate(groups, fmt.get("rows_per_page", 15))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = fmt["name"]
    for key, dim in tpl.ws.column_dimensions.items():
        ws.column_dimensions[key].width = dim.width
    merges, row, number = [], 1, 0
    data_height = tpl.ws.row_dimensions[tpl.data_row].height or 15
    qr_col = fmt.get("qr_column")
    used_crcs = set()

    for p, page in enumerate(pages):
        for hr in tpl.header_rows:
            tpl.copy_row(hr, ws, row, merges=merges)
            row += 1
        for rec in page:
            number += 1
            tpl.copy_row(tpl.data_row, ws, row, with_values=False)
            for col, field in tpl.columns.items():
                val = number if field == "number" else cell_value(field, rec.get(field))
                if val is not None:
                    ws.cell(row, col).value = val
            if qr_col and rec.get("rsbsa"):
                add_qr(ws, row, qr_col, rec["rsbsa"], data_height, fmt.get("qr_size_pt", 22), used_crcs)
            row += 1
        if p < len(pages) - 1:
            ws.row_breaks.append(Break(id=row - 1))

    for fr in tpl.footer_rows:
        tpl.copy_row(fr, ws, row, merges=merges)
        row += 1
    for r1, c1, r2, c2 in merges:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    # print setup (matches the template PDFs): exactly 100% scale so borders stay
    # crisp and the 15 beneficiary rows fill the page.
    ws.print_area = f"A1:{get_column_letter(tpl.last_col)}{row - 1}"
    ws.page_setup.orientation = ORIENTATION
    ws.page_setup.paperSize = PAPER_SIZE
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.scale = 100
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = 0.38
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 8

    wb.save(out_path)
    return len(pages)


def clean_stem(path):
    """'La Trinidad Additional 10-2-26 - AR' -> 'La Trinidad Additional 10-2-26'."""
    return re.sub(r"[\s_-]+(AR|STUB)$", "", Path(path).stem, flags=re.I).strip() or Path(path).stem


def generate(source, filename, formats=None, log=print):
    """Format one masterlist in memory.
    source: path or file-like object of the masterlist .xlsx; filename: its name.
    Yields (output file name, xlsx bytes, page count) for every selected format."""
    records, found = read_masterlist(source)
    if not records:
        raise ValueError("no beneficiary rows found")
    groups = group_by_barangay(records)
    log(f"{Path(filename).name}: {len(records)} beneficiaries, {len(groups)} barangays")
    for fmt in formats or FORMATS:
        tpl = Template(fmt["template"])
        missing = sorted({f for f in tpl.columns.values() if f not in found and f != "number"})
        if missing:
            log(f"   note: masterlist has no {', '.join(m.upper() for m in missing)} column; "
                f"left blank in {fmt['name']}")
        buf = io.BytesIO()
        npages = build(groups, tpl, fmt, buf)
        yield f"{clean_stem(filename)} - {fmt['name']} FORMAT.xlsx", buf.getvalue(), npages, groups


def process(master_path, out_dir=OUTPUT_DIR, formats=None, log=print):
    """Format one masterlist into every selected format, saved in out_dir.
    Returns (list of output paths, barangay groups)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    outputs, groups = [], {}
    for name, data, npages, groups in generate(master_path, master_path, formats, log):
        out = out_dir / name
        try:
            out.write_bytes(data)
        except PermissionError:
            raise PermissionError(f"cannot write {out.name} - close it in Excel and try again")
        log(f"   {name.rsplit(' - ', 1)[-1][:-12]:<5} {npages:>4} pages -> {out}")
        outputs.append(out)
    return outputs, groups


def main(argv):
    files = [Path(a) for a in argv] or sorted(
        p for p in INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        print(f"No masterlists found. Put .xlsx files in: {INPUT_DIR}")
        return 1
    failed = 0
    for f in files:
        try:
            _, groups = process(f)
            for b, recs in groups.items():
                print(f"       {b:<20} {len(recs):>5}")
        except Exception as e:
            failed += 1
            print(f"[ERROR] {f.name}: {e}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
