"""
PAFFP FORMATTER  (SYSTEM UPLOAD + AR + STUB)
============================================
Turns beneficiary masterlists (any column layout) into three files:
    <name> - SYSTEM UPLOAD.xlsx   flat list for uploading to the system
    <name> - AR FORMAT.xlsx       payroll, 15 per page, with signatories
    <name> - STUB FORMAT.xlsx     20 per page, QR code of the RSBSA number

HOW TO USE
  - Easiest: double-click "PAFFP FORMATTER.bat" and pick the masterlist files.
  - Command line:  python paffp_formatter.py "masterlist1.xlsx" "masterlist2.xlsx"
    (with no files given, every .xlsx in the MASTERLISTS folder is formatted)
  - Web: see README.md (Vercel).
  A workbook with several data sheets (e.g. UNCLAIMED, ADDITIONAL) gives one set
  of files per sheet: "<file> - <sheet> - AR FORMAT.xlsx", ...

MASTERLISTS
  Columns are matched by their heading text, in any order (see FIELD_ALIASES).
  Missing columns are left blank (AMOUNT uses DEFAULT_AMOUNT). A name
  extension column (JR, SR, III) is added to the first name. A single
  full-name column ("DELA CRUZ, JUAN P.") is split into its parts.

TEMPLATES ("System Uploading Template.xlsx", "AR FORMAT.xlsx", "STUB FORMAT.xlsx")
  - Rows 1..N   : header block; the last of these rows holds the column headings.
  - Row N+1     : one EMPTY formatted row, used as the format of every beneficiary row.
  - Rows after  : (optional) footer such as signatories, copied once at the end.
  Each template column is filled by matching its heading. Edit text such as
  signatory names directly in the templates; changes are picked up automatically.

RULES APPLIED
  - Every barangay starts on a new page; NO. continues 1, 2, 3 ... to the end.
  - AR/STUB: the header block repeats on every page, which holds exactly
    rows_per_page rows (AR 15, STUB 20) except a barangay's last page.
  - SYSTEM UPLOAD: one continuous list, headings repeat when printed.
  - AR only: signatories after the last beneficiary. STUB: rows + QR codes only.
"""
import io
import re
import sys
import warnings
import zlib
from copy import copy
from datetime import datetime
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
#   name:          short id (used by the app/web checkboxes)
#   suffix:        added to the output file name
#   layout:        "pages" = header block repeated on every page (AR/STUB),
#                  "list"  = one heading row, continuous list (system upload)
#   rows_per_page: beneficiaries per page for "pages" (a barangay's last page may have fewer)
#   qr_column:     column (1 = A) that gets a QR code of the RSBSA number, or None
#   qr_size_pt:    printed size of the QR code (points; 72 pt = 1 inch)
FORMATS = [
    {"name": "SYSTEM", "label": "System upload", "suffix": "SYSTEM UPLOAD",
     "template": BASE_DIR / "System Uploading Template.xlsx", "layout": "list"},
    {"name": "AR", "label": "AR format", "suffix": "AR FORMAT",
     "template": BASE_DIR / "AR FORMAT.xlsx", "layout": "pages", "rows_per_page": 15},
    {"name": "STUB", "label": "STUB format", "suffix": "STUB FORMAT",
     "template": BASE_DIR / "STUB FORMAT.xlsx", "layout": "pages", "rows_per_page": 20,
     "qr_column": 1, "qr_size_pt": 22},
]

DEFAULT_AMOUNT = 2325       # used when the masterlist has no AMOUNT (None = leave blank)
SORT_BY_SURNAME = False     # True = sort beneficiaries alphabetically inside each barangay
PAPER_SIZE = 14             # 14 = Folio / Long bond 8.5x13 in, 5 = Legal, 9 = A4, 1 = Letter
ORIENTATION = "landscape"
QR_LEFT_PT = 4              # gap between the cell's left border and the QR code

# Heading text -> field. Matching ignores case, spaces and punctuation, so
# "RSBSA NO." / "RSBSA No" / "rsbsa_no" all work, and columns may be in any
# order. Used for both the masterlist and the template headings.
# To support a new heading spelling, add it (lower case, letters/digits only).
FIELD_ALIASES = {
    "number": ["no", "nos", "number", "count", "seqno", "itemno"],
    "rsbsa": ["rsbsano", "rsbsa", "rsbsanumber", "rsbsaid", "rsbsarefno", "rsbsareferenceno",
              "rsbsareferencenumber", "rsbsasystemgeneratedno", "referenceno", "referencenumber", "refno"],
    "surname": ["surname", "lastname", "familyname", "lname", "apelyido"],
    "first": ["firstname", "givenname", "first", "fname", "pangalan"],
    "middle": ["middlename", "middle", "mi", "middleinitial", "mname"],
    "ext": ["extname", "nameextension", "extensionname", "nameext", "ext", "extension", "suffix",
            "qualifier"],
    "fullname": ["name", "fullname", "completename", "nameofbeneficiary", "beneficiaryname",
                 "beneficiary", "nameoffarmer", "farmername", "nameoffarmerfisherfolk", "names"],
    "birthdate": ["birthdate", "dateofbirth", "birthday", "bday", "dob", "bdate"],
    "barangay": ["barangay", "brgy", "barangayname", "nameofbarangay", "bgy", "farmeraddressbgy",
                 "addressbgy", "addressbrgy", "addressbarangay", "farmeraddressbarangay"],
    "municipality": ["municipality", "citymunicipality", "municipalitycity", "cityormunicipality",
                     "town", "city", "mun", "municipal", "lgu", "farmeraddressmun", "addressmun",
                     "addressmunicipality", "farmeraddressmunicipality", "addresscitymun"],
    "amount": ["amount", "amt", "amountreceived", "cashassistance"],
    "idtype": ["idtypeidnoofauthorizedclaimants", "idtype", "idno", "validid"],
}
NAME_FIELDS = ("rsbsa", "surname", "first", "fullname")   # a masterlist needs at least one
LABELS = {"number": "NO", "rsbsa": "RSBSA NO.", "surname": "SURNAME", "first": "FIRST NAME",
          "middle": "MIDDLE NAME", "ext": "NAME EXT.", "birthdate": "BIRTHDATE",
          "barangay": "BARANGAY", "municipality": "MUNICIPALITY", "amount": "AMOUNT",
          "idtype": "ID TYPE"}
NO_NOTE = {"number", "idtype"}   # columns that are normally empty / generated
EMU_PER_PT = 12700


def norm(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def match_field(heading, level=2):
    """Field for a heading. level 0 = exact alias only, 1 = also heading starts
    with an alias ("BARANGAY NAME"), 2 = also alias inside heading ("NAME OF BRGY")."""
    h = norm(heading)
    if not h:
        return None
    for field, aliases in FIELD_ALIASES.items():
        if h in aliases:
            return field
    loose = [(f, a) for f, a in FIELD_ALIASES.items() if f not in ("number", "fullname", "ext")]
    if level >= 1:
        for field, aliases in loose:
            if any(h.startswith(a) for a in aliases if len(a) > 3):
                return field
    if level >= 2:
        for field, aliases in loose:
            if any(a in h for a in aliases if len(a) > 4):
                return field
    return None


def map_columns(cells):
    """cells: [(heading text, column)]. Returns {field: column}; exact matches win
    over looser ones, and each field/column is used once."""
    cols, used = {}, set()
    for level in (0, 1, 2):
        for text, col in cells:
            field = match_field(text, level)
            if field and field not in cols and col not in used:
                cols[field] = col
                used.add(col)
    return cols


def find_heading_row(ws, max_scan=30):
    """Row (within the first max_scan rows) whose cells match the most fields."""
    best = (0, None, {})
    for r in range(1, min(ws.max_row, max_scan) + 1):
        cells = [(c.value, c.column) for c in ws[r] if c.value not in (None, "")]
        cols = map_columns(cells)
        score = len(cols) + (2 if "rsbsa" in cols else 0)
        if len(cols) >= 2 and any(f in cols for f in NAME_FIELDS) and score > best[0]:
            best = (score, r, cols)
    return best[1], best[2]


# ---------------------------------------------------------------- template
class Template:
    def __init__(self, path):
        if not Path(path).exists():
            raise FileNotFoundError(f"template not found: {Path(path).name}")
        self.ws = openpyxl.load_workbook(path).active
        ws = self.ws
        self.heading_row, cols = find_heading_row(ws, max_scan=ws.max_row)
        if self.heading_row is None:
            raise ValueError(f"no column-heading row found in {Path(path).name}")
        self.data_row = self.heading_row + 1
        self.header_rows = list(range(1, self.heading_row + 1))
        self.footer_rows = list(range(self.data_row + 1, ws.max_row + 1))
        heads = [c for c in ws[self.heading_row] if c.value not in (None, "")]
        self.last_col = max(c.column for c in heads)
        self.max_col = max(ws.max_column, self.last_col)
        # which field goes into which template column
        self.columns = {col: field for field, col in cols.items()}

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
def split_full_name(name):
    """'DELA CRUZ, JUAN P.' -> (DELA CRUZ, JUAN, P.); 'JUAN P. DELA CRUZ' is split
    as first words / last word."""
    name = re.sub(r"\s+", " ", str(name or "")).strip()
    if not name:
        return None, None, None
    if "," in name:
        surname, rest = [s.strip() for s in name.split(",", 1)]
        words = rest.split()
        if len(words) > 1 and (len(words[-1].rstrip(".")) <= 2 or words[-1].endswith(".")):
            return surname, " ".join(words[:-1]), words[-1]
        return surname, rest, None
    words = name.split()
    if len(words) == 1:
        return words[0], None, None
    return words[-1], " ".join(words[:-1]), None


def read_sheet(ws, hr, cols):
    """Beneficiary records of one sheet. Returns (records, sorted fields found)."""
    heading_text = {f: norm(ws.cell(hr, c).value) for f, c in cols.items()}
    records = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        rec = {f: (row[c - 1] if c - 1 < len(row) else None) for f, c in cols.items()}
        rec = {f: (re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v) for f, v in rec.items()}
        rec = {f: (None if v == "" else v) for f, v in rec.items()}
        if not any(rec.get(f) for f in NAME_FIELDS):
            continue  # blank / filler row
        if any(norm(rec.get(f)) == heading_text[f] for f in NAME_FIELDS if f in rec):
            continue  # heading repeated inside the list
        if any(norm(rec.get(f)) in ("total", "grandtotal", "subtotal") for f in NAME_FIELDS + ("number",)):
            continue  # total line
        if "fullname" in rec and not rec.get("surname") and not rec.get("first"):
            s, f, m = split_full_name(rec["fullname"])
            rec["surname"], rec["first"] = s, f
            rec["middle"] = rec.get("middle") or m
        ext = str(rec.get("ext") or "").strip()
        if ext and ext.upper() not in ("N/A", "NA", "NONE", "-") and rec.get("first"):
            first = str(rec["first"])
            if not norm(first).endswith(norm(ext)):
                rec["first"] = f"{first} {ext}"
        records.append(rec)

    found = set(cols)
    if "fullname" in found:
        found |= {"surname", "first"}
        if any(r.get("middle") for r in records):
            found.add("middle")
    return records, sorted(found)


def read_masterlist(source):
    """All data sheets of a masterlist workbook, in sheet order.
    Columns may be in any order and any may be missing.
    Returns [(sheet title, records, fields found)] (sheets without beneficiaries skipped)."""
    wb = openpyxl.load_workbook(source, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        hr, cols = find_heading_row(ws)
        if hr:
            records, found = read_sheet(ws, hr, cols)
            if records:
                sheets.append((ws.title, records, found))
    if not sheets:
        raise ValueError("could not find beneficiaries under column headings "
                         "(e.g. RSBSA NO., SURNAME, FIRST NAME)")
    return sheets


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
    """Split every barangay into pages of rows_per_page (None = whole barangay)."""
    pages = []
    for recs in groups.values():
        step = rows_per_page or len(recs)
        for i in range(0, len(recs), step):
            pages.append(recs[i:i + step])
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


def copy_print_setup(src, dst):
    """Use the template's own page setup (system upload template)."""
    for attr in ("orientation", "paperSize", "scale", "fitToWidth", "fitToHeight"):
        val = getattr(src.page_setup, attr)
        if val is not None:
            setattr(dst.page_setup, attr, val)
    dst.sheet_properties.pageSetUpPr.fitToPage = bool(src.sheet_properties.pageSetUpPr
                                                      and src.sheet_properties.pageSetUpPr.fitToPage)
    for attr in ("left", "right", "top", "bottom", "header", "footer"):
        setattr(dst.page_margins, attr, getattr(src.page_margins, attr))
    dst.print_options.horizontalCentered = src.print_options.horizontalCentered


def build(groups, tpl, fmt, out):
    """Write one output workbook to out (path or file-like). Returns a short summary."""
    paged = fmt.get("layout", "pages") == "pages"
    pages = paginate(groups, fmt.get("rows_per_page") if paged else None)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = fmt["name"] if paged else tpl.ws.title
    for key, dim in tpl.ws.column_dimensions.items():
        ws.column_dimensions[key].width = dim.width
    merges, row, number = [], 1, 0
    data_height = tpl.ws.row_dimensions[tpl.data_row].height or 15
    qr_col = fmt.get("qr_column")
    used_crcs = set()

    for p, page in enumerate(pages):
        if paged or p == 0:
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

    ws.print_area = f"A1:{get_column_letter(tpl.last_col)}{row - 1}"
    if paged:
        # matches the template PDFs: exactly 100% scale so borders stay crisp
        # and the beneficiary rows fill the page
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
    else:
        copy_print_setup(tpl.ws, ws)
        ws.print_title_rows = f"{tpl.header_rows[0]}:{tpl.header_rows[-1]}"
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 8

    wb.save(out)
    return f"{len(pages)} pages" if paged else f"{number} rows"


def clean_stem(path):
    """'La Trinidad Additional 10-2-26 - AR' -> 'La Trinidad Additional 10-2-26'."""
    stem = Path(str(path)).stem
    return re.sub(r"[\s_-]+(AR|STUB)$", "", stem, flags=re.I).strip() or stem


def safe_name(text):
    return re.sub(r'[\\/:*?"<>|]+', "-", str(text)).strip()


def generate(source, filename, formats=None, log=print, default_amount=DEFAULT_AMOUNT):
    """Format one masterlist workbook in memory.
    source: path or file-like object of the .xlsx; filename: its name.
    Yields (output file name, xlsx bytes, summary, barangay groups) for every
    data sheet x selected format."""
    sheets = read_masterlist(source)
    stem = clean_stem(filename)
    templates = {f["name"]: Template(f["template"]) for f in formats or FORMATS}
    for title, records, found in sheets:
        found = list(found)
        base = stem if len(sheets) == 1 else f"{stem} - {safe_name(title)}"
        groups = group_by_barangay(records)
        log(f"{Path(str(filename)).name}" + (f" [{title}]" if len(sheets) > 1 else "")
            + f": {len(records)} beneficiaries, {len(groups)} barangays")
        log(f"   columns found: {', '.join(LABELS[f] for f in found if f in LABELS)}")
        if default_amount not in (None, "", 0):
            blanks = [r for r in records if r.get("amount") in (None, "")]
            for r in blanks:
                r["amount"] = default_amount
            if blanks:
                log(f"   AMOUNT {default_amount:,} used for {len(blanks)} beneficiar"
                    f"{'y' if len(blanks) == 1 else 'ies'} without an amount")
                found.append("amount")
        for fmt in formats or FORMATS:
            tpl = templates[fmt["name"]]
            missing = [f for f in tpl.columns.values() if f not in found and f not in NO_NOTE]
            if missing:
                log(f"   note: masterlist has no {', '.join(LABELS[m] for m in missing)} column; "
                    f"left blank in {fmt['label']}")
            buf = io.BytesIO()
            summary = build(groups, tpl, fmt, buf)
            yield f"{base} - {fmt['suffix']}.xlsx", buf.getvalue(), summary, groups


def process(master_path, out_dir=OUTPUT_DIR, formats=None, log=print, default_amount=DEFAULT_AMOUNT):
    """Format one masterlist workbook into every selected format, saved in out_dir.
    Returns (list of output paths, barangay groups of the last sheet)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    outputs, groups = [], {}
    for name, data, summary, groups in generate(master_path, master_path, formats, log, default_amount):
        out = out_dir / name
        try:
            out.write_bytes(data)
        except PermissionError:
            raise PermissionError(f"cannot write {out.name} - close it in Excel and try again")
        log(f"   {summary:>10} -> {out.name}")
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
            process(f)
        except Exception as e:
            failed += 1
            print(f"[ERROR] {f.name}: {e}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
