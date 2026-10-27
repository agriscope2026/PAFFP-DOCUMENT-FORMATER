# PAFFP Formatter (System Upload + AR + STUB)

Turns beneficiary masterlists (any column order, missing columns allowed) into three files:

- `<name> - SYSTEM UPLOAD.xlsx` – one continuous list for uploading to the system
- `<name> - AR FORMAT.xlsx` – payroll, 15 per page, with signatories
- `<name> - STUB FORMAT.xlsx` – 20 per page with QR codes of the RSBSA number, no signatories

A workbook with several data sheets (e.g. UNCLAIMED, ADDITIONAL) gets one set per sheet.
If the masterlist has no AMOUNT, the default amount (2325, changeable in the app/web page) is used.

## New format (second tab)

Upload the source workbook (like `sample to format.xlsx`). Sheets named `STUB <MUNICIPALITY>`
(e.g. `STUB HUNGDUAN`) are used when the workbook has any; other sheets are then ignored.
Without them, every sheet with column headings is read instead (e.g. a system upload/export with
`rsbsa_no, first_name, last_name, farmer_address_bgy, farmer_address_mun, ...`): columns are found
by their heading text in any order (spellings: `FIELD_ALIASES` in `paffp_formatter.py`), and the
beneficiaries are split by the MUNICIPALITY column. Each municipality gets:

- `<MUNICIPALITY> - AR FORMAT.pdf` – layout of `AR FORMAT NEW.pdf`: 14 per page, signatories at the end
- `<MUNICIPALITY> - STUB FORMAT.pdf` – layout of `STUB FORMAT NEW.xlsx`: 12 per page, QR code of the
  RSBSA NO. in its own last column (QR CODE), left/right margins from the template (0.75 in)
- `<MUNICIPALITY> - SYSTEM UPLOAD.xlsx` – `SYSTEM UPLOAD.xlsx` filled in (EXT. NAME is added to First Name)

Every barangay starts on a new page. The **Province** box on the tab sets the Province line in the
AR and STUB heading (e.g. IFUGAO); left empty, the Province written on each STUB sheet is used.
Region comes from each STUB sheet's title rows.
Code: `newformat.py` (desktop and `api/newformat.py` on the web). Signatory names and the AR layout
are at the top of `newformat.py`. Fonts: Cambria on Windows; on the web the bundled Caladea
(`fonts/`, SIL Open Font License), which is designed to match Cambria.
Desktop needs: `pip install pymupdf qrcode fonttools`.

| Use | How |
|---|---|
| Desktop | Double-click `PAFFP FORMATTER.bat`, pick masterlists, press **CREATE**. Files go to `OUTPUT/`. |
| Command line | `python paffp_formatter.py "masterlist.xlsx"` (no arguments = every file in `MASTERLISTS/`) |
| Web (Vercel) | Open the site, add masterlists, press **Create**; each downloads as a .zip. |

Templates: `System Uploading Template.xlsx`, `AR FORMAT.xlsx`, `STUB FORMAT.xlsx` — edit their text
(signatories, Region/Province) in Excel; keep them empty of beneficiary data.
Rows per page, QR size, default amount and recognised heading spellings: top of `paffp_formatter.py`.

## Deploy to Vercel

Project layout used by Vercel:

- `public/index.html` – the web page
- `api/generate.py` – Python function (POST a masterlist, returns a zip)
- `api/newformat.py` – Python function for the New format tab (reads the workbook, renders one file per request)
- `public/vendor/jszip.min.js` – zips the New format files in the browser
- `paffp_formatter.py` and the three templates – bundled into the function (`vercel.json`)
- `requirements.txt` – Python packages
- `.vercelignore` / `.gitignore` – keep `MASTERLISTS/`, `OUTPUT/`, `_backup/` and PDFs off the server and out of Git

**Option A – Vercel CLI** (from this folder):

```
npx vercel login
npx vercel          # preview deployment; accept the defaults (framework: Other)
npx vercel --prod   # production
```

**Option B – GitHub:** push this folder to a GitHub repository and import it at vercel.com/new
(Framework Preset: **Other**, no build command).

**Protect the site.** Masterlists contain personal data. In the Vercel project go to
*Settings → Deployment Protection* and turn on **Vercel Authentication** (or password protection)
so only your team can open it. Uploaded files are processed in memory and not stored.

Limits: one masterlist per request, up to 4 MB (a 906-beneficiary list is ~0.25 MB and takes ~6 s).
