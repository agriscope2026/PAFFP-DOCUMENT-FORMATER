# PAFFP Formatter (System Upload + AR + STUB)

Turns beneficiary masterlists (any column order, missing columns allowed) into three files:

- `<name> - SYSTEM UPLOAD.xlsx` – one continuous list for uploading to the system
- `<name> - AR FORMAT.xlsx` – payroll, 15 per page, with signatories
- `<name> - STUB FORMAT.xlsx` – 20 per page with QR codes of the RSBSA number, no signatories

A workbook with several data sheets (e.g. UNCLAIMED, ADDITIONAL) gets one set per sheet.
If the masterlist has no AMOUNT, the default amount (2325, changeable in the app/web page) is used.

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
