# Bill Categorizer

Desktop app for Sunation Energy that:

1. Loads a local employee list (`data/employees.xlsx`).
2. Pulls the live Salesforce report on-demand (report ID `00OUX000006He412AC`) to keep names and departments fresh.
3. Accepts a bill drag-and-dropped into the window — **PDF, Excel/CSV, or scanned image**.
4. Matches every line item to an employee by name (fuzzy-tolerant), looks up the department, and sums costs into the five categories:
   **Residential · Commercial · Service · Roofing · Executive**.
5. Flags unmatched lines so they can be reviewed manually.
6. Exports the breakdown to **Excel** and **PDF**.

Cross-platform — runs on **macOS, Windows, and Linux** from the same Python source.

## Category mapping

Per the rule *"Maintenance, IT, Finance, Sales all fall into Residential"*, the default mapping is:

| Category     | Departments |
|--------------|--------------|
| Residential  | Residential Engineering, Residential Installation, Residential Sales, Finance, Maintenance, Information Systems, Human Resources, Marketing, Operations, Processing, Procurement, Lead Qualification, Corporate Office, Warehouse |
| Commercial   | Commercial, Commercial Engineering, Commercial Installation, Commercial Project Management, Commercial Sales |
| Service      | Service Field, Service Office |
| Roofing      | Roofing |
| Executive    | Executive Office |

Edit `category_map.json` to change any of these — no rebuild needed.

## One-time setup

1. Install **Python 3.10 or later** — https://www.python.org/downloads/ (on macOS, check that `python3 --version` works in Terminal).
2. (Optional — only if you want OCR for scanned image bills) install Tesseract:
   - macOS: `brew install tesseract`
   - Windows: https://github.com/UB-Mannheim/tesseract/wiki
3. Open `config.json` and fill in your Salesforce credentials:
   - `username` — your Salesforce login email
   - `password` — your Salesforce password
   - `security_token` — reset via Salesforce → *Settings → My Personal Information → Reset My Security Token*; it's emailed to you
   - `domain` — `"login"` for production, `"test"` for sandbox
4. (Optional) replace `assets/logo.png` with the official Sunation logo file (same filename).

## Run the app

### macOS
Double-click **`run.command`** in Finder. The first launch creates a virtualenv and installs dependencies (~2 min); every launch after that is instant.

If macOS blocks it with "cannot be opened because it is from an unidentified developer":
- Right-click `run.command` → **Open** → **Open** in the confirmation dialog. (One-time only.)

### Windows
Double-click **`run.bat`**.

### Linux
From a terminal: `./run.sh`

## Distribute to Windows users via USB (GitHub Actions)

A GitHub Actions workflow builds `BillCategorizer.exe` on Windows automatically — no Windows machine required on your end.

**One-time setup (on your Mac):**

```bash
cd BillCategorizer
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/<your-user>/<your-repo>.git
git branch -M main
git push -u origin main
```

**Every time you push to `main`:** GitHub runs `.github/workflows/build.yml`, builds the `.exe` on a Windows runner, and attaches `BillCategorizer-windows.zip` as a downloadable artifact.

**Download the artifact:**
1. On GitHub, open your repo → **Actions** tab.
2. Click the latest "Build Windows EXE" run.
3. Scroll to the bottom → **Artifacts** → **BillCategorizer-windows** → downloads as a zip.

**Publish a versioned release** (optional, cleaner for handing out to users):
```bash
git tag v1.0.0
git push --tags
```
The workflow publishes a GitHub Release with the zip attached — users get a permanent download link.

**Put it on a USB:**
1. Download `BillCategorizer-windows.zip` from the Actions artifact.
2. Unzip. The folder contains `BillCategorizer.exe`, `config.json`, `category_map.json`, `data\`, `assets\`, `reports\`, and a brief `HOW-TO-RUN.txt`.
3. (Optional, if you want the USB pre-configured) open `config.json` in a text editor and fill in your Salesforce credentials — otherwise each user will need to do this themselves.
4. Copy the entire folder to the USB stick.
5. Each Windows user double-clicks `BillCategorizer.exe`. First launch may show a blue SmartScreen warning — click **More info → Run anyway**.

**Important: never commit `config.json` with real credentials to GitHub.** The template that ships in the repo has empty credential fields; keep it that way. Fill in credentials only after downloading the built zip, or use per-user config on the USB.

## Optional: build a standalone executable locally

You don't need this if `run.command` / `run.bat` works for you — but if you want a single double-clickable app that doesn't depend on a system Python, use these:

### macOS — produces `dist/BillCategorizer.app`
```
./build.sh
```

### Windows — produces `dist\BillCategorizer.exe`
```
build.bat
```

### Linux — produces `dist/BillCategorizer`
```
./build.sh
```

**Important:** PyInstaller doesn't cross-compile. Run `build.sh` on a Mac to get the `.app`, run `build.bat` on Windows to get the `.exe`. The `.app` produced on your Mac won't run on Windows machines, and vice versa.

## Day-to-day use

1. Launch the app.
2. Click **Refresh from Salesforce** to pull the latest report (new hires, department moves, departures will update the local EE list). Do this as often as you need.
3. Drag a bill onto the window (or use **Open Bill…**).
4. Review the totals-by-category table and the unmatched-lines panel.
5. Click **Export Excel** or **Export PDF** to save a shareable report into the `reports/` folder (or wherever you choose).

## Files

| File | Purpose |
|------|---------|
| `app.py` | Main GUI |
| `sf_client.py` | Salesforce REST API client |
| `ee_matcher.py` | Employee list loader + fuzzy name matching |
| `bill_parser.py` | PDF / XLSX / CSV / image parsers |
| `categorizer.py` | Department → category mapping |
| `exporters.py` | Excel and PDF export |
| `config.json` | Salesforce credentials + app settings (edit me) |
| `category_map.json` | Department-to-category map (edit me) |
| `data/employees.xlsx` | Local employee list (auto-refreshed by Salesforce pull) |
| `assets/logo.png` | Logo shown in the header and on the PDF report |
| `run.command` | macOS double-click launcher |
| `run.bat` | Windows double-click launcher |
| `run.sh` | Linux launcher |
| `build.sh` | Build a standalone .app / Linux binary |
| `build.bat` | Build a standalone .exe on Windows |
| `requirements.txt` | Python dependencies |

## Troubleshooting

- **"Python is not installed"** — install Python 3.10+ from python.org and try again.
- **macOS "cannot be opened" warning** — right-click `run.command` → **Open** once to approve.
- **"Salesforce credentials are empty"** — fill in `config.json`.
- **`INVALID_LOGIN`** — your security token is stale or has a stray space; reset and re-paste from the email.
- **OCR fails on scanned bills** — install Tesseract (see step 2).
- **Drag-and-drop doesn't work** — use the **Open Bill…** button; it's identical.
- **Too many false matches** — raise `fuzzy_match_threshold` in `config.json` (default 85, try 90).
- **Too many unmatched names** — lower it (try 75).
