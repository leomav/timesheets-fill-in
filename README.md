# Timesheet Filler

A small web tool that fills in the monthly NTUA / EUREKA3D-XR timesheets for you.
Instead of typing hours into every cell by hand, you drop in the empty `.xlsx`
files, say how many hours go to each work package, and download the completed
sheets — ready to sign.

Everything runs **inside your browser**. The files never get uploaded anywhere,
so the names, VAT numbers and tax IDs on the sheets stay on your own computer.

---

## For users — how to use it

1. **Open the tool.** Go to the published page
   (`https://leomav.github.io/timesheets-fill-in/`).
   The first time, give it a few seconds to load the Excel engine.

2. **Add your timesheets.** Click **Choose timesheet files** (or drag them onto
   the box) and pick the empty monthly files you were given, e.g.
   `..._2026_05.xlsx`. You can add several months at once.

3. **Enter the hours.** Each file becomes a card showing the month, how many
   working days it has, and any public holidays already marked on the sheet.
   For every work package, type its name (e.g. `WP3`) and the number of hours.
   Use **+ add work package** if a month has more than one.

   > Example: your coordinator says *“put 120 hours for WP3 and 45 for WP4 in
   > June”* → on the June card, add a `WP3 / 120` row and a `WP4 / 45` row.

4. **Watch the total.** Each card shows `Total: Xh / Yh`. If you go over the
   month's capacity it turns red and blocks generating — lower the hours.

5. **Generate & download.** Click the button and the filled `.xlsx` downloads
   with the same filename. Open it, check it, sign it.

### How the hours get placed

- Hours are spread **evenly across the working days**, at most **8 hours a day**.
- **Weekends and public holidays are skipped.** Any day the sheet already marks
  in its notes column (e.g. *Πρωτομαγιά*, *Αγίου Πνεύματος*, *Δευτέρα του Πάσχα*)
  is treated as a day off and left empty.
- Hours go in the **Research** column; the work-package name goes in the
  **Comments** column — exactly like the reference template.
- Each work package gets its own block of days, so a single day is never split
  between two packages.

> **Tip:** the tool trusts the holidays already printed on the sheet. If a month
> you receive is missing a holiday it should have, add that note to the sheet
> before dropping it in.

---

## For developers — the technical side

A dependency-free static site. The interesting part is that it runs the Python
Excel library **openpyxl in the browser** via
[Pyodide](https://pyodide.org/) (WebAssembly), so there is no backend and files
are processed entirely client-side.

### Why this design

- **openpyxl instead of a JS Excel library.** These sheets carry an integrity
  token in cell `O2` plus formulas in columns K/M. openpyxl writes only the cell
  *values* it's told to and leaves everything else byte-for-byte intact; common
  JS libraries tend to rewrite styles/formulas and would break the token.
- **In-browser (Pyodide) instead of a server.** The sheets contain personal data.
  Keeping everything client-side means nothing is uploaded, and hosting is just
  static files (GitHub Pages).

### Files

| File | Role |
|------|------|
| `index.html` / `style.css` | UI |
| `app.js` | Boots Pyodide, installs openpyxl, bridges JS ↔ Python, handles upload/download |
| `fill.py` | All timesheet logic — **single source of truth**; the browser fetches this exact file |
| `tests.py` | CPython test suite for `fill.py` |
| `.github/workflows/deploy.yml` | Builds nothing; publishes the repo to GitHub Pages |
| `.nojekyll` | Tells Pages to serve every file as-is |

### The sheet layout (`fill.py` constants)

Day rows are 13–43. Column C (Research) receives the hours, column J (Comments)
holds the WP label or a pre-printed holiday, columns K/M are formulas, `O2`/`O1`
are the integrity token. A day counts as workable only if it's a weekday **and**
its column J is empty.

### Run locally

Pyodide fetches `fill.py`, so `file://` won't work — serve over HTTP:

```bash
python3 -m http.server 8000
# open http://localhost:8000
```

### Run the tests

The tests read the real sample workbooks, which are **not committed** (personal
data — see `.gitignore`). Place the empty `.xlsx` files under `cv_examples/`
locally, then:

```bash
pip install openpyxl
python3 tests.py
```

### Deployment

Pushing to `main` triggers `.github/workflows/deploy.yml`, which publishes the
site to GitHub Pages. Enable it once under **Settings → Pages → Build and
deployment → GitHub Actions**.
