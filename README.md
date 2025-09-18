README_CONTENT = """# IBI-PDF Report Generator

This project generates a **multi-page Hebrew PDF report** for credit-portfolio monitoring.  
It loads quarterly Excel workbooks, computes KPIs and portfolio metrics, and renders branded slides (PNG + PDF) with **tables, donut charts, and comparison blocks** using **Pillow**.

> **Language / RTL**: The report is in Hebrew, right-to-left. All text is passed through `fix_hebrew()` (uses `arabic_reshaper` + `bidi`) to ensure correct rendering and alignment.

---

## Report contents

For each **case** (`case_id`) the pipeline builds up to **six slides**:

1. **Executive Overview** — KPIs block over a template background.  
2. **Reclassification of Securities** — two aligned blocks (previous vs. current quarter) with counts and percentages.  
3. **Material Exposures – Periodic Comparison** — three side-by-side tables (Borrower / Sectors / Borrower Groups) comparing current and previous quarters.  
4. **Problematic Debts – Detailed Table** — includes a `סיווג פורום חוב` column; always shows rows like **“מסופק”** and **“Total”**.  
5. **Problematic Debts – Donut charts** — three donuts (Geography, Collateral, Liquidity) plus a segmented selector.  
6. **White slide with notes / secondary metrics** — used for displaying additional calculations such as *“אחוז מתיק האשראי הלא מוחרג”*.  

All slides include **branding** (logo + footer), with safe margins so content never collides with the footer.

---

## Data inputs

By default the script looks in `DataToPDF/`:

- `Data.xlsx` — **current** quarter  
- `DataOld.xlsx` — **previous** quarter  
- `DataOld2Q.xlsx` — **two quarters back**  
- `clients_cases.xlsx` — mapping file for client-to-case configuration

**Excel sheets must contain Hebrew columns**, such as:
- מנפיק/לווה: `תאור מנפיק` / `תיאור מנפיק` / `לווה`
- ענף: `תאור ענף` / `תיאור ענף` / `ענף`
- קבוצת לווים: `תאור קבוצת לווים` / `שם קבוצת לווים` / `קבוצת לווים`
- ערכים מספריים: `שווי נייר`, `סכום`, `אחוזים`, וכו׳

If column names differ, extend the candidate lists in `_pick_col(...)` so the script can resolve them dynamically.

---

## Rendering pipeline (high-level)

- **Canvas**: fixed `1600×900` (16:9).  
- **Drawing engine**: Pillow (`Image`, `ImageDraw`, `ImageFont`).  
- **RTL text**: via `fix_hebrew()` → `arabic_reshaper` + `bidi.get_display`.  
- **Tables**: `_draw_table` / `_draw_table_full` — responsive headers/rows, auto-fit text, ellipsizing when needed, consistent grid lines.  
- **Donuts**: `_draw_donut` draws percentage rings; data prepared in `_donut_config_for_bucket()`.  
- **Branding**: `add_branding()` puts `branding/logo.png` top-left and `branding/about.png` bottom footer.  
- **Safe bottom margin**: controlled by `SAFE_BOTTOM` so tables/charts never overlap footer.  

---

## Installing & running

**Requirements**

- Python **3.12+** (works on 3.10+, tested on 3.12/3.13)  
- Packages: `pillow`, `pandas`, `openpyxl`, `PyPDF2`, `arabic-reshaper`, `python-bidi`

```powershell
pip install -r requirements.txt
# or, if requirements.txt is missing, minimally:
pip install pillow pandas openpyxl pypdf2 arabic-reshaper python-bidi
```

**Run for one or more case IDs**
```powershell
python main_ibi.py 16396 16397 16398
```

**Run for all clients**
```powershell
python main_ibi.py
```
Outputs go to `outputs/`:
- `output_<case_id>.png` — first slide as PNG
- `output_<case_id>.pdf` — multi‑page PDF per case
- `combined_reports<client name>.pdf` — merged PDF of each clients

---

## Where to put branding

Create a folder:
```
branding/
  logo.png    # company logo (transparent PNG recommended)
  about.png   # footer strip / about block (PNG)
```

---

## Layout & spacing

- **Canvas:** `1600 × 900` px (16:9). All coordinates assume this size.
- **Safe bottom (footer):** kept via `SAFE_BOTTOM` and `add_branding(..., pad_bottom=26)`.
- **Tables spacing:** Between tables we use 18–24 px; between the last table and the footer we ensure ≥ 151 px**.
- **Side‑by‑side blocks:** In “Reclassification of Securities”, the two blocks are horizontally aligned using fixed center anchors `left_x = W//2 - 400`, `right_x = W//2 + 400` and identical box widths.

---

## Data mapping, filters & logic

- **Missing columns** → `KeyError: Column(s) [None] do not exist`  
  Add or rename source columns, or extend the candidate list in `_find_col(...)`.
- **Problematic debts filter** → only includes categories in BAD_TYPES = {"השגחה מיוחדת","במעקב מיוחד","מסופק","בפיגור"}.
- **Text overflow in headers** → The code uses a fit helper to shrink or ellipsize. You can tune header font size or the `max_width` passed to the helper.
- **RTL shows reversed (e.g., "ריינ רואת")** → Make sure every cell text goes through `fix_hebrew(text)` before drawing.
- **Excluded securities** → defined in EXCLUDED_SUB_AFIK set (list of numeric codes).
- **Metrics per case** → calculated within each case (not globally).
- **Quarterly comparison** → current quarter (Data.xlsx) vs. previous quarter (DataOld.xlsx) — plus optional two-quarters back.
  

---

## Troubleshooting

- `AttributeError: 'ImageDraw' object has no attribute 'textsize'`  
  Use Pillow ≥ 9.2 and prefer `textbbox` (the code already does).
- Fonts look “off”  
  Ensure Arial (or a Hebrew‑compatible font) is available; `load_font()` falls back but results may vary.
- Branding overlaps content  
  Keep `SAFE_BOTTOM` large enough; `add_branding(..., pad_bottom=26)` already shifts content safely above the footer.
- Hebrew appears reversed  
  ensure all strings go through fix_hebrew.
- Missing columns  
  extend candidate lists in _pick_col
- Branding overlaps  
  adjust SAFE_BOTTOM or footer height
- Font issues  
  verify Hebrew-compatible fonts (Arial/DejaVu)

---

## Project structure (suggested)

```
ibi-pdf/
  branding/
    logo.png
    about.png
  DataToPDF/
    Data.xlsx
    DataOld.xlsx
    DataOld2Q.xlsx
    clients_cases.xlsx
  outputs/
  templates/
    template1.png
  main_ibi.py
  requirements.txt
  README.md
```
