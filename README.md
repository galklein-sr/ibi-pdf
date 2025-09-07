# IBI-PDF Report Generator

This Project generates a multi‑page Hebrew PDF report for credit-portfolio monitoring. It reads several Excel workbooks, computes summary metrics, and renders branded slides (PNG/PDF) with tables and donut charts using Pillow.

> **Language/RTL:** The report is Hebrew and right‑to‑left (RTL). Text rendering is done with a small helper (`fix_hebrew`) and careful alignment so content reads correctly.

---

## What the report includes

The pipeline builds up to five slides per **case** (`case_id`):

1. **Executive Overview** — KPIs block over a background template.
2. **Reclassification of Securities** — two aligned blocks (previous quarter vs. current quarter) with counts/percentages.
3. **Material Exposures – Periodic Comparison** — three responsive tables (Borrower / Sectors / Borrower Groups) shown for current **and** previous quarters side-by-side.
4. **Problematic Debts (table)** — a detailed table; under the “סיווג פורום חוב” column two special rows are always present: **“מסופק”** and **“Total”**.
5. **Problematic Debts (donut + selector)** — three donut charts (Geography, Collateral, Liquidity) and a centered segmented selector (“סכום קבוצת לווים”, “תאור נייר”, “תאור קבוצת לווים”).

All slides include **branding** (logo + footer) except the very first slide if configured so in `main()`.

---

## Data inputs

Place Excel files under the project root (defaults are in the code):

- `DataToPDF/Data.xlsx` — **current** quarter
- `DataPrev/Data_prev.xlsx` — **previous** quarter (optional)
- `DataPrevPrev/Data_prev_prev.xlsx` — **previous‑previous** quarter (optional)

Each sheet must include (Hebrew) columns like:
- מנפיק/לווה: `תאור מנפיק` / `תיאור מנפיק` / `לווה`
- ענף: `תאור ענף` / `תיאור ענף` / `ענף`
- קבוצת לווים: `תאור קבוצת לווים` / `שם קבוצת לווים` / `קבוצת לווים`
- סכום/שווי נייר/אחוזים, וכו׳ (numeric columns used by aggregations)

> If column names differ, update the candidate lists passed to `_find_col(...)` so the code can locate the right column at runtime.

### Bucket selection per case
In `main()` you’ll see the mapping used to choose the bucket label per `case_id`:

```python
case_bucket = {
    16396: ["ארם עד 50"],
    16397: ["ארם 50-60"],
    16398: ["ארם 60 ומעלה"],
}
```

---

## How rendering works (high‑level)

- **Layout engine:** All slides are drawn with Pillow (`PIL.Image`, `ImageDraw`). Text size is measured with `draw.textbbox`, and RTL is handled via `fix_hebrew()` plus explicit horizontal alignment.
- **Tables:** `_draw_table` / `_draw_table_full` paint responsive headers and rows, keep consistent padding, and draw thin grid lines between columns/rows. A small fitting helper ensures header text doesn’t overflow: it either shrinks slightly or ellipsizes (depending on the call site).
- **Donut charts:** `_draw_donut` renders a ring chart with a numeric label (e.g., “100%”). Data for the donuts is computed in `_donut_config_for_bucket(bucket_label, df_curr)` and then drawn three times with different titles.
- **Branding:** `add_branding(img, ...)` places `branding/logo.png` in the top‑left and `branding/about.png` (footer) at the bottom. Both are scaled proportionally with `logo_rel_h` / `footer_rel_h` and respect a global **SAFE_BOTTOM** margin so content never collides with the footer.

---

## Installing & running

**Requirements**
- Python **3.10+** (tested on 3.12)
- Windows (PowerShell) or any OS supported by Pillow

**Create a virtual environment & install**
```powershell
# from repository root
python -m venv .venv
. .venv/Scripts/Activate.ps1
pip install -r requirements.txt
# or, if requirements.txt is missing, minimally:
pip install pillow pandas openpyxl pypdf2
```

**Run for one or more case IDs**
```powershell
python main_ibi.py 16396 16397 16398
```
Outputs go to `outputs/`:
- `output_<case_id>.png` — first slide as PNG
- `output_<case_id>.pdf` — multi‑page PDF per case
- `combined_reports.pdf` — merged PDF of all cases

---

## Where to put branding

Create a folder:
```
branding/
  logo.png    # company logo (transparent PNG recommended)
  about.png   # footer strip / about block (PNG)
```



Example:
```python
# First slide (no logo, footer yes)
add_branding(exec_img, show_logo=False, show_footer=True,
             logo_rel_h=0.07, footer_rel_h=0.038, pad_bottom=26, wipe_footer_bg=True)

# Other slides (logo + footer)
for _img in dist_pages + tables_imgs + bad_pages:
    add_branding(_img, show_logo=True, show_footer=True,
                 logo_rel_h=0.07, footer_rel_h=0.038, pad_bottom=26, wipe_footer_bg=True)
```

---

## Layout & spacing

- **Canvas:** `1600 × 900` px (16:9). All coordinates assume this size.
- **Safe bottom (footer):** kept via `SAFE_BOTTOM` and `add_branding(..., pad_bottom=26)`.
- **Tables spacing:** Between tables we use 18–24 px; between the last table and the footer we ensure ≥ 151 px**.
- **Side‑by‑side blocks:** In “Reclassification of Securities”, the two blocks are horizontally aligned using fixed center anchors `left_x = W//2 - 400`, `right_x = W//2 + 400` and identical box widths.

---

## Data mapping tips

- **Missing columns** → `KeyError: Column(s) [None] do not exist`  
  Add or rename source columns, or extend the candidate list in `_find_col(...)`.
- **Text overflow in headers**  
  The code uses a fit helper to shrink or ellipsize. You can tune header font size or the `max_width` passed to the helper.
- **RTL shows reversed (e.g., "ריינ רואת")**  
  Make sure every cell text goes through `fix_hebrew(text)` before drawing.

---

## Troubleshooting

- `AttributeError: 'ImageDraw' object has no attribute 'textsize'`  
  Use Pillow ≥ 9.2 and prefer `textbbox` (the code already does).
- Fonts look “off”  
  Ensure Arial (or a Hebrew‑compatible font) is available; `load_font()` falls back but results may vary.
- Branding overlaps content  
  Keep `SAFE_BOTTOM` large enough; `add_branding(..., pad_bottom=26)` already shifts content safely above the footer.

---

## Project structure (suggested)

```
ibi-pdf/
  branding/
    logo.png
    about.png
  DataToPDF/
    Data.xlsx
  DataPrev/
    Data_prev.xlsx
  DataPrevPrev/
    Data_prev_prev.xlsx
  outputs/
  main_ibi.py
  requirements.txt
  README.md
```
