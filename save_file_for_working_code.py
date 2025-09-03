# -*- coding: utf-8 -*-
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from bidi.algorithm import get_display
import arabic_reshaper
from PyPDF2 import PdfMerger
import sys, os

# =========================
# Config
# =========================
DATA_CURRENT_PATH = 'DataToPDF/Data.xlsx'
DATA_PREV_PATH    = 'DataToPDF/DataOld.xlsx'
DATA_PREV_PREV_PATH = 'DataToPDF/DataOld2Q.xlsx'
SHEET_NAME        = 'Sheet1'
TEMPLATE_IMAGE    = 'templates/template1.png'
OUTPUT_DIR        = 'outputs'

BAD_TYPES = {'השגחה מיוחדת','במעקב מיוחד','מסופק','בפיגור'}
EXCLUDED_SUB_AFIK = {210,211,220,230,240,241,310,311,312,315,326,354,360,405,407,408,409,411,420,425,602,606,
                     242,243,316,325,328,329,330,331,332,333,334,335,336,337,338,339,340,341,342,343,344,345,
                     346,359,387,391,392,395,398,399,412}
DEFAULT_IDS = [16396, 16397, 16398]

# =========================
# Helpers
# =========================
def fix_hebrew(text: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except Exception:
        return str(text)

# ===== new function ======
def load_font(size: int) -> ImageFont.FreeTypeFont:
    for path in [
        "arial.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/ARIAL.TTF",
        "NotoSansHebrew-Regular.ttf",
        "fonts/NotoSansHebrew-Regular.ttf",
    ]:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "--%"
    try:
        return f"{float(v)*100:.2f}%"
    except Exception:
        return "--%"
    
def ellipsize(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_w: int) -> str:
    t = fix_hebrew(str(text))
    if draw.textbbox((0, 0), t, font=font)[2] <= max_w:
        return t
    # חותכים ומוסיפים '…'
    base = t
    while len(t) > 1 and draw.textbbox((0, 0), t + "…", font=font)[2] > max_w:
        t = t[:-1]
    return t + "…"


def _draw_text_fit(draw, text, x, y, w, h,
                   base_size=20, min_size=12,
                   color=(255,255,255), align="center"):
    """
    מצייר טקסט בתוך תיבה (x,y,w,h) כך שייכנס תמיד:
    - מקטין פונט באופן הדרגתי עד min_size
    - ואם עדיין רחב מדי, עושה אליפסיס … בתוך הרוחב המותר.
    align: "left" | "center" | "right"
    """
    raw = fix_hebrew(str(text))
    size = int(base_size)
    f = load_font(size)

    # נסה להקטין עד שנכנס או עד גודל מינימלי
    while size >= min_size:
        f = load_font(size)
        tw, th = _text_size(draw, raw, f)
        if tw <= (w - 8):  # שוליים קטנים משני הצדדים
            break
        size -= 1

    # אם עדיין לא נכנס — חתוך עם אליפסיס
    if _text_size(draw, raw, f)[0] > (w - 8):
        raw = ellipsize(draw, raw, f, w - 8)

    tw, th = _text_size(draw, raw, f)
    if align == "left":
        tx = x + 4
    elif align == "right":
        tx = x + w - tw - 4
    else:  # center
        tx = x + (w - tw) // 2
    ty = y + (h - th) // 2
    draw.text((tx, ty), raw, font=f, fill=color)
    
    
def _fit_text_to_width(draw, text, font, max_width, min_size=12):
    """מקטין גודל פונט עד שנכנס ברוחב, ואם עדיין גדול—חותך עם '…'."""
    t = "" if text is None else str(text)
    size = getattr(font, "size", 22)
    f = font
    # קודם מקטינים פונט
    while size > min_size:
        w = draw.textbbox((0,0), t, font=f)[2]
        if w <= max_width:
            return f, t
        size -= 1
        f = load_font(size)
    # אם עדיין לא נכנס—חיתוך עם אליפסיס
    while t and draw.textbbox((0,0), t + "…", font=f)[2] > max_width:
        t = t[:-1]
    return f, (t + "…") if t else ""


def _text_center_in_rect(draw, text, font, rect, fill=(255,255,255)):
    """מציב טקסט ממורכז בתוך מלבן (left,top,right,bottom)."""
    l,t,r,b = rect
    w = r - l
    h = b - t
    bbox = draw.textbbox((0,0), text, font=font)
    tw = bbox[2]-bbox[0]
    th = bbox[3]-bbox[1]
    x = l + (w - tw)//2
    y = t + (h - th)//2
    draw.text((x,y), text, font=font, fill=fill)

    
    
    # ========================= new page , table two =========================

def _fmt_num(n, digits=2):
    try:
        v = float(n)
    except Exception:
        return "-"
    # אלפי מפרידים עם נקודה עשרונית
    s = f"{v:,.{digits}f}"
    return s

def _draw_table_full(draw, x, y, w, headers, rows, header_font, cell_font, col_fracs):
    """
    headers: רשימת טקסטים (משמאל לימין!)
    rows:    רשימת טורים בטופל באותו סדר של headers
    col_fracs: סכום≈1, אורך כמו headers
    """
    # גדלים
    row_h = 40
    head_h = 42
    pad    = 10
    hfill  = (12,52,87)   # כחול כהה לכותרת
    htext  = (255,255,255)
    grid   = (220,220,220)
    zebra1 = (255,255,255)
    zebra2 = (245,245,245)

    # רוחבי עמודות
    col_ws = [int(w*f) for f in col_fracs]
    # תיקון סכום רוחבים
    col_ws[-1] = w - sum(col_ws[:-1])

    # --- כותרת טבלה ---
    col_x = x
    for i, h in enumerate(headers):
        cw = col_ws[i]
        draw.rectangle([col_x, y, col_x+cw, y+head_h], fill=hfill)
        txt = fix_hebrew(h)
        # טקסט במרכז תא כותרת
        tb = draw.textbbox((0,0), txt, font=header_font)
        draw.text((col_x + (cw - (tb[2]-tb[0]))//2, y + (head_h - (tb[3]-tb[1]))//2),
                  txt, font=header_font, fill=htext)
        col_x += cw

    # קווי הפרדה אנכיים בכותרת
    col_x = x
    for i in range(len(headers)+1):
        draw.line([(col_x, y), (col_x, y+head_h)], fill=grid, width=1)
        if i < len(headers):
            col_x += col_ws[i]
    # קו תחתון לכותרת
    draw.line([(x, y+head_h), (x+w, y+head_h)], fill=grid, width=1)

    # --- שורות נתונים ---
    cy = y + head_h
    for r_i, row in enumerate(rows):
        bg = zebra1 if (r_i % 2 == 0) else zebra2
        draw.rectangle([x, cy, x+w, cy+row_h], fill=bg)

        col_x = x
        for i, cell in enumerate(row):
            
            cw = col_ws[i]
            
            txt = fix_hebrew("" if cell is None else str(cell))
            # לוודא התאמת טקסט לרוחב
            txt = ellipsize(draw, txt, cell_font, max_w=cw - 2*pad)
            # יישור: מספרים לימין, תיאור לשמאל
            is_num_col = i not in (len(row)-1,)  # כל העמודות חוץ מהאחרונה (תיאור) מספריות
            tb = draw.textbbox((0,0), txt, font=cell_font)
            if is_num_col:
                tx = col_x + cw - pad - (tb[2]-tb[0])
            else:
                tx = col_x + pad
            ty = cy + (row_h - (tb[3]-tb[1]))//2
            draw.text((tx, ty), txt, font=cell_font, fill=(20,20,20))
            # קווי הפרדה
            draw.line([(col_x, cy), (col_x, cy+row_h)], fill=grid, width=1)
            col_x += cw
            

            
            
        # קו ימין של השורה
        draw.line([(x+w, cy), (x+w, cy+row_h)], fill=grid, width=1)
        # קו תחתון
        draw.line([(x, cy+row_h), (x+w, cy+row_h)], fill=grid, width=1)

        cy += row_h

    return cy

# ==== Last line RTL ====
def _draw_footer_total(draw: ImageDraw.ImageDraw, y: int, label_he: str, amount_str: str, W: int, font) -> None:
    """
    מצייר את 'סה\"כ אשראי לא מוחרג' מימין ואת המספר משמאל,
    כאשר שני הפריטים ממורכזים יחד כבלוק אחד.
    """
    # הופכים עברית כדי שלא תהפך
    label_fixed = fix_hebrew(label_he)

    # מודדים רוחבים
    lb = draw.textbbox((0, 0), label_fixed, font=font)
    ab = draw.textbbox((0, 0), amount_str, font=font)
    w_label = lb[2] - lb[0]
    w_amount = ab[2] - ab[0]

    gap = 18  # רווח קטן בין המספר לטקסט
    block_w = w_amount + gap + w_label
    x0 = (W - block_w) // 2  # ממרכזים את כל הבלוק

    # המספר משמאל (LTR)
    draw.text((x0, y), amount_str, font=font, fill=(20, 20, 20))
    # הכיתוב מימין (RTL) – מצויר אחרי המספר וה־gap
    draw.text((x0 + w_amount + gap, y), label_fixed, font=font, fill=(20, 20, 20))


def render_bad_debts_page(account_name_display: str,
                          bucket_label: str,
                          df_curr: pd.DataFrame,
                          date_str: str | None = None) -> Image.Image:
    """
    יוצר עמוד 'תיאור חובות בעייתיים...' עבור בקט מסוים (ארם עד 50 / 50-60 / 60+)
    """
    # --- בסיס ---
    W, H = 1600, 900
    img  = Image.new("RGB", (W, H), (245,245,245))
    draw = ImageDraw.Draw(img)

    title_f   = load_font(56)
    sub_f     = load_font(40)
    header_f  = load_font(22)
    cell_f    = load_font(20)
    small_f   = load_font(22)

    # --- כותרות עמוד ---
    t1 = "תיאור חובות בעייתיים בתיק אשראי לא מוחרג"
    draw.text((W//2 - draw.textbbox((0,0), fix_hebrew(t1), font=title_f)[2]//2, 60),
              fix_hebrew(t1), font=title_f, fill=(20,20,20))
    draw.text((W//2 - draw.textbbox((0,0), fix_hebrew(bucket_label), font=sub_f)[2]//2, 120),
              fix_hebrew(bucket_label), font=sub_f, fill=(20,20,20))

    # --- סינון נתונים ---
    df = df_curr.copy()
    # סינון לפי ארם
    df = _filter_aram_bucket(df, bucket_label)
    # השמטת 'מוחרגים'
    if 'sub_afik' in df.columns:
        df = df[~df['sub_afik'].isin(EXCLUDED_SUB_AFIK)]
    # רק סיווגים בעייתיים
    col_forum = _find_col(df, ["סיווג פורום חוב","debt_forum_type"])
    if col_forum:
        df = df[df[col_forum].isin(BAD_TYPES)]

    # מציאת עמודות נוספות (שמות חלופיים נפוצים)
    col_desc   = _find_col(df, ["תיאור נייר","תאור נייר","שם נייר","security_name"])
    col_qty    = _find_col(df, ["כמות","quantity"])
    col_value  = _find_col(df, ["שווי נייר","sec_value","שווי שוק"])
    col_pctNE  = _find_col(df, ["pct_non_excluded","אחוז מאשראי לא מוחרג"])
    col_maalot = _find_col(df, ["דירוג מעלות לנייר","דרוג מעלות לנייר","דירוג מעלות","דרוג מעלות"])
    col_midrug = _find_col(df, ["דירוג מידרג לנייר","דרוג מידרג לנייר","דירוג מידרג","דרוג מידרג"])
    col_machem = _find_col(df, ["מח\"מ מחושב","מחמ מחושב","מחקמ מחושב","מח\"מ"])
    col_yield  = _find_col(df, ["תשואה ברוטו","תשואה","yld"])

    # חישוב pct_non_excluded אם חסר
    if col_pctNE is None and col_value is not None:
        tot_ne = pd.to_numeric(df[col_value], errors="coerce").fillna(0).sum()
        df["_tmp_pct_ne"] = pd.to_numeric(df[col_value], errors="coerce").fillna(0) / (tot_ne if tot_ne else 1)
        col_pctNE = "_tmp_pct_ne"

    # מיון לפי שווי נייר (אם קיים)
    if col_value:
        df["_sort"] = pd.to_numeric(df[col_value], errors="coerce").fillna(0)
        df = df.sort_values("_sort", ascending=False)
    else:
        df = df.copy()

    # בניית שורות (משמאל לימין!)
    rows = []
    for _, r in df.iterrows():
        rows.append((
            # סיווג פורום חוב (עמודה שמאלית)
            r.get(col_forum, ""),
            _fmt_num(r.get(col_yield, ""), 2),
            _fmt_num(r.get(col_machem, ""), 2),
            str(r.get(col_midrug, "")),
            str(r.get(col_maalot, "")),
            f"{(float(r.get(col_pctNE, 0))*100):.2f}%" if col_pctNE else "--%",
            _fmt_num(r.get(col_value, ""), 2),
            _fmt_num(r.get(col_qty, ""), 2),
            r.get(col_desc, "")
        ))

    # הוספת שורת Total
    if len(rows) > 0:
        sum_qty   = pd.to_numeric(df[col_qty], errors="coerce").fillna(0).sum() if col_qty else 0
        sum_value = pd.to_numeric(df[col_value], errors="coerce").fillna(0).sum() if col_value else 0
        sum_pct   = pd.to_numeric(df[col_pctNE], errors="coerce").fillna(0).sum() if col_pctNE else 0
        rows.append((
            "Total",
            _fmt_num(df[col_yield].astype(float).mean() if col_yield else 0, 2) if col_yield else "-",
            _fmt_num(df[col_machem].astype(float).mean() if col_machem else 0, 2) if col_machem else "-",
            "", "",
            f"{sum_pct*100:.2f}%" if col_pctNE else "--%",
            _fmt_num(sum_value, 2),
            _fmt_num(sum_qty, 2),
            ""   # תאור – ריק בשורת הסיכום
        ))
        
        
    if len(rows) == 0:
        rows.append((fix_hebrew("מסופק"), "", "", "", "", "", "", "", ""))
        rows.append(("Total",  "", "", "", "", "", "", "", "")) 

    # ציור הטבלה
    table_w = 1500
    left_x  = (W - table_w)//2
    top_y   = 200

    headers = [
        "סיווג פורום חוב","תשואה ברוטו","מח\"מ מחושב","דרוג מידרג לנייר",
        "דרוג מעלות לנייר","אחוז מאשראי לא מוחרג","שווי נייר","כמות","תאור נייר"
    ]
    # פרופורציות רוחב (משמאל לימין)
    col_fracs = [0.10, 0.10, 0.10, 0.10, 0.11, 0.13, 0.10, 0.09, 0.17]

    end_y = _draw_table_full(draw, left_x, top_y, table_w,
                             headers, rows, header_f, cell_f, col_fracs)

    # טקסט הסבר מתחת לטבלה
    note = "נכון לתאריך הדוח. אין בתיק חשיפה נוספת לנכסי חוב או נכסים אחרים שהונפקו על ידי מנפיקים אלה"
    tb = draw.textbbox((0,0), fix_hebrew(note), font=small_f)
    draw.text((W//2 - (tb[2]-tb[0])//2, end_y + 60), fix_hebrew(note), font=small_f, fill=(30,30,30))

    # סה״כ אשראי לא מוחרג בתחתית
    tot_ne_value = 0
    if col_value is not None:
        tot_ne_value = pd.to_numeric(df[col_value], errors="coerce").fillna(0).sum()
    bottom_text = "סה\"כ אשראי לא מוחרג"
    bt = fix_hebrew(bottom_text)
    tb1 = draw.textbbox((0,0), bt, font=small_f)
    val = _fmt_num(tot_ne_value, 0)
    tb2 = draw.textbbox((0,0), val, font=small_f)
    gap = 35
    cx = W//2 - ( (tb1[2]-tb1[0]) + gap + (tb2[2]-tb2[0]) )//2
    yb = H - 90
    draw.text((cx, yb), bt, font=small_f, fill=(20,20,20))
    draw.text((cx + (tb1[2]-tb1[0]) + gap, yb), val, font=small_f, fill=(20,20,20))

    return img


# ========== פשעק 3 ==========

def render_bad_distributions_page(
    account_name_display: str,
    bucket_label: str,
    df_curr: pd.DataFrame,
    date_str: str | None = None
) -> Image.Image:
    """עמוד 3: שלוש טבלאות התפלגות + הערה מרכזית + שורת בחירה"""
    # בסיס וקנבס
    W, H = 1600, 900
    img  = Image.new("RGB", (W, H), (245,245,245))
    draw = ImageDraw.Draw(img)

    # פונטים (אותו סגנון כמו בעמודים הקודמים)
    title_f  = load_font(56)
    sub_f    = load_font(40)
    header_f = load_font(22)
    cell_f   = load_font(20)
    donut_title_f = load_font(28)
    donut_label_f = load_font(18)

    # כותרות עמוד
    t1 = "תיאור חובות בעייתיים בתיק אשראי לא מוחרג"
    _draw_centered(draw, t1,           W//2, 60,  title_f, (20,20,20))
    _draw_centered(draw, bucket_label, W//2, 120, sub_f,   (20,20,20))

    # סינון נתונים כמו בעמוד 2
    df = df_curr.copy()
    df = _filter_aram_bucket(df, bucket_label)
    if 'sub_afik' in df.columns:
        df = df[~df['sub_afik'].isin(EXCLUDED_SUB_AFIK)]

    # עמודת סכום לפילוח (שווי נייר)
    val_col = _find_col(df, ["שווי נייר","שווי","value"])
    def _top_rows(group_col: str | None) -> list[tuple[str,str,str]]:
        if not group_col or not val_col or df.empty:
            return []
        g = (df.groupby(group_col, dropna=False)[val_col]
                .sum().sort_values(ascending=False).head(3))
        total = float(g.sum()) if float(g.sum()) != 0 else 1.0
        rows = []
        for label, val in g.items():
            pct = float(val) / total
            rows.append((fmt_pct(pct), fmt_km(val), str(label)))
        return rows

    # שורות לכל טבלה
    rows_right = _top_rows(_find_col(df, ["דרג קבוע לנייר","דרוג קבוע לנייר","דירוג קבוע"]))
    rows_mid   = _top_rows(_find_col(df, ["תאור קבוצת לווים","שם קבוצת לווים","קבוצת לווים"]))
    rows_left  = _top_rows(_find_col(df, ["תאור ענף","ענף"]))

    # אם אין נתונים – שורה אחת ריקה כדי לשמור על המבנה
    def _fallback(r): 
        return r if r else [("","", "")]
    rows_right = _fallback(rows_right)
    rows_mid   = _fallback(rows_mid)
    rows_left  = _fallback(rows_left)

    # מיקומים/מידות לטבלאות (3 טב' בשורה)
    y0 = 200
    gap = 60
    w_small = 440
    x_left  = 80
    x_mid   = x_left + w_small + gap
    x_right = x_mid + w_small + gap

    # בכל הטבלאות סדר העמודות הוא משמאל לימין:
    #   [אחוז מאשראי לא מוחרג, שווי נייר, עמודת תיאור]
    headers_left  = ["אחוז מאשראי לא מוחרג","שווי נייר","תאור ענף"]
    headers_mid   = ["אחוז מאשראי לא מוחרג","שווי נייר","תאור קבוצת לווים"]
    headers_right = ["אחוז מאשראי לא מוחרג","שווי נייר","דרג קבוע לנייר"]
    fracs = [0.28, 0.22, 0.50]   # חלוקת רוחבים: אחוז/שווי/תיאור

    # ציור שלוש הטבלאות
    _draw_table_full(draw, x_left,  y0, w_small, headers_left,  rows_left,  header_f, cell_f, fracs)
    _draw_table_full(draw, x_mid,   y0, w_small, headers_mid,   rows_mid,   header_f, cell_f, fracs)
    _draw_table_full(draw, x_right, y0, w_small, headers_right, rows_right, header_f, cell_f, fracs)

    # טקסט אמצעי מתחת לטבלאות
    note = "נכון לתאריך הדוח, אין בתיק חשיפה נוספת לנכסי חוב או נכסים אחרים שהונפקו על ידי קבוצת הלווים הנ\"ל"
    _draw_centered(draw, note, W//2, 360, cell_f, (30,30,30))

    # שורת הטבלה האמצעית (בחירה) – כותרות בלבד ושורה אחת ריקה
    headers_line = ["סכום קבוצת לווים","תאור נייר","תאור קבוצת לווים"]  # משמאל לימין
    rows_line    = [("","","")]
    _draw_table_full(draw, W//2 - 480//2, 390, 480, headers_line, rows_line, header_f, cell_f, [0.34,0.33,0.33])

    # return img
    donuts = _donut_config_for_bucket(bucket_label, df_curr)

    # מרכזים ורדיוסים
    cy = 720
    cx_right  = int(W * 0.83)   # ימני: חשיפה גאוגרפית
    cx_mid    = int(W * 0.50)   # אמצעי: ביטחונות
    cx_left   = int(W * 0.17)   # שמאלי: סחירות
    ro, ri = 115, 65

    # פלטות צבע
    pale_blue = [(14, 134, 255)]
    dark_blue = [(19, 30, 138)]

    _draw_donut(
        draw, cx_right, cy, ro, ri, donuts["geo"],
            "התפלגות חוב בעייתי על פי חשיפה גאוגרפית",
            donut_title_f, donut_label_f, palette=pale_blue
            )

    _draw_donut(
        draw, cx_mid, cy, ro, ri, donuts["collateral"],
            "התפלגות חוב בעייתי על פי ביטחונות",
            donut_title_f, donut_label_f, palette=dark_blue
            )
    
    _draw_donut(
        draw, cx_left, cy, ro, ri, donuts["liquidity"],
            "התפלגות חוב בעייתי על פי סחירות",
            donut_title_f, donut_label_f, palette=pale_blue
            )

    return img

    # ====== My new code =======
# ==== helpers for data-tables page ====

# EXCLUDED_SUB_AFIK = [210, 240, 310, 330, 341, 345, 360, 405, 407, 425, 602]

def _text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    # מחזיר (width, height) באמצעות textbbox
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _detect_borrower_col(df):
    return _find_col(df, ["תאור מנפיק", "תיאור מנפיק", "לווה"])

def _detect_sector_col(df):
    return _find_col(df, ["תאור ענף", "תיאור ענף", "ענף"])

def _detect_group_col(df):
    return _find_col(df, ["תאור קבוצת לווים", "שם קבוצת לווים", "קבוצת לווים"])

def _ensure_num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def _draw_centered(draw, text, cx, cy, font, fill=(20, 20, 20)): #(draw, text, cx, y_mid, font, fill):
    t = fix_hebrew(text)
    bbox = draw.textbbox((0, 0), t, font=font)
    x = cx - (bbox[2] - bbox[0]) // 2
    y = cy - (bbox[3] - bbox[1]) // 2
    draw.text((x, y), t, font=font, fill=fill)

    

# --- Donut (דונאט) ---------------------------------------------------------

def _draw_donut(
    draw, cx, cy, outer_r, inner_r,
    segments,                 # [(label, value), ...]
    title, title_font, label_font,
    palette=None
    #donuts = _donut_config_for_bucket(bucket_label, df_curr)
):
    """
    segments: [(label, value), ...] ; אם יש רק פריט אחד => 100%
    palette:  רשימת צבעים לסגמנטים; אם None נגדיר ברירת מחדל
    """
    if not segments:
        segments = [("לא", 0.0)]

    if palette is None:
        palette = [(14, 134, 255), (120, 170, 255), (180, 205, 255), (80, 140, 230)]

    # כותרת מעל הדונאט
    _draw_centered(draw, title, cx, cy - outer_r - 45, title_font, (20, 20, 20))

    total = sum(float(v) for _, v in segments) or 1.0
    bbox = [cx - outer_r, cy - outer_r, cx + outer_r, cy + outer_r]

    # ציור הסגמנטים
    start_deg = -90.0
    for i, (lbl, val) in enumerate(segments):
        sweep = 360.0 * (float(val) / total)
        draw.pieslice(bbox, start_deg, start_deg + sweep, fill=palette[i % len(palette)], outline=None)
        start_deg += sweep

    # חור פנימי
    draw.ellipse([cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r], fill=(245, 245, 245))

    # תווית מתחת (הפריט הגדול ביותר)
    label, val = max(segments, key=lambda t: float(t[1]))
    pct = (float(val) / total) * 100.0
    leader_y = cy + outer_r
    draw.line([(cx, cy + outer_r - 2), (cx, leader_y + 10)], fill=(170, 170, 170), width=2)

    txt = f"{fix_hebrew(label)} {val:.2f} ({pct:.0f}%)"
    tb = draw.textbbox((0, 0), txt, font=label_font)
    draw.text((cx - (tb[2] - tb[0]) // 2, leader_y + 20), txt, font=label_font, fill=(90, 90, 90))




def _donut_config_for_bucket(bucket_label: str, df_curr: pd.DataFrame) -> dict:
    # סינון לבקט ארם הרלוונטי + הסרת מוחרגים
    df = _filter_aram_bucket(df_curr.copy(), bucket_label)
    if 'sub_afik' in df.columns:
        df = df[~df['sub_afik'].isin(EXCLUDED_SUB_AFIK)]

    val_col = _find_col(df, ["שווי נייר", "שווי", "value"])

    def _build_segments(col_candidates):
        col = _find_col(df, col_candidates)
        if not col or df.empty:
            # ברירת מחדל: מקטע יחיד 100%
            return [("100%", 1.0)]
        if not val_col:
            # בלי עמודת סכום – ספר פריטים
            s = df[col].fillna("לא ידוע").value_counts()
            total = float(s.sum()) or 1.0
            top = s.head(3)
            segs = [(str(k), float(v)) for k, v in top.items()]
            other = total - float(top.sum())
            if other > 0:
                segs.append(("אחר", other))
            return segs
        # עם סכום: TOP-3 לפי סכום
        g = df.groupby(col, dropna=False)[val_col].sum().sort_values(ascending=False)
        total = float(g.sum()) or 1.0
        top = g.head(3)
        segs = [(str(k), float(v)) for k, v in top.items()]
        other = total - float(top.sum())
        if other > 0:
            segs.append(("אחר", other))
        return segs

    return {
        "geo":        _build_segments(["חשיפה גאוגרפית", "מרחב גאוגרפי", "אזור גאוגרפי"]),
        "collateral": _build_segments(["סוג ביטחונות", "ביטחונות", "תיאור ביטחונות"]),
        "liquidity":  _build_segments(["סחירות", "דרגת סחירות", "תיאור סחירות"]),
    }


def _draw_section(draw, x, top_y, w_tbl, title, sub_font, header, rows, header_f, cell_f, aligns):
    """
    מצייר כותרת סקשן ממורכזת + טבלה + ריווח נדיב בין סקשנים.
    """
    # כותרת סקשן
    t = fix_hebrew(title)
    tw, th = _text_size(draw, t, sub_font)
    _draw_centered(draw, title, x + w_tbl // 2, top_y + th // 2, sub_font, (0, 0, 0))

    # טבלה עם רווח ברור אחרי הכותרת
    table_top = top_y + th + 16
    _draw_table(draw, x, table_top, w_tbl, 38, header, rows, header_f, cell_f, aligns=aligns)

    # גובה: שורת כותרת אחת + עד 3 שורות + רווח לפני הסקשן הבא
    used_rows = min(3, len(rows))
    next_y = table_top + 38 * (1 + used_rows) + 40
    return next_y

def _draw_segmented_selector(draw, center_x, y, total_w, h, labels, font):
    """מצייר סרגל בחירה מפוצל לשלושה חלקים (ויזואלי בלבד)."""
    x = center_x - total_w//2
    outline = (60, 100, 160)
    fill    = (255,255,255)
    # רקע מעוגל
    try:
        draw.rounded_rectangle([x, y, x+total_w, y+h], radius=12, outline=outline, width=2, fill=fill)
    except:
        # fallback לריבוע רגיל אם PIL ישן
        draw.rectangle([x, y, x+total_w, y+h], outline=outline, width=2, fill=fill)
    # מחיצות אנכיות
    seg_w = total_w // len(labels)
    for i in range(1, len(labels)):
        draw.line([(x + i*seg_w, y), (x + i*seg_w, y+h)], fill=outline, width=2)
    # טקסטים
    for i, lbl in enumerate(labels):
        lx = x + i*seg_w
        lf, lt = _fit_text_to_width(draw, lbl, font, seg_w - 16, min_size=12)
        _text_center_in_rect(draw, lt, lf, (lx+4, y+4, lx+seg_w-4, y+h-4), fill=(30,30,30))
        
        
        
def _draw_segmented_selector(
    draw: ImageDraw.ImageDraw,
    cx: int, y: int, width: int, height: int,
    labels: list[str],
    font,
    active: int = 0,
    border=(15, 72, 125),      # כחול כמו כותרות הטבלאות
    bg=(255, 255, 255),
    text=(30, 30, 30),
    underline=(15, 72, 125),
    radius: int = 12
):
    """
    מצייר סרגל בחירה (segmented control) בכיוון RTL.
    labels נתונה מימין→לשמאל: [label_right, label_center, label_left]
    active = אינדקס פעיל (0=ימני, 1=אמצעי, 2=שמאלי)
    """
    left = cx - width // 2
    top = y
    right = left + width
    bottom = top + height

    # מסגרת חיצונית
    try:
        draw.rounded_rectangle([left, top, right, bottom], radius=radius, outline=border, width=2, fill=bg)
    except Exception:
        draw.rectangle([left, top, right, bottom], outline=border, width=2, fill=bg)

    n = max(1, len(labels))
    seg_w = width / n

    # מצייר מחיצות ותגיות: RTL — עוברים על labels לפי האינדקסים
    for i, raw_label in enumerate(labels):
        # גבולות המקטע i (RTL)
        seg_right = right - int(i * seg_w)
        seg_left  = right - int((i + 1) * seg_w)

        # מחיצות פנימיות (מלבד הקצה הימני והשמאלי)
        if i != 0:
            draw.line([(seg_right, top), (seg_right, bottom)], fill=border, width=2)

        # טקסט – התאמה לרוחב וחישוב מרכז
        label = fix_hebrew(raw_label)
        max_w = int(seg_w) - 12  # padding אופקי
        fit_font, fit_text = _fit_text_to_width(draw, label, font, max_w, min_size=12)

        # מרכז הטקסט בתוך המקטע
        tw, th = _text_size(draw, fit_text, fit_font)
        tx = seg_left + (seg_right - seg_left - tw) // 2
        ty = top + (height - th) // 2
        draw.text((tx, ty), fit_text, font=fit_font, fill=text)

        # קו תחתון לחלק הפעיל
        if i == active:
            draw.line([(seg_left + 6, bottom - 2), (seg_right - 6, bottom - 2)], fill=underline, width=3)




def _filter_aram_bucket(df, bucket_label: str):
    """
    אם קיימת עמודת 'ארם' – נסנן:
    - 'ארם עד 50' => ערך <= 50
    - 'ארם 50-60' => 50 < ערך <= 60
    - 'ארם 60 ומעלה' => ערך > 60
    אם אין – נחזיר df כמו שהוא (לא שוברים הרצה).
    """
    col = _find_col(df, ["ארם", "ARem", "A.R.M", "ARM"])
    if col is None:
        return df
    s = pd.to_numeric(df[col], errors="coerce")
    if "עד 50" in bucket_label:
        return df[s <= 50]
    if "50-60" in bucket_label or "50–60" in bucket_label:
        return df[(s > 50) & (s <= 60)]
    if "60" in bucket_label:
        return df[s > 60]
    return df


def _compute_top3(df: pd.DataFrame, group_col: str) -> list[dict]:
    """
    מחזיר עד 3 רשומות: label, pct_non_excluded, pct_total.
    pct_total נלקח מעמודה 'אחוז משווי תיק לפי שיערוך אחרון' אם קיימת;
    אחרת נחשב לפי sum(val) / sum(val_total).
    """
    if df is None or df.empty or not group_col or group_col not in df.columns:
        return []

    val_col        = _find_col(df, ["שווי נייר", "sec_value", "שווי"])
    pct_total_col  = _find_col(df, ["אחוז משווי תיק לפי שיערוך אחרון", "pct_of_portfolio_leval"])
    sub_afik_col   = _find_col(df, ["קוד אפיק ותת אפיק", "sub_afik"])

    need = [c for c in [val_col, pct_total_col] if c]
    _ensure_num(df, need)

    # דנומי לא-מוחרג (על בסיס value):
    denom_ne = 0.0
    df_ne = df
    if val_col:
        if sub_afik_col and sub_afik_col in df.columns:
            df_ne = df[~df[sub_afik_col].isin(EXCLUDED_SUB_AFIK)].copy()
        denom_ne = float(df_ne[val_col].sum()) or 0.0

    # אגרגציה רק עם עמודות שקיימות בפועל
    agg_dict = {}
    if val_col:
        agg_dict[val_col] = "sum"
    if pct_total_col:
        agg_dict[pct_total_col] = "sum"
    if not agg_dict:
        return []

    agg = (
        df.groupby(group_col, dropna=True)
          .agg(agg_dict)
          .rename(columns={val_col: "sum_val", pct_total_col: "sum_pct"})
          .reset_index()
    )

    # לחשב pct_total אם אין עמודת אחוזים מקורית
    if not pct_total_col and val_col:
        denom_total = float(df[val_col].sum()) or 0.0
        agg["sum_pct"] = 0.0 if denom_total == 0 else (agg["sum_val"] / denom_total)

    # לחשב pct_non_excluded
    if val_col:
        agg_ne = (
            df_ne.groupby(group_col, dropna=True)[val_col]
                 .sum()
                 .rename("sum_val_ne")
                 .reset_index()
        )
        agg = agg.merge(agg_ne, on=group_col, how="left")
        agg["sum_val_ne"] = agg["sum_val_ne"].fillna(0.0)
        agg["pct_ne"] = 0.0 if denom_ne == 0 else (agg["sum_val_ne"] / denom_ne)
    else:
        agg["pct_ne"] = 0.0

    # דירוג: קודם “לא מוחרג”, אחר כך “מכלל התיק”
    agg = agg.sort_values(["pct_ne", "sum_pct"], ascending=False).head(3)

    out = []
    for _, r in agg.iterrows():
        out.append({
            "label": str(r[group_col]),
            "pct_non_excluded": float(r["pct_ne"]),
            "pct_total": float(r["sum_pct"]),
        })
    return out

def _draw_table(draw, x, y, w, row_h, header, rows, header_f, cell_f, aligns=None):
    if aligns is None:
        aligns = ["right"] * len(header)

    header_bg = (15, 72, 127)
    header_fg = (255, 255, 255)

    # פס כותרות
    draw.rectangle([x, y, x + w, y + row_h], fill=header_bg)

    # עמדות עמודות
    col_x = x
    col_positions = []
    for (title, ratio), align in zip(header, aligns):
        cw = int(w * ratio)
        col_positions.append((col_x, cw, align))

        # *** כאן השינוי: טקסט הכותרת מתאים את עצמו לתיבה ***
        _draw_text_fit(draw, title, col_x, y, cw, row_h,
                       base_size=20, min_size=12,
                       color=header_fg, align="center")

        # קו אנכי לבן דק בין עמודות גם על ההדר
        draw.line([(col_x, y), (col_x, y + row_h)], fill=(255, 255, 255), width=1)
        col_x += cw
    draw.line([(x + w, y), (x + w, y + row_h)], fill=(255, 255, 255), width=1)

    # שורות נתונים (עד 3) + קווי הפרדה דקים
    alt1, alt2 = (255, 255, 255), (238, 237, 237)
    y_row = y + row_h
    max_rows = min(3, len(rows))
    for i in range(max_rows):
        bg = alt2 if i % 2 == 0 else alt1
        draw.rectangle([x, y_row, x + w, y_row + row_h], fill=bg)
        draw.line([(x, y_row), (x + w, y_row)], fill=(255, 255, 255), width=1)

        for (cx, cw, align), val in zip(col_positions, rows[i]):
            txt = ellipsize(draw, val, cell_f, cw - 20)
            tw, th = _text_size(draw, txt, cell_f)
            if align == "left":
                tx = cx + 10
            elif align == "center":
                tx = cx + (cw - tw) // 2
            else:
                tx = cx + cw - tw - 10
            draw.text((tx, y_row + (row_h - th) // 2), txt, font=cell_f, fill=(25, 25, 25))

        vx = x
        for _, cw, _ in col_positions:
            draw.line([(vx, y_row), (vx, y_row + row_h)], fill=(255, 255, 255), width=1)
            vx += cw

        y_row += row_h

    draw.line([(x, y_row), (x + w, y_row)], fill=(255, 255, 255), width=1)


def render_data_tables(
    account_name_display: str,
    bucket_label: str,
    curr_df: pd.DataFrame,
    prev_df: pd.DataFrame | None,
    date_curr: str | None,
    date_prev: str | None
) -> Image.Image:

    # ===== פריסה (מרווחים/גדלים) =====
    TITLE_Y       = 70    # כותרת ראשית
    BUCKET_Y      = 145   # "ארם עד 50"
    DATES_Y       = 225   # שורת התאריכים
    FIRST_TABLE_Y = 285   # תחילת הטבלאות (סקשן ראשון בכל צד)

    # ===== איתור עמודות מקור =====
    borrower_col = _find_col(curr_df, ["תאור מנפיק", "תיאור מנפיק", "לווה"])
    sector_col   = _find_col(curr_df, ["תאור ענף", "תיאור ענף", "ענף"])
    group_col    = _find_col(curr_df, ["תאור קבוצת לווים", "שם קבוצת לווים", "קבוצת לווים"])

    # ===== Top-3 לכל טבלה =====
    cur_borrowers = _compute_top3(curr_df, borrower_col) if borrower_col else []
    cur_sectors   = _compute_top3(curr_df, sector_col)   if sector_col   else []
    cur_groups    = _compute_top3(curr_df, group_col)    if group_col    else []

    prv_borrowers = _compute_top3(prev_df, borrower_col) if (prev_df is not None and borrower_col) else []
    prv_sectors   = _compute_top3(prev_df, sector_col)   if (prev_df is not None and sector_col)   else []
    prv_groups    = _compute_top3(prev_df, group_col)    if (prev_df is not None and group_col)    else []

    # ===== קנבס וגליפים =====
    # --- הכנה לציור ---
    W, H = 1600, 1020  # היה 900; הגדלנו כדי שלא ייחתכו שורות הטבלה התחתונה
    img = Image.new("RGB", (W, H), (245, 245, 245))
    draw = ImageDraw.Draw(img)

    # פונטים מעט קטנים יותר
    title_f  = load_font(60)
    sub_f    = load_font(40)
    date_f   = load_font(36)
    header_f = load_font(16)
    cell_f   = load_font(16)

    # כותרות עליונות + ריווחים גדולים יותר
    _draw_centered(draw, "ניתוח חשיפות מהותיות בתיק אשראי – השוואה תקופתית", W // 2, 70, title_f, (20, 20, 20))
    _draw_centered(draw, bucket_label, W // 2, 160, sub_f, (20, 20, 20))   # רווח גדול יותר לעומת הכותרת הראשית

    # פריסת עמודות
    left_x, right_x, w_tbl = 120, 840, 560

    # תאריכים ממורכזים
    if not date_curr: date_curr = fix_hebrew("רבעון נוכחי")
    if not date_prev: date_prev = fix_hebrew("רבעון קודם")
    _draw_centered(draw, date_curr, left_x  + w_tbl // 2, 240, date_f, (20, 20, 20))  # רווח גדול יותר מה־bucket
    _draw_centered(draw, date_prev, right_x + w_tbl // 2, 240, date_f, (20, 20, 20))
    
    # כותרות עמודות (שמאל→ימין) — עם מקום רחב ל"תאור ..."
    header_borrower = [
        ("אחוז מכלל התיק", 0.18),
        ("אחוז מתיק אשראי לא מוחרג", 0.36),  # היה 0.32
        ("תאור מנפיק", 0.46),
    ]
    header_sector = [
        ("אחוז מכלל התיק", 0.18),
        ("אחוז מתיק אשראי לא מוחרג", 0.36),
        ("תאור ענף", 0.46),
    ]
    header_group = [
        ("אחוז מכלל התיק", 0.18),
        ("אחוז מתיק אשראי לא מוחרג", 0.36),
        ("תאור קבוצת לווים", 0.46),
    ]

    # יישור נתונים: שתי עמודות האחוזים לימין, התיאור לשמאל
    aligns = ["right", "right", "left"]

    # סקשנים — עמודה שמאלית (נוכחי)
    yL = 285
    yL = _draw_section(draw, left_x,  yL, w_tbl, "לווה יחיד",   sub_f, header_borrower, cur_borrowers, header_f, cell_f, aligns)
    yL = _draw_section(draw, left_x,  yL, w_tbl, "ענפים",       sub_f, header_sector,   cur_sectors,   header_f, cell_f, aligns)
    _  = _draw_section(draw, left_x,  yL, w_tbl, "קבוצת לווים", sub_f, header_group,   cur_groups,    header_f, cell_f, aligns)

    # סקשנים — עמודה ימנית (קודם)
    yR = 285
    yR = _draw_section(draw, right_x, yR, w_tbl, "לווה יחיד",   sub_f, header_borrower, prv_borrowers, header_f, cell_f, aligns)
    yR = _draw_section(draw, right_x, yR, w_tbl, "ענפים",       sub_f, header_sector,   prv_sectors,   header_f, cell_f, aligns)
    _  = _draw_section(draw, right_x, yR, w_tbl, "קבוצת לווים", sub_f, header_group,   prv_groups,    header_f, cell_f, aligns)

    return img

def render_bad_debts_page_old(
    account_name_display: str,
    bucket_label: str,
    df_curr: pd.DataFrame,
    date_str: str | None = None,
    total_non_excluded_value: float | None = None,  # אפשר להשאיר None בשלב זה
) -> Image.Image:
    """
    עמוד: 'תיאור חובות בעייתיים בתיק אשראי לא מוחרג'
    טבלה אחת רחבה (9 עמודות) + הערה + שורת סה"כ למטה (placeholder בינתיים).
    """

    # --- זיהוי עמודות מקור ---
    col_desc     = _find_col(df_curr, ["תאור נייר", "תיאור נייר", "שם נייר"])
    col_qty      = _find_col(df_curr, ["כמות"])
    col_value    = _find_col(df_curr, ["שווי נייר"])
    col_pct_ne   = _find_col(df_curr, ["אחוז מאשראי לא מוחרג", "אחוז מתיק אשראי לא מוחרג"])
    col_rate_maa = _find_col(df_curr, ["דרוג מעלות לנייר", "דירוג מעלות לנייר", "דירוג מעלות"])
    col_rate_mid = _find_col(df_curr, ["דרוג מידרג לנייר", "דירוג מידרג לנייר", "דירוג מידרג"])
    col_makam    = _find_col(df_curr, ["מח\"מ מחושב", "מח\"מ", "מחמ מחושב"])
    col_yield    = _find_col(df_curr, ["תשואה ברוטו", "תשואה"])
    # col_forum    = _find_col(df_curr, ["סיווג פורום חוב"])
    col_forum    = _find_col(df_curr, ["סיווג פורום חוב", "debt_forum_type"])

    # פונקציות עזר לפורמט
    def _fmt_int(x):
        try:
            return f"{int(round(float(x))):,}".replace(",", ",")
        except Exception:
            return str(x)

    def _fmt_pct(x):
        try:
            return f"{float(x)*100:.2f}%"
        except Exception:
            return str(x)

    # סינון רשומות "מסופק" (כפי שביקשת)
    df_bad = df_curr.copy()
    if col_forum and col_forum in df_bad.columns:
        df_bad = df_bad[df_bad[col_forum].astype(str).str.strip() == "מסופק"]

    # בניית שורות הטבלה
    rows = []
    for _, r in df_bad.iterrows():
        rows.append([
            r[col_desc] if col_desc else "",
            _fmt_int(r[col_qty]) if col_qty else "",
            fmt_km(r[col_value]) if col_value else "",
            _fmt_pct(r[col_pct_ne]) if col_pct_ne else "",
            r[col_rate_maa] if col_rate_maa else "",
            r[col_rate_mid] if col_rate_mid else "",
            r[col_makam] if col_makam else "",
            r[col_yield] if col_yield else "",
            r[col_forum] if col_forum else "",
        ])

    # שורת Total (סיכומים בסיסיים)
    if len(df_bad) > 0:
        tot_qty   = _fmt_int(df_bad[col_qty].sum()) if col_qty else ""
        tot_value = fmt_km(df_bad[col_value].sum()) if col_value else ""
        # אחוזים – נשמור פשוט כסכום (או להשאיר ריק/“—” אם מעדיפים):
        tot_pct   = _fmt_pct(df_bad[col_pct_ne].sum()) if col_pct_ne else ""
    else:
        tot_qty = tot_value = tot_pct = ""
        
        
        # אם אין כלל נתונים – עדיין נציג שתי שורות: 'מסופק' ו-'Total' (העמודה האחרונה היא 'סיווג פורום חוב')
    if len(rows) == 0:
        rows.append(["", "", "", "", "", "", "", "", fix_hebrew("מסופק")])
        rows.append(["", "", "", "", "", "", "", "", "Total"])


    rows.append([
        "Total",               # תאור נייר
        tot_qty,               # כמות
        tot_value,             # שווי נייר
        tot_pct,               # אחוז מאשראי לא מוחרג
        "",                    # דרוג מעלות לנייר
        "",                    # דרוג מידרג לנייר
        "",                    # מח"מ מחושב
        "",                    # תשואה ברוטו
        "מסופק",              # סיווג פורום חוב
    ])

    # === ציור העמוד ===
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), (245, 245, 245))
    draw = ImageDraw.Draw(img)

    title_f   = load_font(60)
    bucket_f  = load_font(40)
    date_f    = load_font(30)
    header_f  = load_font(20)
    cell_f    = load_font(18)

    # כותרות עליונות
    _draw_centered(draw, "תיאור חובות בעייתיים בתיק אשראי לא מוחרג", W//2, 60, title_f, (20,20,20))
    _draw_centered(draw, bucket_label, W//2, 110, bucket_f, (20,20,20))
    _draw_centered(draw, (date_str or fix_hebrew("נכון לתאריך הדוח")), W//2, 150, date_f, (40,40,40))

    # טבלה רחבה אחת
    tbl_x = 60
    tbl_w = W - 120
    # 9 עמודות – חלוקה יחסית; אפשר לכוונן מעט אם צריך
    col_fracs = [0.20, 0.07, 0.11, 0.11, 0.10, 0.10, 0.09, 0.09, 0.13]
    col_widths = [int(tbl_w*f) for f in col_fracs]

    headers = [
        "תאור נייר",
        "כמות",
        "שווי נייר",
        "אחוז מאשראי לא מוחרג",
        "דרוג מעלות לנייר",
        "דרוג מידרג לנייר",
        "מח\"מ מחושב",
        "תשואה ברוטו",
        "סיווג פורום חוב",
    ]
    aligns  = ["left", "right", "right", "right", "center", "center", "center", "right", "center"]


    cur_y = _draw_table_full(draw, tbl_x, 200, tbl_w, headers, rows, header_f, cell_f, col_fracs)

    # הערה מתחת לטבלה (רווח נדיב)
    cur_y += 28
    note = "נכון לתאריך הדוח. אין בתיק חשיפה נוספת לנכסי חוב או נכסים אחרים שהונפקו על ידי מנפיקים אלה"
    _draw_centered(draw, note, W//2, cur_y, cell_f, (30,30,30))

    # "סה\"כ אשראי לא מוחרג" בתחתית
    bottom_y = H - 50
    total_txt = "—" if total_non_excluded_value is None else fmt_km(total_non_excluded_value)
    footer = f"סה\"כ אשראי לא מוחרג  {total_txt}"
    _draw_centered(draw, footer, W//2, bottom_y, bucket_f, (20,20,20))

    return img

# ==== end of my new code ======

def fmt_km(n: float, decimals: int = 2) -> str:
    try:
        n = float(n)
    except (TypeError, ValueError):
        return str(n)
    sign = "-" if n < 0 else ""
    n = abs(n)
    if n >= 1_000_000:
        return f"{sign}{n/1_000_000:.{decimals}f}M"
    elif n >= 1_000:
        return f"{sign}{n/1_000:.{decimals}f}K"
    else:
        return f"{sign}{n:.{decimals}f}"

def norm_case_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return pd.to_numeric(s, errors='coerce').astype('Int64')

def load_quarter(path: str) -> pd.DataFrame | None:
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    rename_map = {
        'מספר תיק': 'case_id',
        'שם חשבון': 'account_name',
        'אחוז משווי תיק לפי שיערוך אחרון': 'pct_of_portfolio_leval',
        'שווי נייר': 'sec_value',
        'סיווג פורום חוב': 'debt_forum_type',
        'קוד אפיק ותת אפיק': 'sub_afik',
        'מספר נייר': 'sec_id',
    }
    df = df.rename(columns=rename_map)
    if 'pct_of_portfolio_leval' in df.columns:
        df['pct_of_portfolio_leval'] = pd.to_numeric(df['pct_of_portfolio_leval'], errors='coerce').fillna(0.0)
    if 'sec_value' in df.columns:
        df['sec_value'] = pd.to_numeric(df['sec_value'], errors='coerce').fillna(0.0)
    if 'case_id' in df.columns:
        df['__case'] = norm_case_series(df['case_id'])
    else:
        return None
    return df

# =========================
# Metric calculators
# =========================
def metrics_exec_summary(df_case: pd.DataFrame) -> dict:                #ריכוז מנהלים
    portfolio_value_total = float(df_case['sec_value'].sum()) 

    excluded_rows = df_case[df_case['sub_afik'].isin(EXCLUDED_SUB_AFIK)] if 'sub_afik' in df_case.columns else df_case.iloc[0:0]
    excluded_pct_of_portfolio = float(excluded_rows['pct_of_portfolio_leval'].sum())

    bad_rows = df_case[df_case['debt_forum_type'].isin(BAD_TYPES)]
    bad_pct_of_portfolio = float(bad_rows['pct_of_portfolio_leval'].sum())

    den_excl_value = float(excluded_rows['sec_value'].sum())
    num_excl_bad_value = float(excluded_rows[excluded_rows['debt_forum_type'].isin(BAD_TYPES)]['sec_value'].sum())
    bad_share_by_value = 0.0 if den_excl_value == 0 else (num_excl_bad_value / den_excl_value)

    bad_value_total = float(bad_rows['sec_value'].sum())

    return {
        'excluded_pct_of_portfolio': excluded_pct_of_portfolio,     #סכום אחוזי התיק (עמודת pct) עבור המוחרגים בלבד (sub_afik ב-EXCLUDED_SUB_AFIK)
        'portfolio_value_total': portfolio_value_total,             #סכום “שווי נייר” בכל התיק (נוכחי)
        'bad_pct_of_portfolio': bad_pct_of_portfolio,               #סכום אחוזי התיק עבור BAD בלבד
        'bad_value_total': bad_value_total,                         #סכום “שווי נייר” עבור BAD בלבד
        'bad_share_by_value': bad_share_by_value,                   #יחס “שווי BAD מתוך כלל שווי המוחרגים” = (שווי BAD במוחרגים) / (שווי כל המוחרגים)
    }


def metric_excluded_bad_pct(df_case_curr: pd.DataFrame, df_case_prev: pd.DataFrame | None) -> tuple[float|None, float|None]:            #שינוי סיווג ניירות - אחוז מתיק האשראי הלא מוּחרג
    """
    יחס BAD-value מתוך שווי מוחרגים (לכל רבעון).
    """
    def calc(df_case: pd.DataFrame) -> float | None:
        if df_case is None or df_case.empty or 'sub_afik' not in df_case.columns:
            return None
        excluded = df_case[df_case['sub_afik'].isin(EXCLUDED_SUB_AFIK)]         #שורות שה-sub_afik שלהן מוחרג
        denom = float(excluded['sec_value'].sum())                              #סכום שווי כל המוחרגים
        if denom <= 0:
            return None
        numer = float(excluded[excluded['debt_forum_type'].isin(BAD_TYPES)]['sec_value'].sum())             #סכום שווי BAD בתוך המוחרגים
        return numer / denom
    return calc(df_case_curr), calc(df_case_prev)

def metric_total_bad_pct(df_case_curr: pd.DataFrame, df_case_prev: pd.DataFrame | None) -> tuple[float|None, float|None]:           #שינוי סיווג ניירות - אחוז מכלל התיק
    """
    אחוז BAD מכלל התיק (סכום עמודת pct).
    """
    def calc(df_case: pd.DataFrame) -> float | None:
        if df_case is None or df_case.empty or 'pct_of_portfolio_leval' not in df_case.columns:
            return None
        return float(df_case[df_case['debt_forum_type'].isin(BAD_TYPES)]['pct_of_portfolio_leval'].sum())
    return calc(df_case_curr), calc(df_case_prev)

def metric_total_bad_count(df_case: pd.DataFrame | None) -> int | None:         # שינוי סיווג ניירות - סה"כ ניירות בעייתיים
    if df_case is None or df_case.empty:
        return None
    bad_df = df_case[df_case['debt_forum_type'].isin(BAD_TYPES)]
    if 'sec_id' in bad_df.columns:
        s = bad_df['sec_id'].dropna().astype(str).str.strip()
        return int(s.nunique())
    # fallback: אם אין עמודה 'sec_id'
    return int(bad_df.shape[0])

def metric_bad_entry_exit(df_case_curr: pd.DataFrame | None,            #שינוי סיווג ניירות - כניסה/יציאה לסיווג BAD
                          df_case_prev: pd.DataFrame | None) -> tuple[int | None, int | None]:
    """
    כניסה/יציאה לסיווג BAD לפי Distinct נייר (sec_id):
      entries = ניירות שהיו not-BAD קודם וכעת BAD
      exits   = ניירות שהיו BAD קודם וכעת not-BAD
    """
    if df_case_curr is None or df_case_curr.empty or df_case_prev is None or df_case_prev.empty:
        return None, None

    def bad_set(df: pd.DataFrame) -> set:
        # קובע לכל נייר האם BAD בהינתן שורות מרובות
        if 'sec_id' in df.columns:
            tmp = df[['sec_id', 'debt_forum_type']].copy()
            tmp['is_bad'] = tmp['debt_forum_type'].isin(BAD_TYPES)
            g = tmp.groupby('sec_id')['is_bad'].max()
            return set(g.index[g])
        # fallback: מפתח אחר אם קיים
        for key_col in ('ISIN', 'שם נייר', 'SecurityID', 'מספר נייר'):
            if key_col in df.columns:
                tmp = df[[key_col, 'debt_forum_type']].copy()
                tmp['is_bad'] = tmp['debt_forum_type'].isin(BAD_TYPES)
                g = tmp.groupby(key_col)['is_bad'].max()
                return set(g.index[g])
        # fallback קיצוני: לפי שורות
        return set(df.index[df['debt_forum_type'].isin(BAD_TYPES)])

    prev_bad = bad_set(df_case_prev)
    curr_bad = bad_set(df_case_curr)

    entries = len(curr_bad - prev_bad)  
    exits   = len(prev_bad - curr_bad) 
    return entries, exits

def metric_class_change_count(df_curr, df_prev):            #שינוי סיווג ניירות - מספר שינוי סיווג
    """
    כמה ניירות (Distinct sec_id) החליפו סיווג בין prev -> curr
    מטפל בריבוי שורות לאותו sec_id ע"י בחירת label אחד
    """
    if df_curr is None or df_curr.empty or df_prev is None or df_prev.empty:
        return None

    key = 'sec_id' if ('sec_id' in df_curr.columns and 'sec_id' in df_prev.columns) else None
    if key is None:
        return None

    def one_label_per_sec(df):
        s = (
            df[[key, 'debt_forum_type']]
            .dropna(subset=[key])
            .assign(
                **{
                    key: lambda x: x[key].astype(str).str.strip(),
                    'debt_forum_type': lambda x: x['debt_forum_type'].astype(str).str.strip()
                }
            )
            .groupby(key)['debt_forum_type']
            .agg(lambda col: col.mode().iloc[0] if not col.mode().empty else col.iloc[0])
        )
        #סדר קבוע
        return s.sort_index()

    curr_labels = one_label_per_sec(df_curr)
    prev_labels = one_label_per_sec(df_prev)

    # רק ניירות שמופיעים בשני הרבעונים
    common_ids = curr_labels.index.intersection(prev_labels.index)
    if common_ids.empty:
        return 0

    # יישור אינדקסים והשוואה
    aligned_curr = curr_labels.reindex(common_ids)
    aligned_prev = prev_labels.reindex(common_ids)

    changes = (aligned_curr != aligned_prev).sum()
    return int(changes)


def metric_late_or_delivered_count(df_case):            #שינוי סיווג ניירות - סה"כ ניירות בפיגור/מסופק
    """Distinct 'sec_id' עבור ניירות שהסיווג שלהם 'בפיגור' או 'מסופק' ברבעון הנתון."""
    if df_case is None or df_case.empty:
        return None
    mask = df_case['debt_forum_type'].isin(['בפיגור', 'מסופק'])
    filtered = df_case[mask]
    if filtered.empty:
        return 0
    if 'sec_id' in filtered.columns:
        s = filtered['sec_id'].dropna().astype(str).str.strip()
        return int(s.nunique())
    return int(filtered.shape[0])  # fallback אם אין sec_id

# =========================
# Renderers
# =========================
def render_exec_slide(base_image_path: str, account_name_display: str, metrics: dict) -> Image.Image:
    image = Image.open(base_image_path).convert("RGB")
    draw = ImageDraw.Draw(image)
    def draw_centered_raw(text, cx, cy, font, fill):
        bbox = draw.textbbox((0, 0), text, font=font)
        x = cx - (bbox[2]-bbox[0]) // 2
        y = cy - (bbox[3]-bbox[1]) // 2
        draw.text((x, y), text, font=font, fill=fill)
    # f32 = ImageFont.truetype("arial.ttf", 32)
    # f42 = ImageFont.truetype("arial.ttf", 42)
    
    f32 = load_font(32)
    f42 = load_font(42)

    
    items = [
        (fix_hebrew(account_name_display), 1030, 50,  f32, (255,255,255), False),
        (f"{metrics['excluded_pct_of_portfolio']*100:.2f}%", 246, 132, f32, (255,255,255), True),
        (fmt_km(metrics['portfolio_value_total']),               470, 132, f32, (255,255,255), True),
        (f"{metrics['bad_pct_of_portfolio']*100:.2f}%",         246, 440, f42, (255,255,255), True),
        (fmt_km(metrics['bad_value_total']),                     470, 440, f42, (255,255,255), True),
        (f"{metrics['bad_share_by_value']*100:.2f}%",          1020, 290, f42, (255,255,255), True),
    ]
    for text, cx, cy, font, color, is_num in items:
        if is_num:
            draw_centered_raw(text, cx, cy, font, color)
        else:
            bbox = draw.textbbox((0, 0), text, font=font)
            x = cx - (bbox[2]-bbox[0]) // 2
            y = cy - (bbox[3]-bbox[1]) // 2
            draw.text((x, y), text, font=font, fill=color)
    return image

def render_white_slide(account_name_display: str,
                       curr_excluded_bad_pct: float | None, prev_excluded_bad_pct: float | None,                                        #אחוז מתיק האשראי הלא מוּחרג
                       curr_total_bad_pct: float | None,    prev_total_bad_pct: float | None,                                           #אחוז מכלל התיק
                       curr_bad_count: int | None,           prev_bad_count: int | None,                                                #סה"כ ניירות בעייתיים
                       curr_bad_entries: int | None,         prev_bad_entries: int | None,                                              #כניסה לסיווג כחוב בעייתי
                       curr_bad_exits: int | None,    prev_bad_exits: int | None,                                                       #יציאה מסיווג כחוב בעייתי
                       curr_class_changes: int | None, prev_class_changes: int | None,                                                  #מספר שינוי סיווג
                       curr_late_or_delivered: int | None = None, prev_late_or_delivered: int | None = None) -> Image.Image:            #סה"כ ניירות בפיגור/מסופק
    W, H = 1600, 900
    slide = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(slide)

    # title_f = ImageFont.truetype("arial.ttf", 60)
    # sub_f   = ImageFont.truetype("arial.ttf", 40)
    # box_t_f = ImageFont.truetype("arial.ttf", 28)
    # pct_f   = ImageFont.truetype("arial.ttf", 56)
    # small_f = ImageFont.truetype("arial.ttf", 26)
    
    title_f = load_font(60)
    sub_f   = load_font(40)
    box_t_f = load_font(28)
    pct_f   = load_font(56)
    small_f = load_font(26)

    

    def draw_centered_text(text, cx, cy, fnt, fill=(0,0,0)):
        t = fix_hebrew(text)
        bbox = draw.textbbox((0,0), t, font=fnt)
        draw.text((cx-(bbox[2]-bbox[0])//2, cy-(bbox[3]-bbox[1])//2), t, font=fnt, fill=fill, align="center")

    def draw_centered_raw(text, cx, cy, fnt, fill=(0,0,0)):
        bbox = draw.textbbox((0,0), text, font=fnt)
        draw.text((cx-(bbox[2]-bbox[0])//2, cy-(bbox[3]-bbox[1])//2), text, font=fnt, fill=fill, align="center")

    draw_centered_text("שינוי סיווג ניירות", W//2, 100, title_f, (30,30,30))
    draw_centered_text(account_name_display, W//2, 170, sub_f, (60,60,60))

    left_x, right_x = W//2 - 400, W//2 + 400
    draw_centered_text("רבעון קודם",  left_x,  240, box_t_f, (50,50,50))
    draw_centered_text("רבעון נוכחי", right_x, 240, box_t_f, (50,50,50))

    def draw_box(x1,y1,x2,y2):
        draw.rectangle((x1,y1,x2,y2), outline=(180,180,180), width=2)

    def draw_percent(cx, top_y, title, value):
        bw, bh = 320, 150
        l = cx - bw//2
        draw_box(l, top_y, l+bw, top_y+bh)
        draw_centered_text(title, cx, top_y+35, box_t_f, (70,70,70))
        if value is None:
            txt = "--%"
            fill = (200,200,200)
        else:
            txt = f"{value*100:.2f}%"
            fill = (0,0,0)
        draw_centered_raw(txt, cx, top_y+95, pct_f, fill)

    # שורה עליונה: אחוז מכלל התיק (BAD על בסיס אחוזים מהמקור)
    draw_percent(left_x,  280, "אחוז מכלל התיק", prev_total_bad_pct)
    draw_percent(right_x, 280, "אחוז מכלל התיק", curr_total_bad_pct)

    # שורה תחתונה: “אחוז מתיק האשראי הלא מוּחרג” (עפ״י ההנחיה: מחושב על המוחרגים)
    draw_percent(left_x,  460, "אחוז מתיק האשראי\nהלא מוּחרג", prev_excluded_bad_pct)
    draw_percent(right_x, 460, "אחוז מתיק האשראי\nהלא מוּחרג", curr_excluded_bad_pct)

    # בלוק סטטיסטיקות תחתון
    def draw_stats(cx, top_y, total_bad_count: int | None, bad_entries_count: int | None = None, bad_exits_count: int | None = None, class_changes_count: int | None = None, late_or_delivered_count: int | None = None):
        bw, bh = 560, 230
        l = cx - bw//2
        draw_box(l, top_y, l+bw, top_y+bh)
        pad = 18; y = top_y + pad

        line1_val = "--" if total_bad_count is None else str(total_bad_count)
        line2_val = "--" if bad_entries_count is None else str(bad_entries_count)
        line3_val = "--" if bad_exits_count is None else str(bad_exits_count)
        line4_val = "--" if class_changes_count is None else str(class_changes_count)
        line5_val = "--" if late_or_delivered_count is None else str(late_or_delivered_count)

        lines = [
            f"סה\"כ ניירות בעייתיים       {line1_val}",
            f"כניסה לסיווג כחוב בעייתי        {line2_val}",
            f"יציאה מסיווג כחוב בעייתי        {line3_val}",
            f"סה\"כ ניירות בפיגור/מסופק    {line5_val}",
            f"מספר שינוי סיווג            {line4_val}",
        ]
        for line in lines:
            draw.text((l+pad, y), fix_hebrew(line), font=small_f, fill=(30,30,30))
            y += 40
            
    draw_stats(left_x,  640, prev_bad_count, prev_bad_entries, prev_bad_exits, prev_class_changes, prev_late_or_delivered)
    draw_stats(right_x, 640, curr_bad_count, curr_bad_entries, curr_bad_exits, curr_class_changes, curr_late_or_delivered)

    return slide


def render_data_tables_demo(account_name_display: str):
    def draw_centered_text(text, cx, cy, fnt, fill=(0,0,0)):
        t = fix_hebrew(text)
        bbox = draw.textbbox((0,0), t, font=fnt)
        draw.text((cx-(bbox[2]-bbox[0])//2, cy-(bbox[3]-bbox[1])//2), t, font=fnt, fill=fill, align="center")
    W, H = 1600, 900
    slide = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(slide)
    
    #Titles
    title_f = ImageFont.truetype("arial.ttf", 60)
    sub_title_f = ImageFont.truetype("arial.ttf", 55)
    table_header_f = ImageFont.truetype("arial.ttf", 28)
    title_text="ניתוח חשיפות מהותיות בתיק אשראי-השוואה תקופתית"
    t = fix_hebrew(title_text)
    draw.text((110,50), t, font=title_f, fill=(0,0,0), align="center")
    draw_centered_text(account_name_display,800,150,sub_title_f)
    
    #Dates
    date_1="30/09/2024"
    date_2="31/12/2024"
    table_1_title="לווה יחיד"
    draw_centered_text(date_1,1200,200,sub_title_f)
    draw_centered_text(date_2,400,200,sub_title_f)
    draw_centered_text(table_1_title,1200,260,table_header_f)
    draw_centered_text(table_1_title,400,260,table_header_f)
    table_header_f = ImageFont.truetype("arial.ttf", 18)
    table_header_r="תאור מנפיק"
    table_header_m="אחוז מתיק אשראי לא מוחרג"
    table_header_l="אחוז מכלל התיק"
    right_pos_x=1200
    left_pos_x=450
    pos_y=300
    white_color = (255,255,255)
    grey_color=(238,237,237)
    blue_color=(15,72,127)
    #right
    draw.rectangle((right_pos_x,pos_y,right_pos_x+300,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_r,right_pos_x+150,pos_y+15,table_header_f, fill=white_color)
    draw.rectangle((right_pos_x-200,pos_y,right_pos_x,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_m,right_pos_x-100,pos_y+15,table_header_f, fill=white_color)
    draw.rectangle((right_pos_x-350,pos_y,right_pos_x-200,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_l,right_pos_x-275,pos_y+15,table_header_f, fill=white_color)
    for i in range(3):
        pos_y=pos_y+40
        desc_text="טקסט של התאור"
        prec_dis_text="{:.2f}".format(0.18*100) + '%'
        prec_tot_text="{:.2f}".format(0.25*100) + '%'
        color = white_color
        if i%2==0:
            color = grey_color
        draw.rectangle((right_pos_x,pos_y,right_pos_x+300,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(desc_text,right_pos_x+150,pos_y+15,table_header_f)
        draw.rectangle((right_pos_x-200,pos_y,right_pos_x,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(prec_dis_text,right_pos_x-100,pos_y+15,table_header_f)
        draw.rectangle((right_pos_x-350,pos_y,right_pos_x-200,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(prec_tot_text,right_pos_x-275,pos_y+15,table_header_f)

    pos_y=300
    #left
    draw.rectangle((left_pos_x,pos_y,left_pos_x+300,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_r,left_pos_x+150,pos_y+15,table_header_f, fill=white_color)
    draw.rectangle((left_pos_x-200,pos_y,left_pos_x,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_m,left_pos_x-100,pos_y+15,table_header_f, fill=white_color)
    draw.rectangle((left_pos_x-350,pos_y,left_pos_x-200,pos_y+40), fill=blue_color, outline=white_color, width=2)
    draw_centered_text(table_header_l,left_pos_x-275,pos_y+15,table_header_f, fill=white_color)
    for i in range(3):
        pos_y=pos_y+40
        desc_text="טקסט של התאור"
        prec_dis_text="{:.2f}".format(0.18*100) + '%'
        prec_tot_text="{:.2f}".format(0.25*100) + '%'
        color = white_color
        if i%2==0:
            color = grey_color
        draw.rectangle((left_pos_x,pos_y,left_pos_x+300,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(desc_text,left_pos_x+150,pos_y+15,table_header_f)
        draw.rectangle((left_pos_x-200,pos_y,left_pos_x,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(prec_dis_text,left_pos_x-100,pos_y+15,table_header_f)
        draw.rectangle((left_pos_x-350,pos_y,left_pos_x-200,pos_y+40), fill=color, outline=color, width=0)
        draw_centered_text(prec_tot_text,left_pos_x-275,pos_y+15,table_header_f)
    return slide

# =========================
# Main
# =========================
def main():
    ids = [int(a) for a in sys.argv[1:]] if len(sys.argv) > 1 else DEFAULT_IDS
    df_curr = load_quarter(DATA_CURRENT_PATH)
    df_prev = load_quarter(DATA_PREV_PATH)
    df_prev_prev = load_quarter(DATA_PREV_PREV_PATH)


    if df_curr is None:
        print("לא נמצא או לא תקין קובץ Data.xlsx")
        return

    df_curr['__case'] = norm_case_series(df_curr['case_id'])
    if df_prev is not None:
        df_prev['__case'] = norm_case_series(df_prev['case_id'])

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    pdf_parts = []

    for case_id in ids:
        case_curr = df_curr[df_curr['__case'] == case_id].copy() #dataframe המכיל את כל השורות עבור התיק הספציפי ברבעון הנוכחי
        if case_curr.empty:
            print(f"CASE {case_id}: אין נתונים בקובץ הנוכחי – מדלג.")
            continue

        account_name = case_curr.iloc[0]['account_name'] if 'account_name' in case_curr.columns else f"תיק {case_id}"

        exec_metrics = metrics_exec_summary(case_curr)

        case_prev = df_prev[df_prev['__case'] == case_id].copy() if df_prev is not None else None   #dataframe המכיל את כל השורות עבור התיק הספציפי ברבעון הקודם
        case_prev_prev = df_prev_prev[df_prev_prev['__case'] == case_id].copy() if df_prev_prev is not None else None   #dataframe המכיל את כל השורות עבור התיק הספציפי ברבעון שלפני הקודם

        # שינוי סיווג – שני מדדים
        case_prev = df_prev[df_prev['__case'] == case_id].copy() if df_prev is not None else None
        case_prev_prev = df_prev_prev[df_prev_prev['__case'] == case_id].copy() if df_prev_prev is not None else None
        curr_excl_bad_pct, prev_excl_bad_pct = metric_excluded_bad_pct(case_curr, case_prev)
        curr_total_bad_pct, prev_total_bad_pct = metric_total_bad_pct(case_curr, case_prev)
        curr_bad_count = metric_total_bad_count(case_curr)
        prev_bad_count = metric_total_bad_count(case_prev)
        curr_bad_entries, curr_bad_exits = metric_bad_entry_exit(case_curr, case_prev)
        prev_bad_entries, prev_bad_exits = metric_bad_entry_exit(case_prev, case_prev_prev)
        curr_class_changes = metric_class_change_count(case_curr, case_prev)
        prev_class_changes = metric_class_change_count(case_prev, case_prev_prev)
        curr_late_or_delivered = metric_late_or_delivered_count(case_curr)
        prev_late_or_delivered = metric_late_or_delivered_count(case_prev)

        # שקפים
        exec_img  = render_exec_slide(TEMPLATE_IMAGE, account_name, exec_metrics)
        white_img = render_white_slide(account_name,
                                       curr_excl_bad_pct, prev_excl_bad_pct,
                                       curr_total_bad_pct, prev_total_bad_pct,
                                       curr_bad_count, prev_bad_count,
                                       curr_bad_entries, prev_bad_entries,               
                                       curr_bad_exits, prev_bad_exits,
                                       curr_class_changes, prev_class_changes,
                                       curr_late_or_delivered, prev_late_or_delivered)
        
        case_bucket = {
            16396: ["ארם עד 50"],
            16397: ["ארם 50-60"],
            16398: ["ארם 60 ומעלה"],  
        }
        buckets_for_this_case = case_bucket.get(case_id, ["ארם עד 50"])
        
        tables_imgs = []
        bad_pages   = []
        dist_pages = []

        # נגדיר פעם אחת את התאריכים שנשתמש בהם
        date_cur_str  = "30/06/2025"
        date_prev_str = "31/03/2025"

        for bucket_label in buckets_for_this_case:
            # עמוד הטבלאות (העמוד הראשון)
            tables_imgs.append(
                render_data_tables(
                    account_name_display=account_name,
                    bucket_label=bucket_label,
                    curr_df=case_curr,
                    prev_df=case_prev,
                    date_curr=date_cur_str,
                    date_prev=date_prev_str,
                )
            )

    # עמוד "תיאור חובות בעייתיים..."
            bad_pages.append(
                render_bad_debts_page(
                    account_name_display=account_name,  # שים לב לשם הפרמטר
                    bucket_label=bucket_label,          # לא 'bl' — משתמשים במשתנה של הלולאה
                    df_curr=case_curr,
                    date_str=date_cur_str,
                )
            )
            
            dist_pages.append(
                render_bad_distributions_page(
                    account_name_display=account_name,
                    bucket_label=bucket_label,
                    df_curr=case_curr,
                    date_str="30/06/2025",   # או משתנה תאריך אם יש
                    # סרגל בחירה במקום טבלה
                )
            )
        # === end of my edit ===
        
        # שמירה
        out_png = os.path.join(OUTPUT_DIR, f'output_{case_id}.png')
        out_pdf = os.path.join(OUTPUT_DIR, f'output_{case_id}.pdf')
        exec_img.save(out_png)

        exec_img.save(out_pdf, save_all=True, append_images=[white_img, *tables_imgs, *bad_pages, *dist_pages])

        pdf_parts.append(out_pdf)
        print(f"CASE {case_id}: created {out_png}, {out_pdf}")

    if pdf_parts:
        merger = PdfMerger()
        for p in pdf_parts:
            merger.append(p)
        merged_path = os.path.join(OUTPUT_DIR, 'combined_reports.pdf')
        merger.write(merged_path)
        merger.close()
        print(f"created merged PDF: {merged_path}")
    else:
        print("did not create any PDF files.")

if __name__ == "__main__":
    main()