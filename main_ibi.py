from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from bidi.algorithm import get_display
import arabic_reshaper
from PyPDF2 import PdfMerger
import sys, os
from pathlib import Path

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


# לוגו + פס יצירת קשר בתחתית
LOGO_PATH   = os.path.join("branding", "logo.png")
FOOTER_PATH = os.path.join("branding", "about.png")

def _safe_open_rgba(path):
    try:
        if path and os.path.exists(path):
            return Image.open(path).convert("RGBA")
    except Exception:
        pass
    return None

def add_branding(
    img: Image.Image,
    show_logo=False,        
    show_footer=True,
    logo_rel_h=0.08,      # גובה לוגו יחסי (קטן יותר כדי לא להפריע)
    footer_rel_h=0.045,   # גובה פוטר יחסי
    pad_left=22, pad_top=18, pad_bottom=22,
    wipe_footer_bg=False  # אם TRUE צובע רצועה בהירה לפני הדבקת הפוטר
):
    W, H = img.size
    draw = ImageDraw.Draw(img)
    
    if show_logo:                                   
        logo = _safe_open_rgba(LOGO_PATH)
        if logo:
            target_h = int(H * logo_rel_h)
            ratio    = target_h / logo.height
            lw       = int(logo.width * ratio)
            max_lw   = int(W * 0.23)
            if lw > max_lw:
                ratio    = max_lw / logo.width
                target_h = int(logo.height * ratio)
                lw       = max_lw
            logo = logo.resize((lw, target_h), Image.LANCZOS)
            img.paste(logo, (pad_left, pad_top), logo)

    # --- פוטר ---
    if show_footer:                                   
        footer = _safe_open_rgba(FOOTER_PATH)
        if footer:
            target_h = int(H * footer_rel_h)
            ratio    = target_h / footer.height
            fw       = int(footer.width * ratio)
            max_fw   = int(W * 0.72)
            if fw > max_fw:
                ratio    = max_fw / footer.width
                target_h = int(footer.height * ratio)
                fw       = max_fw
            footer = footer.resize((fw, target_h), Image.LANCZOS)
            x = (W - footer.width) // 2
            y = H - footer.height - pad_bottom
            if wipe_footer_bg:
                band_pad = 8
                draw.rectangle([0, max(0, y - band_pad), W, H], fill=(245, 245, 245))
            img.paste(footer, (x, y), footer)

# =========================
# Helpers
# =========================
def fix_hebrew(text: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except Exception:
        return str(text)

SAFE_TOP = 80
SAFE_BOTTOM = 110

CM = 37.7952755906
SAFE_BOTTOM = int(4 * CM)

def cm_to_px(cm, ppi=96):
    # המרה פשוטה: 1 אינץ' = 2.54 ס"מ; נניח 96dpi לקנבס 1600x900
    return int(round(cm * ppi / 2.54))

RESERVED_BOTTOM_CM = 4
SAFE_BOTTOM = max(SAFE_BOTTOM, cm_to_px(RESERVED_BOTTOM_CM))

# ===== robust font loader =====
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
    

def _fit_text_to_width(draw, text, font, max_width, min_size=12):
    t = "" if text is None else str(text)
    size = getattr(font, "size", 22)
    f = font
    while size > min_size:
        w = draw.textbbox((0, 0), t, font=f)[2]
        if w <= max_width:
            return f, t
        size -= 1
        f = load_font(size)
    while t and draw.textbbox((0, 0), t + "…", font=f)[2] > max_width:
        t = t[:-1]
    return f, (t + "…") if t else ""


def _fit_or_ellipsis(draw, text, font, max_width, allow_shrink_pts=2):
    """
    שומר על גודל הפונט כמעט קבוע; קודם כל מנסה לקצר עם …,
    ורק אם גם תו אחד + … לא נכנס – מקטין עד 2pt לכל היותר.
    """
    t = "" if text is None else str(text)
    f = font

    # 1) קודם: חתוך עם … בגודל המקורי
    while t and draw.textbbox((0, 0), t + "…", font=f)[2] > max_width:
        t = t[:-1]
    if t:
        return f, (t + "…") if draw.textbbox((0, 0), t + "…", font=f)[2] <= max_width else t

    # 2) אם אפילו תו אחד לא נכנס, הקטן מעט (עד 2pt)
    base = getattr(font, "size", 22)
    for size in range(base - 1, max(base - allow_shrink_pts, 8) - 1, -1):
        f2 = load_font(size)
        t2 = "" if text is None else str(text)
        while t2 and draw.textbbox((0, 0), t2 + "…", font=f2)[2] > max_width:
            t2 = t2[:-1]
        if t2:
            return f2, (t2 + "…")
    # fallback: מחזיר את המקורי כמו שהוא
    return font, (text if isinstance(text, str) else str(text))

def _text_center_in_rect(draw, text, font, rect, fill=(255,255,255)):
    l, t, r, b = rect
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = l + ((r - l) - tw) // 2
    y = t + ((b - t) - th) // 2
    draw.text((x, y), text, font=font, fill=fill)

    

def ellipsize(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_w: int) -> str:
    t = fix_hebrew(str(text))
    if draw.textbbox((0, 0), t, font=font)[2] <= max_w:
        return t
    while len(t) > 1 and draw.textbbox((0, 0), t + "…", font=font)[2] > max_w:
        t = t[:-1]
    return t + "…"

####### DOWN : START OF RECENT EDITS OF HEBREW WORDS #######
def ellipsize_raw(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_w: int) -> str:
    """קיצור טקסט לרוחב נתון *בלי* fix_hebrew (משתמשים כשכבר עשינו עיבוד RTL קודם)."""
    t = str(text)
    if draw.textbbox((0, 0), t, font=font)[2] <= max_w:
        return t
    while len(t) > 1 and draw.textbbox((0, 0), t + "…", font=font)[2] > max_w:
        t = t[:-1]
    return t + "…"

def _is_numeric_text(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    if s == "":
        return False
    s_clean = s.replace(",", "")
    if s_clean.endswith("%"):
        s_clean = s_clean[:-1]
    try:
        float(s_clean)
        return True
    except Exception:
        return False
####### ABOVE : END OF RECENT EDITS OF HEBREW WORDS #######

def _draw_text_fit(draw, text, x, y, w, h,
                   base_size=20, min_size=12,
                   color=(255,255,255), align="center"):
    raw = fix_hebrew(str(text))
    size = int(base_size)
    f = load_font(size)
    while size >= min_size:
        f = load_font(size)
        tw, th = _text_size(draw, raw, f)
        if tw <= (w - 8):
            break
        size -= 1
    if _text_size(draw, raw, f)[0] > (w - 8):
        raw = ellipsize(draw, raw, f, w - 8)
    tw, th = _text_size(draw, raw, f)
    if align == "left":
        tx = x + 4
    elif align == "right":
        tx = x + w - tw - 4
    else:
        tx = x + (w - tw) // 2
    ty = y + (h - th) // 2
    draw.text((tx, ty), raw, font=f, fill=color)

def _fit_text_to_width(draw, text, font, max_width, min_size=12):
    t = "" if text is None else str(text)
    size = getattr(font, "size", 22)
    f = font
    while size > min_size:
        w = draw.textbbox((0,0), t, font=f)[2]
        if w <= max_width:
            return f, t
        size -= 1
        f = load_font(size)
    while t and draw.textbbox((0,0), t + "…", font=f)[2] > max_width:
        t = t[:-1]
    return f, (t + "…") if t else ""

def _text_center_in_rect(draw, text, font, rect, fill=(255,255,255)):
    l,t,r,b = rect
    w = r - l
    h = b - t
    bbox = draw.textbbox((0,0), text, font=font)
    tw = bbox[2]-bbox[0]
    th = bbox[3]-bbox[1]
    x = l + (w - tw)//2
    y = t + (h - th)//2
    draw.text((x,y), text, font=font, fill=fill)
    
    ######## new code for Deleting a table and adding a rectangle
def _draw_segmented_selector(
    draw, cx, y, width, height, labels, font,
    active=0,
    border=(15, 72, 125),  # כחול כמו כותרות הטבלאות
    bg=(255, 255, 255),
    text=(30, 30, 30),
    underline=(15, 72, 125),
    radius=12
):
    """
    מצייר סרגל בחירה (segmented control) בכיוון RTL.
    labels נתונה מימין→לשמאל: [ימין, אמצע, שמאל]
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

    # מחיצות + טקסטים (RTL)
    for i, raw_label in enumerate(labels):
        seg_right = right - int(i * seg_w)
        seg_left  = right - int((i + 1) * seg_w)
        if i != 0:
            draw.line([(seg_right, top), (seg_right, bottom)], fill=border, width=2)

        label = fix_hebrew(raw_label)
        max_w = int(seg_w) - 12
        fit_font, fit_text = _fit_text_to_width(draw, label, font, max_w, min_size=12)

        tw, th = _text_size(draw, fit_text, fit_font)
        tx = seg_left + (seg_right - seg_left - tw) // 2
        ty = top + (height - th) // 2
        draw.text((tx, ty), fit_text, font=fit_font, fill=text)

    # קו הדגשה מתחת לפריט הפעיל (אופציונלי)
    if 0 <= active < n:
        seg_right = right - int(active * seg_w)
        seg_left  = right - int((active + 1) * seg_w)
        draw.line([(seg_left + 8, bottom - 2), (seg_right - 8, bottom - 2)], fill=underline, width=3)
    ####### end of the new code #######

def _fmt_num(n, digits=2):
    try:
        v = float(n)
    except Exception:
        return "-"
    s = f"{v:,.{digits}f}"
    return s

def _draw_table_full(draw, x, y, w, headers, rows, header_font, cell_font, col_fracs):
    row_h = 40
    head_h = 42
    pad    = 10
    hfill  = (12,52,87)
    htext  = (255,255,255)
    grid   = (220,220,220)
    zebra1 = (255,255,255)
    zebra2 = (245,245,245)
    col_ws = [int(w*f) for f in col_fracs]
    col_ws[-1] = w - sum(col_ws[:-1])
    col_x = x
    for i, h in enumerate(headers):
        cw = col_ws[i]
        draw.rectangle([col_x, y, col_x+cw, y+head_h], fill=hfill)
        txt = fix_hebrew(h)
        # tb = draw.textbbox((0,0), txt, font=header_font)
        # draw.text((col_x + (cw - (tb[2]-tb[0]))//2, y + (head_h - (tb[3]-tb[1]))//2),
        #           txt, font=header_font, fill=htext)
        hdr_rect = (col_x + 6, y + 6, col_x + cw - 6, y + head_h - 6)   # שוליים קטנים
        fit_font, fit_text = _fit_or_ellipsis(draw, fix_hebrew(h), header_font, cw - 12)
        _text_center_in_rect(draw, fit_text, fit_font, hdr_rect, fill=htext)
        col_x += cw
    col_x = x
    for i in range(len(headers)+1):
        draw.line([(col_x, y), (col_x, y+head_h)], fill=grid, width=1)
        if i < len(headers):
            col_x += col_ws[i]
    draw.line([(x, y+head_h), (x+w, y+head_h)], fill=grid, width=1)
    cy = y + head_h
    for r_i, row in enumerate(rows):
        bg = zebra1 if (r_i % 2 == 0) else zebra2
        draw.rectangle([x, cy, x+w, cy+row_h], fill=bg)
        col_x = x
        
        
        for i, cell in enumerate(row):
            cw = col_ws[i]
            raw = "" if cell is None else str(cell)
            #  עברית → RTL פעם אחת; מספרים/אנגלית נשארים LTR
            if _is_numeric_text(raw):
                txt = raw
                align_right = True
            else:
                txt = fix_hebrew(raw)
                align_right = False
                
            # קיצור לפי רוחב התא (בלי עיבוד RTL נוסף)
            txt = ellipsize_raw(draw, txt, cell_font, max_w=cw - 2*pad)
             #יישור: מספרים לימין; טקסט לשמאל
            tb = draw.textbbox((0, 0), txt, font=cell_font)
            if align_right:
                tx = col_x + cw - pad - (tb[2] - tb[0])
                
            else:
                tx = col_x + pad
            ty = cy + (row_h - (tb[3]-tb[1]))//2
            draw.text((tx, ty), txt, font=cell_font, fill=(20,20,20))
            draw.line([(col_x, cy), (col_x, cy+row_h)], fill=grid, width=1)
            col_x += cw
                      
        draw.line([(x+w, cy), (x+w, cy+row_h)], fill=grid, width=1)
        draw.line([(x, cy+row_h), (x+w, cy+row_h)], fill=grid, width=1)
        cy += row_h
    return cy

# ==== small helpers ====
def _text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]

def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _ensure_num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df

def _draw_centered(draw, text, cx, cy, font, fill=(20, 20, 20)):
    t = fix_hebrew(text)
    bbox = draw.textbbox((0,0), t, font=font)
    x = cx - (bbox[2]-bbox[0]) // 2
    y = cy - (bbox[3]-bbox[1]) // 2
    draw.text((x, y), t, font=font, fill=fill)

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

def _draw_centered_fit(draw, text, cx, y_center, max_width, base_size=28, min_size=14, color=(20,20,20)):
    """מצייר טקסט ממורכז סביב y_center, ומקטין פונט עד שהטקסט נכנס לרוחב max_width."""
    t = fix_hebrew(text)
    size = int(base_size)
    f = load_font(size)
    # מקטינים עד שנכנס לרוחב
    while size > min_size and draw.textbbox((0, 0), t, font=f)[2] > max_width:
        size -= 1
        f = load_font(size)
    tw, th = _text_size(draw, t, f)
    draw.text((cx - tw // 2, y_center - th // 2), t, font=f, fill=color)

# --- Donut (דונאט) ---
def _draw_donut(
    draw, cx, cy, outer_r, inner_r,
    segments,                 # [(label, value), ...]
    title, title_font, label_font,
    palette=None
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
    # _draw_centered(draw, title, cx, cy - outer_r - 45, title_font, (20, 20, 20))
    # כותרת מעל הדונאט – התאמה אוטומטית לרוחב הדונאט
    title_y = cy - outer_r - 36
    max_w   = int(outer_r * 2)          # קוטר הדונאט
    base_sz = getattr(title_font, "size", 28)
    _draw_centered_fit(draw, title, cx, title_y, max_w, base_size=base_sz, min_size=16, color=(20,20,20))


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

    # תווית מתחת (הפריט הגדול ביותר) + קו דק
    label, val = max(segments, key=lambda t: float(t[1]))
    pct = (float(val) / total) * 100.0
    leader_y = cy + outer_r
    draw.line([(cx, cy + outer_r - 2), (cx, leader_y + 10)], fill=(170, 170, 170), width=2)

    txt = f"{fix_hebrew(label)} {val:.2f} ({pct:.0f}%)"
    tb = draw.textbbox((0, 0), txt, font=label_font)
    draw.text((cx - (tb[2] - tb[0]) // 2, leader_y + 20), txt, font=label_font, fill=(90, 90, 90))


def _donut_config_for_bucket(bucket_label: str, df_curr: pd.DataFrame) -> dict:
    """
    בונה מקטעים לדונאטס מתוך אותו יקום של הטבלאות בשקופית:
    - מסנן לפי ARM (bucket)
    - "לא מוחרג": sub_afik ∈ EXCLUDED_SUB_AFIK
    - חוב בעייתי בלבד: debt_forum_type ∈ BAD_TYPES
    - ערך לחיבור: 'שווי נייר' (sec_value)
    - TOP-3 + 'אחר'
    מיפוי:  geo ← קבוצת לווים, collateral  ← ענף, liquidity ← דירוג
    """
    # 1) ARM bucket
    df = _filter_aram_bucket(df_curr.copy(), bucket_label)

    # 2) "לא מוחרג" – זו בדיוק אותה בחירה כמו בטבלאות
    if 'sub_afik' in df.columns:
        df = df[df['sub_afik'].isin(EXCLUDED_SUB_AFIK)]

    # 3) חוב בעייתי בלבד
    forum_col = _find_col(df, ["debt_forum_type", "סיווג פורום חוב"])
    if forum_col:
        df = df[df[forum_col].isin(BAD_TYPES)]
    else:
        df = df.iloc[0:0]

    # 4) עמודת סכום
    val_col = _find_col(df, ["sec_value", "שווי נייר", "שווי", "שווי שוק"]) or "sec_value"
    df[val_col] = pd.to_numeric(df.get(val_col, 0), errors="coerce").fillna(0.0)

    # 5) עמודות קיבוץ לפי שלוש הטבלאות
    sector_col = _find_col(df, ["תאור ענף", "תיאור ענף", "ענף"])
    group_col  = _find_col(df, ["תאור קבוצת לווים", "שם קבוצת לווים", "קבוצת לווים"])

    # דירוג: קבוע > מעלות > מידרג > 'ללא דירוג'
    rating_fixed = _find_col(df, ["דירוג קבוע לנייר", "דרוג קבוע לנייר", "דרג קבוע לנייר"])
    maalot_col   = _find_col(df, ["דירוג מעלות לנייר", "דרוג מעלות לנייר", "דירוג מעלות", "דרוג מעלות"])
    midrug_col   = _find_col(df, ["דירוג מידרג לנייר", "דרוג מידרג לנייר", "דירוג מידרג", "דרוג מידרג"])

    if rating_fixed:
        df["__rating_label"] = df[rating_fixed].astype(str).str.strip().replace({"": "ללא דירוג"})
    else:
        df["__rating_label"] = ""
        if maalot_col:
            df["__rating_label"] = df[maalot_col].astype(str).str.strip()
        if midrug_col:
            df["__rating_label"] = df["__rating_label"].mask(
                df["__rating_label"].eq("") | df["__rating_label"].isna(),
                df[midrug_col].astype(str).str.strip()
            )
        df["__rating_label"] = df["__rating_label"].replace({"": "ללא דירוג", "nan": "ללא דירוג"})

    def _segments_by(col_name: str | None) -> list[tuple[str, float]]:
        """מחזיר [(label, value)] ל־TOP-3 + 'אחר' לפי סכום שווי נייר."""
        if not col_name or col_name not in df.columns or df.empty:
            return [("אחר", float(df[val_col].sum()))] if not df.empty else [("100%", 1.0)]
        g = df.groupby(col_name, dropna=False)[val_col].sum().sort_values(ascending=False)
        total = float(g.sum()) or 1.0
        top = g.head(3)
        segs = [(str(k), float(v)) for k, v in top.items()]
        other = total - float(top.sum())
        if other > 0:
            segs.append(("אחר", other))
        return segs

    # === המיפוי לדונאטס בדיוק כמו באפיון/צילום ===
    segments_geo        = _segments_by(group_col)         # חשיפה גאוגרפית ← קבוצת לווים
    segments_collateral = _segments_by(sector_col)        # ביטחונות ← ענף
    segments_liquidity  = _segments_by("__rating_label")  # סחירות ← דירוג

    return {
        "geo":        segments_geo,
        "collateral": segments_collateral,
        "liquidity":  segments_liquidity,
    }

# =========================
# Metric calculators
# =========================
#ריכוז מנהלים
def metrics_exec_summary(df_case: pd.DataFrame) -> dict:
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
        'excluded_pct_of_portfolio': excluded_pct_of_portfolio,
        'portfolio_value_total': portfolio_value_total,
        'bad_pct_of_portfolio': bad_pct_of_portfolio,
        'bad_value_total': bad_value_total,
        'bad_share_by_value': bad_share_by_value,
    }

#אחוז מתיק האשראי הלא מוּחרג
def metric_excluded_bad_pct(df_case_curr: pd.DataFrame, df_case_prev: pd.DataFrame | None) -> tuple[float|None, float|None]:
    def calc(df_case: pd.DataFrame) -> float | None:
        if df_case is None or df_case.empty or 'sub_afik' not in df_case.columns:
            return None
        excluded = df_case[df_case['sub_afik'].isin(EXCLUDED_SUB_AFIK)]
        denom = float(excluded['sec_value'].sum())
        if denom <= 0:
            return None
        numer = float(excluded[excluded['debt_forum_type'].isin(BAD_TYPES)]['sec_value'].sum())
        return numer / denom
    return calc(df_case_curr), calc(df_case_prev)

#אחוז מכלל התיק
def metric_total_bad_pct(df_case_curr: pd.DataFrame, df_case_prev: pd.DataFrame | None) -> tuple[float|None, float|None]:
    def calc(df_case: pd.DataFrame) -> float | None:
        if df_case is None or df_case.empty or 'pct_of_portfolio_leval' not in df_case.columns:
            return None
        return float(df_case[df_case['debt_forum_type'].isin(BAD_TYPES)]['pct_of_portfolio_leval'].sum())
    return calc(df_case_curr), calc(df_case_prev)

# סה"כ ניירות בעייתיים
def metric_total_bad_count(df_case: pd.DataFrame | None) -> int | None:
    if df_case is None or df_case.empty:
        return None
    bad_df = df_case[df_case['debt_forum_type'].isin(BAD_TYPES)]
    if 'sec_id' in bad_df.columns:
        s = bad_df['sec_id'].dropna().astype(str).str.strip()
        return int(s.nunique())
    return int(bad_df.shape[0])

# כניסה/יציאה מסיווג כחוב בעייתי
def metric_bad_entry_exit(df_case_curr: pd.DataFrame | None,
                          df_case_prev: pd.DataFrame | None) -> tuple[int | None, int | None]:
    if df_case_curr is None or df_case_curr.empty or df_case_prev is None or df_case_prev.empty:
        return None, None
    def bad_set(df: pd.DataFrame) -> set:
        if 'sec_id' in df.columns:
            tmp = df[['sec_id', 'debt_forum_type']].copy()
            tmp['is_bad'] = tmp['debt_forum_type'].isin(BAD_TYPES)
            g = tmp.groupby('sec_id')['is_bad'].max()
            return set(g.index[g])
        for key_col in ('ISIN', 'שם נייר', 'SecurityID', 'מספר נייר'):
            if key_col in df.columns:
                tmp = df[[key_col, 'debt_forum_type']].copy()
                tmp['is_bad'] = tmp['debt_forום חוב'].isin(BAD_TYPES)
                g = tmp.groupby(key_col)['is_bad'].max()
                return set(g.index[g])
        return set(df.index[df['debt_forum_type'].isin(BAD_TYPES)])
    prev_bad = bad_set(df_case_prev)
    curr_bad = bad_set(df_case_curr)
    entries = len(curr_bad - prev_bad)
    exits   = len(prev_bad - curr_bad)
    return entries, exits

# מספר שינוי סיווג
def metric_class_change_count(df_curr, df_prev):
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
        return s.sort_index()
    curr_labels = one_label_per_sec(df_curr)
    prev_labels = one_label_per_sec(df_prev)
    common_ids = curr_labels.index.intersection(prev_labels.index)
    if common_ids.empty:
        return 0
    aligned_curr = curr_labels.reindex(common_ids)
    aligned_prev = prev_labels.reindex(common_ids)
    changes = (aligned_curr != aligned_prev).sum()
    return int(changes)

#סה"כ ניירות בפיגור/מסופק
def metric_late_or_delivered_count(df_case):
    if df_case is None or df_case.empty:
        return None
    mask = df_case['debt_forum_type'].isin(['בפיגור', 'מסופק'])
    filtered = df_case[mask]
    if filtered.empty:
        return 0
    if 'sec_id' in filtered.columns:
        s = filtered['sec_id'].dropna().astype(str).str.strip()
        return int(s.nunique())
    return int(filtered.shape[0])

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
                       curr_excluded_bad_pct: float | None, prev_excluded_bad_pct: float | None,
                       curr_total_bad_pct: float | None,    prev_total_bad_pct: float | None,
                       curr_bad_count: int | None,           prev_bad_count: int | None,
                       curr_bad_entries: int | None,         prev_bad_entries: int | None,
                       curr_bad_exits: int | None,    prev_bad_exits: int | None,
                       curr_class_changes: int | None, prev_class_changes: int | None,
                       curr_late_or_delivered: int | None = None, prev_late_or_delivered: int | None = None) -> Image.Image:
    W, H = 1600, 900
    slide = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(slide)
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

    def draw_percent(cx, top_y, title, value, bw=260, bh=140):
        l = cx - bw//2
        draw.rectangle((l, top_y, l+bw, top_y+bh), outline=(180,180,180), width=2)
        draw_centered_text(title, cx, top_y+35, box_t_f, (70,70,70))
        if value is None:
            txt, fill = "--%", (200,200,200)
        else:
            txt, fill = f"{value*100:.2f}%", (0,0,0)
        draw_centered_raw(txt, cx, top_y+95, pct_f, fill)

    # שני כרטיסים זה לצד זה בכל צד
    gapx   = 50         # מרווח בין הכרטיסים
    bw     = 260        # תואם לברירת המחדל בפונקציה
    top_y  = 300

# שמאל (רבעון קודם)
    left_c1 = left_x - (bw//2 + gapx//2)   # "אחוז מכלל התיק"
    left_c2 = left_x + (bw//2 + gapx//2)   # "אחוז מתיק האשראי הלא מוּחרג"
    draw_percent(left_c1, top_y, "אחוז מכלל התיק",               prev_total_bad_pct, bw=bw)
    draw_percent(left_c2, top_y, "אחוז מתיק האשראי\nהלא מוּחרג", prev_excluded_bad_pct, bw=bw)

# ימין (רבעון נוכחי)
    right_c1 = right_x - (bw//2 + gapx//2)
    right_c2 = right_x + (bw//2 + gapx//2)
    draw_percent(right_c1, top_y, "אחוז מכלל התיק",               curr_total_bad_pct, bw=bw)
    draw_percent(right_c2, top_y, "אחוז מתיק האשראי\nהלא מוּחרג", curr_excluded_bad_pct, bw=bw)

    
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
    # draw_stats(left_x,  640, prev_bad_count, prev_bad_entries, prev_bad_exits, prev_class_changes, prev_late_or_delivered)
    # draw_stats(right_x, 640, curr_bad_count, curr_bad_entries, curr_bad_exits, curr_class_changes, curr_late_or_delivered)
    bh = 140          # הגובה של כרטיס draw_percent (כמו בברירת המחדל של הפונקציה)
    stats_h = 230     # הגובה של תיבת הסטטיסטיקות בתוך draw_stats
    gap_stats = 32    # המרווח בין שורת הכרטיסים לבין תיבות הסטטיסטיקות (קטן יותר = יותר צמוד)
    stats_top_y = top_y + bh + gap_stats

# לא לדרוס את אזור הפוטר – נשמור רווח מינימלי מלמטה
    safe_y = H - SAFE_BOTTOM
    stats_top_y = min(stats_top_y, safe_y - stats_h)

    draw_stats(left_x,  stats_top_y, prev_bad_count, prev_bad_entries, prev_bad_exits,
            prev_class_changes, prev_late_or_delivered)
    draw_stats(right_x, stats_top_y, curr_bad_count, curr_bad_entries, curr_bad_exits,
            curr_class_changes, curr_late_or_delivered)
    
    return slide

# =========================
# NEW: accurate computations for the 3 new slides
# =========================

def _filter_aram_bucket(df, bucket_label: str):
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

def _top3_for_group(df_case_full: pd.DataFrame, group_col: str) -> list[tuple[str, str, str]]:
    """
    מחזיר עד 3 רשומות בסדר:
    (אחוז מכלל התיק, אחוז מתיק אשראי לא מוחרג, תיאור הקבוצה)
    שני המדדים מחושבים על יקום 'לא מוחרג' לפי EXCLUDED_SUB_AFIK.
    """
    if not group_col or group_col not in df_case_full.columns or df_case_full.empty:
        return []

    # עמודות בסיס
    val_col = 'sec_value' if 'sec_value' in df_case_full.columns else None
    pct_col = 'pct_of_portfolio_leval' if 'pct_of_portfolio_leval' in df_case_full.columns else None

    # סינון ליקום האשראי הלא-מוחרג 
    if 'sub_afik' in df_case_full.columns:
        df_ne = df_case_full[df_case_full['sub_afik'].isin(EXCLUDED_SUB_AFIK)].copy()
    else:
        df_ne = df_case_full.copy()

    # --- אחוז מכלל התיק (מתוך הלא-מוחרג) ---
    if pct_col and not df_ne.empty:
        s_total_pct = df_ne.groupby(group_col, dropna=False)[pct_col].sum()
    else:
        # נפילה חכמה לרזרבה במקרה ואין העמודה – נחשב לפי value/total
        if val_col and not df_ne.empty:
            total_val_ne = float(pd.to_numeric(df_ne[val_col], errors='coerce').fillna(0).sum())
            s_total_pct = (df_ne.groupby(group_col, dropna=False)[val_col]
                              .sum()
                              .div(total_val_ne if total_val_ne else 1.0))
        else:
            s_total_pct = pd.Series(0.0, index=df_case_full.groupby(group_col, dropna=False).size().index)

    # --- אחוז מתיק אשראי לא מוחרג (כבר היה נכון – נשאר) ---
    if val_col and not df_ne.empty:
        s_excl_val = (df_ne.groupby(group_col, dropna=False)[val_col]
                         .sum())
        denom_excl = float(s_excl_val.sum())
        s_excl_pct = s_excl_val.div(denom_excl if denom_excl else 1.0)
    else:
        s_excl_pct = pd.Series(0.0, index=s_total_pct.index)

    # איחוד, מיון והחזרה
    agg = pd.concat(
        [s_total_pct.rename('pct_total'), s_excl_pct.rename('pct_excl')],
        axis=1
    ).fillna(0.0)

    agg = agg.sort_values(['pct_excl', 'pct_total'], ascending=False).head(3)

    out = []
    for label, row in agg.iterrows():
        out.append((fmt_pct(row['pct_total']), fmt_pct(row['pct_excl']), str(label)))

    return out if out else [("","", "")]

def render_bad_debts_page(account_name_display: str,
                          bucket_label: str,
                          df_curr: pd.DataFrame,
                          date_str: str | None = None) -> Image.Image:
    """
    תיאור חובות בעייתיים בתיק אשראי לא מוחרג:
    - מסנן לבקט ARM
    - מסנן ל*מוחרגים* (EXCLUDED_SUB_AFIK)
    - מסנן ל-BAD - "אחוז מאשראי לא מוחרג" = value_row / sum(value בכל המוחרגים בתיק)
    """
    W, H = 1600, 900
    img  = Image.new("RGB", (W, H), (245,245,245))
    draw = ImageDraw.Draw(img)

    title_f   = load_font(56)
    sub_f     = load_font(40)
    header_f  = load_font(22)
    cell_f    = load_font(20)
    small_f   = load_font(22)

    # כותרות
    t1 = "תיאור חובות בעייתיים בתיק אשראי לא מוחרג"
    _draw_centered(draw, t1,           W//2, 60,  title_f, (20,20,20))
    _draw_centered(draw, bucket_label, W//2, 120, sub_f,   (20,20,20))

    # בסיס לחישוב דנומינטור "לא מוחרג": מתוך כל התיק (לפני BAD)
    df_bucket = _filter_aram_bucket(df_curr.copy(), bucket_label)
    if 'sub_afik' in df_bucket.columns:
        df_excluded_all = df_bucket[df_bucket['sub_afik'].isin(EXCLUDED_SUB_AFIK)].copy()
    else:
        df_excluded_all = df_bucket.copy()
    total_ne_value = pd.to_numeric(df_excluded_all.get('sec_value', 0), errors="coerce").fillna(0).sum()

    # כעת הטבלה: מוחרגים + BAD
    col_forum = _find_col(df_bucket, ["debt_forum_type","סיווג פורום חוב"])
    df = df_excluded_all.copy()
    if col_forum:
        df = df[df[col_forum].isin(BAD_TYPES)].copy()

    # עמודות להצגה
    col_desc   = _find_col(df, ["תאור נייר","תיאור נייר","שם נייר","security_name"])
    col_qty    = _find_col(df, ["כמות","quantity"])
    col_value  = _find_col(df, ["sec_value","שווי נייר","שווי שוק"])
    col_maalot = _find_col(df, ["דירוג מעלות לנייר","דרוג מעלות לנייר","דירוג מעלות","דרוג מעלות"])
    col_midrug = _find_col(df, ["דירוג מידרג לנייר","דרוג מידרג לנייר","דירוג מידרג","דרוג מידרג"])
    col_machem = _find_col(df, ["מח\"מ מחושב","מחמ מחושב","מח\"מ"])
    col_yield  = _find_col(df, ["תשואה ברוטו","תשואה","yld"])

    # שורות
    rows = []
    for _, r in df.iterrows():
        val = float(pd.to_numeric(r.get(col_value, 0), errors="coerce")) if col_value else 0.0
        pct_ne = (val / total_ne_value) if total_ne_value else 0.0
        rows.append((
            r.get(col_forum, ""),
            _fmt_num(r.get(col_yield, ""), 2),
            _fmt_num(r.get(col_machem, ""), 2),
            str(r.get(col_midrug, "")),
            str(r.get(col_maalot, "")),
            f"{pct_ne*100:.2f}%" if total_ne_value else "--%",
            _fmt_num(val, 2),
            _fmt_num(r.get(col_qty, ""), 2),
            r.get(col_desc, "")
        ))

    # שורת Total
    if len(rows) > 0:
        sum_qty   = pd.to_numeric(df[col_qty],  errors="coerce").fillna(0).sum() if col_qty  else 0
        sum_value = pd.to_numeric(df[col_value],errors="coerce").fillna(0).sum() if col_value else 0
        sum_pct   = (sum_value/total_ne_value) if total_ne_value else 0.0
        rows.append((
            "Total", "", "", "", "",
            f"{sum_pct*100:.2f}%" if total_ne_value else "--%",
            _fmt_num(sum_value, 2),
            _fmt_num(sum_qty, 2),
            ""
        ))
    else:
        rows.append((fix_hebrew("מסופק"), "", "", "", "", "", "", "", ""))
        rows.append(("Total",  "", "", "", "", "", "", "", ""))

    # טבלה
    table_w = 1500
    left_x  = (W - table_w)//2
    top_y   = 200
    headers = [
        "סיווג פורום חוב","תשואה ברוטו","מח\"מ מחושב","דרוג מידרג לנייר",
        "דרוג מעלות לנייר","אחוז מאשראי לא מוחרג","שווי נייר","כמות","תאור נייר"
    ]
    col_fracs = [0.10, 0.10, 0.10, 0.10, 0.11, 0.13, 0.10, 0.09, 0.17]
    end_y = _draw_table_full(draw, left_x, top_y, table_w,
                             headers, rows, header_f, cell_f, col_fracs)

    # הערה
    note = "נכון לתאריך הדוח. אין בתיק חשיפה נוספת לנכסי חוב או נכסים אחרים שהונפקו על ידי מנפיקים אלה"
    # _draw_centered(draw, note, W//2, end_y + 60, small_f, (30,30,30))
    safe_y = H - SAFE_BOTTOM                  # קו עליון של אזור הפוטר
    note_y = min(end_y + 60, safe_y - 80)     # לא לגעת בפוטר; 80px ריווח לנשימה
    _draw_centered(draw, note, W//2, note_y, small_f, (30,30,30))

    # סה\"כ אשראי לא מוחרג בתחתית (של כלל המוחרגים, לא רק BAD)
    bottom_text = "סה\"כ אשראי לא מוחרג"
    bt = fix_hebrew(bottom_text)
    tb1 = draw.textbbox((0,0), bt, font=small_f)
    val = _fmt_num(total_ne_value, 0)
    tb2 = draw.textbbox((0,0), val, font=small_f)
    gap = 35
    cx = W//2 - ( (tb1[2]-tb1[0]) + gap + (tb2[2]-tb2[0]) )//2
    # yb = H - 90
    safe_y = H - SAFE_BOTTOM
    yb = safe_y - 30  
    
    draw.text((cx, yb), bt, font=small_f, fill=(20,20,20))
    draw.text((cx + (tb1[2]-tb1[0]) + gap, yb), val, font=small_f, fill=(20,20,20))

    return img

def render_bad_debts_page_alt(
    account_name_display: str,
    bucket_label: str,
    df_curr: pd.DataFrame,
    date_str: str | None = None
) -> Image.Image:
    """עמוד 3: שלוש טבלאות התפלגות על BAD בתוך יקום 'לא מוחרג' (EXCLUDED_SUB_AFIK),
    'אחוז מאשראי לא מוחרג' = סכום שווי בקבוצה / סכום שווי BAD בתוך הלא-מוחרג."""
    # קנבס
    W, H = 1600, 900
    img  = Image.new("RGB", (W, H), (245,245,245))
    draw = ImageDraw.Draw(img)

    title_f  = load_font(56)
    sub_f    = load_font(40)
    header_f = load_font(22)
    cell_f   = load_font(20)
    donut_title_f = load_font(28)
    donut_label_f = load_font(18)

    # כותרות
    _draw_centered(draw, "תיאור חובות בעייתיים בתיק אשראי לא מוחרג", W//2, 60,  title_f, (20,20,20))
    _draw_centered(draw, bucket_label,                                       W//2, 120, sub_f,   (20,20,20))

    df = df_curr.copy()
    df = _filter_aram_bucket(df, bucket_label)

    # יקום "לא מוחרג" - סינון
    if 'sub_afik' in df.columns:
        df_ne = df[df['sub_afik'].isin(EXCLUDED_SUB_AFIK)].copy()
    else:
        df_ne = df.copy()

    # ערך (value) ודגל BAD
    val_col  = _find_col(df_ne, ["sec_value","שווי נייר","שווי","שווי שוק"]) or "sec_value"
    df_ne[val_col] = pd.to_numeric(df_ne.get(val_col, 0), errors="coerce").fillna(0.0)

    col_forum = _find_col(df_ne, ["debt_forum_type","סיווג פורום חוב"])
    df_bad = df_ne[df_ne[col_forum].isin(BAD_TYPES)].copy() if col_forum else df_ne.iloc[0:0].copy()

    # דנומינטור נכון: סכום BAD בתוך הלא-מוחרג
    denom_bad_ne = float(df_bad[val_col].sum())

    # ---- בניית עמודות קיבוץ ----
    # דירוג: קודם קבוע; אם אין — ממזגים מעלות/מידרג; אם אין — "ללא דירוג"
    rating_fixed = _find_col(df_bad, ["דירוג קבוע לנייר","דרוג קבוע לנייר","דרג קבוע לנייר"])
    maalot_col   = _find_col(df_bad, ["דירוג מעלות לנייר","דרוג מעלות לנייר","דירוג מעלות","דרוג מעלות"])
    midrug_col   = _find_col(df_bad, ["דירוג מידרג לנייר","דרוג מידרג לנייר","דירוג מידרג","דרוג מידרג"])

    if rating_fixed:
        df_bad["__rating_label"] = df_bad[rating_fixed].astype(str).str.strip().replace({"": "ללא דירוג"})
    else:
        # בוחרים את הראשון שאינו ריק מבין מעלות/מידרג
        df_bad["__rating_label"] = ""
        if maalot_col:
            df_bad["__rating_label"] = df_bad[maalot_col].astype(str).str.strip()
        if midrug_col:
            df_bad["__rating_label"] = df_bad["__rating_label"].mask(
                df_bad["__rating_label"].eq("") | df_bad["__rating_label"].isna(),
                df_bad[midrug_col].astype(str).str.strip()
            )
        df_bad["__rating_label"] = df_bad["__rating_label"].replace({"": "ללא דירוג", "nan": "ללא דירוג"})

    sector_col   = _find_col(df_bad, ["תאור ענף","תיאור ענף","ענף"])
    group_col    = _find_col(df_bad, ["תאור קבוצת לווים","שם קבוצת לווים","קבוצת לווים"])

    # ---- פונקציית TOP-3 לכל קיבוץ, אחוז מתוך denom_bad_ne ----
    def _top_rows(group_col_name: str | None) -> list[tuple[str,str,str]]:
        if not group_col_name or group_col_name not in df_bad.columns or df_bad.empty:
            return []
        agg = (df_bad.groupby(group_col_name, dropna=False)[val_col]
                    .sum().sort_values(ascending=False).head(3))
        rows = []
        for label, v in agg.items():
            pct = (float(v) / denom_bad_ne) if denom_bad_ne > 0 else 0.0
            rows.append((fmt_pct(pct), fmt_km(v), str(label)))
        return rows

    rows_right = _top_rows("__rating_label")         # דירוג
    rows_mid   = _top_rows(group_col)                # קבוצת לווים
    rows_left  = _top_rows(sector_col)               # ענף

    # Fallback אם אין נתונים
    def _fallback(r): 
        return r if r else [("","", "")]
    rows_right = _fallback(rows_right)
    rows_mid   = _fallback(rows_mid)
    rows_left  = _fallback(rows_left)

    # פריסה וציור
    y0 = 200
    gap = 60
    w_small = 440
    x_left  = 80
    x_mid   = x_left + w_small + gap
    x_right = x_mid + w_small + gap

    headers_left  = ["אחוז מאשראי לא מוחרג","שווי נייר","תאור ענף"]
    headers_mid   = ["אחוז מאשראי לא מוחרג","שווי נייר","תאור קבוצת לווים"]
    headers_right = ["אחוז מאשראי לא מוחרג","שווי נייר","דרוג קבוע לנייר"]
    fracs = [0.28, 0.22, 0.50]

    _draw_table_full(draw, x_left,  y0, w_small, headers_left,  rows_left,  header_f, cell_f, fracs)
    _draw_table_full(draw, x_mid,   y0, w_small, headers_mid,   rows_mid,   header_f, cell_f, fracs)
    _draw_table_full(draw, x_right, y0, w_small, headers_right, rows_right, header_f, cell_f, fracs)

    # הערה
    note = "נכון לתאריך הדוח, אין בתיק חשיפה נוספת לנכסי חוב או נכסים אחרים שהונפקו על ידי קבוצת הלווים הנ\"ל"
    # _draw_centered(draw, note, W//2, 360, cell_f, (30,30,30))
    safe_y = H - SAFE_BOTTOM
    _draw_centered(draw, note, W//2, min(360, safe_y - 80), cell_f, (30,30,30))

    donuts = _donut_config_for_bucket(bucket_label, df_curr)
    selector_labels = ["סכום קבוצת לווים", "תאור נייר", "תאור קבוצת לווים"]  # מימין→שמאל
    _draw_segmented_selector(draw, W//2, 390, 480, 44, selector_labels, cell_f, active=1)

    # מרכזים ורדיוסים
    # cy = 720
    ro, ri = 110, 64            
    leader_gap = 36             # מרווח קו/תווית מתחת לעיגול
    cy = H - SAFE_BOTTOM - (ro + leader_gap)
    cx_left   = int(W * 0.17)   # שמאלי: סחירות
    ro, ri = 115, 65
    cx_mid    = int(W * 0.50)   # אמצעי: ביטחונות
    cx_right  = int(W * 0.83)   # ימני: חשיפה גאוגרפית
    
    pale_blue = [(14, 134, 255)]
    dark_blue = [(19, 30, 138)]
    
    _draw_donut(
        draw, cx_right, cy, ro, ri,
        donuts["geo"],
        "התפלגות חוב בעייתי על פי חשיפה גאוגרפית",
        donut_title_f, donut_label_f,
        palette=pale_blue
    )
    
    _draw_donut(
        draw, cx_mid, cy, ro, ri,
        donuts["collateral"],
        "התפלגות חוב בעייתי על פי ביטחונות",
        donut_title_f, donut_label_f,
        palette=dark_blue
    )
    
    _draw_donut(
        draw, cx_left, cy, ro, ri,
        donuts["liquidity"],
        "התפלגות חוב בעייתי על פי סחירות",
        donut_title_f, donut_label_f,
        palette=pale_blue
    )
    return img


def _estimate_section_height(draw, title_font, header_font, cell_font, rows):
    title_h  = _text_size(draw, "כותרת", title_font)[1] + 16
    header_h = 38
    row_h    = 38
    n_rows   = max(3, len(rows))
    gap_after = 40
    
    return title_h + header_h + used_rows * row_h + gap_after

def _draw_section(draw, x, top_y, w_tbl, title, sub_font, header, rows, header_f, cell_f, aligns):
    t = fix_hebrew(title)
    tw, th = _text_size(draw, t, sub_font)
    _draw_centered(draw, title, x + w_tbl // 2, top_y + th // 2, sub_font, (0, 0, 0))
    table_top = top_y + th + 16
    _draw_table(draw, x, table_top, w_tbl, 38, header, rows, header_f, cell_f, aligns=aligns)
    used_rows = min(3, len(rows))
    next_y = table_top + 38 * (1 + used_rows) + 40
    return next_y

def _estimate_section_height(draw, title_font, header_font, cell_font, rows):
    """גובה משוער של סקשן: כותרת + טבלת header+rows."""
    title_h  = _text_size(draw, "כותרת", title_font)[1] + 14  # כותרת + רווח קטן
    header_h = 42
    row_h    = 40
    n_rows   = max(1, len(rows))
    return title_h + header_h + n_rows * row_h + 8


def _draw_table(draw, x, y, w, row_h, header, rows, header_f, cell_f, aligns=None):
    if aligns is None:
        aligns = ["right"] * len(header)
    header_bg = (15, 72, 127)
    header_fg = (255, 255, 255)
    draw.rectangle([x, y, x + w, y + row_h], fill=header_bg)
    col_x = x
    col_positions = []
    for (title, ratio), align in zip(header, aligns):
        cw = int(w * ratio)
        col_positions.append((col_x, cw, align))
        _draw_text_fit(draw, title, col_x, y, cw, row_h,
                       base_size=20, min_size=12,
                       color=header_fg, align="center")
        draw.line([(col_x, y), (col_x, y + row_h)], fill=(255, 255, 255), width=1)
        col_x += cw
    draw.line([(x + w, y), (x + w, y + row_h)], fill=(255, 255, 255), width=1)
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

def render_bad_distributions_page(
    account_name_display: str,
    bucket_label: str,
    df_curr: pd.DataFrame,
    df_prev: pd.DataFrame | None,
    date_curr: str | None = None,
    date_prev: str | None = None,
) -> Image.Image:
    """ניתוח חשיפות מהותיות – השוואה תקופתית (שמאל=נוכחי, ימין=קודם)."""
    W, H = 1600, 900
    img  = Image.new("RGB", (W, H), (245,245,245))
    draw = ImageDraw.Draw(img)

    title_f  = load_font(56)
    sub_f    = load_font(40)
    date_f   = load_font(36)
    header_f = load_font(22)
    cell_f   = load_font(20)

    # כותרות עליונות
    _draw_centered(draw, "ניתוח חשיפות מהותיות בתיק אשראי – השוואה תקופתית", W // 2, 70,  title_f, (20, 20, 20))
    _draw_centered(draw, bucket_label,                                      W // 2, 120, sub_f,   (20, 20, 20))

    # פילוח לפי ARM לבקט
    dfL = _filter_aram_bucket(df_curr.copy(), bucket_label)
    dfR = _filter_aram_bucket(df_prev.copy(), bucket_label) if df_prev is not None else None

    # עמודות מועמדות
    def cols(dfx):
        if dfx is None: 
            return (None, None, None)
        borrower_col = _find_col(dfx, ["תאור מנפיק", "תיאור מנפיק", "לווה"])
        sector_col   = _find_col(dfx, ["תאור ענף", "תיאור ענף", "ענף"])
        group_col    = _find_col(dfx, ["תאור קבוצת לווים", "שם קבוצת לווים", "קבוצת לווים"])
        return borrower_col, sector_col, group_col

    bL, sL, gL = cols(dfL)
    bR, sR, gR = cols(dfR) if dfR is not None else (None, None, None)

    # Top-3 לכל צד
    left_borrowers = _top3_for_group(dfL, bL) if bL else [("","","")]
    left_sectors   = _top3_for_group(dfL, sL) if sL else [("","","")]
    left_groups    = _top3_for_group(dfL, gL) if gL else [("","","")]

    right_borrowers = _top3_for_group(dfR, bR) if (dfR is not None and bR) else [("","","")]
    right_sectors   = _top3_for_group(dfR, sR) if (dfR is not None and sR) else [("","","")]
    right_groups    = _top3_for_group(dfR, gR) if (dfR is not None and gR) else [("","","")]

    # פריסה
    left_x, right_x, w_tbl = 120, 840, 560
    _draw_centered(draw, (date_curr or fix_hebrew("רבעון נוכחי")), left_x  + w_tbl // 2, 190, date_f, (20, 20, 20))
    _draw_centered(draw, (date_prev or fix_hebrew("רבעון קודם")),  right_x + w_tbl // 2, 190, date_f, (20, 20, 20))

    header_borrower = [("אחוז מכלל התיק", 0.18), ("אחוז מתיק אשראי לא מוחרג", 0.36), ("תאור מנפיק", 0.46)]
    header_sector   = [("אחוז מכלל התיק", 0.18), ("אחוז מתיק אשראי לא מוחרג", 0.36), ("תאור ענף",   0.46)]
    header_group    = [("אחוז מכלל התיק", 0.18), ("אחוז מתיק אשראי לא מוחרג", 0.36), ("תאור קבוצת לווים", 0.46)]
    aligns = ["right", "right", "left"]

# ----------------------
# פריסת 3 הסקשנים בצד שמאל (נוכחי)
# ----------------------
    TOP_START    = 200           # היכן להתחיל ציור העמודה
    GAP_DEFAULT  = 36            # רווח רצוי בין סקשנים
    SAFE_BOTTOM  = 140           # חייב להיות אותו ערך כמו ב-add_branding

# גבהים משוערים של כל שלושת הסקשנים (לפי עד 3 שורות בפועל)
    hL1 = _estimate_section_height(draw, sub_f, header_f, cell_f, left_borrowers)
    hL2 = _estimate_section_height(draw, sub_f, header_f, cell_f, left_sectors)
    hL3 = _estimate_section_height(draw, sub_f, header_f, cell_f, left_groups)

# כמה מקום באמת יש לנו עד הפוטר
    safe_height_L = (H - SAFE_BOTTOM) - TOP_START

# מתחילים עם רווח ברירת מחדל; אם לא נכנס — מצמצמים באופן אחיד
    gapL = GAP_DEFAULT
    totalL = hL1 + hL2 + hL3 + 2 * gapL
    if totalL > safe_height_L:
        gapL = max(16, (safe_height_L - (hL1 + hL2 + hL3)) // 2)

# מיקומים סופיים
    yL1 = TOP_START
    yL2 = yL1 + hL1 + gapL
    yL3 = yL2 + hL2 + gapL

# ודואגים שהסקשן האחרון לא יחרוג מתחת ל-SAFE_BOTTOM
    max_yL3 = H - SAFE_BOTTOM - hL3
    if yL3 > max_yL3:
        shift = yL3 - max_yL3
        yL1 -= shift
        yL2 -= shift
        yL3 -= shift

# ציור שלושת הסקשנים (שמאל)
    _draw_section(draw, left_x,  yL1, w_tbl, "לווה יחיד",   sub_f, header_borrower, left_borrowers, header_f, cell_f, aligns)
    _draw_section(draw, left_x,  yL2, w_tbl, "ענפים",       sub_f, header_sector,   left_sectors,   header_f, cell_f, aligns)
    _draw_section(draw, left_x,  yL3, w_tbl, "קבוצת לווים", sub_f, header_group,    left_groups,    header_f, cell_f, aligns)

# ----------------------
# פריסת 3 הסקשנים בצד ימין (קודם)
# ----------------------
    hR1 = _estimate_section_height(draw, sub_f, header_f, cell_f, right_borrowers)
    hR2 = _estimate_section_height(draw, sub_f, header_f, cell_f, right_sectors)
    hR3 = _estimate_section_height(draw, sub_f, header_f, cell_f, right_groups)

    safe_height_R = (H - SAFE_BOTTOM) - TOP_START
    gapR = GAP_DEFAULT
    totalR = hR1 + hR2 + hR3 + 2 * gapR
    if totalR > safe_height_R:
        gapR = max(16, (safe_height_R - (hR1 + hR2 + hR3)) // 2)

    yR1 = TOP_START
    yR2 = yR1 + hR1 + gapR
    yR3 = yR2 + hR2 + gapR

    max_yR3 = H - SAFE_BOTTOM - hR3
    if yR3 > max_yR3:
        shift = yR3 - max_yR3
        yR1 -= shift
        yR2 -= shift
        yR3 -= shift

# ציור שלושת הסקשנים (ימין)
    _draw_section(draw, right_x, yR1, w_tbl, "לווה יחיד",   sub_f, header_borrower, right_borrowers, header_f, cell_f, aligns)
    _draw_section(draw, right_x, yR2, w_tbl, "ענפים",       sub_f, header_sector,   right_sectors,   header_f, cell_f, aligns)
    _draw_section(draw, right_x, yR3, w_tbl, "קבוצת לווים", sub_f, header_group,    right_groups,    header_f, cell_f, aligns)

    return img

# =========================
# Main (unchanged except calling the new slides)
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
        case_curr = df_curr[df_curr['__case'] == case_id].copy()
        if case_curr.empty:
            print(f"CASE {case_id}: אין נתונים בקובץ הנוכחי – מדלג.")
            continue

        account_name = case_curr.iloc[0]['account_name'] if 'account_name' in case_curr.columns else f"תיק {case_id}"

        # ריכוז מנהלים 
        exec_metrics = metrics_exec_summary(case_curr)

        # שינוי סיווג 
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

        # שקפים קיימים
        exec_img  = render_exec_slide(TEMPLATE_IMAGE, account_name, exec_metrics)
        white_img = render_white_slide(account_name,
                                       curr_excl_bad_pct, prev_excl_bad_pct,
                                       curr_total_bad_pct, prev_total_bad_pct,
                                       curr_bad_count, prev_bad_count,
                                       curr_bad_entries, prev_bad_entries,
                                       curr_bad_exits, prev_bad_exits,
                                       curr_class_changes, prev_class_changes,
                                       curr_late_or_delivered, prev_late_or_delivered)

        # ====== שלושת השקפים החדשים ======
        case_bucket = {
            16396: ["ארם עד 50"],
            16397: ["ארם 50-60"],
            16398: ["ארם 60 ומעלה"],
        }
        buckets_for_this_case = case_bucket.get(case_id, ["ארם עד 50"])

        tables_imgs = []
        bad_pages   = []
        dist_pages  = []

        date_cur_str  = "30/06/2025"
        date_prev_str = "31/03/2025"

        for bucket_label in buckets_for_this_case:
        # 1) השוואה תקופתית: שמאל=נוכחי, ימין=קודם
            dist_pages.append(
                render_bad_distributions_page(
                    account_name_display=account_name,
                    bucket_label=bucket_label,
                    df_curr=case_curr,
                    df_prev=case_prev,          
                    date_curr=date_cur_str,     # תאריך נוכחי
                    date_prev=date_prev_str,    # תאריך קודם
                )
            )

        # 2) תיאור חובות בעייתיים (הראשונה)
            bad_pages.append(
                render_bad_debts_page(
                    account_name_display=account_name,
                    bucket_label=bucket_label,
                    df_curr=case_curr,
                    date_str=date_cur_str,
                )
            )

        # 3) תיאור חובות בעייתיים – השקופית השנייה (הטבלאית)
            bad_pages.append(
                render_bad_debts_page_alt(
                    account_name_display=account_name,
                    bucket_label=bucket_label,
                    df_curr=case_curr,
                    date_str=date_cur_str,
                )
            )
            
        for _img in [white_img, *dist_pages, *tables_imgs, *bad_pages]:
            add_branding(
                _img,
                show_logo=True,
                show_footer=True,
                logo_rel_h=0.07, footer_rel_h=0.038,
                pad_bottom=26, wipe_footer_bg=True
            )
        # שמירה
        out_png = os.path.join(OUTPUT_DIR, f'output_{case_id}.png')
        out_pdf = os.path.join(OUTPUT_DIR, f'output_{case_id}.pdf')
        exec_img.save(out_png)
        exec_img.save(out_pdf, save_all=True, append_images=[white_img, *dist_pages, *tables_imgs, *bad_pages])
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
        
    print("LOGO exists:", Path(LOGO_PATH).exists(), LOGO_PATH)
    print("FOOTER exists:", Path(FOOTER_PATH).exists(), FOOTER_PATH)

if __name__ == "__main__":
    main()