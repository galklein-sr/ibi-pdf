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
    f32 = ImageFont.truetype("arial.ttf", 32)
    f42 = ImageFont.truetype("arial.ttf", 42)
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

    title_f = ImageFont.truetype("arial.ttf", 60)
    sub_f   = ImageFont.truetype("arial.ttf", 40)
    box_t_f = ImageFont.truetype("arial.ttf", 28)
    pct_f   = ImageFont.truetype("arial.ttf", 56)
    small_f = ImageFont.truetype("arial.ttf", 26)

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



    draw_stats(left_x,  680, prev_bad_count, prev_bad_entries, prev_bad_exits, prev_class_changes, prev_late_or_delivered)
    draw_stats(right_x, 680, curr_bad_count, curr_bad_entries, curr_bad_exits, curr_class_changes, curr_late_or_delivered)

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

        # שמירה
        out_png = os.path.join(OUTPUT_DIR, f'output_{case_id}.png')
        out_pdf = os.path.join(OUTPUT_DIR, f'output_{case_id}.pdf')
        exec_img.save(out_png)
        exec_img.save(out_pdf, save_all=True, append_images=[white_img])
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