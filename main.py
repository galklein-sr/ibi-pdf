from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from bidi.algorithm import get_display
import arabic_reshaper
import pdfkit
from jinja2 import Template
from PyPDF2 import PdfMerger
import sys, os

config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
#Load Excel Sheets 
ids = [int(a) for a in sys.argv[1:]] if len(sys.argv) > 1 else [16396, 16397, 16398]

data_excel_file = 'DataToPDF/data.xlsx'
data_sheet_name = 'Sheet1'  # You can also use an integer index, e.g., 0
full_df = pd.read_excel(data_excel_file, sheet_name=data_sheet_name)

def _norm_case_series(s):
    s = s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)  
    return pd.to_numeric(s, errors='coerce').astype('Int64')
full_df['__case'] = _norm_case_series(full_df['מספר תיק'])

os.makedirs('outputs', exist_ok=True)

pdf_files = [] #איסוף קבצי PDF למיזוג

#print(data_df.head(5))
#Extract Values:
for case_id in ids:
    case_id = int(case_id)
    data_df = full_df[ full_df['__case'] == case_id ].copy()
    print(f"CASE {case_id}: rows = {len(data_df)}")
    if data_df.empty:
        print(f"לא נמצאו נתונים לתיק {case_id}, מדלג.")
        continue

    account_name = data_df.iloc[0]['שם חשבון']
    account_name = arabic_reshaper.reshape(account_name)
    account_name = get_display(account_name)

    problematic_debt = data_df['אחוז מהתיק'].sum()

    data_df = data_df.rename(columns={"אחוז משווי תיק לפי שיערוך אחרון": "p_o_p_b_leval"})
    data_df["p_o_p_b_leval"] = data_df["p_o_p_b_leval"].fillna(0)
    p_o_p_value = data_df['p_o_p_b_leval'].sum()

    data_df = data_df.rename(columns={"שווי נייר": "p_value"})
    data_df["p_value"] = data_df["p_value"].fillna(0)
    p_value = data_df['p_value'].sum()

    data_df = data_df.rename(columns={'סיווג פורום חוב': "debt_forum_type"})
    discared_forum_types = ['השגחה מיוחדת','במעקב מיוחד','מסופק','בפיגור']
    data_dt_filter = data_df[data_df['debt_forum_type'].isin(discared_forum_types)]
    p_o_p_filter_value = data_dt_filter['p_o_p_b_leval'].sum()
    p_debt_value = data_dt_filter['p_value'].sum()

    data_df["p_o_afik"] = 0 if p_value == 0 else data_df["p_value"] / p_value
    data_dt_filter["p_o_afik"] = 0 if p_value == 0 else data_dt_filter["p_value"] / p_value
    p_o_afik_filter = data_dt_filter['p_o_afik'].sum()

    data_df = data_df.rename(columns={'קוד אפיק ותת אפיק': "sub_afik"})
    discarded_sub_afik = [210,240,310,330,341,345,360,405,407,425,602]
    data_dt_subafik_filter = data_df[data_df['sub_afik'].isin(discarded_sub_afik)]
    p_o_p_afik_filter_value = data_dt_subafik_filter['p_o_p_b_leval'].sum()

    image = Image.open('templates/template1.png').convert("RGB")
    draw = ImageDraw.Draw(image)

    texts = [
        {
            'text': account_name,
            'position': (1030, 50),
            'font_path': 'arial.ttf',
            'font_size': 32,
            'color': (255, 255, 255)  
        },
        {
            'text': "{:.2f}".format(p_o_p_afik_filter_value*100) + '%',
            'position': (246, 132),
            'font_path': 'arial.ttf',
            'font_size': 32,
            'color': (255, 255, 255)  
        },
        {
            'text': f"{p_value/1_000_000:.2f}M",
            'position': (470, 132),
            'font_path': 'arial.ttf',
            'font_size': 32,
            'color': (255, 255, 255)  
        },
        {
            'text': "{:.2f}".format(p_o_p_filter_value*100) + '%',
            'position': (246, 440),
            'font_path': 'arial.ttf',
            'font_size': 42,
            'color': (255, 255, 255)  
        },
        {
            'text': "{:.2f}".format(p_debt_value) ,
            'position': (470, 440),
            'font_path': 'arial.ttf',
            'font_size': 42,
            'color': (255, 255, 255)  
        }
        ,
        {
            'text': "{:.2f}".format(p_o_afik_filter*100) + '%',
            'position': (1020, 290),
            'font_path': 'arial.ttf',
            'font_size': 42,
            'color': (255, 255, 255)  
        }
    ]
# Add each text to the image
    for item in texts:
        font = ImageFont.truetype(item['font_path'], item['font_size'])
        cx, cy = item['position']
        bbox = draw.textbbox((0, 0), item['text'], font=font)
        x = cx - (bbox[2]-bbox[0]) // 2
        y = cy - (bbox[3]-bbox[1]) // 2
        draw.text((x, y), item['text'], font=font, fill=item['color'])



# Save the edited image (unique name)
    output_image = f'outputs/output_{case_id}.png'
    output_pdf   = f'outputs/output_{case_id}.pdf'
    image.save(output_image)
    image.save(output_pdf)
    pdf_files.append(output_pdf)
    print(f"created: {output_image}, {output_pdf}")

# Merge all PDF files into one
if pdf_files:
    merger = PdfMerger()
    for p in pdf_files:
        merger.append(p)
    merger.write('outputs/combined_reports.pdf')
    merger.close()
    print("created merged PDF's file: outputs/combined_reports.pdf")
else:
    print("did not create any PDF files.")