from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from bidi.algorithm import get_display
import arabic_reshaper
import pdfkit
from jinja2 import Template
from PyPDF2 import PdfMerger

config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
#Load Excel Sheets 
filter = 16396

data_excel_file = 'DataToPDF/data.xlsx'
data_sheet_name = 'Sheet1'  # You can also use an integer index, e.g., 0
data_df = pd.read_excel(data_excel_file, sheet_name=data_sheet_name)
data_df = data_df[data_df['מספר תיק'] == filter]

#print(data_df.head(5))
#Extract Values:
account_name = data_df.iloc[0]['שם חשבון'] 
account_name = arabic_reshaper.reshape(account_name)
account_name = get_display(account_name)
problematic_debt = data_df['אחוז מהתיק'].sum()
#אחוז משווי תיק
data_df = data_df.rename(columns={"אחוז משווי תיק לפי שיערוך אחרון": "p_o_p_b_leval"})
data_df["p_o_p_b_leval"] = data_df["p_o_p_b_leval"].fillna(0)
p_o_p_value = data_df['p_o_p_b_leval'].sum()
print(p_o_p_value)
#שווי נייר
data_df = data_df.rename(columns={"שווי נייר": "p_value"})
data_df["p_value"] = data_df["p_value"].fillna(0)
p_value = data_df['p_value'].sum()
#print(p_value)
#חוב בעייתי
data_df = data_df.rename(columns={'סיווג פורום חוב': "debt_forum_type"})
data_dt_filter = data_df[data_df['debt_forum_type'].isin(['השגחה מיוחדת','במעקב מיוחד','מסופק','בפיגור'])]
p_o_p_filter_value = data_dt_filter['p_o_p_b_leval'].sum()
p_debt_value = data_dt_filter['p_value'].sum()
#אחוז מאפיק
data_df["p_o_afik"] = data_df["p_value"] / p_value
p_o_afik = data_df['p_o_afik'].sum()
#אחוז אשראי מוחרג
data_df = data_df.rename(columns={'קוד אפיק ותת אפיק': "sub_afik"})
data_dt_subafik_filter = data_df[data_df['sub_afik'].isin([405, 210, 240, 310,330,345,425])]
p_o_p_afik_filter_value = data_dt_subafik_filter['p_o_p_b_leval'].sum()
#print(p_o_afik)
data_df.to_excel('DataToPDF/data_df.xlsx', index=False)


# Load the template image
image = Image.open('templates/template1.png')
draw = ImageDraw.Draw(image)

# Example text settings (add as many as you need)
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
        'text': '10.00%',
        'position': (1020, 290),
        'font_path': 'arial.ttf',
        'font_size': 42,
        'color': (255, 255, 255)  
    }
]

# Add each text to the image
for item in texts:
    font = ImageFont.truetype(item['font_path'], item['font_size'])
    center_x, center_y = item['position'][0], item['position'][1]
    bbox = draw.textbbox((0, 0), item['text'], font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = center_x - text_width // 2
    y = center_y - text_height // 2
    draw.text((x, y), item['text'], font=font, fill=item['color'])
    #draw.text(item['position'], item['text'], font=font, fill=item['color'])

# Save the edited image
output_image = 'outputs/output.png'
output_pdf = 'outputs/output.pdf'
options = {
        'page-size': 'A4',
        'orientation': 'Landscape',
        'encoding': 'UTF-8',
        'margin-top': '5mm',
        'margin-bottom': '5mm',
        'margin-left': '5mm',
        'margin-right': '5mm'
    }
image.save(output_pdf)

print("Text added and image saved as 'output.png'")
