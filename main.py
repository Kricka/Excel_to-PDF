from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path



filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format='A4')
    pdf.add_page()

    pdf.set_font(family="Times",style="B",size=16)

    filename=Path(filepath).stem
    invoice_nr=filename.split("-")[0]
    invoice_date = filename.split("-")[1]

    pdf.cell(w=50,h=8,txt=f"Invoice nr. {invoice_nr} ",ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date} ",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns=df.columns
    columns=[item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times",style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"{columns[0]}", border=1)
    pdf.cell(w=50, h=8, txt=f"{columns[1]}", border=1)
    pdf.cell(w=35, h=8, txt=f"{columns[2]}", border=1)
    pdf.cell(w=40, h=8, txt=f"{columns[3]}", border=1)
    pdf.cell(w=40, h=8, txt=f"{columns[4]}", border=1, ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30,h=8, txt=f"{row['product_id']}",border=1)
        pdf.cell(w=50, h=8, txt=f"{row['product_name']}",border=1)
        pdf.cell(w=35, h=8, txt=f"{row['amount_purchased']}",border=1)
        pdf.cell(w=40, h=8, txt=f"{row['price_per_unit']}",border=1)
        pdf.cell(w=40, h=8, txt=f"{row['total_price']}",border=1,ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt=f"{df['total_price'].sum()}", border=1, ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30,h=8,txt=f"The total price is {df['total_price'].sum()}",ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30,h=8,txt=f"Konak Zlatibor ",ln=1)
    pdf.image("logo.png")



    pdf.output(f"PDFs/{filename}.pdf")













