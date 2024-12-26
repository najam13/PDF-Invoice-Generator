import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("Invoice/*xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f'Invoice_nr.{invoice_nr}', ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f'date.{date}', ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Add a header
    column = df.columns
    column = [item.replace("-", " ").title() for item in column]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(column[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(column[1]), border=1)
    pdf.cell(w=40, h=8, txt=str(column[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(column[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(column[4]), border=1, ln=1)

    #Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="B", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum =df["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(""), border=1)
    pdf.cell(w=70, h=8, txt=str(""), border=1)
    pdf.cell(w=40, h=8, txt=str(""), border=1)
    pdf.cell(w=30, h=8, txt=str(""), border=1)
    pdf.cell(w=30, h=8, txt=str(str(total_sum)), border=1)

    #Add total sum sentence
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=f'The total price is total sum = {total_sum}', ln=1)

    #Add company name and logo
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=27, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")







