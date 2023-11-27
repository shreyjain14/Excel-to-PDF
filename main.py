from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    file_name, date = filename.split('-')

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No.{file_name}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=2)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]

    pdf.set_font(family="Times", size=8, style="B")
    pdf.set_fill_color(80, 80, 80)

    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    total_purchase_amount = 0

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_fill_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

        total_purchase_amount += row["total_price"]

    pdf.set_font(family="Times", size=8, style="B")
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_purchase_amount), border=1, ln=1)

    pdf.ln()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"The total price is {total_purchase_amount}.")

    pdf.output(f"PDFs/{filename}.pdf")
