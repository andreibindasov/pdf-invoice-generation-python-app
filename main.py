import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for fp in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(fp).stem
    invoice_nr, invoice_date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=12, style="IB")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    df = pd.read_excel(fp, sheet_name="Sheet 1")

    # Add a header
    columns = list(df.columns)
    columns = [c.replace("_", " ").title() for c in columns]

    pdf.set_font(family="Arial", size=9, style="IB")
    pdf.set_text_color(11, 9, 222)
    pdf.cell(w=30, h=9, txt=columns[0], border=1, align="L")
    pdf.cell(w=70, h=9, txt=columns[1], border=1, align="L")
    pdf.cell(w=30, h=9, txt=columns[2], border=1, align="C")
    pdf.cell(w=30, h=9, txt=columns[3], border=1, align="C")
    pdf.cell(w=30, h=9, txt=columns[4], border=1, ln=1, align="C")

    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=12, style="")
        pdf.set_text_color(81, 81, 81)
        pdf.cell(w=30, h=9, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=9, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=9, txt=str(row["amount_purchased"]), border=1, align="C")
        pdf.cell(w=30, h=9, txt=str(row["price_per_unit"]), border=1, align="C")
        pdf.cell(w=30, h=9, txt=str(row["total_price"]), border=1, ln=1, align="C")

    pdf.output(f"PDFs/{invoice_nr}.pdf")
