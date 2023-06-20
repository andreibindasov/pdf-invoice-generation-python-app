import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for fp in filepaths:
    df = pd.read_excel(fp, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(fp).stem
    invoice_nr = filename.split("-")[0]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr.{invoice_nr}")

    pdf.output(f"PDFs/{invoice_nr}.pdf")
