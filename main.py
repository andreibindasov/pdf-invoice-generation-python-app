import pandas as pd
from fpdf import FPDF
import glob

filepaths = glob.glob("invoices/*.xlsx")

for fp in filepaths:
    df = pd.read_excel(fp, sheet_name="Sheet 1")
    print(df)
