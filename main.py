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

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Arial", size=13, style="B")
    pdf.set_text_color(81, 81, 81)
    pdf.cell(w=30, h=9, txt="", border=1)
    pdf.cell(w=70, h=9, txt="", border=1)
    pdf.cell(w=30, h=9, txt="", border=1, align="C")
    pdf.cell(w=30, h=9, txt="", border=1, align="C")
    pdf.cell(w=30, h=9, txt=str(total_sum), border=1, ln=1, align="C")

    # Add total footnote
    pdf.set_font(family="Times", size=14, style="I")
    pdf.set_text_color(11, 9, 222)
    pdf.cell(w=0, h=9, txt=f"The total amount due ${total_sum}", ln=1)
    pdf.cell(w=45, h=9, txt="by :andrei:bindasov")
    pdf.image("stamp.png", w=15)

    # Multi-cell example
    pdf.set_font(family="Courier", size=7, style="I")
    pdf.set_text_color(190, 190, 253)
    content = """
       11 And I will establish my covenant with you; neither shall all flesh be cut off any more by the waters of a flood; 
       neither shall there any more be a flood to destroy the earth.
       12 And God said, This is the token of the covenant which I make between me and you and every living creature that 
       is with you, for perpetual generations:
       13 I do set my bow in the cloud, and it shall be for a token of a covenant between me and the earth.
       14 And it shall come to pass, when I bring a cloud over the earth, that the bow shall be seen in the cloud:
       15 And I will remember my covenant, which is between me and you and every living creature of all flesh; and the 
       waters shall no more become a flood to destroy all flesh.
       """

    pdf.multi_cell(w=0, h=6, txt=content)

    pdf.output(f"PDFs/{invoice_nr}.pdf")
