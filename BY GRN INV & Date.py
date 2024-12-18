from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import re
import fitz  # PyMuPDF
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, asksaveasfilename

# Step 1: Upload all GRN PDF files
print("Select all your GRN PDF files:")
Tk().withdraw()  # Hide the root window
pdf_files = askopenfilenames(filetypes=[("PDF files", "*.pdf")])

# Step 2: Extract Invoice Dates and Reference Numbers from the PDFs
data = []

def extract_invoice_data(pdf_file):
    invoice_date = None
    invoice_ref = None
    try:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            text = page.extract_text()
            date_match = re.search(r"Invoice Date:\s*(\d{2}.\d{2}.\d{4})", text)
            if date_match:
                invoice_date = date_match.group(1)
            if "Invoice Ref #:" in text:
                start_idx = text.find("Invoice Ref #:") + len("Invoice Ref #:")
                end_idx = text.find("\n", start_idx)
                invoice_ref = text[start_idx:end_idx].strip()
    except Exception as e:
        print(f"Error extracting data from {pdf_file}: {str(e)}")
    return invoice_date, invoice_ref

# Collect Invoice Dates and Reference Numbers
for pdf_file in pdf_files:
    invoice_date, invoice_ref = extract_invoice_data(pdf_file)
    data.append({
        "Filename": pdf_file,
        "Invoice Date": invoice_date,
        "Invoice Ref": invoice_ref
    })

# Step 3: Export extracted data to Excel
df = pd.DataFrame(data)
excel_filename = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
df.to_excel(excel_filename, index=False)
print(f"Extracted data saved to {excel_filename}")

# Step 4: Prompt user to edit the Excel file and load the corrected data
input("Please edit the Excel file with correct data if needed, then save and close it. Press Enter to continue...")

# Load the corrected data from Excel
df_corrected = pd.read_excel(excel_filename)
corrected_data = df_corrected.to_dict(orient="records")

# Function to sanitize filenames to match Windows format (remove forbidden characters)
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '', filename)

# Step 5: Function to clear 'Invoice Date:' and 'Invoice Ref #' fields and then add the correct data using PyMuPDF
def update_invoice_data_in_pdf(pdf_filename, new_invoice_date, new_invoice_ref):
    # Open the original PDF with PyMuPDF
    pdf_document = fitz.open(pdf_filename)

    # Load the Verdana font files
    verdana_font = "C:/Users/sachin3.agarwal/Downloads/verdana-font-family/verdana.ttf"
    verdana_bold_font = "C:/Users/sachin3.agarwal/Downloads/verdana-font-family/verdana-bold.ttf"

    # Iterate through each page to find and clear "Invoice Date:" and "Invoice Ref #" fields
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text_instances_date = page.search_for("Invoice Date:")
        text_instances_ref = page.search_for("Invoice Ref #:")

        if text_instances_date:
            for inst in text_instances_date:
                x0, y0, x1, y1 = inst
                bg_color = (243/255, 243/255, 243/255)
                page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1), color=bg_color, fill=bg_color)
                adjusted_y0 = y1 - 1
                page.insert_text((x0, adjusted_y0), "Invoice Date:", fontsize=9, fontname="verdana-bold", fontfile=verdana_bold_font, color=(0, 0, 0))
                page.insert_text((x0 + 68, adjusted_y0), new_invoice_date, fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))
        else:
            page.insert_text((50, 50), f"Invoice Date: {new_invoice_date}", fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))

        if text_instances_ref:
            for inst in text_instances_ref:
                x0, y0, x1, y1 = inst
                bg_color = (243/255, 243/255, 243/255)
                page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1), color=bg_color, fill=bg_color)
                adjusted_y0 = y1 - 2
                page.insert_text((x0, adjusted_y0), "Invoice Ref #:", fontsize=9, fontname="verdana-bold", fontfile=verdana_bold_font, color=(0, 0, 0))
                page.insert_text((x0 + 73, adjusted_y0), new_invoice_ref, fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))
        else:
            page.insert_text((50, 70), f"Invoice Ref #: {new_invoice_ref}", fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))

    output_filename = pdf_filename.replace(".pdf", "_updated.pdf")
    pdf_document.save(output_filename)
    pdf_document.close()

    print(f"Updated PDF saved as: {output_filename}")
    return output_filename

# Step 6: Process each uploaded PDF file using the corrected data
updated_files = []
for item in corrected_data:
    pdf_file = item["Filename"]
    new_invoice_date = pd.to_datetime(item["Invoice Date"], dayfirst=True).strftime('%d/%m/%Y') if pd.notnull(item["Invoice Date"]) else None
    new_invoice_ref = item["Invoice Ref"]
    try:
        updated_file = update_invoice_data_in_pdf(pdf_file, new_invoice_date, new_invoice_ref)
        updated_files.append(updated_file)
    except Exception as e:
        print(f"Error updating {pdf_file}: {str(e)}")

# Step 7: Provide Download Links for Updated PDFs
if updated_files:
    print("Updated PDF files:")
    for updated_pdf in updated_files:
        print(updated_pdf)
else:
    print("No PDF files were updated.")