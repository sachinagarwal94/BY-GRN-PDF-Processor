from flask import Flask, request, send_file, jsonify, render_template, flash, redirect, url_for, after_this_request
import pandas as pd
import re
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
import io
import os
import shutil
import threading
import time
import zipfile

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Needed for flash messages
UPLOAD_FOLDER = 'uploaded_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def update_invoice_data_in_pdf(pdf_filename, new_invoice_date, new_invoice_ref):
    pdf_document = fitz.open(pdf_filename)
    verdana_font = "Fonts/verdana.ttf"
    verdana_bold_font = "Fonts/verdana-bold.ttf"

    # Create a new PDF writer
    new_pdf_document = fitz.open()

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        new_page = new_pdf_document.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, pdf_document, page_num)

        text_instances_date = page.search_for("Invoice Date:")
        text_instances_ref = page.search_for("Invoice Ref #:")

        if text_instances_date:
            for inst in text_instances_date:
                x0, y0, x1, y1 = inst
                bg_color = (243 / 255, 243 / 255, 243 / 255)
                new_page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1), color=bg_color, fill=bg_color)
                adjusted_y0 = y1 - 1
                new_page.insert_text((x0, adjusted_y0), "Invoice Date:", fontsize=9, fontname="verdana-bold", fontfile=verdana_bold_font, color=(0, 0, 0))
                new_page.insert_text((x0 + 68, adjusted_y0), new_invoice_date, fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))
        else:
            new_page.insert_text((50, 50), f"Invoice Date: {new_invoice_date}", fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))

        if text_instances_ref:
            for inst in text_instances_ref:
                x0, y0, x1, y1 = inst
                bg_color = (243 / 255, 243 / 255, 243 / 255)
                new_page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1), color=bg_color, fill=bg_color)
                adjusted_y0 = y1 - 2
                new_page.insert_text((x0, adjusted_y0), "Invoice Ref #:", fontsize=9, fontname="verdana-bold", fontfile=verdana_bold_font, color=(0, 0, 0))
                new_page.insert_text((x0 + 73, adjusted_y0), new_invoice_ref, fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))
        else:
            new_page.insert_text((50, 70), f"Invoice Ref #: {new_invoice_ref}", fontsize=9, fontname="verdana", fontfile=verdana_font, color=(0, 0, 0))

    output_filename = pdf_filename.replace(".pdf", "_updated.pdf")
    new_pdf_document.save(output_filename)  # Save the new PDF
    new_pdf_document.close()
    pdf_document.close()

    print(f"Updated PDF saved as: {output_filename}")
    return output_filename
    
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        files = request.files.getlist('pdf_files')
        if not files or files[0].filename == '':
            flash('No files selected for uploading', 'danger')
            return redirect(url_for('home'))

        data = []

        def extract_invoice_data(pdf_file_path):
    invoice_date = None
    invoice_ref = None
    try:
        reader = PdfReader(pdf_file_path)
        for page in reader.pages:
            text = page.extract_text()
            date_match = re.search(r"Invoice Date:\s*(\d{2}/\d{2}/\d{4})", text)
            if date_match:
                invoice_date = date_match.group(1)
            if "Invoice Ref #:" in text:
                start_idx = text.find("Invoice Ref #:") + len("Invoice Ref #:")
                end_idx = text.find("\n", start_idx)
                invoice_ref = text[start_idx:end_idx].strip()
    except Exception as e:
        print(f"Error extracting data from {pdf_file_path}: {str(e)}")
    return invoice_date, invoice_ref
    
        for file in files:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            if os.path.isdir(file_path):
                flash(f"Error: '{file.filename}' is a directory, not a file.", 'danger')
                continue
            file.save(file_path)
            invoice_date, invoice_ref = extract_invoice_data(file_path)
            data.append({
                "Filename": file_path,
                "Invoice Date": invoice_date,
                "Invoice Ref": invoice_ref
            })

        if not data:
            flash('No valid files were uploaded.', 'danger')
            return redirect(url_for('home'))

        df = pd.DataFrame(data)
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        @after_this_request
        def cleanup(response):
            def delete_files():
                time.sleep(300)  # Wait for 1 minute before deleting files
                shutil.rmtree(UPLOAD_FOLDER)
                os.makedirs(UPLOAD_FOLDER)

            threading.Thread(target=delete_files).start()
            return response

        return send_file(output, download_name="GRN_Data.xlsx", as_attachment=True)
    except Exception as e:
        flash(f"Error processing files: {str(e)}", 'danger')
        return redirect(url_for('home'))

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    try:
        excel_file = request.files['excel_file']
        if not excel_file:
            flash('No Excel file selected for uploading', 'danger')
            return redirect(url_for('home'))

        df_corrected = pd.read_excel(excel_file)
        corrected_data = df_corrected.to_dict(orient="records")

        def sanitize_filename(filename):
            return re.sub(r'[<>:"/\\|?*]', '', filename)

        updated_files = []
        for item in corrected_data:
            pdf_file_path = item["Filename"]
            new_invoice_date = pd.to_datetime(item["Invoice Date"], dayfirst=True).strftime('%d/%m/%Y') if pd.notnull(item["Invoice Date"]) else None
            new_invoice_ref = item["Invoice Ref"]
            try:
                updated_file = update_invoice_data_in_pdf(pdf_file_path, new_invoice_date, new_invoice_ref)
                updated_files.append(updated_file)
            except Exception as e:
                print(f"Error updating {pdf_file_path}: {str(e)}")

        zip_path = os.path.join(UPLOAD_FOLDER, 'updated_pdfs.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in updated_files:
                zipf.write(file, os.path.basename(file))

        @after_this_request
        def cleanup(response):
            def delete_files():
                time.sleep(600)  # Wait for 1 minute before deleting files
                shutil.rmtree(UPLOAD_FOLDER)
                os.makedirs(UPLOAD_FOLDER)

            threading.Thread(target=delete_files).start()
            return response

        return send_file(zip_path, download_name='updated_pdfs.zip', as_attachment=True)
    except Exception as e:
        flash(f"Error uploading and processing Excel file: {str(e)}", 'danger')
        return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
