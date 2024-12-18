from flask import Flask, request, send_file, jsonify, after_this_request, render_template_string
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
UPLOAD_FOLDER = 'uploaded_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return render_template_string('''
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>GRN PDF Processor</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f4f4f9;
                    margin: 0;
                    padding: 0;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                }
                .container {
                    background-color: #fff;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    max-width: 500px;
                    width: 100%;
                }
                h1 {
                    font-size: 24px;
                    margin-bottom: 20px;
                    color: #333;
                }
                form {
                    display: flex;
                    flex-direction: column;
                }
                input[type="file"] {
                    margin-bottom: 10px;
                }
                input[type="submit"] {
                    background-color: #007bff;
                    color: #fff;
                    border: none;
                    padding: 10px;
                    border-radius: 4px;
                    cursor: pointer;
                }
                input[type="submit"]:hover {
                    background-color: #0056b3;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>GRN PDF Processor</h1>
                <form action="/process" method="post" enctype="multipart/form-data">
                    <label for="pdf_files">Upload GRN PDF Files:</label>
                    <input type="file" name="pdf_files" multiple>
                    <input type="submit" value="Upload and Process">
                </form>
                <h2>Upload Edited Excel File</h2>
                <form action="/upload_excel" method="post" enctype="multipart/form-data">
                    <label for="excel_file">Upload Edited Excel File:</label>
                    <input type="file" name="excel_file">
                    <input type="submit" value="Upload and Process Excel">
                </form>
            </div>
        </body>
        </html>
    ''')


@app.route('/process', methods=['POST'])
def process():
    try:
        files = request.files.getlist('pdf_files')
        data = []

        def extract_invoice_data(pdf_file):
            invoice_date = None
            invoice_ref = None
            try:
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    text = page.extract_text()
                    date_match = re.search(
                        r"Invoice Date:\s*(\d{2}.\d{2}.\d{4})", text)
                    if date_match:
                        invoice_date = date_match.group(1)
                    if "Invoice Ref #:" in text:
                        start_idx = text.find("Invoice Ref #:") + len(
                            "Invoice Ref #:")
                        end_idx = text.find("\n", start_idx)
                        invoice_ref = text[start_idx:end_idx].strip()
            except Exception as e:
                print(
                    f"Error extracting data from {pdf_file.filename}: {str(e)}"
                )
            return invoice_date, invoice_ref

        for file in files:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            invoice_date, invoice_ref = extract_invoice_data(file_path)
            data.append({
                "Filename": file_path,
                "Invoice Date": invoice_date,
                "Invoice Ref": invoice_ref
            })

        df = pd.DataFrame(data)
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        return send_file(output,
                         download_name="GRN_Data.xlsx",
                         as_attachment=True)
    except Exception as e:
        print(f"Error processing files: {str(e)}")
        return jsonify({"error": str(e)}), 500


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    try:
        excel_file = request.files['excel_file']
        df_corrected = pd.read_excel(excel_file)
        corrected_data = df_corrected.to_dict(orient="records")

        def sanitize_filename(filename):
            return re.sub(r'[<>:"/\\|?*]', '', filename)

        def update_invoice_data_in_pdf(pdf_filename, new_invoice_date,
                                       new_invoice_ref):
            pdf_document = fitz.open(pdf_filename)
            verdana_font = "Fonts/verdana.ttf"
            verdana_bold_font = "Fonts/verdana-bold.ttf"

            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text_instances_date = page.search_for("Invoice Date:")
                text_instances_ref = page.search_for("Invoice Ref #:")

                if text_instances_date:
                    for inst in text_instances_date:
                        x0, y0, x1, y1 = inst
                        bg_color = (243 / 255, 243 / 255, 243 / 255)
                        page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1),
                                       color=bg_color,
                                       fill=bg_color)
                        adjusted_y0 = y1 - 1
                        page.insert_text((x0, adjusted_y0),
                                         "Invoice Date:",
                                         fontsize=9,
                                         fontname="verdana-bold",
                                         fontfile=verdana_bold_font,
                                         color=(0, 0, 0))
                        page.insert_text((x0 + 68, adjusted_y0),
                                         new_invoice_date,
                                         fontsize=9,
                                         fontname="verdana",
                                         fontfile=verdana_font,
                                         color=(0, 0, 0))
                else:
                    page.insert_text((50, 50),
                                     f"Invoice Date: {new_invoice_date}",
                                     fontsize=9,
                                     fontname="verdana",
                                     fontfile=verdana_font,
                                     color=(0, 0, 0))

                if text_instances_ref:
                    for inst in text_instances_ref:
                        x0, y0, x1, y1 = inst
                        bg_color = (243 / 255, 243 / 255, 243 / 255)
                        page.draw_rect(fitz.Rect(x0, y0, x1 + 100, y1),
                                       color=bg_color,
                                       fill=bg_color)
                        adjusted_y0 = y1 - 2
                        page.insert_text((x0, adjusted_y0),
                                         "Invoice Ref #:",
                                         fontsize=9,
                                         fontname="verdana-bold",
                                         fontfile=verdana_bold_font,
                                         color=(0, 0, 0))
                        page.insert_text((x0 + 73, adjusted_y0),
                                         new_invoice_ref,
                                         fontsize=9,
                                         fontname="verdana",
                                         fontfile=verdana_font,
                                         color=(0, 0, 0))
                else:
                    page.insert_text((50, 70),
                                     f"Invoice Ref #: {new_invoice_ref}",
                                     fontsize=9,
                                     fontname="verdana",
                                     fontfile=verdana_font,
                                     color=(0, 0, 0))

            output_filename = pdf_filename.replace(".pdf", "_updated.pdf")
            pdf_document.save(output_filename)
            pdf_document.close()

            print(f"Updated PDF saved as: {output_filename}")
            return output_filename

        updated_files = []
        for item in corrected_data:
            pdf_file_path = item["Filename"]
            new_invoice_date = pd.to_datetime(
                item["Invoice Date"],
                dayfirst=True).strftime('%d/%m/%Y') if pd.notnull(
                    item["Invoice Date"]) else None
            new_invoice_ref = item["Invoice Ref"]
            try:
                updated_file = update_invoice_data_in_pdf(
                    pdf_file_path, new_invoice_date, new_invoice_ref)
                updated_files.append(updated_file)
            except Exception as e:
                print(f"Error updating {pdf_file_path}: {str(e)}")

        @after_this_request
        def cleanup(response):

            def delete_files():
                time.sleep(60)  # Wait for 10 minutes before deleting files
                shutil.rmtree(UPLOAD_FOLDER)
                os.makedirs(UPLOAD_FOLDER)

            threading.Thread(target=delete_files).start()
            return response

        if updated_files:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                for file_path in updated_files:
                    zip_file.write(file_path, os.path.basename(file_path))

            zip_buffer.seek(0)
            return send_file(zip_buffer,
                             download_name='updated_files.zip',
                             as_attachment=True)

        else:
            return jsonify({"message": "No PDF files were updated."})
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
import logging

# Configure logging
logging.basicConfig(filename='error.log', level=logging.ERROR)

# Your Flask app setup
from flask import Flask, request

app = Flask(__name__)

# Your routes and handlers
@app.route('/')
def home():
    return "Welcome to the app!"

@app.route('/process', methods=['POST'])
def process_endpoint():
    try:
        # Logic to process request
        # Example: process_file(request.form['file_path'])
        return "File processed successfully", 200
    except Exception as e:
        logging.error(f"Error in /process endpoint: {e}")
        return "An error occurred while processing the file", 500

@app.route('/upload_excel', methods=['POST'])
def upload_excel_endpoint():
    try:
        # Logic to upload and handle Excel files
        # Example: process_file(request.files['excel_file'])
        return "File uploaded and processed successfully", 200
    except Exception as e:
        logging.error(f"Error in /upload_excel endpoint: {e}")
        return "An error occurred while uploading the file", 500

if __name__ == '__main__':
    app.run(debug=True)
