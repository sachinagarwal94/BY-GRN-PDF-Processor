import subprocess
import os
from flask import Flask, request, send_file, jsonify, after_this_request, render_template_string
import pandas as pd
import re
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
import io
import shutil
import threading
import time
import zipfile

# Clone the GitHub repository containing the font files
subprocess.run(['git', 'clone', 'https://github.com/yourusername/yourrepository.git'])

app = Flask(__name__)
UPLOAD_FOLDER = 'uploaded_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Define the font paths
verdana_font_path = os.path.join('yourrepository', 'Fonts', 'verdana.ttf')
verdana_bold_font_path = os.path.join('yourrepository', 'Fonts', 'verdana_bold.ttf')

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
                footer {
                    margin-top: 20px;
                    font-size: 12px;
                    color: #777;
                    text-align: center;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h2>GRN PDF Processor</h2>
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
                <footer>
                    <p>Files will be deleted after 1 minute for privacy reasons.</p>
                    <p>Created with ❤️ by 
                        <a href="mailto:mumbai.sachin@gmail.com" style="text-decoration: none; color: inherit;">Sachin Agarwal</a>
                    </p>
                </footer>
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
            for page in pdf_document:
                page.insert_text((50, 50), f"Invoice Date: {new_invoice_date}",
                                 fontname="helv", fontsize=10)
                page.insert_text((50, 70), f"Invoice Ref #: {new_invoice_ref}",
                                 fontname="helv", fontsize=10)
            pdf_document.save(pdf_filename)

        for record in corrected_data:
            old_filename = record["Filename"]
            sanitized_filename = sanitize_filename(
                f"corrected_{old_filename}")
            new_file_path = os.path.join(UPLOAD_FOLDER, sanitized_filename)
            shutil.copy(old_filename, new_file_path)

            update_invoice_data_in_pdf(new_file_path, record["Invoice Date"],
                                       record["Invoice Ref"])

        @after_this_request
        def cleanup(response):
            files = os.listdir(UPLOAD_FOLDER)
            for file in files:
                file_path = os.path.join(UPLOAD_FOLDER, file)
                os.remove(file_path)
            return response

        return jsonify({"message": "Files processed successfully"}), 200

    except Exception as e:
        print(f"Error uploading excel file: {str(e)}")
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8080)
