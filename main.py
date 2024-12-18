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
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    background-color: #e9ecef;
                    margin: 0;
                    padding: 0;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    color: #333;
                }
                .container {
                    background-color: #ffffff;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
                    max-width: 600px;
                    width: 100%;
                    text-align: center;
                }
                h1, h2 {
                    font-weight: 600;
                    margin-bottom: 15px;
                }
                form {
                    display: flex;
                    flex-direction: column;
                    gap: 15px;
                }
                input[type="file"] {
                    padding: 10px;
                    border: 1px solid #ccc;
                    border-radius: 5px;
                    font-size: 14px;
                    color: #555;
                }
                input[type="submit"] {
                    background-color: #28a745;
                    color: #fff;
                    border: none;
                    padding: 12px;
                    border-radius: 5px;
                    cursor: pointer;
                    font-size: 16px;
                    opacity: 0.6;
                    pointer-events: none;
                }
                input[type="submit"]:hover {
                    background-color: #218838;
                }
                footer {
                    margin-top: 20px;
                    font-size: 14px;
                    color: #6c757d;
                    text-align: center;
                }
                a {
                    color: #007bff;
                    text-decoration: none;
                }
                a:hover {
                    text-decoration: underline;
                }
                .error-message {
                    color: red;
                    font-size: 14px;
                    display: none;
                    background-color: #ffc107;
                    padding: 10px;
                    border-radius: 5px;
                    text-align: center;
                    font-weight: bold;
                    animation: shake 0.5s ease-in-out infinite;
                }
                @keyframes shake {
                    0% { transform: translateX(0); }
                    20% { transform: translateX(-10px); }
                    40% { transform: translateX(10px); }
                    60% { transform: translateX(-10px); }
                    80% { transform: translateX(10px); }
                    100% { transform: translateX(0); }
                }
            </style>
            <script>
                function validateFiles() {
                    const fileInput = document.getElementById('pdf_files');
                    const uploadButton = document.getElementById('submitBtn');
                    const uploadButton2 = document.getElementById('submitBtn2');
                    const errorMessage = document.getElementById('error-message');

                    if (fileInput.files.length === 0) {
                        errorMessage.style.display = 'block';
                        uploadButton.style.opacity = 0.6;
                        uploadButton.style.pointerEvents = 'none';
                        uploadButton2.style.opacity = 0.6;
                        uploadButton2.style.pointerEvents = 'none';
                    } else {
                        errorMessage.style.display = 'none';
                        uploadButton.style.opacity = 1;
                        uploadButton.style.pointerEvents = 'auto';
                        uploadButton2.style.opacity = 1;
                        uploadButton2.style.pointerEvents = 'auto';
                    }
                }
            </script>
        </head>
        <body>
            <div class="container">
                <h2>Welcome to the GRN PDF Processor</h2>
                <h2>Upload GRN .PDF Files</h2>
                <form action="/process" method="post" enctype="multipart/form-data">
                    <input type="file" id="pdf_files" name="pdf_files" multiple onchange="validateFiles()">
                    <input type="submit" id="submitBtn" value="Upload and Process">
                    <div id="error-message" class="error-message">Please select at least one file to upload! 🫣</div>
                </form>
                <h2>Upload Edited Excel File</h2>
                <form action="/upload_excel" method="post" enctype="multipart/form-data">
                    <input type="file" name="excel_file">
                    <input type="submit" id="submitBtn2" value="Upload and Process Excel">
                </form>
                <footer>
                    <p>Files will be deleted after 1 minute for privacy reasons.</p>
                    <p>Created with ❤️ by 
                        <a href="mailto:mumbai.sachin@gmail.com">Sachin Agarwal</a>
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
            verdana_font = "Fonts/verdana.ttf"
            verdana_bold_font = "Fonts/verdana_bold.ttf"
            for page in pdf_document:
                page.insert_text((50, 50), f"Invoice Date: {new_invoice_date}",
                                 fontname=verdana_font, fontsize=10)
                page.insert_text((50, 70), f"Invoice Ref #: {new_invoice_ref}",
                                 fontname=verdana_font, fontsize=10)
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
