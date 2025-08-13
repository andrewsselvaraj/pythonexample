import pdfplumber
import pandas as pd
import os
from flask import Flask, request, render_template, send_file

app = Flask(__name__)

# Function to extract text from PDF and save it as Excel
def pdf_to_excel(pdf_path, excel_path):
    data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                data.extend(table)  # Append extracted table data
    
    if data:
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False, header=False)
        return True
    return False

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part"
        file = request.files['file']
        if file.filename == '':
            return "No selected file"
        
        pdf_path = os.path.join("uploads", file.filename)
        excel_path = pdf_path.replace(".pdf", ".xlsx")
        
        file.save(pdf_path)
        
        if pdf_to_excel(pdf_path, excel_path):
            return send_file(excel_path, as_attachment=True)
        else:
            return "No table data found in the PDF."
    
    return '''
    <!doctype html>
    <title>Upload PDF</title>
    <h1>Upload PDF to Convert to Excel</h1>
    <form method=post enctype=multipart/form-data>
        <input type=file name=file>
        <input type=submit value=Upload>
    </form>
    '''

if __name__ == '__main__':
    os.makedirs("uploads", exist_ok=True)
    app.run(debug=True)
