from flask import Flask, request, render_template, send_file
from docxtpl import DocxTemplate
import pandas as pd
import os
import io
from zipfile import ZipFile

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge():
    doc_file = request.files['template']
    excel_file = request.files['data']

    if not doc_file or not excel_file:
        return "Please upload both files.", 400

    df = pd.read_excel(excel_file)
    data_list = df.to_dict(orient='records')

    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, 'a') as zip_file:
        for data in data_list:
            doc = DocxTemplate(doc_file)
            doc.render(data)
            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)
            filename = f"{data.get('name')}_letter.docx"
            zip_file.writestr(filename, output_stream.read())

    zip_buffer.seek(0)
    return send_file(zip_buffer, mimetype='application/zip', download_name='merged_letters.zip', as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # default to 5000 if PORT not set
    app.run(host='0.0.0.0', port=port)