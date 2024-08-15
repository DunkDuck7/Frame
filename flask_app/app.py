from flask import Flask, render_template, request, send_file, session
import pandas as pd
from data_converter import json_to_excel
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for using sessions

# Route to download the Excel file
@app.route('/download')
def download_file():
    # Retrieve the file name from the session
    file_name = session.get('file_name', 'DataFrame.xlsx')
    
    # Provide the Excel file as a download
    if os.path.exists(file_name):
        return send_file(file_name, as_attachment=True)
    else:
        return 'File not found', 404

# Route to render the main page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle the form submission and load the selected JSON file
@app.route('/load_data', methods=['POST'])
def load_data():
    file = request.files.get('json_file')
    
    if file:
        # Check MIME type of the uploaded file
        mime_type = file.content_type
        if mime_type == 'application/json':
            # CONVERT FILE TYPE FROM FileStorage to Str
            file = file.filename
            # Convert JSON to Excel
            df = json_to_excel(file)
            excel_path = 'DataFrame.xlsx'
            df.to_excel(excel_path, index=False)
            
            # Store the file name in the session
            session['file_name'] = excel_path
            
            table = df.to_html(classes='data', header="true", index=False)
            return render_template('index.html', table=table)
        else:
            return f'Invalid file type: {mime_type}. Please select a JSON file.', 400
    else:
        return 'No file was selected', 400

if __name__ == '__main__':
    app.run(debug=True)
