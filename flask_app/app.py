from flask import Flask, render_template, request, send_file, session
import pandas as pd
from data_converter import json_to_python, python_to_excel, file_path      
import os
import pyzipper

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for using sessions

# Route to download the Excel file
@app.route('/download')
def download_file():
    # Get the filename and password from the query parameters
    filename = request.args.get('filename')
    password = request.args.get('password')

    # Use the provided filename or fallback to a default name
    if not filename:
        filename = 'DataFrame.xlsx'
    else:
        filename += '.xlsx'
    
    # Check if the file exists
    if os.path.exists(file_path):
        if password:
            # Create a password-protected zip file
            zip_filename = filename.replace('.xlsx', '.zip')
            with pyzipper.AESZipFile(zip_filename, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                zf.setpassword(password.encode('utf-8'))
                zf.write(file_path, filename)
            return send_file(zip_filename, as_attachment=True)
        else:
            # Send the file as is, without password protection
            return send_file(file_path, as_attachment=True, download_name=filename)
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
    json_data = file.filename

    if file:
        # Check MIME type of the uploaded file
        mime_type = file.content_type

        if mime_type == 'application/json':
            python_data = json_to_python(json_data)
            df = python_to_excel(python_data)    
            table = df.to_html(classes='data', header="true", index=False)
            return render_template('index.html', table=table)
        
        else:
            error_message = 'Invalid file type. Please select a JSON file.'
            return render_template('index.html', error_message=error_message)

    else:
        error_message = 'No file was selected'
        return render_template('index.html', error_message=error_message)

if __name__ == '__main__':
    app.run(debug=True)
