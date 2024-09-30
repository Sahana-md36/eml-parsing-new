import base64
import datetime
import email
import email.header
import os
import eml_parser
from flask import Flask, request, jsonify
from flask_cors import CORS
from extract_text_wordpdf import (
    extract_doc,
    process_pdf,
    process_pdf_upload,
    extract_text_from_txt,
    extract_text_from_image,
    extract_text_from_csv,
    extract_text_from_xlsx,
    extract_text_from_html,process_image_jpg
)
from extract_emailbody import read_email
from extractmsg import extract_text_from_msg
from extract_msg_body import read_email_content
from extract_text_from_doc import extract_text_from_doc

app = Flask(__name__)
CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    elif isinstance(obj, email.header.Header):
        raise Exception('object cannot be of type email.header.Header')
    elif isinstance(obj, bytes):
        return obj.decode('utf-8', errors='ignore')
    raise TypeError(f'Type "{str(type(obj))}" not serializable')

def parse_email(eml_file_name, output_folder_path='email_attachments'):
    def clear_output_folder(path):
        """Clears the output folder to remove previous attachments"""
        if os.path.exists(path) and os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    os.remove(os.path.join(root, file))
                for dir in dirs:
                    os.rmdir(os.path.join(root, dir))
        os.makedirs(path, exist_ok=True)

    def recursively_extract_attachments(eml_file_name, output_folder_path):
        ep = eml_parser.EmlParser(include_attachment_data=True)
        print(f'Parsing: {eml_file_name}')
        with open(eml_file_name, 'rb') as f:
            m = ep.decode_email_bytes(f.read())
        attachments = []

        # For regular attachments
        if 'attachment' in m:
            for a in m['attachment']:
                out_filepath = os.path.join(output_folder_path, a['filename'])
                print(f'\tWriting attachment: {out_filepath}')
                with open(out_filepath, 'wb') as a_out:
                    a_out.write(base64.b64decode(a['raw']))
                attachments.append({'filename': a['filename'], 'path': out_filepath})
            print('Regular attachments extracted')

        return attachments

    clear_output_folder(output_folder_path)

    attachments = recursively_extract_attachments(eml_file_name, output_folder_path)

    parsed_attachments = []
    for attachment in attachments:
        file_path = attachment['path']
        file_name = attachment['filename']
        filetype = os.path.splitext(file_name)[1][1:].lower()

        try:
            if file_name.endswith('.docx'):
                print(f"Extracting text from docx file: {file_name}")
                docx_text = extract_doc(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': docx_text if docx_text else 'Invalid attachment'})
            
            elif file_name.endswith('.doc'):
                print(f"Extracting text from doc file: {file_name}")
                doc_text = extract_text_from_doc(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': doc_text if doc_text else 'Invalid attachment'})
            
            elif file_name.endswith('.pdf'):
                print(f"Extracting text from pdf file: {file_name}")
                with open(file_path, 'rb') as pdf_file:
                    pdf_data = pdf_file.read()
                pdf_text = process_pdf_upload(pdf_data)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': pdf_text if pdf_text else 'Invalid attachment'})

            elif file_name.endswith('.txt'):
                print(f"Extracting text from txt file: {file_name}")
                txt_text = extract_text_from_txt(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': txt_text if txt_text else 'Invalid attachment'})

            elif file_name.endswith('.csv'):
                print(f"Extracting text from csv file: {file_name}")
                csv_text = extract_text_from_csv(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': csv_text if csv_text else 'Invalid attachment'})

            elif file_name.endswith('.xlsx'):
                print(f"Extracting text from xlsx file: {file_name}")
                xlsx_text = extract_text_from_xlsx(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': xlsx_text if xlsx_text else 'Invalid attachment'})

            elif file_name.endswith('.html'):
                print(f"Extracting text from html file: {file_name}")
                html_text = extract_text_from_html(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': html_text if html_text else 'Invalid attachment'})

            elif file_name.endswith('.jpg') or file_name.endswith('.jpeg') or file_name.endswith('.png'):
                print(f"Extracting text from image file: {file_name}")
                # with open(file_path, 'rb') as image_file:
                #     image_data = image_file.read()
                image_text = process_image_jpg(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': image_text if image_text else 'Poor quality image or invalid attachment'})

            elif file_name.startswith('part-000'):
                print(f"Extracting text from MSG file: {file_name}")
                msg_text = read_email_content(file_path)
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': msg_text if msg_text else 'Invalid attachment'})

            else:
                print(f"Unsupported file format: {file_name}")
                parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': 'Invalid attachment'})

        except Exception as e:
            print(f"Error parsing {file_name}: {e}")
            parsed_attachments.append({'filename': file_name, 'filetype': filetype, 'content': 'Invalid attachment'})

    email_details = read_email(eml_file_name)

    if 'Body' not in email_details or not email_details['Body'].strip():
        email_details['Body'] = 'Unavailable'

    email_details['Attachments'] = parsed_attachments if parsed_attachments else []

    print(f"Read email details: {email_details}")
    print(f"Parsed document text: {parsed_attachments}")

    return email_details

@app.route("/")
def home():
    return "EML file parsing API"

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file found"})
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "File not uploaded"})

    if file and file.filename.endswith('.eml'):
        eml_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(eml_file_name)
        result = parse_email(eml_file_name)
        return jsonify(result)

    elif file and file.filename.endswith('.msg'):
        msg_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(msg_file_name)
        result = extract_text_from_msg(msg_file_name)
        return jsonify({"result": result})

    elif file and file.filename.endswith('.pdf'):
        pdf_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(pdf_file_name)
    
    # Open the PDF file in binary mode and read its content
        with open(pdf_file_name, 'rb') as pdf_file:
            pdf_data = pdf_file.read()
    
        result = process_pdf_upload(pdf_data)  # Pass the binary data to process_pdf
        return jsonify({"result": result})

    
    elif file and file.filename.endswith('.doc'):
        doc_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(doc_file_name)
        result = extract_text_from_doc(doc_file_name)
        return jsonify({"result": result})

    else:
        return jsonify({"error": "Unsupported file type"})

if __name__ == '__main__':
    app.run(debug=True)
