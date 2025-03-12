import base64
import datetime
import email
import email.header
import os
import re
import eml_parser
from flask import Flask, request, jsonify
from flask_cors import CORS
from extract_text_wordpdf import (
    extract_doc,
    process_pdf_upload,
    extract_text_from_txt,
    extract_text_from_csv,
    extract_text_from_xlsx,
    extract_text_from_html, process_image_jpg
)
from extract_emailbody import read_email
from extractmsg import extract_text_from_msg
from extract_msg_body import read_email_content
from extract_text_from_doc import extract_text_from_doc
from bs4 import BeautifulSoup
import requests

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

def extract_links_from_html(body):
    """Extracts hyperlinks from anchor elements in the email body."""
    soup = BeautifulSoup(body, 'html.parser')
    links = []

    # Find all anchor tags that contain a hyperlink
    for button in soup.find_all('a', href=True):
        links.append(clean_url(button['href']))

    return links

def extract_links_from_text(body):
    """Extracts all valid URLs from plain text using regular expressions."""
    url_pattern = re.compile(r'https?://[^\s]+')
    return url_pattern.findall(body)

def clean_url(url):
    """Cleans a URL by removing trailing unwanted characters such as > or ] if they exist."""
    return url.rstrip('>').rstrip(']')

def clean_filetype(filetype):
    """Cleans the file type by removing any unwanted characters such as '>' or '<'."""
    return filetype.split('?')[0].split('#')[0].strip('.').lower()  # Clean and extract file extension

def process_external_link(url):
    """Fetches the URL content and extracts text based on document type."""
    # Send GET request to the URL
    response = requests.get(url)
    content_type = response.headers.get('Content-Type')

    # Based on the content type, handle different document formats
    if 'pdf' in content_type:
        pdf_text = process_pdf_upload(response.content)  # Process PDF content
        return pdf_text
    elif 'html' in content_type:
        html_text = extract_text_from_html(response.content)  # Process HTML content
        return html_text
    elif 'doc' in content_type or 'docx' in content_type:
        docx_text = extract_doc(response.content)  # Process DOCX content
        return docx_text
    elif 'csv' in content_type:
        csv_text = extract_text_from_csv(response.content)  # Process CSV content
        return csv_text
    elif 'txt' in content_type:
        return response.text  # Text content directly
    else:
        return 'Unsupported document format'
    
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

    # Extract the email body
    email_details = read_email(eml_file_name)

    if 'Body' not in email_details or not email_details['Body'].strip():
        email_details['Body'] = 'Unavailable'

    # Extract links from email body
    links = extract_links_from_text(email_details['Body'])
    extracted_links_content = []
    hyperlink_counter = 1  # Initialize counter for hyperlink filenames

    # For each link found, fetch content and store it
    for link in links:
        print(f"Processing link: {link}")
        link_content = process_external_link(link)
        
        if link_content and link_content != "Unsupported document format":
            # Extract the file type from the link
            filetype = link.split('.')[-1].lower()  # Default filetype extraction using extension
            
            # Clean the filetype to avoid unwanted characters
            filetype = clean_filetype(filetype)
            filetype = filetype.replace('>', '')
            
            extracted_links_content.append({
                "filename": f"hyperlink-{hyperlink_counter}",
                "filetype": filetype,
                "content": link_content
            })
            hyperlink_counter += 1  # Increment hyperlink counter

    # If ButtonLinksContent has any data, move it to Attachments
    if extracted_links_content:
        email_details['Attachments'] = extracted_links_content
    else:
        # No links found, add regular attachments
        email_details['Attachments'] = parsed_attachments if parsed_attachments else []

    # Final result to return
    print(f"Read email details: {email_details}")
    print(f"Parsed document text: {parsed_attachments}")

    return email_details

def clean_text(text):
    """Removes excessive newlines and formats the text properly."""
    if isinstance(text, dict):
        return {key: clean_text(value) for key, value in text.items()}
    elif isinstance(text, list):
        return [clean_text(item) for item in text]
    elif isinstance(text, str):
        text = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        text = re.sub(r'\s+', ' ', text).strip()
        return text
    return text



@app.route("/")
def home():
    return "EML file parsing API"

# @app.route('/upload', methods=['POST'])
# def upload_file():
#     if 'file' not in request.files:
#         return jsonify({"error": "No file found"})
#     file = request.files['file']
#     if file.filename == '':
#         return jsonify({"error": "File not uploaded"})

#     if file and file.filename.endswith('.eml'):
#         eml_file_name = os.path.join('uploaded', file.filename)
#         os.makedirs('uploaded', exist_ok=True)
#         file.save(eml_file_name)
#         result = parse_email(eml_file_name)
#         return jsonify(result)


#     elif file and file.filename.endswith('.msg'):
#         msg_file_name = os.path.join('uploaded', file.filename)
#         os.makedirs('uploaded', exist_ok=True)
#         file.save(msg_file_name)
#         result = extract_text_from_msg(msg_file_name)
#         return jsonify({"result": result})


#     elif file and file.filename.endswith('.pdf'):
#         pdf_file_name = os.path.join('uploaded', file.filename)
#         os.makedirs('uploaded', exist_ok=True)
#         file.save(pdf_file_name)
    
#     # Open the PDF file in binary mode and read its content
#         with open(pdf_file_name, 'rb') as pdf_file:
#             pdf_data = pdf_file.read()
    
#         result = process_pdf_upload(pdf_data)  # Pass the binary data to process_pdf
#         return jsonify({"result": result})


    
#     elif file and file.filename.endswith('.doc'):
#         doc_file_name = os.path.join('uploaded', file.filename)
#         os.makedirs('uploaded', exist_ok=True)
#         file.save(doc_file_name)
#         result = extract_text_from_doc(doc_file_name)
#         return jsonify({"result": result})


#     else:
#         return jsonify({"error": "Unsupported file type"})

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
        cleaned_result = clean_text(result)
        return jsonify({"result": cleaned_result})

    elif file and file.filename.endswith('.msg'):
        msg_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(msg_file_name)
        result = extract_text_from_msg(msg_file_name)
        cleaned_result = clean_text(result)
        return jsonify({"result": cleaned_result})

    elif file and file.filename.endswith('.pdf'):
        pdf_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(pdf_file_name)

        with open(pdf_file_name, 'rb') as pdf_file:
            pdf_data = pdf_file.read()

        result = process_pdf_upload(pdf_data)
        cleaned_result = clean_text(result)
        return jsonify({"result": cleaned_result})

    elif file and file.filename.endswith('.doc'):
        doc_file_name = os.path.join('uploaded', file.filename)
        os.makedirs('uploaded', exist_ok=True)
        file.save(doc_file_name)
        result = extract_text_from_doc(doc_file_name)
        cleaned_result = clean_text(result)
        return jsonify({"result": cleaned_result})

    else:
        return jsonify({"error": "Unsupported file type"})


if __name__ == '__main__':
    app.run(debug=True)
