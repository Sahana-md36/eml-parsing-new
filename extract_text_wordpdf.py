import base64
import os
import time
import extract_msg
import pandas as pd
from io import BytesIO
from bs4 import BeautifulSoup
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from msrest.authentication import CognitiveServicesCredentials
from dotenv import load_dotenv
import fitz
import pypandoc
import json
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
load_dotenv()

subscription_key = os.getenv('subscription_key')
endpoint = os.getenv('endpoint')
computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))

form_recognizer_key = os.getenv('AZURE_FORM_RECOGNIZER_KEY')
form_recognizer_endpoint = os.getenv('AZURE_FORM_RECOGNIZER_ENDPOINT')
form_recognizer_client = DocumentAnalysisClient(form_recognizer_endpoint, AzureKeyCredential(form_recognizer_key))

def extract_doc(docx_data):
    """Extract text from DOCX file content."""
    with open("temp.docx", "wb") as f:
        f.write(docx_data)
    output = pypandoc.convert_file("temp.docx", 'rst')
    os.remove("temp.docx")
    return output

def extract_text_from_txt(txt_data):
    """Extract text from a TXT file."""
    return txt_data.decode('utf-8')

def extract_text_from_csv(csv_data):
    """Extract text from a CSV file."""
    try:
        df = pd.read_csv(BytesIO(csv_data))
        return df.to_string(index=False)
    except Exception as e:
        print(f"Error extracting text from CSV: {e}")
        return ""

def extract_text_from_xlsx(xlsx_data):
    """Extract text from an XLSX file."""
    try:
        df = pd.read_excel(BytesIO(xlsx_data))
        return df.to_string(index=False)
    except Exception as e:
        print(f"Error extracting text from XLSX: {e}")
        return ""

def extract_text_from_html(html_data):
    """Extract text from an HTML file."""
    try:
        soup = BeautifulSoup(html_data, 'html.parser')
        return soup.get_text()
    except Exception as e:
        print(f"Error extracting text from HTML: {e}")
        return ""

def extract_text_from_pdf(pdf_data):
    """Extract text from a PDF file."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    full_text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        full_text += page.get_text()
    return full_text

def is_text_based_pdf(pdf_data):
    """Check if a PDF file is text-based or scanned."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        if page.get_text().strip():
            return True
    return False


def convert_pdf_to_images(pdf_data):
    """Convert PDF pages to images with higher DPI for better OCR results."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    image_paths = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=600)
        image_path = f"page_{page_num + 1}.png"
        pix.save(image_path)
        image_paths.append(image_path)
    return image_paths

def extract_text_from_image(image_path):
    """Extract text from an image file using Azure Vision OCR."""
    try:
        with open(image_path, "rb") as image_stream:
            ocr_result = computervision_client.read_in_stream(image_stream, raw=True)

        operation_location = ocr_result.headers["Operation-Location"]
        operation_id = operation_location.split("/")[-1]

        while True:
            result = computervision_client.get_read_result(operation_id)
            if result.status not in ['notStarted', 'running']:
                break
            time.sleep(1)

        if result.status == OperationStatusCodes.succeeded:
            text = ""
            for read_result in result.analyze_result.read_results:
                for line in read_result.lines:
                    text += line.text + " "
            return text
        else:
            print("OCR failed: insufficient image quality.")
            return ""
    except Exception as e:
        print(f"Image is invalid for text extraction: {e}")
        return ""

def extract_selection_marks_and_text(pdf_data):
    """Extract selection marks and text lines from the document."""
    try:
        poller = form_recognizer_client.begin_analyze_document("prebuilt-document", pdf_data)
        result = poller.result()

        selection_marks = []
        text_lines = []

        for page in result.pages:
            for selection_mark in page.selection_marks:
                selection_marks.append({
                    "Page": page.page_number,
                    "State": selection_mark.state,
                    "Polygon": selection_mark.polygon
                })
                print(f"Selection Mark: Page {page.page_number}, State {selection_mark.state}, Polygon {selection_mark.polygon}")
            for line in page.lines:
                text_lines.append({
                    "Page": page.page_number,
                    "Text": line.content,
                    "Polygon": line.polygon
                })
                print(f"Text Line: Page {page.page_number}, Text {line.content}, Polygon {line.polygon}")

        return selection_marks, text_lines
    except Exception as e:
        print(f"Error extracting selection marks and text: {e}")
        return [], []

def associate_checkboxes_with_options(selection_marks, text_lines):
    """Associate checkboxes with their nearest text options."""
    checkboxes = []

    for selection_mark in selection_marks:
        nearest_text = None
        min_distance = float('inf')
        for line in text_lines:
            if line["Page"] == selection_mark["Page"]:
                line_center_x = sum([point[0] for point in line["Polygon"]]) / len(line["Polygon"])
                line_center_y = sum([point[1] for point in line["Polygon"]]) / len(line["Polygon"])
                selection_mark_center_x = sum([point[0] for point in selection_mark["Polygon"]]) / len(selection_mark["Polygon"])
                selection_mark_center_y = sum([point[1] for point in selection_mark["Polygon"]]) / len(selection_mark["Polygon"])
                distance = ((line_center_x - selection_mark_center_x) ** 2 + (line_center_y - selection_mark_center_y) ** 2) ** 0.5
                if distance < min_distance:
                    min_distance = distance
                    nearest_text = line["Text"]

        checkboxes.append({
            "Page": selection_mark["Page"],
            "State": selection_mark["State"],
            "Option": nearest_text
        })

    return checkboxes

def format_checkboxes_as_json(checkboxes):
    """Format checkboxes as JSON objects."""
    return json.dumps(checkboxes, indent=4)

def analyze_document_with_form_recognizer(pdf_data):
    """Analyze the PDF document using Azure Form Recognizer to extract tables and checkboxes."""
    try:
        selection_marks, text_lines = extract_selection_marks_and_text(pdf_data)
        checkboxes = associate_checkboxes_with_options(selection_marks, text_lines)

        poller = form_recognizer_client.begin_analyze_document("prebuilt-document", pdf_data)
        result = poller.result()

        tables = []
        for table in result.tables:
            table_data = []
            for cell in table.cells:
                while len(table_data) <= cell.row_index:
                    table_data.append([""] * table.column_count)
                table_data[cell.row_index][cell.column_index] = cell.content
            tables.append(table_data)

        return {
            "tables": tables,
            "checkboxes": checkboxes
        }
    except Exception as e:
        print(f"Error analyzing document with Form Recognizer: {e}")
        return {
            "tables": [],
            "checkboxes": []
        }




def process_pdf(pdf_data):
    """Process the PDF file to extract text, tables, and checkboxes."""
    try:
        if is_text_based_pdf(pdf_data):
            print("The PDF is text-based. Extracting text and analyzing for tables and checkboxes...")
            text = extract_text_from_pdf(pdf_data)
            analysis_results = analyze_document_with_form_recognizer(pdf_data)
            return {
                "text": text,
                "tables": analysis_results.get("tables", []),
                "checkboxes": analysis_results.get("checkboxes", [])
            }
        else:
            print("The PDF contains scanned images. Performing OCR and analyzing for tables and checkboxes...")
            image_paths = convert_pdf_to_images(pdf_data)
            full_text = ""
            for image_path in image_paths:
                text = extract_text_from_image(image_path)
                full_text += text
                os.remove(image_path)

            analysis_results = analyze_document_with_form_recognizer(pdf_data)
            return {
                "text": full_text,
                "tables": analysis_results.get("tables", []),
                "checkboxes": analysis_results.get("checkboxes", [])
            }
    except Exception as e:
        print(f"Error during PDF processing: {e}")
        return {}

def extract_text_from_attachment(attachment):
    """Extract text content from an attachment."""
    file_name = attachment.longFilename if attachment.longFilename else attachment.shortFilename
    file_data = attachment.data

    if file_name.lower().endswith('.docx'):
        print(f"Extracting text from docx file: {file_name}")
        return extract_doc(file_data)

    elif file_name.lower().endswith('.pdf'):
        print(f"Extracting text from pdf file: {file_name}")
        return process_pdf(file_data)

    elif file_name.lower().endswith('.txt'):
        print(f"Extracting text from txt file: {file_name}")
        return extract_text_from_txt(file_data)

    elif file_name.lower().endswith('.csv'):
        print(f"Extracting text from csv file: {file_name}")
        return extract_text_from_csv(file_data)

    elif file_name.lower().endswith('.xlsx'):
        print(f"Extracting text from xlsx file: {file_name}")
        return extract_text_from_xlsx(file_data)

    elif file_name.lower().endswith('.html'):
        print(f"Extracting text from html file: {file_name}")
        return extract_text_from_html(file_data)

    elif file_name.lower().endswith(('.jpg', '.jpeg', '.png')):
        print(f"Extracting text from image file: {file_name}")
        with open("temp_image", "wb") as f:
            f.write(file_data)
        text = extract_text_from_image("temp_image")
        os.remove("temp_image")
        return text

    else:
        print(f"Unsupported file type: {file_name}. Returning base64 encoded content.")
        return base64.b64encode(file_data).decode('utf-8')

def extract_text_from_msg(file_path):
    """Extract text content and attachments from an MSG file."""
    try:
        msg = extract_msg.Message(file_path)
        attachments = []

        for attachment in msg.attachments:
            file_name = attachment.longFilename if attachment.longFilename else attachment.shortFilename
            attachment_info = {
                "filename": file_name,
                "content": extract_text_from_attachment(attachment),
                "filetype": file_name.split('.')[-1]
            }
            attachments.append(attachment_info)

        return {
            "Subject": msg.subject,
            "From": msg.sender,
            "To": msg.to,
            "Date": msg.date,
            "Body": msg.body,
            "Attachments": attachments
        }
    except Exception as e:
        print(f"Error extracting details from MSG: {e}")
        return {"error": "Invalid attachment or MSG file."}


def extract_attachment(file_path):
    try:
        with open(file_path, 'rb') as f:
            msg = extract_msg(f)
            return msg.body  
    except Exception as e:
        print("Error:", e)
        return None
