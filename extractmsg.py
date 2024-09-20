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

load_dotenv()

# Set up cognitive credentials
subscription_key = os.getenv('subscription_key')
endpoint = os.getenv('endpoint')
computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))


def extract_doc(docx_data):
    """Extract text from DOCX file content."""
    with open("temp.docx", "wb") as f:
        f.write(docx_data)
    output = pypandoc.convert_file("temp.docx", 'rst')
    os.remove("temp.docx")
    return output


def extract_text_from_txt(txt_data):
    """Extract text from a txt file."""
    return txt_data.decode('utf-8')


def extract_text_from_csv(csv_data):
    """Extract text from a CSV file."""
    try:
        df = pd.read_csv(BytesIO(csv_data))
        text = df.to_string(index=False)
        return text
    except Exception as e:
        print(f"Error extracting text from CSV: {e}")
        return ""


def extract_text_from_xlsx(xlsx_data):
    """Extract text from an XLSX file."""
    try:
        df = pd.read_excel(BytesIO(xlsx_data))
        text = df.to_string(index=False)
        return text
    except Exception as e:
        print(f"Error extracting text from XLSX: {e}")
        return ""


def extract_text_from_html(html_data):
    """Extract text from an HTML file."""
    try:
        soup = BeautifulSoup(html_data, 'html.parser')
        text = soup.get_text()
        return text
    except Exception as e:
        print(f"Error extracting text from HTML: {e}")
        return ""


def extract_text_from_pdf(pdf_data):
    """Extract text from a PDF file."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    full_text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        full_text += text
    return full_text


def convert_pdf_to_images(pdf_data):
    """Convert PDF pages to images."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    image_paths = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
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
            print("Sorry, the image quality is not sufficient for text extraction. Please try again with a clearer image.")
            return ""
    except Exception as e:
        print(f"Image is invalid for text extraction: {e}")
        return ""


def is_text_based_pdf(pdf_data):
    """Check if a PDF file is text-based or scanned."""
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        if text.strip():
            return True
    return False


def process_pdf(pdf_data):
    """Process the PDF file to extract text."""
    if is_text_based_pdf(pdf_data):
        print("The PDF is text-based. Extracting text...")
        return extract_text_from_pdf(pdf_data)
    else:
        print("The PDF contains scanned images. Performing OCR...")
        try:
            image_paths = convert_pdf_to_images(pdf_data)
            full_text = ""
            for image_path in image_paths:
                text = extract_text_from_image(image_path)
                full_text += text
                os.remove(image_path)
            return full_text
        except Exception as e:
            print(f"Error during OCR extraction: {e}")
            return ""


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

    elif file_name.lower().endswith('.jpg') or file_name.lower().endswith('.jpeg') or file_name.lower().endswith('.png'):
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
            _, file_extension = os.path.splitext(file_name.lower())
            if file_extension in ['.jpg', '.jpeg', '.png']:
                continue
            
            attachment_info = {
                "filename": file_name,
                "content": extract_text_from_attachment(attachment),
                "filetype": file_extension[1:]
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
