from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup
from email.header import decode_header, make_header

def read_eml_file(file_path):
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    return msg

def extract_visible_text_from_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator='\n', strip=True)
    return text

def get_email_text(msg):
    text_parts = []

    def extract_text(part):
        content_type = part.get_content_type()
        charset = part.get_content_charset() or 'utf-8'
        
        if content_type == 'text/plain':
            print(f"Extracting text from text/plain part")
            text = part.get_payload(decode=True).decode(charset, errors='replace')
            text_parts.append(text)
        elif content_type == 'text/html':
            print(f"Extracting text from text/html part")
            html = part.get_payload(decode=True).decode(charset, errors='replace')
            visible_text = extract_visible_text_from_html(html)
            text_parts.append(visible_text)
        elif part.is_multipart():
            for subpart in part.iter_parts():
                extract_text(subpart)
    
    extract_text(msg)
    
    if text_parts:
        if text_parts:
            combined_text = text_parts[0]
            return combined_text if combined_text else "Unavailable"
    else:
        return "Email Body is Unavailable"
    
def decode_mime_words(s):
    return str(make_header(decode_header(s)))
    
def extract_email_details(msg):
    email_details = {
        'Subject': decode_mime_words(msg['subject']) if msg['subject'] else 'Not available',
        'From': decode_mime_words(msg['from']) if msg['from'] else 'Not available',
        'To': decode_mime_words(msg['to']) if msg['to'] else 'Not available',
        'Date': msg['date'] if msg['date'] else 'Not available',
        'Body': get_email_text(msg) if get_email_text else 'Not available'
    }
    return email_details


def read_email(file_path):
    msg = read_eml_file(file_path)
    email_text = extract_email_details(msg)
    return email_text
