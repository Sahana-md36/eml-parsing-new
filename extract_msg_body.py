from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup

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
            text = part.get_payload(decode=True).decode(charset, errors='replace')
            text_parts.append(text)
        elif content_type == 'text/html':
            html = part.get_payload(decode=True).decode(charset, errors='replace')
            visible_text = extract_visible_text_from_html(html)
            text_parts.append(visible_text)
        elif part.is_multipart():
            for subpart in part.iter_parts():
                extract_text(subpart)

    extract_text(msg)

    if text_parts:
        combined_text = '\n'.join(text_parts)
        return combined_text if combined_text else "Unavailable"
    else:
        return "Email Body is Unavailable"

def read_email_content(file_path):
    msg = read_eml_file(file_path)
    email_text = get_email_text(msg)
    return email_text


