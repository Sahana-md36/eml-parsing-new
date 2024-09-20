import warnings
from spire.doc import *

def extract_text_from_doc(doc_file):
    warnings.filterwarnings("ignore")
    document = Document()
    document.LoadFromFile(doc_file)
    document_text = document.GetText()
    document.Close()
    document_text = document_text.replace("\r\nEvaluation Warning: The document was created with Spire.Doc for Python.", "")

    return document_text
