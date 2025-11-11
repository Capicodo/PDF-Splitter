import fitz
from ContactData import ContactData


class Report:
    def __init__(self, pli_id: int, document: str, contact_data: ContactData):
        self.pli_id: int = pli_id
        self.document: fitz.Document = document
        self.contact_data = contact_data
