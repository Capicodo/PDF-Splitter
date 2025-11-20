"""
Report
------

Lightweight container that represents a single per-person monthly report
file and its associated contact information.

Author: Mu Dell'Oro
Version: v2.0
Date: 12.11.2025
Git: https://github.com/Capicodo/PDF-Splitter.git
"""

import fitz
from ContactData import ContactData


class Report:
    """
    Container for a generated per-person report file and contact info.

    Attributes:
        pli_id: Piluweri ID (int) for the person.
        document: Filesystem path or PyMuPDF document reference for the
            generated per-person PDF.
        contact_data: The associated ``ContactData`` instance.
    """

    def __init__(self, pli_id: int, document: str, contact_data: ContactData):
        """
        Initialize a ``Report`` instance.

        Args:
            pli_id: Piluweri ID for this report.
            document: Path to the generated PDF file (or a fitz.Document
                reference where used by calling code).
            contact_data: ``ContactData`` associated with this report.
        """
        self.pli_id: int = pli_id
        self.document: fitz.Document = document
        self.contact_data = contact_data
