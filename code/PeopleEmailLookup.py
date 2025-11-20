"""
PeopleEmailLookup
------------------

Helper functions for reading a contact CSV and resolving delivery
preferences and email addresses for Piluweri IDs (PLI-#).

Author: Mu Dell'Oro
Version: v2.0 
Date: 12.11.2025
Git: https://github.com/Capicodo/PDF-Splitter.git
"""

import csv
from typing import List

from ContactData import ContactData

csv_data: List[dict] = []


def extract_pli_id(name: str) -> int:
    """
    Extract the leading integer PLI ID from the beginning of a ``Dienstplan``
    text field.

    Examples:
        "32 casdf vadsf" -> 32
        "6 adsfa adf" -> 6
        "68 safd adf" -> 68

    Args:
        name: The text extracted from the Dienstplan field which should start
            with the numeric PLI ID.

    Returns:
        The extracted PLI ID as an integer.

    Raises:
        ValueError: If the first token is not an integer.
    """
    # Split the string by spaces and take the first part
    first_part = name.split()[0]

    try:
        pli_id = int(first_part)
        return pli_id
    except ValueError:
        raise ValueError(f"No valid PLI ID found in '{name}'")


def init(path: str):
    """
    Load the contact CSV into memory for later lookups.

    This function reads the CSV file at ``path`` using ``csv.DictReader`` and
    stores the rows in the module-level ``csv_data`` list.

    Args:
        path: Filesystem path to the CSV file encoded in UTF-8.

    Raises:
        Exception: If the file cannot be read or parsed.
    """
    global csv_data
    try:
        with open(path, newline="", encoding="utf-8") as csv_fh:
            csv_data = list(csv.DictReader(csv_fh))
    except Exception as e:
        csv_data = []  # fallback to empty list if reading fails
        raise Exception(f"Fehler beim Lesen der CSV: {e}")


def sheets_formated_str_to_bool(s: str) -> bool:
    """
    Convert spreadsheet-like TRUE/FALSE strings to booleans.

    Args:
        s: Input string such as 'TRUE' or 'FALSE' (case-insensitive).

    Returns:
        True if the string represents a truthy value, False if falsy.

    Raises:
        ValueError: If the string cannot be recognized as a boolean.
    """
    s = s.strip().lower()
    if s == "true":
        return True
    elif s == "false":
        return False
    else:
        raise ValueError(f"Cannot convert '{s}' to boolean")


def get_data_from_pli_id(pli_id: int) -> ContactData:
    """
    Find contact data for a given PLI ID from the loaded CSV rows.

    The CSV is expected to contain a column named "PLI - #" which stores the
    Piluweri ID. When a matching row is found this function builds and
    returns a ``ContactData`` instance.

    Args:
        pli_id: Piluweri ID to search for.

    Returns:
        A ``ContactData`` object for the matched row.

    Raises:
        Exception: If no matching deliver information is found.
    """
    pli_id_str = str(pli_id)
    print(f"COMPARE ID: --{pli_id_str}--")

    for row in csv_data:
        if row["PLI - #"] == pli_id_str:
            print(f"Found Contact Data. -> {row.get('Papierbericht')}, {row.get('Mail-Adresse')}")
            contact_data = ContactData(
                sheets_formated_str_to_bool(row["Papierbericht"]),
                row["Mail-Adresse"],
                pli_id,
                row["Rufname"],
                row["Nachname"],
            )

            print(contact_data.deliver_via_paper, contact_data.email)

            return contact_data

    raise Exception("❌ No Deliver Information found")
