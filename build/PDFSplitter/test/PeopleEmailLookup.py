import csv
import os
import re


import re

import re

import re


def getNameParts(name: str):
    """
    Extracts first name and surname from a name string that may include
    company prefixes like 'GbR' or 'OHG', surname particles like 'von', 'van', 'de', etc.,
    strips leading/trailing spaces, and removes spaces after apostrophes.
    """
    # Remove space after apostrophes
    name = re.sub(r"'\s+", "'", name.strip())

    # Common business prefixes
    business_words = {"GBR", "OHG", "KG", "UG", "GMBH"}

    # Surname particles
    surname_particles = {
        "von",
        "van",
        "de",
        "del",
        "della",
        "da",
        "di",
        "du",
        "la",
        "le",
        "zu",
        "zum",
        "zur",
    }

    # Split into parts
    parts = re.split(r"\s+", name)

    # Remove business prefixes
    clean_parts = [p for p in parts if p.upper() not in business_words]

    if not clean_parts:
        return ("", "")

    if len(clean_parts) == 1:
        return ("", clean_parts[0].strip())

    # Heuristic for first name and surname
    first_name = clean_parts[0]
    surname = " ".join(clean_parts[1:])

    for i, word in enumerate(clean_parts[1:], start=1):
        if word.lower() in surname_particles:
            first_name = " ".join(clean_parts[:i])
            surname = " ".join(clean_parts[i:])
            break

    return (first_name.strip(), surname.strip())


def getDataFromName(name: str):
    """This function finds the information, where the monthly report should be send, or if it should be printed

    TODO

    """

    first_name, surname = getNameParts(name)
    
    try:

        # Pfad relativ zum Skript
        script_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(script_dir, "Kontaktdaten.csv")

        with open(csv_path, newline="", encoding="utf-8") as f:

            reader = csv.DictReader(f)

            for row in reader:

                current_first_name = row["Rufname"]
                current_surname = row["Nachname"]
                # current_email = row["Mail-Adresse"]

                if current_first_name == first_name:
                    if current_surname == surname:
                        return True

    except Exception as e:

        print(f"Fehler beim Verarbeiten der CSV: {e}")

    return False
