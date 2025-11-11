import csv

from ContactData import ContactData

csv_data = []


def extract_pli_id(name: str) -> int:
    """
    Extracts the leading number (PLI-#) from a name string.


    Examples:
        "32 casdf vadsf" -> 32
        "6 adsfa adf" -> 6
        "68 safd adf" -> 68
    """
    # Split the string by spaces and take the first part
    first_part = name.split()[0]

    # Convert it to an integer
    try:
        pli_id = int(first_part)
        return pli_id
    except ValueError:
        raise ValueError(f"No valid PLI ID found in '{name}'")


def init(path: str):
    global csv_data
    # Read the CSV once and store it in a variable
    try:
        with open(path, newline="", encoding="utf-8") as csv_fh:
            csv_data = list(csv.DictReader(csv_fh))
    except Exception as e:
        csv_data = []  # fallback to empty list if reading fails
        raise Exception(f"Fehler beim Lesen der CSV: {e}")


def sheets_formated_str_to_bool(s: str) -> bool:
    """
    Converts a string like 'TRUE'/'FALSE' (case-insensitive) to a boolean.
    Raises ValueError if the string is not recognized.
    """
    s = s.strip().lower()
    if s == "true":
        return True
    elif s == "false":
        return False
    else:
        raise ValueError(f"Cannot convert '{s}' to boolean")


def getDataFromPLIID(pli_id: int) -> ContactData:
    """This function finds the information, where the monthly report should be send, or if it should be printed

    TODO

    """

    pli_id_str = str(pli_id)
    print(f"COMPARE ID: --{pli_id_str}--")

    for row in csv_data:

        if row["PLI - #"] == pli_id_str:
            print(
                f"Found Contact Data. -> {row["Papierbericht"]}, {row["Mail-Adresse"]}"
            )
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
