"""
Monatsbericht Automat Test
----------------------

Script to split a large exported monthly-report PDF (from Timoto) into
per-person PDFs, sort them according to delivery preference (email or paper),
and optionally send email attachments via Outlook. The script extracts the
Piluweri ID (PLI-#) from the "Dienstplan" field printed on the PDF, uses a
contact CSV to determine delivery preferences and addresses, and logs all
console output to a central log file.

Background, goals and process description are recorded in the project wiki
and README. This module implements the user-facing CLI orchestration and PDF
splitting logic.

Author: Mu Dell'Oro
Version: v2.1 TESTING VERSION (27.11.2025)
Date: 27.11.2025
GitHub: https://github.com/Capicodo/PDF-Splitter.git

"""

########################################
############# IMPORTS ##################
########################################

import datetime
import locale
import logging
import os
import random
import sys
import time
from typing import Dict
import fitz  # PyMuPDF
import pyfiglet
import re

import win32com.client as win32

from ContactData import ContactData
from Report import Report

from PeopleEmailLookup import get_data_from_pli_id, extract_pli_id, init

########################################
############# GLOBALS ##################
########################################

logger: logging.Logger

sort_by_deliver_method: bool = True

contact_failures = []
contact_data_list = []

reports: Dict[int, Report] = {}

regex_name_finding_pattern = r"Name:\s*(.*?)\n"
regex_dienstplan_finding_pattern = r"Dienstplan:\s*(.*?)\n"

raw_report_file_path: str
destination_folder_path: str
contact_data_csv_path: str

raw_report_doc: fitz.Document

outlook: win32.CDispatch
accounts = None

year = ""
month_name = ""

########################################
############ FUNCTIONS #################
########################################


def setup_logging(log_file="log.txt", override_print=True):
    """
    Set up logging to file (with timestamp) and console (without timestamp).

    Args:
        log_file (str): Path to the log file.
        override_print (bool): If True, overrides the built-in print() to log automatically.

    Returns:
        logging.Logger: Configured logger instance.
    """
    # Create logger
    logger = logging.getLogger("my_logger")
    logger.setLevel(logging.INFO)  # log everything INFO and above
    logger.propagate = False  # avoid duplicate logs if root logger exists

    # Clear existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    # --- File handler with timestamp ---
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_formatter = logging.Formatter("%(asctime)s | %(message)s")
    file_handler.setFormatter(file_formatter)

    # --- Console handler without timestamp ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_formatter = logging.Formatter("%(message)s")
    console_handler.setFormatter(console_formatter)

    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # Optional: override print()
    if override_print:

        def custom_print(*args, **kwargs):
            sep = kwargs.get("sep", " ")
            end = kwargs.get("end", "\n")
            message = sep.join(str(a) for a in args) + end
            logger.info(
                message.rstrip()
            )  # remove extra newline since logger adds its own

        # Override built-in print
        builtins = __import__("builtins")
        builtins.print = custom_print

    return logger


def setup_date_month_year():
    """
    Set the global `year` and `month_name` variables to represent the previous
    month.

    The function configures the German locale for month name formatting if
    possible, computes the previous month relative to the current date, and
    stores the localized month name and 4-digit year into the module-level
    globals `month_name` and `year`.

    Returns:
        None

    Notes:
        If the configured German locale is not available this function will
        attempt a fallback and will print an error message but continue.
    """
    global year
    global month_name
    try:

        # Set locale to German
        try:
            locale.setlocale(locale.LC_TIME, "German_Germany.1252")
        except locale.Error as e:

            print(f"German locale not available on this system. {e}")
            locale.setlocale(locale.LC_TIME, "de_DE.UTF-8")

        # Get current date
        now = datetime.datetime.now()

        # Compute the previous month
        if now.month == 1:
            prev_month = 12
            prev_year = now.year - 1
        else:
            prev_month = now.month - 1
            prev_year = now.year

        # Create a date object for the previous month (use day=1)
        prev_date = datetime.datetime(prev_year, prev_month, 1)
        # Get full month name in German
        month_name = prev_date.strftime("%B")
        # Get 4-digit year
        year = prev_date.strftime("%Y")

    except Exception as e:
        print(f"Fehler im erkennen des Monats und des Jahres")


def clean_path(path: str) -> str:
    """Remove quotes and surrounding whitespace from a file or folder path.

    Args:
        path: Raw file or folder path provided by user input (may include
            surrounding quotes from drag-and-drop behavior).

    Returns:
        A cleaned path string with surrounding whitespace and quotes removed.
    """
    return path.strip().strip('"').strip("'")


def input_paths():
    """
    Prompt the user for input file paths and initialize required resources.

    The function asks the user (via console input) for the raw monthly report
    PDF path, the destination folder for the split PDFs, and the CSV path for
    contact data. It attempts to initialize the contact data lookup, creates
    the destination folder if necessary, and opens the raw PDF with PyMuPDF.

    Globals set:
        rawReportFilePath, destinationFolderPath, contact_data_csv_path,
        raw_report_doc, sort_by_deliver_method

    Raises:
        SystemExit: If file access or PDF opening fails the function prints an
            error message and exits the program.
    """

    global sort_by_deliver_method

    global raw_report_file_path
    global destination_folder_path
    global contact_data_csv_path
    global raw_report_doc

    try:
        raw_report_file_path = input(
            "Pfad zum rohen Monatsbericht eingeben oder per Drag & Drop in das Fenster ziehen. \nAnschließend mit Enter bestätigen. \n\nPfad: "
        )
        raw_report_file_path = clean_path(raw_report_file_path)
        print(f"\n✅ Eingabepfad erkannt: {raw_report_file_path}\n")

        destination_folder_path = input(
            "\nPfad zum Zielordner für die individuellen PDFs eingeben oder per Drag & Drop in das Fenster ziehen. \nAnschließend mit Enter bestätigen. \n\nPfad: "
        )
        destination_folder_path = clean_path(destination_folder_path)
        print(f"\n✅ Zielordner erkannt: {destination_folder_path}\n")

        contact_data_csv_path = input(
            "\nPfad zur Kontaktdaten(CSV)-Datei eingeben oder per Drag & Drop in das Fenster ziehen \nAnschließend mit Enter bestätigen. \n\nPfad: "
        )
        contact_data_csv_path = clean_path(contact_data_csv_path)
        print(f"\n✅ Eingabepfad erkannt: {contact_data_csv_path}\n")

        try:
            init(contact_data_csv_path)
            print("✅ Kontaktdaten erfolgreich initialisiert")
        except Exception as e:
            sort_by_deliver_method = False
            destination_folder_path += f"/Kontaktdatenlos_und_Unsortiert"
            print(f"❌ FEHLER BEIM DATEI-ZUGRIFF: {e}")
            print(f"ℹ️ Es wird ohne Kontaktdatenliste gearbeitet")

        os.makedirs(destination_folder_path, exist_ok=True)
        print("✅ Zielordner erstellt oder bereits vorhandenen gefunden")

        raw_report_doc = fitz.open(raw_report_file_path)
        print("✅ PDF erfolgreich geöffnet\n\n")

    except Exception as e:
        print(f"❌ FEHLER BEIM DATEI-ZUGRIFF: {e}")
        input("Zum Beenden beliebige Taste drücken...")
        raise SystemExit


def regex_search_text(_regex, _text):
    """
    Search `_text` for `_regex` and return the first capture group if found.

    Args:
        _regex: A regular expression pattern with at least one capture group.
        _text: The text to search.

    Returns:
        The stripped first capture group as a string if a match is found,
        otherwise ``None``.

    Notes:
        Any internal regex errors are caught, logged to console and ``None`` is
        returned.
    """
    try:
        match = re.search(_regex, _text)
        if match:
            return match.group(1).strip()
        return None
    except Exception as e:
        print(f"❌ Regex-Fehler: {e}")
        return None


def create_report(
    start_page_index, end_page_index, person_name, contact_data: ContactData = None
):
    """
    Create a per-person PDF by slicing the raw report and save it to disk.

    The function extracts pages from the global `raw_report_doc` starting at
    `_newNamePageIndex` up to `_pageIndex` (inclusive), writes the new PDF to
    a subfolder of `destinationFolderPath` depending on delivery preference
    (``send``, ``print`` or ``unsorted``), and registers a `Report` object in
    the module-level `reports` dictionary when ``contact_data`` is provided.

    Args:
        _newNamePageIndex: First page index for the person's report (0-based).
        _pageIndex: Last page index for the person's report (0-based).
        _name: Person's name used to build the target filename.
        contact_data: Optional ``ContactData`` used to determine destination
            folder and to create a `Report` entry.

    Returns:
        None

    Notes:
        Filenames are sanitized for Windows and any errors during save are
        printed to console.
    """

    group_folder_path: str = destination_folder_path

    if contact_data:
        group_folder_path += rf"\print" if contact_data.deliver_via_paper else rf"\send"
    else:
        group_folder_path += rf"\unsorted"

    os.makedirs(group_folder_path, exist_ok=True)

    new_doc = fitz.open()

    try:
        safe_name = re.sub(
            r'[<>:"/\\|?*]', "_", person_name
        )  # sanitize for Windows filenames
        joined_path = os.path.join(
            group_folder_path,
            f"Monatsbericht_{safe_name}_{start_page_index+1}-{end_page_index+1}.pdf",
        )

        new_doc.insert_pdf(
            raw_report_doc, from_page=start_page_index, to_page=end_page_index
        )
        new_doc.save(joined_path)

        if contact_data:

            pli_id: int = contact_data.pli_id
            new_report: Report = Report(pli_id, joined_path, contact_data)
            reports[pli_id] = new_report

        print(f"💾 Datei gespeichert: {joined_path}")
    except Exception as e:

        print(f"❌ Fehler beim Speichern: {e}")


def get_page_person_infos(_index):
    """
    Extract the name and PLI ID from a page in the raw report PDF.

    Args:
        _index: Page index (0-based) to scan in `raw_report_doc`.

    Returns:
        A tuple of ``(name, pli_id)`` where ``name`` is the extracted person
        name string and ``pli_id`` is an integer PLI ID if successfully
        extracted; ``pli_id`` will be ``None`` on failure to parse.

    Raises:
        Exception: If the expected name field is not found on the page.
    """

    currentPage = raw_report_doc[_index]
    currentText = currentPage.get_text()

    print("-------------START Scanning--------------")
    currentName = regex_search_text(regex_name_finding_pattern, currentText)

    if currentName:
        print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")

    currentDienstplan = regex_search_text(regex_dienstplan_finding_pattern, currentText)

    if currentName:
        print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")
    print("-------------END Scanning----------------")

    try:

        pli_id: int = extract_pli_id(currentDienstplan)
        print(f"Found PLI ID:-->{pli_id}<--")

    except Exception as e:

        print("❌❌❌ Fehler beim Bearbeiten der PLI-# ❌❌❌")
        print(e)
        return currentName, None

    return currentName, pli_id


def get_searched_contact_data(pli_id):
    """
    Retrieve contact data for a given PLI ID using the CSV-based lookup.

    Args:
        pli_id: Piluweri ID used to search the contact CSV.

    Returns:
        A ``ContactData`` instance corresponding to the PLI ID.

    Raises:
        Exception: If no deliver information is found or the lookup raises an
            error. The original error will be propagated with the PLI ID added
            for context.
    """

    try:
        contact_data = get_data_from_pli_id(pli_id)
        print(
            f"✅✅✅✅✅✅✅ For PLI-#: {pli_id} was deliver-information successfully found ✅✅✅✅✅✅"
        )
        print("")
    except Exception as e:
        print(f"⚠️⚠️⚠️⚠️ For PLI-#: {pli_id} was NO deliver-information found! ⚠️⚠️⚠️⚠️")
        raise Exception(f"{e}, {pli_id}")

    return contact_data


def iterate_pages():
    """
    Iterate over pages in the raw PDF and split them into per-person PDFs.

    The function walks through every page of the global `raw_report_doc`,
    detects changes in the `Name:` field to determine page boundaries for a
    single person's report, attempts to resolve contact data for each person
    and calls ``create_report`` to write the per-person PDF files. Found
    contact entries are appended to `contact_datas` and lookup failures to
    `contact_fails`.

    Returns:
        None
    """

    last_name, last_pli_id = get_page_person_infos(0)
    lastNewNamePageIndex = 0

    for pageIndex in range(raw_report_doc.page_count):

        currentName, current_pli_id = get_page_person_infos(pageIndex)

        if last_name != currentName:

            contact_data = None

            if sort_by_deliver_method:

                try:
                    contact_data = get_searched_contact_data(last_pli_id)
                    contact_data_list.append(contact_data)
                except Exception as e:
                    contact_failures.append(
                        f"⚠️ Für {last_name} war Kontaktdatensuche fehlerhaft: {e} \n⚠️ Die PDF wurde in den unsorted-Ordner gelegt!⚠️"
                    )

            print("\n\n")
            print(
                f"🎯 Seitenwechsel bei Seite {pageIndex+1} → Neuer Name: {currentName}"
            )

            create_report(lastNewNamePageIndex, pageIndex - 1, last_name, contact_data)

            lastNewNamePageIndex = pageIndex

        last_name = currentName
        last_pli_id = current_pli_id

        print("")

    print("\n\n✅✅✅ PDFs wurden erstellt ✅✅✅\n\n")

    if contact_data_list:
        print(
            f"\n\n✅✅✅ {len(contact_data_list)} Kontaktdaten wurden gefunden: ✅✅✅\n\n"
        )
    for current_contact_data in contact_data_list:

        print(f"✅ {current_contact_data.__dict__}")

    if contact_failures:
        print(
            f"\n⚠️⚠️⚠️ {len(contact_failures)} Kontaktdaten wurden nicht gefunden: ⚠️⚠️⚠️\n\n"
        )
        for current_fail in contact_failures:
            print(f"⚠️ NICHT GEFUNDEN: {current_fail}")


def get_answer_yes_no():
    """
    Prompt the user for a yes/no answer and return a boolean.

    The prompt accepts 'y' or 'n' (case-insensitive). It will repeat until a
    valid answer is entered.

    Returns:
        True if the user entered 'y', False if the user entered 'n'.
    """

    while True:

        print("\n (Y -> Ja) | (N -> Nein)")
        answer: str = input("\nEingabe:")

        if str.lower(answer) == "y":
            return True
        elif str.lower(answer) == "n":
            return False
        else:
            print("\n ❌ Ungültige Eingabe. Bitte erneut Eingeben")


def print_banner():
    """
    Print an ASCII banner and version/introduction text to the console.

    The banner includes the module version and a short note about supported
    features.
    """

    print(
        "\033[32m"
        + """
 ____  ____  _____   ____        _ _ _   _
|  _ \\|  _ \\|  ___| / ___| _ __ | (_) |_| |_ ___ _ __
| |_) | | | | |_    \\___ \\| '_ \\| | | __| __/ _ \\ '__|
|  __/| |_| |  _|    ___) | |_) | | | |_| ||  __/ |
|_|   |____/|_|     |____/| .__/|_|_|\\__|\\__\\___|_|
                          |_|
                  _             __  __
                 | |__  _   _  |  \\/  |_   _
                 | '_ \\| | | | | |\\/| | | | |
                 | |_) | |_| | | |  | | |_| |
                 |_.__/ \\__, | |_|  |_|\\__,_|
                        |___/
"""
    )

    print()
    print("v2.1 TEST VERSION")
    print("27.11.2025")
    print("Diese Version unterstützt das Teilen und Senden der Monatsberichte")
    print("Das Drucken wird in dieser Version noch NICHT unterstützt")
    print("Dies ist eine TESTVERSION, es wird nur an Mu Dell'Oro gesendet")
    print("\033[0m")


def send_emails():
    """
    Send emails for all reports that are configured to be delivered by email.

    The function lists the recipients, prompts the user for the sender account
    and confirmation, then iterates over created `Report` instances and sends
    the reports to recipients who prefer email delivery via Outlook.

    Returns:
        None
    """

    global accounts
    global outlook

    print_people_getting_emailed()

    outlook = win32.Dispatch("outlook.application")
    accounts = outlook.Session.Accounts

    sender_email = input("\nGib nun die Absender-Email an:\n")
    loop_check_sender(sender_email)

    print(f"\n❗Willst du wirklich JETZT die Berichte senden?")
    print(f"❗Diese Aktion kann nicht revidiert werden❗\n")

    decision: bool = get_answer_yes_no()

    if decision:
        print("ℹ️ Starting sending Emails")
        for report in reports.values():
            if not report.contact_data.deliver_via_paper:
                if report.contact_data.last_name == "Dell'Oro":
                    send_report_to(report, report.contact_data.email, sender_email)

    print("\n\n✔️ Die Emails wurden gesendet ✔️")
    print("⚠️ Schaue in deinem Postfach nach, ob die Emails wirklich rausgegangen sind!")


def print_people_getting_emailed():
    """
    Print a list of people who will receive monthly reports by email.

    Uses the module-level `contact_datas` list and filters for entries where
    `deliver_via_paper` is ``False``.
    """

    print(f"\nAn die folgenden Personen werden Monatsberichte gesendet:\n")

    for current_contact_data in [
        current_contact_data
        for current_contact_data in contact_data_list
        if not current_contact_data.deliver_via_paper
    ]:
        print(
            f"✅ {current_contact_data.first_name} {current_contact_data.last_name} | {current_contact_data.email}"
        )


def send_report_to(report: Report, recipient_email: str, sender_email: str):
    """
    Compose and send a single report as an Outlook email message.

    Args:
        report: A `Report` object holding the document and contact info.
        recipient_email: The recipient email address.
        sender_email: The sender email address configured in Outlook.

    Raises:
        Exception: Re-raises any exception thrown while composing or sending
            the email after printing a short error notice.
    """

    try:

        mail = outlook.CreateItem(0)

        set_sender(mail, sender_email)

        mail.To = recipient_email

        print("Monat:", month_name)
        print("Jahr:", year)
        mail.Subject = f"Monatsbericht {month_name} {year}"

        mail.Display(False)
        # time.sleep(0.05)
        # Generate a temporary signature by creating a new mail (this preserves Outlook’s default signature)
        # The default signature is automatically included when accessing mail.HTMLBody before setting it manually.
        signature = mail.HTMLBody  # This retrieves the default Outlook signature

        custom_body = f"""
        <p>Hallo {report.contact_data.first_name} {report.contact_data.last_name},</p>
        <p>Anbei findest Du Deinen aktuellen Monatsbericht.<br>
        Diese Nachricht wurde automatisch erstellt. Falls Schwierigkeiten auftreten, wende Dich bitte an mich.<br>
        <br>
        Falls in der Spalte 'Bemerkung/Projekt' 'Falsche Buchung' oder 'Durch System korrigiert' steht, kontaktiere bitte Deinen
        Verantwortlichen.</p>
        <br>
        <p>Viele Grüße</p>
        <br>
        """

        # Append your custom message *before* the signature
        mail.HTMLBody = custom_body + signature

        mail.Attachments.Add(report.document)
        mail.Send()

        print(
            f"✅ Bericht von {report.contact_data.first_name} {report.contact_data.last_name} erfolgreich zu {recipient_email} gesendet"
        )
    except Exception as e:
        print(f"❌ Error sending Email ❌ \n {e}")
        raise e


def loop_check_sender(sender_email):
    """
    Repeatedly prompt for a sender address until it is validated by
    ``check_sender``.

    Args:
        sender_email: Initial sender email to validate.

    Returns:
        None
    """

    while True:
        try:
            check_sender(sender_email)
            break
        except Exception as e:
            print(e)
            print("⚠️ Bitte versuche es erneut\n")
            sender_email = input("Bitte gib eine gültige Absenderadresse ein:\n")


def check_sender(sender_email: str):
    """
    Verify that the provided sender email exists in the global Outlook
    `accounts` collection.

    Args:
        sender_email: Email address to validate.

    Raises:
        Exception: If the sender email is not found in `accounts`.
    """

    for account in accounts:
        if account.SmtpAddress.lower() == sender_email.lower():
            return

    raise Exception(
        "\n❌ Die Eingegebene Email konnte nicht in deinen Outlook-Konten gefunden werden ❌"
    )


def set_sender(mail, sender_email: str):
    """
    Set the Outlook ``SendUsingAccount`` on the provided mail item.

    Args:
        mail: An Outlook mail item object returned by ``outlook.CreateItem``.
        sender_email: Email address to set as the sending account.

    Raises:
        Exception: If no matching Outlook account is found for the provided
            sender email.
    """

    for account in accounts:
        if account.SmtpAddress.lower() == sender_email.lower():
            mail._oleobj_.Invoke(
                *(64209, 0, 8, 0, account)
            )  # This sets SendUsingAccount
            return

    raise Exception("\n❌ Problem Beim Setzen der Sender Email ❌")


########################################
############### MAIN ###################
########################################


def main():

    global logger
    logger = setup_logging("log.txt")  # prints automatically go to console + file

    setup_date_month_year()

    print_banner()

    input_paths()

    try:
        iterate_pages()

        print(
            f"\n\nWillst du {f'⚠️⚠️ Trotz {len(contact_failures)} Kontaktdaten-Fehlern ⚠️⚠️ \n' if contact_failures else ''}alle digital zu verarbeitenden Monatsberichte per EMAIL SENDEN? 📧"
        )

        decision: bool = get_answer_yes_no()
        if decision:
            send_emails()

    except Exception as e:
        print(f"❌ FEHLER BEIM ITERIEREN: {e}")
        print("❌❌❌ PDFs wurden nicht oder fehlerhaft erstellt ❌❌❌")

    input("\n\n\n\nZum BEENDEN des Programms beliebige Taste drücken...")


if __name__ == "__main__":
    main()
