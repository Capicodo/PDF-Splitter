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

from PeopleEmailLookup import getDataFromPLIID, extract_pli_id, init


logger: logging.Logger

sort_by_deliver_method: bool = True

contact_fails = []
contact_datas = []

reports: Dict[int, Report] = {}

reGexNameFindingPattern = r"Name:\s*(.*?)\n"
reGexDienstplanFindingPattern = r"Dienstplan:\s*(.*?)\n"

rawReportFilePath: str
destinationFolderPath: str
contact_data_csv_path: str

raw_report_doc: fitz.Document

outlook: win32.CDispatch
accounts = None

year = ""
month_name = ""


def setup_logging(log_file="log.txt", override_print=True):
    """
    Set up logging to file (with timestamp) and console (without timestamp).

    Parameters:
        log_file (str): Path to the log file.
        override_print (bool): If True, overrides the built-in print() to log automatically.
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
    """Remove quotes and surrounding whitespace from a file/folder path."""
    return path.strip().strip('"').strip("'")


def input_paths():

    global sort_by_deliver_method

    global rawReportFilePath
    global destinationFolderPath
    global contact_data_csv_path
    global raw_report_doc

    try:
        rawReportFilePath = input(
            "Pfad zum rohen Monatsbericht eingeben oder per Drag & Drop in das Fenster ziehen. \nAnschließend mit Enter bestätigen. \n\nPfad: "
        )
        rawReportFilePath = clean_path(rawReportFilePath)
        print(f"\n✅ Eingabepfad erkannt: {rawReportFilePath}\n")

        destinationFolderPath = input(
            "\nPfad zum Zielordner für die individuellen PDFs eingeben oder per Drag & Drop in das Fenster ziehen. \nAnschließend mit Enter bestätigen. \n\nPfad: "
        )
        destinationFolderPath = clean_path(destinationFolderPath)
        print(f"\n✅ Zielordner erkannt: {destinationFolderPath}\n")

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
            destinationFolderPath += f"/Kontaktdatenlos_und_Unsortiert"
            print(f"❌ FEHLER BEIM DATEI-ZUGRIFF: {e}")
            print(f"ℹ️ Es wird ohne Kontaktdatenliste gearbeitet")

        os.makedirs(destinationFolderPath, exist_ok=True)
        print("✅ Zielordner erstellt oder bereits vorhandenen gefunden")

        raw_report_doc = fitz.open(rawReportFilePath)
        print("✅ PDF erfolgreich geöffnet\n\n")

    except Exception as e:
        print(f"❌ FEHLER BEIM DATEI-ZUGRIFF: {e}")
        input("Zum Beenden beliebige Taste drücken...")
        raise SystemExit


def regexSearchText(_regex, _text):
    try:
        match = re.search(_regex, _text)
        if match:
            return match.group(1).strip()
        return None
    except Exception as e:
        print(f"❌ Regex-Fehler: {e}")
        return None


def create_report(
    _newNamePageIndex, _pageIndex, _name, contact_data: ContactData = None
):

    group_folder_path: str = destinationFolderPath

    if contact_data:
        group_folder_path += rf"\print" if contact_data.deliver_via_paper else rf"\send"
    else:
        group_folder_path += rf"\unsorted"

    os.makedirs(group_folder_path, exist_ok=True)

    new_doc = fitz.open()

    try:
        safe_name = re.sub(
            r'[<>:"/\\|?*]', "_", _name
        )  # sanitize for Windows filenames
        joinedPath = os.path.join(
            group_folder_path,
            f"Monatsbericht_{safe_name}_{_newNamePageIndex+1}-{_pageIndex+1}.pdf",
        )

        new_doc.insert_pdf(
            raw_report_doc, from_page=_newNamePageIndex, to_page=_pageIndex
        )
        new_doc.save(joinedPath)

        if contact_data:

            pli_id: int = contact_data.pli_id
            new_report: Report = Report(pli_id, joinedPath, contact_data)
            reports[pli_id] = new_report

        print(f"💾 Datei gespeichert: {joinedPath}")

    except Exception as e:

        print(f"❌ Fehler beim Speichern: {e}")


def getPagePersonInfos(_index):

    currentPage = raw_report_doc[_index]
    currentText = currentPage.get_text()

    print("-------------START Scanning--------------")
    currentName = regexSearchText(reGexNameFindingPattern, currentText)

    if currentName:
        print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")

    currentDienstplan = regexSearchText(reGexDienstplanFindingPattern, currentText)

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

    try:

        contact_data = getDataFromPLIID(pli_id)

        print(
            f"✅✅✅✅✅✅✅ For PLI-#: {pli_id} was deliver-information successfully found ✅✅✅✅✅✅"
        )
        print("")

    except Exception as e:

        print(f"⚠️⚠️⚠️⚠️ For PLI-#: {pli_id} was NO deliver-information found! ⚠️⚠️⚠️⚠️")
        raise Exception(f"{e}, {pli_id}")

    return contact_data


def iteratePages():

    last_name, last_pli_id = getPagePersonInfos(0)
    lastNewNamePageIndex = 0

    for pageIndex in range(raw_report_doc.page_count):

        currentName, current_pli_id = getPagePersonInfos(pageIndex)

        if last_name != currentName:

            contact_data = None

            if sort_by_deliver_method:

                try:
                    contact_data = get_searched_contact_data(last_pli_id)
                    contact_datas.append(contact_data)
                except Exception as e:
                    contact_fails.append(
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

    if contact_datas:
        print(
            f"\n\n✅✅✅ {len(contact_datas)} Kontaktdaten wurden gefunden: ✅✅✅\n\n"
        )
    for current_contact_data in contact_datas:

        print(f"✅ {current_contact_data.__dict__}")

    if contact_fails:
        print(f"\n⚠️⚠️⚠️ {len(contact_fails)} Kontaktdaten wurden nicht gefunden: ⚠️⚠️⚠️\n\n")
        for current_fail in contact_fails:
            print(f"⚠️ NICHT GEFUNDEN: {current_fail}")


def getAnswerYesNo():

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
""")

    print()
    print("v2.0")
    print("12.11.2025")
    print("Diese Version unterstützt das Teilen und Senden der Monatsberichte")
    print("Das Drucken wird in dieser Version noch NICHT unterstützt")
    print("\033[0m")

def send_emails():

    global accounts
    global outlook

    print_people_getting_emailed()

    outlook = win32.Dispatch("outlook.application")
    accounts = outlook.Session.Accounts

    sender_email = input("\nGib nun die Absender-Email an:\n")
    loop_check_sender(sender_email)

    print(f"\n❗Willst du wirklich JETZT die Berichte senden?")
    print(f"❗Diese Aktion kann nicht revidiert werden❗\n")

    decision: bool = getAnswerYesNo()

    if decision:
        print("ℹ️ Starting sending Emails")
        for report in reports.values():
            if not report.contact_data.deliver_via_paper:
                send_report_to(report, report.contact_data.email, sender_email)

    print("\n\n✔️ Die Emails wurden gesendet ✔️")
    print("⚠️ Schaue in deinem Postfach nach, ob die Emails wirklich rausgegangen sind!")


def print_people_getting_emailed():

    print(f"\nAn die folgenden Personen werden Monatsberichte gesendet:\n")

    for current_contact_data in [
        current_contact_data
        for current_contact_data in contact_datas
        if not current_contact_data.deliver_via_paper
    ]:
        print(
            f"✅ {current_contact_data.first_name} {current_contact_data.last_name} | {current_contact_data.email}"
        )


def send_report_to(report: Report, recipient_email: str, sender_email: str):

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
        Diese Nachricht wurde automatisch erstellt. Falls Schwierigkeiten auftreten, wende Dich bitte an mich.</p>
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

    while True:
        try:
            check_sender(sender_email)
            break
        except Exception as e:
            print(e)
            print("⚠️ Bitte versuche es erneut\n")
            sender_email = input("Bitte gib eine gültige Absenderadresse ein:\n")


def check_sender(sender_email: str):

    for account in accounts:

        if account.SmtpAddress.lower() == sender_email.lower():
            return

    raise Exception(
        "\n❌ Die Eingegebene Email konnte nicht in deinen Outlook-Konten gefunden werden ❌"
    )


def set_sender(mail, sender_email: str):

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
        iteratePages()

        print(
            f"\n\nWillst du {f'⚠️⚠️ Trotz {len(contact_fails)} Kontaktdaten-Fehlern ⚠️⚠️ \n' if contact_fails else ''}alle digital zu verarbeitenden Monatsberichte per EMAIL SENDEN? 📧"
        )

        decision: bool = getAnswerYesNo()
        if decision:
            send_emails()

    except Exception as e:
        print(f"❌ FEHLER BEIM ITERIEREN: {e}")
        print("❌❌❌ PDFs wurden nicht oder fehlerhaft erstellt ❌❌❌")

    input("\n\n\n\nZum BEENDEN des Programms beliebige Taste drücken...")


if __name__ == "__main__":
    main()
