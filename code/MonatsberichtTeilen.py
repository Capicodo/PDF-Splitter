import os
import fitz  # PyMuPDF
import pyfiglet
import re

import win32com.client as win32

from ContactData import ContactData

from PeopleEmailLookup import getDataFromPLIID, extract_pli_id, init


sort_by_deliver_method: bool = True

contact_fails = []
contact_datas = []

reGexNameFindingPattern = r"Name:\s*(.*?)\n"
reGexDienstplanFindingPattern = r"Dienstplan:\s*(.*?)\n"

rawReportFilePath: str
destinationFolderPath: str
contact_data_csv_path: str

report_doc: fitz.Document


def clean_path(path: str) -> str:
    """Remove quotes and surrounding whitespace from a file/folder path."""
    return path.strip().strip('"').strip("'")


def input_paths():

    global sort_by_deliver_method

    global rawReportFilePath
    global destinationFolderPath
    global contact_data_csv_path
    global report_doc

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

        report_doc = fitz.open(rawReportFilePath)
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


def getPagePersonInfos(_index):

    currentPage = report_doc[_index]
    currentText = currentPage.get_text()

    # print("-------------START Scanning--------------")
    currentName = regexSearchText(reGexNameFindingPattern, currentText)

    if currentName:
        # print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
        None
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")

    currentDienstplan = regexSearchText(reGexDienstplanFindingPattern, currentText)

    if currentName:
        None
        # print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")
    # print("-------------END Scanning----------------")

    return currentName, currentDienstplan


def createIndividualPDF(
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

        new_doc.insert_pdf(report_doc, from_page=_newNamePageIndex, to_page=_pageIndex)
        new_doc.save(joinedPath)

        print(f"💾 Datei gespeichert: {joinedPath}")

    except Exception as e:

        print(f"❌ Fehler beim Speichern: {e}")


def search_contact_data(_name):

    pli_id: int = extract_pli_id(_name)
    print(f"Current PLI ID:-->{pli_id}<--")

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

    lastName, last_dienstplan = getPagePersonInfos(0)
    lastNewNamePageIndex = 0

    for pageIndex in range(report_doc.page_count):

        currentName, current_dienstplan = getPagePersonInfos(pageIndex)

        if lastName != currentName:

            contact_data = None

            if sort_by_deliver_method:
                try:
                    contact_data = search_contact_data(last_dienstplan)
                    contact_datas.append(contact_data)
                except Exception as e:
                    contact_fails.append(
                        f"⚠️ Für {lastName} war Kontaktdatensuche fehlerhaft: {e} \n⚠️ Die PDF wurde in den unsorted-Ordner gelegt!⚠️"
                    )

            print("\n\n")
            print(
                f"🎯 Seitenwechsel bei Seite {pageIndex+1} → Neuer Name: {currentName}"
            )

            createIndividualPDF(
                lastNewNamePageIndex, pageIndex - 1, lastName, contact_data
            )

            lastNewNamePageIndex = pageIndex

        lastName = currentName
        last_dienstplan = current_dienstplan

        print("")

    if contact_datas:
        print(
            f"\n\n✅✅✅ {len(contact_datas)} Kontaktdaten wurden gefunden: ✅✅✅\n\n"
        )
    for current_contact_data in contact_datas:

        print(f"✅ {current_contact_data.__dict__}")

    if contact_fails:
        print("\n\n")
        print(
            f"\n\n⚠️⚠️⚠️ {len(contact_fails)} Kontaktdaten wurden nicht gefunden: ⚠️⚠️⚠️\n\n"
        )
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
            print("\n ❌ Ungültige Eingabe. Du wirst erneut zur Eingabe aufgefordert.")


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
"""
        + "\033[0m"
    )

    print()
    print()
    print()


def send_emails():

    print(f"\nAn die folgenden Personen werden die Monatsberichte gesendet:\n")

    for current_contact_data in [
        current_contact_data
        for current_contact_data in contact_datas
        if not current_contact_data.deliver_via_paper
    ]:
        print(f"✅ {current_contact_data.__dict__}")

    send_example_email()


def send_example_email():

    test_example_goal_email: str = "calvin.delloro@piluweri.de"
    test_example_sender_email: str = "dev@ite-pli.de"
    test_example_sender_email = input("\nGib nun die Absender-Email an:\n")

    try:

        outlook: win32.CDispatch = win32.Dispatch("outlook.application")

        accounts = outlook.Session.Accounts

        mail = outlook.CreateItem(0)

        try_loop_set_sender(test_example_sender_email, accounts, mail)

        mail.To = test_example_goal_email
        mail.Subject = "PDFS Python Script Test"
        mail.Body = "PDFS Python Script Test Body"
        mail.HTMLBody = "<h2>HTML Message body</h2>"  # this field is optional

        # To attach a file to the email (optional):
        # attachment = "Path to the attachment"

        full_path = next(
            os.path.join(destinationFolderPath, "send", f)
            for f in os.listdir(os.path.join(destinationFolderPath, "send"))
        )

        mail.Attachments.Add(full_path)
        mail.Send()

    except Exception as e:
        print(f"❌ Error sending Email ❌ \n {e}")


def try_loop_set_sender(sender_email, accounts, mail):

    while True:
        try:
            set_sender(accounts, mail, sender_email)
            break
        except Exception as e:
            print(e)
            print("⚠️ Bitte versuche es erneut\n")
            sender_email = input("Bitte gib eine gültige Absenderadresse ein:\n")


def set_sender(accounts, mail, sender_email: str):

    for account in accounts:

        if account.SmtpAddress.lower() == sender_email.lower():
            mail._oleobj_.Invoke(
                *(64209, 0, 8, 0, account)
            )  # This sets SendUsingAccount
            return

    raise Exception(
        "\n❌ Die Eingegebene Email konnte nicht in deinen Outlook-Konten gefunden werden ❌"
    )


########################################
############### MAIN ###################
########################################


print_banner()

input_paths()

try:
    iteratePages()
    print("\n\n✅✅✅ PDFs wurden erstellt ✅✅✅\n\n")

except Exception as e:
    print(f"❌ FEHLER BEIM ITERIEREN: {e}")
    print("❌❌❌ PDFs wurden nicht oder fehlerhaft erstellt ❌❌❌")


print(
    "\n\nWillst du JETZT alle digital zu verarbeitenden Monatsberichte per Email senden?"
)

decision: bool = getAnswerYesNo()
if decision:
    send_emails()

input("\n\n\n\nZum BEENDEN des Programms beliebige Taste drücken...")
