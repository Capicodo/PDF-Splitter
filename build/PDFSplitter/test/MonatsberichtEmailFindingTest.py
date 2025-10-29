import os
import fitz  # PyMuPDF
import re

from ContactData import ContactData

from PeopleEmailLookup import getDataFromPLIID, extract_pli_id, init


sort_by_deliver_method: bool = True


def clean_path(path: str) -> str:
    """Remove quotes and surrounding whitespace from a file/folder path."""
    return path.strip().strip('"').strip("'")


import pyfiglet


ascii_banner = pyfiglet.figlet_format("PDF Splitter\n                 by Mu")
# Colorize and reset
green_banner = f"\033[32m{ascii_banner}\033[0m"
print(green_banner)
print()
print()
print()


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
        destinationFolderPath += f"/Unsortiert"
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

reGexNameFindingPattern = r"Name:\s*(.*?)\n"
reGexDienstplanFindingPattern = r"Dienstplan:\s*(.*?)\n"


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

    contact_fails = []
    contact_datas = []

    for pageIndex in range(report_doc.page_count):

        currentName, current_dienstplan = getPagePersonInfos(pageIndex)

        if lastName != currentName:

            contact_data = None

            if sort_by_deliver_method:
                try:
                    contact_data = search_contact_data(last_dienstplan)
                    # print(
                    #     f"Print Contact Data: {contact_data.deliver_via_paper}, {contact_data.email}"
                    # )
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

    if contact_data:
        print(
            f"\n\n✅✅✅ {len(contact_fails)} Kontaktdaten wurden gefunden: ✅✅✅\n\n"
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


try:
    iteratePages()
    print("\n\n✅✅✅ PDFs wurden erstellt ✅✅✅\n\n")

except Exception as e:
    print(f"❌ FEHLER BEIM ITERIEREN: {e}")
    print("❌❌❌ PDFs wurden nicht oder fehlerhaft erstellt ❌❌❌")


input("\n\n\n\nZum BEENDEN des Programms beliebige Taste drücken...")
