import os
import fitz  # PyMuPDF
import re

from ContactData import ContactData

from PeopleEmailLookup import getDataFromName, getNameParts

def clean_path(path: str) -> str:
    """Remove quotes and surrounding whitespace from a file/folder path."""
    return path.strip().strip('"').strip("'")


try:
    rawFilePath = input(
        "Pfad zum rohen Monatsbericht eingeben oder per Drag & Drop in das Fenster ziehen \nAnschließend mit Enter bestätigen \n\nPfad: "
    )
    rawFilePath = clean_path(rawFilePath)
    print(f"\n✅ Eingabepfad erkannt: {rawFilePath}\n")

    destinationFolderPath = input(
        "\nPfad zum Zielordner für die individuellen PDFs eingeben oder per Drag & Drop in das Fenster ziehen \nAnschließend mit Enter bestätigen \n\nPfad: "
    )
    destinationFolderPath = clean_path(destinationFolderPath)
    print(f"\n✅ Zielordner erkannt: {destinationFolderPath}")

    os.makedirs(destinationFolderPath, exist_ok=True)
    print("✅ Zielordner erstellt oder bereits vorhandenen gefunden")

    doc = fitz.open(rawFilePath)
    print("✅ PDF erfolgreich geöffnet\n\n")

except Exception as e:
    print(f"❌ FEHLER BEIM DATEI-ZUGRIFF: {e}")
    input("Zum Beenden beliebige Taste drücken...")
    raise SystemExit

reGexNameFindingPattern = r"Name:\s*(.*?)\n"


def regexSearchForName(_text):
    try:
        match = re.search(reGexNameFindingPattern, _text)
        if match:
            return match.group(1).strip()
        return None
    except Exception as e:
        print(f"❌ Regex-Fehler: {e}")
        return None


def getPageName(_index):
    currentPage = doc[_index]
    currentText = currentPage.get_text()

    print("-------------START Name Scanning--------------")
    currentName = regexSearchForName(currentText)
    if currentName:
        print(f"✅ Seite {_index+1}: Name gefunden → {currentName}")
    else:
        raise Exception(f"❌ Kein Name auf Seite {_index+1} gefunden ❌")
    print("-------------END Name Scanning----------------")
    return currentName


def createIndividualPDF(_newNamePageIndex, _pageIndex, _name):
    # new_doc = fitz.open()
    try:
        safe_name = re.sub(
            r'[<>:"/\\|?*]', "_", _name
        )  # sanitize for Windows filenames
        joinedPath = os.path.join(
            destinationFolderPath,
            f"Monatsbericht_{safe_name}_{_newNamePageIndex+1}-{_pageIndex+1}.pdf",
        )
        
        # new_doc.insert_pdf(doc, from_page=_newNamePageIndex, to_page=_pageIndex)
        # new_doc.save(joinedPath)
        
        # print(f"💾 Datei gespeichert: {joinedPath}")
        
    except Exception as e:
        
        print(f"❌ Fehler beim Speichern: {e}")

def search_contact_data(_name):

    
    first_name, surname = getNameParts(_name)
    print(f"Current Name Parts:--{first_name}|{surname}--")

    deliverInformationFound = getDataFromName(_name)
    
    if (not deliverInformationFound):
        print(f"⚠️⚠️⚠️⚠️ For Name: {first_name} {surname} was NO deliver Information found! ⚠️⚠️⚠️⚠️")
        raise Exception(f"_{first_name}_{surname}_") 
    
    print(f"✅✅✅✅✅✅✅ For Name: {first_name} {surname} was deliver Information successfull found ✅✅✅✅✅✅")
    return ContactData(False, "test@testmail.test")

def iteratePages():
    
    lastName = getPageName(0)
    lastNewNamePageIndex = 0

    contact_fails = []
    
    for pageIndex in range(doc.page_count):
        
        currentName = getPageName(pageIndex)
        
        if lastName != currentName:
               
            try:
                contact_data = search_contact_data(lastName)
                print(f"Print Test Contact Data: {contact_data.__dict__}")
            except Exception as e:
                contact_fails.append(e)
                
            print ("\n\n") 
            print(
                f"🎯 Seitenwechsel bei Seite {pageIndex+1} → Neuer Name: {currentName}"
            )
            createIndividualPDF(lastNewNamePageIndex, pageIndex - 1, lastName)
            lastNewNamePageIndex = pageIndex
        lastName = currentName
        
        print("")
        
    if contact_fails:
        print(f"\n\n⚠️⚠️⚠️ {len(contact_fails)} Kontaktdaten wurden nicht gefunden: ⚠️⚠️⚠️\n\n")
        for current_fail in contact_fails:
            print(f"⚠️ NICHT GEFUNDEN: {current_fail}")

try:
    iteratePages()
    print("\n\n✅✅✅ PDFs wurden erfolgreich erstellt ✅✅✅\n\n")
    
except Exception as e:
    print(f"❌ FEHLER BEIM ITERIEREN: {e}")
    print("❌❌❌ PDFs wurden nicht oder fehlerhaft erstellt ❌❌❌")

input("\n\n\n\nZum BEENDEN des Programms beliebige Taste drücken...")
