import os
import fitz  # PyMuPDF
import re

try:
    rawFilePath = input(
        "Pfad zum rohen Monatsbericht eingeben oder per Drag & Drop in das Fenster ziehen \nAnschließend mit Enter bestätigen \nPfad: "
    )
    print("✅")
    destinationFolderPath = input(
        "Pfad zum Zielordner für die individuellen PDFs eingeben oder per Drag & Drop in das Fenster ziehen \nAnschließend mit Enter bestätigen \nPfad: "
    )
    os.makedirs(destinationFolderPath, exist_ok=True)
    print("✅")
    doc = fitz.open(rawFilePath)
except Exception as e:
    print(f"❌ EXCEPTION HANDLING INPUT: {e} ❌")

reGexNameFindingPattern = r"Name:\s*(.*?)\n"


def getPageName(_index):

    currentPage = doc[_index]
    currentText = currentPage.get_text()

    print("-------------START Name Scanning--------------")

    currentName = regexSearchForName(currentText)
    if currentName:
        print(f"✅ Page {_index+1} has a valid Name field. Name is: {currentName}")

    else:
        raise Exception("❌ ERROR – Name not Found ❌")

    print("-------------END Name Scanning----------------")

    return currentName


def regexSearchForName(_text):
    try:
        # Search for the name using the regex pattern
        match = re.search(reGexNameFindingPattern, _text)
        if match:
            currentName = match.group(1)
            if currentName:
                return currentName
            else:
                return None
        else:
            return None

    except re.error as e:
        # Handle regex-related errors
        print(f"❌ Regex error occurred: {e}")
        return None

    except Exception as e:
        # Catch any other errors in the name extraction process
        print(f"❌ An unexpected error occurred in name extraction: {e}")
        return None


def iteratePages():

    # lastName = getPageName(doc[1])
    lastName = "GbR Alexej Bergmann"
    lastNewNamePageIndex = 0

    for pageIndex in range(doc.page_count):
        currentName = getPageName(pageIndex)
        if lastName != currentName:
            print(f"🎯 Page {pageIndex+1} has a name change! New Name is {currentName}")

            createIndividualPDF(lastNewNamePageIndex, pageIndex - 1, lastName)
            lastNewNamePageIndex = pageIndex

        lastName = currentName

        for index in range(0, 2):
            print("")


def createIndividualPDF(_newNamePageIndex, _pageIndex, _name):

    new_doc = fitz.open()

    try:
        # Correct the file path concatenation
        joinedPath = os.path.join(
            destinationFolderPath,
            f"Monatsbericht_{_name}_{_newNamePageIndex+1}-{_pageIndex+1}.pdf",
        )

        new_doc.insert_pdf(doc, from_page=_newNamePageIndex, to_page=_pageIndex)

        new_doc.save(joinedPath)
        print(f"File saved to: {joinedPath}")

    except Exception as e:
        print(f"Error saving file: {e}")
        raise Exception(f"Error saving file: {e}")


try:
    iteratePages()
    print("✅✅✅  PDFs wurden erfolgreich erstellt  ✅✅✅")
except Exception as e:
    print(f"❌ EXCEPTION WHILE ITERATING: {e} ❌")
    print("❌❌❌  PDFs wurden nicht oder fehlerhaft erstellt  ❌❌❌")


input("Zum BEENDEN des Programms beliebige Taste drücken")
