import win32print
import win32api

# Path to your file
file_path = r"C:\Monatsbericht_OHG Mu Dell' Oro_114-115.pdf"

try:
    
    # Get default printer
    printer_name = win32print.GetDefaultPrinter()

    # Print the file
    win32api.ShellExecute(
        0,
        "print",
        file_path,
        f'/d:"{printer_name}"',
        ".",
        0
    )
except Exception as e:
    print(e)

input("beliebige Taste drücken")