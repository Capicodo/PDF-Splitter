import csv
import os

def main():
    try:
        # Pfad relativ zum Skript
        script_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(script_dir, "Kontaktdaten.csv")

        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row["Rufname"]
                email = row["Mail-Adresse"]
                print(name, email)
    except Exception as e:
        print(f"Fehler beim Verarbeiten der CSV: {e}")
    finally:
        input("Drücke Enter, um das Programm zu beenden...")

if __name__ == "__main__":
    main()
