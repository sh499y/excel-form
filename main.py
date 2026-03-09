from openpyxl import load_workbook
import warnings

import shutil
from datetime import datetime
import os


'''
Import Wykaz.xlsx
'''
Wykaz = "Wykaz.xlsx"
Wykaz_SHEET = "Szablon 23.11.0.0"

main_wb = load_workbook(Wykaz)
main_ws = main_wb[Wykaz_SHEET]


'''
Skanowanie i Wczytywanie Eksporty
'''
EXPORTS_DIR = "Eskporty"


def wczytaj_eksporty(folder=EXPORTS_DIR):
    eksporty = []

    if not os.path.isdir(folder):
        print(f"Folder z eksportami nie istnieje: {folder}")
        return eksporty

    for nazwa_pliku in sorted(os.listdir(folder)):
        if nazwa_pliku.startswith("~$"):
            continue
        if not nazwa_pliku.lower().endswith(".xlsx"):
            continue

        sciezka = os.path.join(folder, nazwa_pliku)
        wb = load_workbook(sciezka, data_only=True)
        ws = wb.active

        eksporty.append({
            "nazwa": nazwa_pliku,
            "sciezka": sciezka,
            "workbook": wb,
            "worksheet": ws,
        })

    return eksporty


'''
Tworzy backup w foldrze Backup
'''
def backup():
    oryginal = "Wykaz.xlsx"
    backup_dir = "Backup"

    os.makedirs(backup_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_path = os.path.join(backup_dir, f"Wykaz_backup_{timestamp}.xlsx")

    shutil.copy2(oryginal, backup_path)

    print("Backup zapisany:", backup_path)



if __name__ == '__main__':
    #backup()
    eksporty = wczytaj_eksporty()

    for eksport in eksporty:
        print(f"Wczytano: {eksport['nazwa']} | arkusz: {eksport['worksheet'].title}")
