from openpyxl import load_workbook
import warnings

import shutil
from datetime import datetime
import os

#Usuwa nie potrzebne blendy
warnings.filterwarnings("ignore", category=UserWarning)


'''
Import Wykaz.xlsx
'''
Wykaz = "Wykaz.xlsx"
Wykaz_SHEET = "Szablon 23.11.0.0"

main_wb = load_workbook(Wykaz)
main_ws = main_wb[Wykaz_SHEET]

'''
Testwoe wyswietlanie main
'''


def Test_main():
    for row in main_ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=10):
        for cell in row:
            print(cell.value, end=" ")
        print()

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
NR KOlumny kota bedzie odczytywana
2 = kolumna B, 3 = kolumna C, 4 = kolumna D
'''
NR_KOLUMNY = 2

def drukuj_druga_kolumne(eksporty):
    for eksport in eksporty:
        ws = eksport["worksheet"]
        print(f"\nPlik: {eksport['nazwa']}")

        for nr_wiersza, row in enumerate(
            ws.iter_rows(min_col=NR_KOLUMNY, max_col=NR_KOLUMNY, values_only=True),
            start=1,
        ):
            wartosc = row[0]

            if wartosc is None:
                continue
            if str(wartosc).strip() == "":
                continue
            if str(wartosc).strip().lower() == "nr rej":
                continue

            print(f"Wiersz {nr_wiersza}: {wartosc}")


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
    #drukuj_druga_kolumne(eksporty)
    Test_main()