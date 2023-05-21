# Dieses kleine Programm erstellt eine Excel-Tabelle und listet dort alle Unterordner eines Ordners auf,
# gibt diesen eine ID und stellt die Größe des jeweiligen Ordners dar. Außerdem wird in der Tabelle ein
# kleiner Rechner erstellt, der die Größen der Ordner zusammenrechnet. Ausgewählt wird mit einem "x" in
# der Spalte D (Auswahl). Die Pfade werden vom Benutzer abgefragt.

import os
import xlsxwriter
import time

try:
    folder_path = input("Bitte Pfad des Ordners angeben, dessen Unterordner aufgelistet werden sollen (z.B.: C:/Filme): ")
    excel_path = input("Bitte Pfad der Excel-Tabelle mit Endung angeben (z.B.: D:/Filmesammlung.xlsx): ")

    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({"bold": True})
    left = workbook.add_format({"align": "left"})

    worksheet.write("B2", "Kleines x in Spalte D eintragen, um Datei auszuwählen.", bold)
    worksheet.write("B4", "Größe (GB) Ausgewählt:", bold)
    worksheet.write_formula("B5", '=SUMIF(D8:D300,"=x",C8:C300)', left)
    worksheet.write("A7", "ID", bold)
    worksheet.write("B7", "Datei", bold)
    worksheet.write("C7", "Größe (GB)", bold)
    worksheet.write("D7", "Auswahl", bold)

    worksheet.set_column(0, 0, 4)
    worksheet.set_column(1, 1, 65)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 3, 8)


    def get_folder_size(folder_path):
        total_size = 0
        for path, dirs, files in os.walk(folder_path):
            for f in files:
                fp = os.path.join(path, f)
                total_size += os.path.getsize(fp)
        return total_size


    for i, subfolder in enumerate(os.listdir(folder_path)):
        worksheet.write(i+8, 0, i+1)
        worksheet.write(i+8, 1, subfolder)
        size_mb = get_folder_size(os.path.join(folder_path, subfolder)) / 1024 / 1024 / 1024
        size_mb = round(size_mb, 2)
        worksheet.write(i+8, 2, size_mb)

    workbook.close()

except Exception as error:
    print(error)

finally:
    time.sleep(3)