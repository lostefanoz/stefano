from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter import *

import os
from openpyxl import load_workbook, Workbook


def seleziona_file():
    global file_paths

    file_dialog = filedialog.askopenfilename(title="Seleziona file", filetypes=(("File Excel", "*.xlsx"),))
    if file_dialog:
        file_paths.append(file_dialog)
        listbox_files.insert(END, os.path.basename(file_dialog))


def rimuovi_file():
    global file_paths

    index = listbox_files.curselection()
    if index:
        listbox_files.delete(index)
        file_paths.pop(index[0])


def confronta_valori():
    global file_paths

    if len(file_paths) < 2:
        messagebox.showwarning("Attenzione", "Seleziona almeno due file da confrontare.")
        return

    colonna_prezzo = "NETTO"  # Modifica la colonna del prezzo se necessario
    risultati_confronto = {}

    progress_bar["maximum"] = len(file_paths)

    for i, file_path in enumerate(file_paths, start=1):
        wb = load_workbook(file_path)
        sheet = wb.active

        progress_label.set(f"Confrontando file {i}/{len(file_paths)}")
        root.update()

        # Aggiorna il dizionario column_index
        column_index = {
            "Codice EAN": 0,
            "Codice": 1,
            "Descrizione": None,  # Inizialmente impostato su None
            "NETTO": 3
        }

        # Trova l'indice corretto per "Descrizione" in base all'header delle colonne
        for row in sheet.iter_rows(min_row=1, max_row=2, values_only=True):
            for index, nome_colonna in enumerate(row):
                if nome_colonna and nome_colonna.lower() in ["descrizione", "descrizione articolo"]:
                    column_index["Descrizione"] = index
                    break

        # Verifica se la colonna "Descrizione" è stata trovata
        if column_index["Descrizione"] is None:
            messagebox.showwarning("Errore", "Colonna 'Descrizione' non trovata nei file.")
            return

        # Itera sulle righe del foglio a partire dal secondo riga
        for row in sheet.iter_rows(min_row=2, values_only=True):
            codice_ean = row[column_index["Codice EAN"]] if "Codice EAN" in column_index else None
            codice = row[column_index["Codice"]] if "Codice" in column_index else None
            descrizione = row[column_index["Descrizione"]]
            prezzo = row[column_index[colonna_prezzo]] if colonna_prezzo in column_index else None


        # Verifica se la colonna "Descrizione" è stata trovata
        if column_index["Descrizione"] is None:
            messagebox.showwarning("Errore", "Colonna 'Descrizione' non trovata nei file.")
            return

        # Itera sulle righe del foglio
        for row in sheet.iter_rows(min_row=2, values_only=True):
            codice_ean = row[column_index["Codice EAN"]] if "Codice EAN" in column_index else None
            codice = row[column_index["Codice"]] if "Codice" in column_index else None
            descrizione = row[column_index["Descrizione"]]
            prezzo = row[column_index[colonna_prezzo]] if colonna_prezzo in column_index else None

            # Gestione dei valori non numerici nella colonna del prezzo
            if prezzo is not None:
                prezzo = float(prezzo)
            else:
                continue

            # Resto del codice...

            if codice_ean is not None and codice is not None and prezzo is not None:
                codice_ean = str(codice_ean)
                codice = str(codice)
                prezzo = float(prezzo)

                # Verifica se il prodotto è già presente nel risultato
                if (codice_ean, codice) in risultati_confronto:
                    # Aggiorna il prezzo minimo se necessario
                    if prezzo < risultati_confronto[(codice_ean, codice)]["Prezzo Minimo"]:
                        risultati_confronto[(codice_ean, codice)]["Prezzo Minimo"] = prezzo
                        risultati_confronto[(codice_ean, codice)]["Descrizione"] = descrizione
                else:
                    # Aggiungi il prodotto al risultato
                    risultati_confronto[(codice_ean, codice)] = {
                        "Descrizione": descrizione,
                        "Prezzo Minimo": prezzo
                    }

        progress_bar["value"] = i
        root.update()

    # Crea il nuovo file Excel con il prezzo minimo e il prodotto corrispondente
    nuovo_file = Workbook()
    nuovo_foglio = nuovo_file.active

    # Intestazioni delle colonne
    nuovo_foglio.append(["Codice EAN", "Codice", "Descrizione", "Prezzo Minimo"])

    # Popola il nuovo file con i dati
    for risultato in risultati_confronto.values():
        codice_ean = risultato["Codice EAN"]
        codice = risultato["Codice"]
        descrizione = risultato["Descrizione"]
        prezzo_minimo = risultato["Prezzo Minimo"]

        nuovo_foglio.append([codice_ean, codice, descrizione, prezzo_minimo])

    # Salva il nuovo file
    nuovo_file.save("C:/tmp/confronto.xlsx")


file_paths = []
column_index = {}

root = Tk()
root.title("Confronta File")
root.geometry("400x300")
root.resizable(False, False)

# Lista dei file selezionati
frame_files = Frame(root)
frame_files.pack(pady=10)

label_files = Label(frame_files, text="File Selezionati:")
label_files.pack()

listbox_files = Listbox(frame_files, width=50, height=5)
listbox_files.pack(side=LEFT, fill=Y)

scrollbar_files = Scrollbar(frame_files, orient=VERTICAL)
scrollbar_files.pack(side=RIGHT, fill=Y)

listbox_files.config(yscrollcommand=scrollbar_files.set)
scrollbar_files.config(command=listbox_files.yview)

frame_buttons = Frame(root)
frame_buttons.pack(pady=10)

button_seleziona = Button(frame_buttons, text="Seleziona File", command=seleziona_file)
button_seleziona.pack(side=LEFT, padx=5)

button_rimuovi = Button(frame_buttons, text="Rimuovi File", command=rimuovi_file)
button_rimuovi.pack(side=LEFT, padx=5)

button_confronta = Button(frame_buttons, text="Confronta", command=confronta_valori)
button_confronta.pack(side=LEFT, padx=5)

# Progress bar
progress_label = StringVar()
progress_label.set("")
progress_label_text = Label(root, textvariable=progress_label)
progress_label_text.pack(pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack()

root.mainloop()

