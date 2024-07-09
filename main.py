import tkinter as tk
from tkinter import ttk  # Importiere ttk für Combobox
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
from datetime import datetime
import locale
import os


# Pfad zur Datei, die die laufende Nummer speichert
nummer_datei = "laufende_nummer.txt"
locale.setlocale(locale.LC_ALL, 'de_DE.utf8')

# Funktion, um die laufende Nummer aus der Datei zu laden oder die Datei zu erstellen
def lade_laufende_nummer():
    if os.path.exists(nummer_datei):
        with open(nummer_datei, "r") as file:
            return int(file.read().strip())
    else:
        # Erstelle die Datei mit der Zahl 1, falls sie nicht existiert
        with open(nummer_datei, "w") as file:
            file.write("1")
        return 1

# Funktion, um die laufende Nummer in der Datei zu speichern
def speichere_laufende_nummer(nummer):
    with open(nummer_datei, "w") as file:
        file.write(str(nummer))

# Funktion zum Laden der Identifier aus der Excel-Datei
def lade_identifier_aus_excel(excel_path):
    try:
        # Lese die Excel-Datei
        df = pd.read_excel(excel_path)

        # Überprüfe, ob die Spalte '<Name>' vorhanden ist
        if '<Name>' not in df.columns:
            messagebox.showerror("Fehler", "Die Spalte '<Name>' wurde in der Excel-Datei nicht gefunden.")
            print(f"Vorhandene Spalten: {df.columns.tolist()}")
            return []

        # Extrahiere eindeutige Werte aus der Spalte '<Name>'
        identifier_list = df['<Name>'].unique().tolist()

        return identifier_list

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Laden der Identifier: {str(e)}")
        print(e)
        return []

# Lade die laufende Nummer beim Start
laufende_nummer = lade_laufende_nummer()

# Funktion zum Ersetzen der Platzhalter und Speichern des Dokuments
def replace_and_save(doc_path, excel_path, identifier, output_path):
    global laufende_nummer
    try:
        # Lese die Excel-Datei und identifiziere die richtige Zeile
        df = pd.read_excel(excel_path)
        row = df[df['<Name>'] == identifier].iloc[0]

        # Erstelle ein Dictionary mit den Platzhaltern und ihren Ersatzwerten
        replacements = row.to_dict()
        replacements['<laufende Nummer>'] = laufende_nummer
        replacements['<Datum>'] = datetime.now().strftime("%d.%m.%Y")
        replacements['<Monat>'] = datetime.now().strftime("%B")

        # Öffne das Word-Dokument
        doc = Document(doc_path)

        # Ersetze die Platzhalter im Dokument
        for paragraph in doc.paragraphs:
            for placeholder, replacement in replacements.items():
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(replacement))

        # Speichere das geänderte Dokument
        doc.save(output_path)

        # Inkrementiere die laufende Nummer
        laufende_nummer += 1
        speichere_laufende_nummer(laufende_nummer)

        messagebox.showinfo("Erfolg", "Platzhalter wurden erfolgreich ersetzt und das Dokument wurde gespeichert.")

    except Exception as e:
        messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten: {str(e)}")

# Funktion zum Auswählen einer Datei über das Dateiauswahlfenster
def select_file(file_type):
    file_types = {
        "word": [("Word Dateien", "*.docx")],
        "excel": [("Excel Dateien", "*.xlsx")],
        "save": [("Word Dateien", "*.docx")]
    }
    return filedialog.askopenfilename(filetypes=file_types[file_type]) if file_type != "save" else filedialog.asksaveasfilename(defaultextension=".docx", filetypes=file_types[file_type])

# Funktion zum Ausführen der Platzhalterersetzung
def run_replacement():
    doc_path = "./files/Wordvorlage.docx"
    excel_path ="./files/Daten.xlsx"
    identifier = combobox_identifier.get()
    output_path = entry_output_path.get()

    if not doc_path or not excel_path or not identifier or not output_path:
        messagebox.showwarning("Warnung", "Bitte füllen Sie alle Felder aus.")
        return

    replace_and_save(doc_path, excel_path, identifier, output_path)

# Funktion zum Filtern der Optionen in der Combobox basierend auf dem eingegebenen Text
def filter_options(event):
    # Aktualisiere die Combobox-Werte basierend auf dem eingegebenen Text
    current_text = combobox_identifier.get()
    filtered_options = [option for option in identifier_options if current_text.lower() in option.lower()]
    combobox_identifier['values'] = filtered_options



def lade_dateipfad(save_file):
    if os.path.exists(save_file):
        with open(nummer_datei, "r") as file:
            return file.read().strip()
    else:
        # Erstelle die Datei mit der Zahl 1, falls sie nicht existiert
        with open(save_file, "w") as file:
            file.write("")
        return 1

# Erstelle das Hauptfenster der UI
root = tk.Tk()
root.title("Word Platzhalter Ersetzer")


# Erstelle die UI-Elemente

tk.Label(root, text="Name:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
identifier_options = lade_identifier_aus_excel("./files/Daten.xlsx")
combobox_identifier = ttk.Combobox(root, values=identifier_options, state="normal", width=47)
combobox_identifier.grid(row=2, column=1, padx=10, pady=10)
combobox_identifier.bind("<KeyRelease>", filter_options)

tk.Label(root, text="Speichern als:").grid(row=3, column=0, padx=10, pady=10, sticky='e')
entry_output_path = tk.Entry(root, width=50)
entry_output_path.grid(row=3, column=1, padx=10, pady=10)
tk.Button(root, text="Durchsuchen", command=lambda: entry_output_path.insert(0, select_file("save"))).grid(row=3, column=2, padx=10, pady=10)

tk.Button(root, text="Ersetzen", command=run_replacement).grid(row=4, column=1, padx=10, pady=20)

# Starte die Haupt-UI-Schleife
root.mainloop()
