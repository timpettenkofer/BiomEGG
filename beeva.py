import os
import pandas as pd
import openpyxl
from tkinter import Tk, filedialog, messagebox


def select_file(prompt):
    """Dateiauswahl für CSV, XLS oder XLSX."""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=prompt,
        filetypes=[("Excel- oder CSV-Dateien", "*.xlsx *.xls *.csv"), ("Alle Dateien", "*.*")]
    )
    root.destroy()
    return file_path


def save_file(prompt):
    """Zielpfad für Ausgabedatei wählen."""
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        title=prompt,
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")]
    )
    root.destroy()
    return file_path


def show_message(title, message):
    """Einfaches Popup anzeigen."""
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def read_input_file(file_path, sheet_name=None):
    """CSV oder Excel einlesen."""
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == ".csv":
        return pd.read_csv(file_path, encoding="utf-8", sep=";")
    elif ext in [".xlsx", ".xls"]:
        engine = "openpyxl" if ext == ".xlsx" else None
        return pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
    else:
        raise ValueError(f"Nicht unterstütztes Dateiformat: {ext}")


try:
    # Dateiauswahl
    mastr_file = select_file("Bitte die Datei mit dem Auszug aus dem Marktstammdatenregister auswählen.")
    eeg_file = select_file("Bitte die Datei mit den Zuschlägen der EEG-Biomasseausschreibung auswählen.")

    # Einlesen der Dateien
    mastr_df = read_input_file(mastr_file)
    eeg_sheet2_df = read_input_file(eeg_file, sheet_name=1)
    eeg_sheet3_df = read_input_file(eeg_file, sheet_name=2)

    # Zuschlagsnummern extrahieren
    def extract_numbers(cell):
        if pd.isna(cell):
            return []
        numbers = cell.replace(", ", "; ").split("; ")
        return [num for num in numbers if num.startswith("EEG") or num.startswith("SEE")]

    eeg_sheet3_df['Nummern'] = eeg_sheet3_df.iloc[:, 1].apply(extract_numbers)
    eeg_sheet3_df['Zuschlagsnummer'] = eeg_sheet3_df.iloc[:, 0]
    eeg_numbers = set(num for sublist in eeg_sheet3_df['Nummern'] for num in sublist)

    def find_matching_zuschlagsnummer(row):
        mastr_numbers = {row['MaStR-Nr. der EEG-Anlage'], row['MaStR-Nr. der Einheit']}
        for num in mastr_numbers:
            if num in eeg_numbers:
                zuschlagsnummer = eeg_sheet3_df.loc[
                    eeg_sheet3_df['Nummern'].apply(lambda x: num in x), 'Zuschlagsnummer'].values
                return zuschlagsnummer[0] if len(zuschlagsnummer) > 0 else None
        return None

    # Übereinstimmungen suchen
    mastr_df['Zuschlagsnummer'] = mastr_df.apply(find_matching_zuschlagsnummer, axis=1)
    matching_rows = mastr_df.dropna(subset=['Zuschlagsnummer'])

    # Falls keine Übereinstimmungen
    if matching_rows.empty:
        show_message("Hinweis", "Keine Übereinstimmungen gefunden. Es wurde keine Datei erstellt.")
    else:
        # Dateispeicherort wählen
        output_file = save_file("Bitte den Speicherort für die Ausgabedatei angeben.")

        # Spaltenreihenfolge festlegen
        columns_order = ['Zuschlagsnummer', 'MaStR-Nr. der EEG-Anlage', 'MaStR-Nr. der Einheit'] + \
                        [col for col in matching_rows.columns if col not in [
                            'Zuschlagsnummer', 'MaStR-Nr. der EEG-Anlage', 'MaStR-Nr. der Einheit',
                            'Gemarkung', 'Flurstück', 'Gemeindeschlüssel', 'Anzahl der Solar-Module',
                            'Hauptausrichtung der Solar-Module', 'Name des Windparks',
                            'Nabenhöhe der Windenergieanlage', 'Rotordurchmesser der Windenergieanlage',
                            'Hersteller der Windenergieanlage', 'Typenbezeichnung',
                            'Nutzbare Speicherkapazität in kWh', 'Lage der Einheit']]

        # Speichern
        final_df = matching_rows[columns_order]
        final_df.to_excel(output_file, index=False)
        show_message("Erfolg", "Die Auswertung wurde erfolgreich durchgeführt und die Datei wurde gespeichert.")

except Exception as e:
    show_message("Fehler", f"Ein Fehler ist aufgetreten:\n\n{e}")
