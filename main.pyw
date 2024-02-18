import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import messagebox
import os

# Funzione per caricare lo stato di voto da un file JSON specifico per il foglio
def load_vote_state(sheet_name):
    try:
        with open(f"{sheet_name}_vote_state.json", "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}

# Funzione per salvare lo stato di voto su un file JSON specifico per il foglio
def save_vote_state(vote_state, sheet_name):
    with open(f"{sheet_name}_vote_state.json", "w") as file:
        json.dump(vote_state, file)

# Funzione per ottenere i dati dal foglio di Google Sheets
def get_sheet_data(sheet_name):
    try:
        # Costruisci il percorso completo del file di credenziali
        credentials_path = os.path.join(os.path.dirname(__file__), "google_sheet_credentials.json")

        # Credenziali OAuth 2.0
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(creds)

        # Apri il foglio di Google Sheets utilizzando il nome specifico del foglio
        sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1JeLOdLMehwVYjFlscW88s17TQzYwuPgHEjNze8L1XZY/edit?usp=sharing").worksheet(sheet_name)

        # Ottieni tutti i valori dal foglio
        data = sheet.get_all_values()

        return sheet, data
    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante il recupero dei dati dal foglio Google Sheets: {str(e)}")
        return None, None

# Funzione per formattare il numero con la virgola anziché col punto
def format_number(value):
    return value.replace(".", ",")

# Funzione per salvare il voto di un singolo cantante sul foglio di Google Sheets
def save_vote(row, vote_state, sheet, sheet_name):
    try:
        # Ottieni l'ID del concorrente
        contestant_id = sheet.cell(row, 1).value

        # Controlla se il concorrente è già stato votato
        if vote_state.get(contestant_id, False):
            messagebox.showerror("Errore", "Questo concorrente è già stato votato.")
            return

        # Controlla se ci sono già voti per questo concorrente sul foglio
        contestant_votes = sheet.row_values(row)[1:5]  # Prendi i voti dalle colonne B, C, D, E
        if any(contestant_votes):
            messagebox.showerror("Errore", "Questo concorrente ha già ricevuto dei voti.")
            return

        # Ottieni i valori dalle caselle di testo corrispondenti al cantante
        entry_b = entries[row, 2]  # Colonna B
        entry_c = entries[row, 3]  # Colonna C
        entry_d = entries[row, 4]  # Colonna D
        entry_e = entries[row, 5]  # Colonna E

        # Verifica se sono stati inseriti voti
        if not entry_b.get() or not entry_c.get() or not entry_d.get() or not entry_e.get():
            messagebox.showerror("Errore", "Si prega di inserire tutti i voti prima di votare.")
            return

        # Validazione e formattazione dei voti
        for entry in [entry_b, entry_c, entry_d, entry_e]:
            vote = entry.get()
            # Se il voto contiene un punto, lo sostituiamo con una virgola
            if "." in vote:
                vote = format_number(vote)
                entry.delete(0, tk.END)  # Cancella il contenuto della casella di testo
                entry.insert(tk.END, vote)  # Inserisce il voto formattato nella casella di testo
            # Verifichiamo che il voto sia compreso tra 1 e 10
            vote = float(vote.replace(",", "."))  # Convertiamo il voto in float
            if vote < 1 or vote > 10:
                messagebox.showerror("Errore", "I voti devono essere compresi tra 1 e 10.")
                return

        # Aggiorna il foglio di Google Sheets con i voti del cantante
        sheet.update_cell(row, 2, entry_b.get())
        sheet.update_cell(row, 3, entry_c.get())
        sheet.update_cell(row, 4, entry_d.get())
        sheet.update_cell(row, 5, entry_e.get())

        # Imposta lo stato di voto per questo concorrente a True
        vote_state[contestant_id] = True

        # Salva lo stato di voto su file
        save_vote_state(vote_state, sheet_name)

        messagebox.showinfo("Successo", f"Voti per il concorrente {sheet.cell(row, 1).value} aggiornati con successo su Google Sheets!")
    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante l'aggiornamento dei voti su Google Sheets: {str(e)}")

# Connessione a Google Sheets e ottenimento dei dati
def connect_to_sheet():
    sheet_name = sheet_name_entry.get()
    if sheet_name:
        sheet, sheet_data = get_sheet_data(sheet_name)
        if sheet_data:
            global entries
            entries = {}
            vote_state = load_vote_state(sheet_name)  # Carica lo stato dei voti specifico per il foglio
            for row, data_row in enumerate(sheet_data, start=1):
                for col, value in enumerate(data_row, start=1):
                    if row == 1:  # Se è la riga delle etichette
                        if col == 1 or col == 6:  # Etichette per le parti non editabili
                            label = tk.Label(root, text=value, bg="lightgray", relief="ridge", width=entry_width)
                            label.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                        else:  # Etichette non editabili
                            label = tk.Label(root, text=value, bg="lightgray", relief="ridge", width=entry_width)
                            label.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                    else:  # Se non è la riga delle etichette
                        if col == 1 or col == 6:  # Etichette per le parti non editabili
                            label = tk.Label(root, text=value, bg="lightgray", relief="ridge", width=entry_width)
                            label.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                        elif col >= 2 and col <= 5:  # Caselle di testo per le colonne B, C, D, E
                            entry = tk.Entry(root, width=entry_width)  
                            entry.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                            entry.insert(tk.END, value)
                            entries[(row, col)] = entry
                        else:  # Caselle di testo non editabili
                            entry = tk.Entry(root, width=entry_width, state='readonly')
                            entry.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                            entry.insert(tk.END, value)
                            entries[(row, col)] = entry

            # Creazione dei pulsanti VOTA
            for row, data_row in enumerate(sheet_data, start=1):
                if row != 1:  # Se non è la riga delle etichette
                    vote_button = tk.Button(root, text="Vota", command=lambda row=row: save_vote(row, vote_state, sheet, sheet_name))
                    vote_button.grid(row=row, column=len(sheet_data[0]) + 1, padx=(0, 5), pady=5, sticky="nsew")

            # Aggiungi peso alle righe e alle colonne per uniformare la dimensione degli elementi
            for i in range(len(sheet_data) + 2):
                root.grid_rowconfigure(i, weight=1)
            for i in range(len(sheet_data[0]) + 1):
                root.grid_columnconfigure(i, weight=1)
        else:
            messagebox.showerror("Errore", "Nome del foglio non valido o foglio vuoto.")
    else:
        messagebox.showerror("Errore", "Si prega di inserire il nome del foglio.")

root = tk.Tk()
root.title("Modifica Foglio Excel")

# Uniforma le dimensioni delle caselle di testo
entry_width = 15
entry_height = 2

# Etichetta e casella di testo per il nome del foglio
sheet_name_label = tk.Label(root, text="Nome Giudice:")
sheet_name_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
sheet_name_entry = tk.Entry(root, width=entry_width)
sheet_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

# Pulsante per connettersi al foglio di lavoro
connect_button = tk.Button(root, text="Connetti", command=connect_to_sheet)
connect_button.grid(row=0, column=2, padx=5, pady=5)

# Creazione di un dizionario per memorizzare le caselle di testo
entries = {}

root.mainloop()
