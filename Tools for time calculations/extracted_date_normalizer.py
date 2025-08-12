import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def correggi_mese(mese):
    mesi_corretti = {
        "gennaro": 1, "genaro": 1, "gennio": 1, "gennaio": 1, "gen": 1,
        "frebbraio": 2, "febraio": 2, "febbraio": 2, "feb": 2,
        "marso": 3, "marzo": 3, "mar": 3,
        "aprile": 4, "apr": 4, "april": 4,
        "maggio": 5, "mag": 5,
        "guigno": 6, "giunio": 6, "giugno": 6, "giu": 6, "june": 6,
        "luglio": 7, "lug": 7, "july": 7,
        "agosto": 8, "ago": 8,
        "settempre": 9, "settembre": 9, "set": 9,
        "ottobre": 10, "ott": 10,
        "novembre": 11, "nov": 11,
        "dicembre": 12, "decembre": 12, "dic": 12
    }
    return mesi_corretti.get(mese.lower())

def correggi_citta(citta):
    correzioni = {
        "Versaglies": "Versailles"
    }
    return correzioni.get(citta.capitalize(), citta.capitalize())

def estrai_date_citta(testo):
    pattern = r"([A-ZÀ-Ž][a-zà-ž]+(?:\s+[A-Z][a-zà-ž]+)*)\s+(\d{1,2})[.\s]?\s*([a-zA-Z]+)"
    matches = re.findall(pattern, testo)
    risultati = []
    
    for citta, giorno, mese in matches:
        mese_numerico = correggi_mese(mese)
        if mese_numerico:
            citta_corretta = correggi_citta(citta)
            try:
                risultati.append((citta_corretta, int(giorno), mese_numerico))
            except ValueError:
                print(f"Errore nella conversione della data: {giorno}-{mese}")
    
    return risultati

def calcola_differenze(data_riferimento, testo_estrazione):
    date_estratte = estrai_date_citta(testo_estrazione)
    risultati = []
    
    for citta, giorno, mese in date_estratte:
        anno = data_riferimento.year
        
        try:
            data_estratta = datetime(anno, mese, giorno)
            if data_estratta > data_riferimento:
                data_estratta = data_estratta.replace(year=anno - 1)

            differenza = (data_riferimento - data_estratta).days
            risultati.append(f"{citta}: {differenza} giorni")
        except ValueError:
            print(f"Errore nel creare la data per: {citta} {giorno}-{mese}-{anno}")
    
    return "; ".join(risultati) if risultati else None

def processa_file(file_path):
    df = pd.read_excel(file_path, engine="openpyxl")
    differenze = []
    
    for _, row in df.iterrows():
        try:
            data_riferimento = pd.to_datetime(row["date_A"], format="%Y-%m-%d", errors="coerce")
            if pd.isna(data_riferimento):
                differenze.append("Data non valida")
            else:
                differenze.append(calcola_differenze(data_riferimento, str(row["date_estratte"])))
        except Exception as e:
            differenze.append(f"Errore: {str(e)}")
    
    df["Differenza Giorni"] = differenze
    output_path = file_path.replace(".xlsx", "_output.xlsx")
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Elaborazione completata. File salvato come {output_path}")

def carica_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Seleziona il file Excel", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        processa_file(file_path)

carica_file()
