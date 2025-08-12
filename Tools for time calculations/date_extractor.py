import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

def correggi_mese(mese):
    mesi_corretti = {
        "gennaro": "Gennaio", "genaro": "Gennaio", "gennio": "Gennaio",
        "frebbraio": "Febbraio", "febraio": "Febbraio",
        "marso": "Marzo", "mar": "Marzo",
        "aprile": "Aprile", "apr": "Aprile",
        "maggio": "Maggio",
        "guigno": "Giugno", "giunio": "Giugno", "giugno": "Giugno", "giu": "Giugno",
        "luglio": "Luglio", "lug": "Luglio",
        "agosto": "Agosto", "ago": "Agosto",
        "settempre": "Settembre", "settembre": "Settembre", "set": "Settembre",
        "ottobre": "Ottobre", "ott": "Ottobre",
        "novembre": "Novembre", "nov": "Novembre",
        "dicembre": "Dicembre", "Decembre": "Dicembre", "decembre": "Dicembre", "dic": "Dicembre",
        "april": "Aprile", "june": "Giugno", "july": "Luglio"  
    }
    return mesi_corretti.get(mese.lower(), mese.capitalize())

def estrai_date_testo(testo):    
    pattern = r"([A-ZÀ-Ž]+)\s+(\d{1,2})[.\s]?\s*([a-zA-Z]+)"
    date_estratte = []

    for match in re.findall(pattern, testo):
        citta, giorno, mese = match
        mese_corretto = correggi_mese(mese)
        date_estratte.append(f"{citta.capitalize()} {giorno} {mese_corretto}")
    
    return ", ".join(date_estratte)

def carica_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Seleziona il file Excel", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        processa_file(file_path)

def processa_file(file_path):
    df = pd.read_excel(file_path, engine="openpyxl")  
    df["date_estratte"] = df["transcription"].astype(str).apply(estrai_date_testo)
    output_path = file_path.replace(".xlsx", "_output.xlsx")
    df.to_excel(output_path, index=False, engine="openpyxl")  
    print(f"Elaborazione completata. File salvato come {output_path}")

carica_file()
