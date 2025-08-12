import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re

def estrai_dati(hubsdif):
    pattern = r"(\w+): (\d+) giorni"
    return re.findall(pattern, hubsdif)

def calcola_media_giorni(file_path, sheet_name=0, output_file="output.xlsx"):
    # Carica il file XLSX
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Assicurati che la colonna 'hubsdif' sia presente
    if 'Differenza Giorni' not in df.columns:
        raise ValueError("La colonna 'hubsdif' non è presente nel file.")
    
    # Estrazione delle città e giorni
    dati_estratti = []
    for entry in df['Differenza Giorni'].dropna():
        dati_estratti.extend(estrai_dati(entry))
    
    # Creazione del DataFrame con città e giorni separati
    df_estratto = pd.DataFrame(dati_estratti, columns=['Città', 'Giorni'])
    df_estratto['Giorni'] = pd.to_numeric(df_estratto['Giorni'], errors='coerce')
    
    # Raggruppamento per città e calcolo della media
    risultato = df_estratto.groupby('Città')['Giorni'].mean().reset_index()
    
    # Ordina il risultato
    risultato = risultato.sort_values(by='Giorni', ascending=False)
    
    # Salva il file di output
    risultato.to_excel(output_file, index=False)
    print(f"Risultato salvato in {output_file}")

def seleziona_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Esempio di utilizzo
file_path = seleziona_file()
if file_path:
    calcola_media_giorni(file_path)