import os
import re
import json
import requests
import xlrd
import openpyxl
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime

BASE_URL = "https://www.aena.es"
STATS_URL = "https://www.aena.es/es/estadisticas/informes-mensuales.html"
DATA_DIR = "data"
XLS_DIR = os.path.join(DATA_DIR, "xls")
JSON_PATH = os.path.join(DATA_DIR, "data.json")

def ensure_dirs():
    os.makedirs(XLS_DIR, exist_ok=True)

def _extract_arrivals(values) -> int | None:
    """
    Estrae il numero degli arrivi (Llegadas).
    Nei file Excel Aena, i numeri grandi per riga sono tipicamente in quest'ordine: 
    1. Pasajeros Totales, 2. Llegadas, 3. Salidas.
    Saltando il primo numero grande trovato, otteniamo automaticamente gli Arrivi.
    """
    first_found = False
    for v in values:
        if isinstance(v, (int, float)) and v > 10000:
            if not first_found:
                first_found = True
                continue # Salta il Totale
            return int(v) # Ritorna gli Arrivi (Llegadas)
    return None

def parse_with_xlrd(file_path: str) -> int | None:
    try:
        wb = xlrd.open_workbook(file_path)
        for sheet in wb.sheets():
            # SALVATAGGIO CRITICO: Ignora i fogli "Acumulado" per evitare di prendere i dati annuali sommati
            if "ACUM" in sheet.name.upper():
                continue
                
            for row_idx in range(sheet.nrows):
                row_values = sheet.row_values(row_idx)
                if any(isinstance(cell, str) and ("PALMA DE MALLORCA" in cell.upper() or cell.upper() == "PALMA") for cell in row_values):
                    return _extract_arrivals(row_values)
    except Exception as e:
        print(f"Errore xlrd su {file_path}: {e}")
    return None

def parse_with_openpyxl(file_path: str) -> int | None:
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheetname in wb.sheetnames:
            # SALVATAGGIO CRITICO: Ignora i fogli "Acumulado" per evitare di prendere i dati annuali sommati
            if "ACUM" in sheetname.upper():
                continue
                
            sheet = wb[sheetname]
            for row in sheet.iter_rows(values_only=True):
                if row and any(isinstance(cell, str) and ("PALMA DE MALLORCA" in cell.upper() or str(cell).upper() == "PALMA") for cell in row):
                    return _extract_arrivals(row)
    except Exception as e:
        print(f"Errore openpyxl su {file_path}: {e}")
    return None

def process_file(file_path: str) -> int | None:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xls':
        return parse_with_xlrd(file_path)
    elif ext == '.xlsx':
        return parse_with_openpyxl(file_path)
    return None

def scrape_and_update():
    ensure_dirs()
    
    print("Scaricando la pagina delle statistiche...")
    response = requests.get(STATS_URL)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')

    links = soup.find_all('a', href=True)
    xls_links = []
    
    # Trova tutti i link a file Excel
    for a in links:
        href = a['href']
        if href.lower().endswith('.xls') or href.lower().endswith('.xlsx'):
            full_url = urljoin(BASE_URL, href)
            xls_links.append(full_url)

    print(f"Trovati {len(xls_links)} file Excel. Estrazione dei mesi e anni...")
    
    # Carichiamo i dati esistenti se ci sono
    if os.path.exists(JSON_PATH):
        with open(JSON_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = {}

    pattern = re.compile(r'/(\d{4})_(\d{2})\.(xlsx|xls)', re.IGNORECASE)
    
    for url in xls_links:
        match = pattern.search(url)
        if match:
            year = match.group(1)
            month = match.group(2)
            ext = match.group(3)
            
            # Se l'anno non è nel dict, lo inizializziamo con 12 zeri
            if year not in data:
                data[year] = [0] * 12
                
            month_idx = int(month) - 1
            
            # Se abbiamo già un dato > 0 per quel mese/anno, saltiamo per evitare download inutili
            if data[year][month_idx] > 0:
                continue
                
            filename = f"{year}_{month}.{ext}"
            file_path = os.path.join(XLS_DIR, filename)
            
            # Scarica il file se non esiste
            if not os.path.exists(file_path):
                print(f"Scaricando {filename}...")
                try:
                    r = requests.get(url)
                    r.raise_for_status()
                    with open(file_path, 'wb') as f:
                        f.write(r.content)
                except Exception as e:
                    print(f"Errore durante il download di {url}: {e}")
                    continue

            # Estrai gli arrivi dal file scaricato
            print(f"Estrazione dati da {filename}...")
            arrivals = process_file(file_path)
            
            if arrivals:
                data[year][month_idx] = arrivals
                print(f"Trovato: {year}-{month} -> {arrivals} arrivi")
            else:
                print(f"Nessun dato trovato per Palma in {filename}")

    # Salva il JSON aggiornato
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    print("Dati salvati in data.json")

if __name__ == "__main__":
    scrape_and_update()
            
