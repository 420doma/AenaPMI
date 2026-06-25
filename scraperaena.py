#!/usr/bin/env python3
"""
Aena Mallorca Passenger Scraper
Scarica i rapporti XLS/XLSX mensili da Aena ed estrae i dati passeggeri
dell'aeroporto di Palma de Mallorca (PMI).

Uso: python3 scraper.py [--force] [--years 2019-2026]
"""

import argparse
import json
import time
import re
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

DATA_DIR  = Path("data")
XLS_CACHE = DATA_DIR / "xls"
DATA_FILE = DATA_DIR / "passengers.json"
LOG_FILE  = DATA_DIR / "scraper.log"

DATA_DIR.mkdir(exist_ok=True)
XLS_CACHE.mkdir(exist_ok=True)

# ─── Costanti ─────────────────────────────────────────────────────────────────

BASE_URL   = "https://www.aena.es"
MONTHS_URL = "https://www.aena.es/es/estadisticas/informes-mensuales.html"

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":          "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,it;q=0.7",
    "Referer":         MONTHS_URL,
}

MONTHS_ES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,
    "mayo":5,"junio":6,"julio":7,"agosto":8,
    "septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
}

# Termini usati per identificare la riga di Palma nei file XLS
PALMA_KEYWORDS = ["PALMA", "PMI", "LEPA", "MALLORCA"]

# ─── HTTP ─────────────────────────────────────────────────────────────────────

def http_get(url, retries=3, delay=2.0):
    """Fetch URL with retry logic."""
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HTTP_HEADERS, timeout=30)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            log(f"    ⚠️  [{attempt+1}/{retries}] {e}")
            if attempt < retries - 1:
                time.sleep(delay)
    return None

# ─── HTML parsing ─────────────────────────────────────────────────────────────

def get_xls_links_for_year(year: int) -> dict:
    """
    Scrapa la pagina Aena per l'anno dato.
    Ritorna {mese_num: url_xls}.
    """
    r = http_get(f"{MONTHS_URL}?anio={year}")
    if not r:
        return {}

    soup = BeautifulSoup(r.text, "html.parser")

    # Tutti i link XLS con "blobwhere" in ordine di pagina
    xls_urls = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "blobwhere" not in href:
            continue
        label = a.get_text(strip=True).upper()
        if "XLS" not in label and "XLSX" not in label:
            continue
        if href.startswith("/"):
            href = BASE_URL + href
        xls_urls.append(href)

    if not xls_urls:
        return {}

    # Trova l'ordine dei mesi cercando "informe {mes}" nel testo della pagina
    full_text  = soup.get_text(separator=" ").lower()
    ordered_months = []
    search_from = 0

    for _ in range(12):
        best_pos, best_month = len(full_text), None
        for mes, num in MONTHS_ES.items():
            pos = full_text.find(f"informe {mes}", search_from)
            if 0 <= pos < best_pos:
                best_pos, best_month = pos, num
        if best_month is None:
            break
        ordered_months.append(best_month)
        search_from = best_pos + 1

    result = {}
    for i, month_num in enumerate(ordered_months):
        if i < len(xls_urls):
            result[month_num] = xls_urls[i]

    return result

# ─── XLS parsing ──────────────────────────────────────────────────────────────

def is_old_xls_format(path: Path) -> bool:
    """Controlla i magic bytes: OLE2 = vecchio .xls, PK = .xlsx."""
    with open(path, "rb") as f:
        return f.read(4) == b"\xd0\xcf\x11\xe0"

def _row_has_palma(cells_upper: list) -> bool:
    joined = " ".join(cells_upper)
    return any(kw in joined for kw in PALMA_KEYWORDS)

def _first_large_number(values, skip_first=True) -> int | None:
    """
    Ritorna il primo numero > 10.000 trovato nella lista.
    Gestisce in modo sicuro float nativi e stringhe formattate.
    """
    for i, val in enumerate(values):
        if skip_first and i == 0:
            continue
        if val is None:
            continue
        
        # Se è già un tipo numerico nativo, lo usiamo direttamente senza toccare i decimali
        if isinstance(val, (int, float)):
            n = int(val)
            if n > 10_000:
                return n
            continue

        try:
            s = str(val).strip().replace("\xa0", "").replace(" ", "")
            if not s:
                continue
            
            # Proviamo prima la conversione pulita standard
            try:
                n = int(float(s))
                if n > 10_000:
                    return n
                continue
            except ValueError:
                pass

            # Rilevamento e pulizia di formati stringa complessi (es: 1.234.567 o 1.234.567,00)
            if "," in s and "." in s:
                if s.rfind(",") > s.rfind("."):  # Formato europeo
                    s = s.replace(".", "").replace(",", ".")
                else:  # Formato US
                    s = s.replace(",", "")
            elif "," in s:
                if re.match(r"^\d{1,3}(,\d{3})+$", s):
                    s = s.replace(",", "")
                else:
                    s = s.replace(",", ".")
            elif "." in s:
                if re.match(r"^\d{1,3}(\.\d{3})+$", s):
                    s = s.replace(".", "")
            
            n = int(float(s))
            if n > 10_000:
                return n
        except (ValueError, TypeError):
            continue
    return None

def parse_with_openpyxl(path: Path) -> int | None:
    try:
        import openpyxl
        # Passiamo il file come stream binario per bypassare i controlli sulla stringa dell'estensione
        with open(path, "rb") as f:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            try:
                for ws in wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        cells_up = [str(c or "").upper().strip() for c in row]
                        if _row_has_palma(cells_up):
                            return _first_large_number(list(row))
            finally:
                wb.close()
    except Exception as e:
        log(f"    openpyxl error: {e}")
    return None

def parse_with_xlrd(path: Path) -> int | None:
    try:
        import xlrd
        wb = xlrd.open_workbook(str(path), on_demand=True)
        try:
            for ws in wb.sheets():
                for rx in range(ws.nrows):
                    row = ws.row(rx)
                    cells_up = [str(c.value).upper().strip() for c in row]
                    if _row_has_palma(cells_up):
                        numeric_vals = []
                        for c in row:
                            if c.ctype == xlrd.XL_CELL_NUMBER:
                                numeric_vals.append(c.value)
                            elif c.ctype == xlrd.XL_CELL_TEXT:
                                numeric_vals.append(c.value)
                            else:
                                numeric_vals.append(None)
                        return _first_large_number(numeric_vals)
        finally:
            wb.release_resources()
    except Exception as e:
        log(f"    xlrd error: {e}")
    return None

def extract_passengers(path: Path) -> int | None:
    """Prova entrambi i parser in base al formato reale del file."""
    if is_old_xls_format(path):
        pax = parse_with_xlrd(path)
        if pax is None:  # fallback
            pax = parse_with_openpyxl(path)
    else:
        pax = parse_with_openpyxl(path)
        if pax is None:  # fallback
            pax = parse_with_xlrd(path)
    return pax

# ─── Scraping principale ──────────────────────────────────────────────────────

def scrape_year(year: int, force: bool = False) -> dict:
    current_year = datetime.now().year
    is_current   = (year == current_year)

    log(f"\n📅  {year}")
    links = get_xls_links_for_year(year)

    if not links:
        log("    Nessun link trovato")
        return {}

    log(f"    Trovati {len(links)} mesi")
    results = {}

    for month in sorted(links.keys()):
        cache_path_xls = XLS_CACHE / f"{year}_{month:02d}.xls"
        cache_path_xlsx = XLS_CACHE / f"{year}_{month:02d}.xlsx"
        
        # Cerca se esiste già una delle due estensioni in cache
        cache_path = None
        if cache_path_xls.exists():
            cache_path = cache_path_xls
        elif cache_path_xlsx.exists():
            cache_path = cache_path_xlsx

        # Ri-scarica se forzato, mancante o anno corrente
        needs_download = (cache_path is None) or force or is_current
        if needs_download:
            r = http_get(links[month])
            if r:
                # Sceglie l'estensione corretta analizzando l'header del file scaricato
                if r.content[:4] == b"\x50\x4b\x03\x04":  # Header ZIP -> XLSX
                    chosen_path = cache_path_xlsx
                    other_path = cache_path_xls
                else:
                    chosen_path = cache_path_xls
                    other_path = cache_path_xlsx
                
                # Rimuove il file dell'estensione opposta per evitare duplicati sporchi
                if other_path.exists():
                    other_path.unlink()
                    
                chosen_path.write_bytes(r.content)
                cache_path = chosen_path
                time.sleep(1.5)
            else:
                log(f"    ❌  {month:02d}: download fallito")
                continue

        pax = extract_passengers(cache_path)
        if pax is not None:
            results[month] = pax
            log(f"    ✅  {month:02d}: {pax:>12,} passeggeri")
        else:
            log(f"    ⚠️   {month:02d}: parsing fallito — controlla {cache_path}")

    return results

# ─── Utils ────────────────────────────────────────────────────────────────────

def log(msg: str):
    print(msg)
    with open(LOG_FILE, "a") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")

def load_existing_data() -> dict:
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text())
        except json.JSONDecodeError:
            return {}
    return {}

# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Aena Mallorca Scraper")
    parser.add_argument("--force", action="store_true",
                        help="Ri-scarica tutti i file anche se già in cache")
    parser.add_argument("--years", type=str, default="2019-",
                        help="Range anni, es: 2019-2026 o 2019- (default: 2019-oggi)")
    args = parser.parse_args()

    current_year = datetime.now().year
    match = re.match(r"^(\d{4})-(\d{4})?$", args.years)
    if match:
        start = int(match.group(1))
        end   = int(match.group(2)) if match.group(2) else current_year
    else:
        start, end = 2019, current_year

    years = range(start, end + 1)

    log(f"\n{'='*50}")
    log(f"🛫  Aena Mallorca Scraper — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    log(f"    Anni: {start}–{end} | Force: {args.force}")
    log(f"{'='*50}")

    data = load_existing_data()

    for year in years:
        result = scrape_year(year, force=args.force)
        if result:
            data[str(year)] = {str(k): v for k, v in result.items()}

    # Salva i dati strutturati finali
    DATA_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False))
    total_points = sum(len(v) for v in data.values())
    log(f"\n✅  Dati salvati: {DATA_FILE}")
    log(f"    Anni disponibili: {sorted(data.keys())}")
    log(f"    Punti dati totali: {total_points}")

if __name__ == "__main__":
    main()
