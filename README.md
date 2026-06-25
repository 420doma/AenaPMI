# 🛫 Mallorca Airport Stats

Dashboard locale per le statistiche passeggeri dell'aeroporto di Palma de Mallorca (PMI),
con dati ufficiali da [aena.es](https://www.aena.es/es/estadisticas/).

---

## Struttura del progetto

```
mallorca-stats/
├── scraper.py          ← scarica e analizza i file XLS da Aena
├── app.py              ← server Flask (dashboard web)
├── templates/
│   └── index.html      ← dashboard con Chart.js
├── data/               ← creata automaticamente
│   ├── passengers.json ← dati estratti (cache)
│   ├── scraper.log     ← log delle esecuzioni
│   └── xls/            ← file XLS scaricati (cache)
└── requirements.txt
```

---

## Installazione su Raspberry Pi 5

### 1. Installa le dipendenze Python

```bash
cd ~/mallorca-stats

# Crea un virtual environment (consigliato)
python3 -m venv .venv
source .venv/bin/activate

# Installa le librerie
pip install -r requirements.txt
```

### 2. Prima esecuzione — scarica i dati

```bash
# Anni 2019 → oggi (può richiedere 5-10 minuti)
python3 scraper.py

# Solo anni specifici
python3 scraper.py --years 2022-2026

# Forza il re-download di tutto
python3 scraper.py --force
```

### 3. Avvia il server

```bash
python3 app.py
```

Apri il browser su: **http://localhost:5000**
Dal telefono via Tailscale: **http://100.x.x.x:5000**

---

## Aggiornamento automatico (cron)

Per aggiornare i dati automaticamente ogni mese:

```bash
crontab -e
```

Aggiungi questa riga (esegui il 5 di ogni mese alle 08:00):
```
0 8 5 * * cd /home/pi/mallorca-stats && /home/pi/mallorca-stats/.venv/bin/python3 scraper.py >> /home/pi/mallorca-stats/data/cron.log 2>&1
```

---

## Eseguire come servizio systemd

Crea il file `/etc/systemd/system/mallorca-stats.service`:

```ini
[Unit]
Description=Mallorca Airport Stats Dashboard
After=network.target

[Service]
Type=simple
User=pi
WorkingDirectory=/home/pi/mallorca-stats
ExecStart=/home/pi/mallorca-stats/.venv/bin/python3 app.py
Restart=on-failure
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Attiva il servizio:
```bash
sudo systemctl daemon-reload
sudo systemctl enable mallorca-stats
sudo systemctl start mallorca-stats
sudo systemctl status mallorca-stats
```

---

## Note tecniche

- **Formato file**: Aena pubblica file `.xls` (vecchio formato OLE2).
  Il parser prova prima `xlrd` (vecchio xls) poi `openpyxl` (xlsx) come fallback.
- **Identificazione Palma**: cerca le parole chiave `PALMA`, `PMI`, `LEPA`, `MALLORCA`
  nella riga del file Excel.
- **Cache**: i file XLS sono salvati in `data/xls/` — non vengono riscaricati
  a meno che non si usi `--force` o sia l'anno corrente.
- **Dati disponibili**: da 2019 a oggi. Modificare `--years` in `scraper.py`
  per scaricare anni precedenti (disponibili dal 2004 su Aena).

---

## Troubleshooting

### "parse fallito" per qualche mese
Il formato XLS di Aena può variare negli anni. Controlla il log:
```bash
cat data/scraper.log | grep "⚠️"
```
E apri manualmente il file `data/xls/ANNO_MESE.xls` con LibreOffice
per vedere la struttura e adattare la funzione `_first_large_number()`.

### Port 5000 già in uso
```bash
python3 app.py  # oppure
FLASK_RUN_PORT=8080 flask --app app run --host=0.0.0.0
```

### Errore di rete durante lo scraping
Aena può restituire errori temporanei. Riprova — il sistema ha un retry automatico
di 3 tentativi per ogni file.
