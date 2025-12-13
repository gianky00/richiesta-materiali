# 📋 RDA Viewer - Sistema di Gestione Richieste di Acquisto

Sistema completo per la gestione automatizzata delle Richieste di Acquisto (RDA), con interfaccia grafica moderna e bot di automazione per l'elaborazione delle email.

![Versione](https://img.shields.io/badge/versione-2.0.0-blue)
![Platform](https://img.shields.io/badge/piattaforma-Windows-green)
![Python](https://img.shields.io/badge/python-3.8+-yellow)

---

## 📌 Caratteristiche Principali

### Interfaccia Grafica (GUI)
- ✅ **Tema chiaro moderno** - Design pulito e professionale
- ✅ **Dashboard interattiva** - Panoramica rapida delle RDA
- ✅ **Statistiche dettagliate** - Analisi dei dati con grafici
- ✅ **RDA Scadute** - Visualizzazione immediata delle richieste in ritardo
- ✅ **Ricerca avanzata** - Filtri multipli per trovare le RDA
- ✅ **Esportazione CSV** - Export dei dati per analisi esterne
- ✅ **Auto-sincronizzazione** - Aggiornamento automatico all'avvio
- ✅ **Menu contestuale** - Apertura rapida dei PDF allegati

### Bot di Automazione
- ✅ **Scansione email Outlook** - Elaborazione automatica degli allegati
- ✅ **Parsing PDF intelligente** - Estrazione dati dalle RDA
- ✅ **Archiviazione PDF** - Salvataggio organizzato dei documenti
- ✅ **Aggiornamento Excel** - Sincronizzazione con il registro esistente
- ✅ **Alert automatici** - Calcolo livelli di urgenza
- ✅ **Email di riepilogo** - Notifica RDA scadute

---

## 🏗️ Architettura

```
richiesta materiali/
├── main_gui.py          # Applicazione GUI principale
├── main_bot.py          # Bot di automazione
├── run_sync.py          # Sincronizzazione manuale
├── build_exe.py         # Script creazione EXE
├── requirements.txt     # Dipendenze Python
├── src/
│   ├── config.py        # Configurazione centralizzata
│   ├── database.py      # Gestione SQLite
│   ├── excel_manager.py # Automazione Excel
│   ├── email_scanner.py # Scansione Outlook
│   ├── pdf_parser.py    # Parsing documenti PDF
│   └── utils.py         # Funzioni di utilità
├── DATABASE/
│   ├── database_RDA.xlsm # Registro Excel
│   └── database_RDA.db   # Database SQLite
└── RDA_PDF/             # Archivio PDF
```

---

## 🚀 Installazione

### Metodo 1: Eseguibile Standalone (Consigliato)

**Non richiede Python installato!**

1. Scarica la cartella `dist` contenente gli eseguibili
2. Esegui `RDA_Viewer.exe` per aprire l'interfaccia
3. Schedula `RDA_Bot.exe` con Task Scheduler per l'automazione

### Metodo 2: Da Sorgente (Per Sviluppatori)

1. **Clona/Scarica il progetto**

2. **Installa Python 3.8+** (se non presente)
   - Scarica da [python.org](https://www.python.org/downloads/)

3. **Installa le dipendenze**
   ```powershell
   pip install -r requirements.txt
   ```

4. **Avvia l'applicazione**
   ```powershell
   python main_gui.py
   ```

### Creazione Eseguibile

Per creare gli EXE standalone:

```powershell
python build_exe.py
```

Gli eseguibili verranno creati nella cartella `dist/`.

---

## ⚙️ Configurazione

Modifica il file `src/config.py` per personalizzare:

```python
# Percorsi
NETWORK_BASE_PATH = r"\\server\Condivisa\RICHIESTE MATERIALI"

# Impostazioni Outlook
TARGET_FOLDER_NAME = "MAGO"
SENDER_EMAIL = "magonet@coemi.it"
DAYS_TO_CHECK = 60

# Email Alert
EMAIL_RECIPIENT = "destinatario@esempio.it"
EMAIL_SUBJECT = "RIEPILOGO RDA SCADUTE"
```

---

## 📖 Guida all'Uso

### Interfaccia GUI

#### Tab Dati RDA
- **Ricerca**: Digita nella barra di ricerca per filtrare in tempo reale
- **Ordinamento**: Clicca sulle intestazioni colonne per ordinare
- **PDF**: Doppio click o tasto destro → "Apri PDF"
- **Copia**: Tasto destro → "Copia riga" per copiare negli appunti

#### Tab Dashboard
Mostra una panoramica con:
- Totale RDA
- Articoli totali
- RDA scadute
- Top richiedenti

#### Tab Statistiche
Analisi dettagliate:
- Distribuzione livelli alert
- Top 10 materiali richiesti
- Distribuzione APF

#### Tab RDA Scadute
Lista delle RDA con alert attivo, colorate per urgenza:
- 🔴 **Rosso**: Alert alto (≥10 settimane)
- 🟡 **Giallo**: Alert medio (5-9 settimane)
- 🔵 **Blu**: Alert basso (1-4 settimane)

#### Tab Ricerca Avanzata
Filtri disponibili:
- Numero RDA
- Richiedente
- Range date
- APF
- Solo scadute

Possibilità di esportare i risultati in CSV.

### Bot di Automazione

Il bot può essere eseguito:

1. **Manualmente**:
   ```powershell
   python main_bot.py
   ```

2. **Schedulato** con Windows Task Scheduler:
   - Programma: `RDA_Bot.exe` (o `pythonw.exe main_bot.py`)
   - Trigger: Giornaliero alle 08:00

---

## 📊 Formato Dati

### Colonne Database

| Colonna | Descrizione |
|---------|-------------|
| N° RDA | Numero identificativo (es. 25/01812) |
| Articolo | Codice articolo/commessa |
| Descrizione | Descrizione del materiale |
| UM | Unità di misura |
| Quantità | Quantità richiesta |
| APF | Codice APF |
| Data RDA | Data della richiesta |
| Data Consegna | Data prevista consegna |
| Alert | Livello di urgenza (settimane) |
| Richiedente | Nome del richiedente |

### Formattazione Numeri

- I numeri interi vengono mostrati senza decimali (2.0 → 2)
- I codici testuali come "25/039" rimangono invariati
- I decimali usano la virgola come separatore italiano

---

## 🔧 Risoluzione Problemi

### "Impossibile connettersi al database"
- Verificare che il percorso di rete sia accessibile
- Controllare permessi di lettura/scrittura

### "Outlook non disponibile"
- Verificare che Microsoft Outlook sia installato e configurato
- Eseguire Outlook almeno una volta prima del bot

### "Excel non si apre"
- Verificare che Microsoft Excel sia installato
- Controllare che il file .xlsm non sia già aperto

### Applicazione lenta
- La prima sincronizzazione può richiedere tempo
- Verificare la connessione di rete

---

## 📝 Changelog

### Versione 2.0.0
- 🆕 Interfaccia completamente ridisegnata con tema chiaro
- 🆕 Dashboard con statistiche
- 🆕 Tab RDA scadute con colorazione
- 🆕 Ricerca avanzata con filtri multipli
- 🆕 Esportazione CSV
- 🆕 Auto-sincronizzazione all'avvio
- 🔧 Migliorata formattazione numeri
- 🔧 Separazione completa GUI/Bot
- 🔧 Supporto eseguibile standalone

### Versione 1.0.0
- Versione iniziale

---

## 👥 Supporto

Per problemi o richieste:
- Aprire una issue nel repository
- Contattare l'amministratore di sistema

---

## 📄 Licenza

Uso interno aziendale. Tutti i diritti riservati.

---

*Sviluppato con ❤️ per semplificare la gestione delle Richieste di Acquisto*
