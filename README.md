# DHL Spedizioni

Utility desktop per la creazione di documenti CSV doganali da caricare nei sistemi dello spedizioniere.

## Download

Vai alla sezione [Releases](../../releases) e scarica l'ultimo file per il tuo sistema operativo:

| Sistema | File |
|---------|------|
| Windows | `DHL.exe` |
| macOS   | `DHL-mac.dmg` |

Nessuna installazione richiesta — l'eseguibile include tutte le dipendenze.

---

## Utilizzo

### Prima esecuzione
All'avvio viene chiesto di selezionare una **cartella di lavoro**: qui vengono salvati l'anagrafica prodotti e tutti i documenti esportati.

### Visualizza Anagrafica
Apre `anagrafica_spedizioni.csv` in una tabella modificabile.

- I campi sono **bloccati** di default. Per modificare un valore clicca sulla cella e conferma la domanda di sicurezza.
- La modifica viene salvata su disco immediatamente, senza ulteriori conferme.
- Usa **Aggiungi prodotto** per inserire nuovi prodotti e **Elimina prodotto** per rimuovere la riga selezionata.

**Colonne anagrafica:** Rif. | Tipo | Descrizione | Cod. Doganale | U.M. | Altro Prezzo | Origine

### Crea Documento
Compone un nuovo documento di spedizione.

1. Inserisci il **nome file** (es. `2025-600`) — il file CSV verrà salvato automaticamente con questo nome.
2. Clicca **Aggiungi riga** per ogni prodotto da spedire.
3. Seleziona la **Descrizione** dal menu a tendina: tutti gli altri campi vengono pre-compilati dall'anagrafica.
4. Compila **Q.tà / Peso** e **Prezzo**; scegli la **Valuta** (EUR / USD / CHF, default EUR).
5. Il documento si **auto-salva** a ogni modifica (appena il nome file è presente).
6. Usa **Stampa PDF** per generare e aprire un PDF pronto per la stampa.

**Formato output CSV:** nessuna intestazione, delimitatore `;`, 11 colonne:
```
Rif.;Tipo;Descrizione;Cod. Doganale;Q.tà / Peso;U.M.;Prezzo;Valuta;Altro Prezzo;;Origine
```

---

## Sviluppo locale

```bash
# Dipendenze
pip install reportlab

# Avvio
python main.py
```

### Build eseguibili

```bash
pip install pyinstaller reportlab
pyinstaller dhl.spec
# Risultato: dist/DHL.exe  (Windows)  oppure  dist/DHL.app  (macOS)
```

### Rilascio su GitHub

```bash
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions compila automaticamente Windows `.exe` e macOS `.dmg` e li pubblica come assets della release.

---

## Struttura progetto

```
dhl/
├── main.py                    # Intera applicazione
├── anagrafica_spedizioni.csv  # Anagrafica prodotti (pre-compilata)
├── requirements.txt
├── dhl.spec                   # PyInstaller
└── .github/workflows/
    └── release.yml            # CI/CD build & release
```
