# CRManager

Applicazione locale per gestire Change Request SAP organizzate per cliente e progetto.

## Contenuto repository

- `SAPCRManager/`
  - app Flask principale
  - dashboard Kanban per stato trasporto
  - export Excel della lista CR
- `avvia_sap_cr_manager.bat`
  - launcher Windows con doppio click

## Funzionalita principali

- Struttura `Cliente -> Progetto -> CR`
- Stati CR:
  - `Sviluppo`
  - `Quality`
  - `Produzione`
- Colori per stato trasporto
- Board Kanban globale e board Kanban per progetto
- Drag and drop per cambiare stato
- Export Excel filtrato
- Salvataggio locale su file JSON

## Avvio rapido Windows

Fai doppio click su:

```bat
avvia_sap_cr_manager.bat
```

Il launcher:

- controlla che `py` sia disponibile
- installa le dipendenze se necessario
- avvia l'app locale
- apre il browser su `http://127.0.0.1:5000`

## Avvio manuale

```bash
cd SAPCRManager
py -m pip install -r requirements.txt
py app.py
```

## Persistenza dati locale

I dati operativi vengono salvati in:

- `SAPCRManager/data_store.json`

Questo file e ignorato dal repository e non viene pushato su GitHub.

## Note

- La cartella `Home13/` non fa parte del deliverable pubblicato: e stata usata come riferimento stilistico locale.
- La documentazione tecnica dell'app si trova anche in `SAPCRManager/README.md`.