# SAP CR Manager

Applicazione Flask locale per gestire Change Request SAP organizzate per cliente e progetto.

## Funzionalita

- Persistenza locale su `data_store.json`
- Struttura gerarchica `Cliente -> Progetto -> CR`
- Campi CR: ID CR, utente creatore, descrizione, note operative, stato trasporto
- Stati colore:
  - `Sviluppo` -> verde
  - `Quality` -> viola
  - `Produzione` -> rosso
- Dashboard con contatori e filtri per stato o testo libero
- Export Excel della lista CR, anche con filtri correnti applicati

## Avvio locale

```bash
py -m pip install -r requirements.txt
py app.py
```

Oppure su Windows con doppio click dal workspace root:

```bat
avvia_sap_cr_manager.bat
```

L'app risponde su `http://127.0.0.1:5055`.

Per cambiare porta puoi impostare la variabile ambiente `SAP_CR_MANAGER_PORT` prima dell'avvio.