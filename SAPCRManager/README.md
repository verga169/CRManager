# SAP CR Manager

Applicazione Flask locale per gestire Change Request SAP organizzate per cliente e progetto.

## Funzionalita

- Persistenza locale su `data_store.json`
- Struttura gerarchica `Cliente -> Progetto -> CR`
- Campi CR: ordine rilascio, tipo (`Workbench`/`Customizing`), Richiesta, utente creatore, descrizione, note operative, stato trasporto
- Stati colore:
  - `Sviluppo` -> verde
  - `Quality` -> viola
  - `Produzione` -> rosso
- Ordinamento CR per sequenza di rilascio (ordine crescente), con ordine univoco all'interno del progetto
- Dashboard con contatori e filtri per stato, tipo CR o testo libero
- Export per singolo progetto direttamente dalla card progetto:
  - Excel (.xlsx)
  - PDF (.pdf) con layout tabellare formattato

## Avvio locale

```bash
py -m pip install -r requirements.txt
py app.py
```

Oppure su Windows con doppio click dal workspace root:

```bat
avvia_sap_cr_manager.bat
```

Il launcher avvia sempre in modalita sviluppo con auto-reload.

L'app risponde su `http://127.0.0.1:5055`.

Per cambiare porta puoi impostare la variabile ambiente `SAP_CR_MANAGER_PORT` prima dell'avvio.