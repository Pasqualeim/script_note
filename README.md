# 📌 Compare Excel Fixed - README

## 📖 Descrizione

Questo script Python permette di **comparare e analizzare due file Excel**, verificando le corrispondenze tra componenti e note estratte, applicando colori per evidenziare le note impattate e pulendo la colonna "Note Number" dai duplicati prima di salvare il file finale.

Lo script utilizza **Pandas** per l'elaborazione dei dati e **OpenPyXL** per la manipolazione diretta dei file Excel.

---

## 🛠️ Funzionalità principali

- **Selezione manuale dei file**: L'utente può scegliere i file Excel da comparare.
- **Applicazione di colori**: Le note impattate vengono colorate di **rosso**.
- **Pulizia della colonna "Note Number"**: I valori duplicati consecutivi vengono rimossi automaticamente.
- **Salvataggio dei risultati**: I file finali vengono salvati in una posizione scelta dall'utente.

---

## 📂 File generati

- **`Note Extraction Updated.xlsx`** → File aggiornato con i colori applicati.
- **`Impacted_Notes.xlsx`** → File contenente solo le note impattate, già pulito dai duplicati.

---

## 📂 Origine del file delle note impattate (Impacted_Notes.xlsx)

1. 📂 Il file si troverà nel seguente percorso in cartelle suddivise per trimestri o mensilitá (in base alle esigenze dei clienti):
 ```bash
...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\Note 2025
```
  
2. 📂 Attualmente viene generato da uno script (ongoing) situato in:
```bash
...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\script_python
```

---

## 🏗️ Requisiti

Assicurati di avere installate le seguenti librerie Python prima di eseguire lo script:

```bash
pip install pandas openpyxl
```
### Struttura del file delle componenti

Deve contenere solo tre colonne:
- Component
- Release
- SPLevel

Per il Kernel, la colonna Component deve avere una delle seguenti nomenclature:

- KERNEL
- KRNL64UC
- KRNLUC
  
---

## 🚀 Come utilizzare lo script

1. **Esegui il file Python** dalla cartella in cui è presente:

   ```bash
   python merge_and_clean_notes.py
   ```

2. **Seleziona i file**:
   
   - **Componenti** → Contiene le componenti del sistema.
   - **Note Extraction.xlsx** → Contiene le note estratte.

3. **Seleziona la posizione per salvare i file di output**.

4. **Lo script analizzerà i dati e genererà i file aggiornati**:

   - Colorando le note impattate.
   - Eliminando i duplicati dalla colonna "Note Number".

5. **Controlla i file generati** e utilizzali per ulteriori analisi.

---

## 📌 Struttura del codice

- **`select_file(prompt)`**: Mostra una finestra di dialogo per selezionare i file di input.
- **`select_save_location(default_name)`**: Mostra una finestra di dialogo per scegliere dove salvare i file.
- **`convert_version_format(version)`**: Converte le versioni in formato numerico.
- **`extract_sp_level(sp_value)`**: Estrae il valore numerico della colonna SP-Level.
- **`check_release_and_patch(component_row, note_row)`**: Verifica se il componente rientra nei range di versione.
- **`clean_impacted_notes(ws_red_notes)`**: Rimuove i duplicati consecutivi dalla colonna "Note Number".
- **`apply_color_to_note_number(components_df, notes_df, notes_file)`**:
  - Analizza i dati, applica i colori alle note impattate.
  - Pulisce la colonna "Note Number".
  - Salva i file aggiornati.

---

## 🔄 Esempio di utilizzo

```python
📌 Controllo componente: SEM-BW | SPLevel: 21
🔴 Impattato: A1448 colorato di rosso
📌 Controllo componente: SAP_BW | SPLevel: 22
🔴 Impattato: A1494 colorato di rosso
✅ Salvataggio completato: C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Note Extraction_Updated.xlsx e C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Impacted_Notes.xlsx
```

---

## 🛠️ Possibili miglioramenti

- **Ottimizzazione della gestione della memoria** per file molto grandi.
- **Aggiunta di un'interfaccia grafica (GUI)** per semplificare l'interazione.

---

## 📜 Licenza

Questo progetto è rilasciato sotto la licenza MIT. Puoi modificarlo e distribuirlo liberamente.

---

## 📧 Contatti

Per domande o suggerimenti, contattami su GitHub! 😊

