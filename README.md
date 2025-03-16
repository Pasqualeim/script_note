# 📌 Tool Security Notes - README

## 📚 Descrizione

Questo script Python permette di **generare, comparare e analizzare due file Excel**, verificando le corrispondenze tra componenti e note estratte, applicando colori per evidenziare le note impattate e creando un file con sole queste ultime. Inoltre, pulisce la colonna **"Note Number"** dai duplicati prima di salvare il file finale.

Lo script utilizza **Pandas** per l'elaborazione dei dati e **OpenPyXL** per la manipolazione diretta dei file Excel.  
**L'interfaccia grafica (GUI) sviluppata in PyQt6 permette un'esperienza utente semplificata con selezione file, stato di avanzamento e output in tempo reale.**

---

## 🛠️ Funzionalità principali

- **Interfaccia Grafica (GUI) con PyQt6**: Possibilità di avviare il confronto da un'applicazione grafica.
- **Selezione manuale dei file**: L'utente può scegliere i file Excel da comparare.
- **Applicazione di colori**: Le note impattate vengono colorate di **rosso**.
- **Pulizia della colonna "Note Number"**: I valori duplicati consecutivi vengono rimossi automaticamente.
- **Salvataggio dei risultati**: I file finali vengono salvati in una posizione scelta dall'utente.
- **Output in tempo reale**: L'output del terminale viene mostrato direttamente nella GUI.
- **Barra di avanzamento**: Indica lo stato di avanzamento del confronto.

---

## 📂 File generati

- **`Note Extraction Updated.xlsx`** → File aggiornato con i colori applicati.
- **`Impacted_Notes.xlsx`** → File contenente solo le note impattate, già pulito dai duplicati.

---

## 📂 Origine del file delle note impattate (Impacted_Notes.xlsx)

1. 📂 Il file si troverà nel seguente percorso in cartelle suddivise per trimestri o mensilitá (in base alle esigenze dei clienti):

   ```
   ...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\Note 2025
   ```

2. 📂 Attualmente viene generato da uno script (ongoing) situato in:

   ```
   ...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\script_python
   ```

---

## 🏠 Requisiti

Assicurati di avere installate le seguenti librerie Python prima di eseguire lo script:

```
pip install pandas openpyxl pyqt6
```

### **Struttura del file delle componenti**

Deve contenere solo tre colonne:
- **Component**
- **Release**
- **SPLevel**

Per il Kernel, la colonna `Component` deve avere una delle seguenti nomenclature:

- `KERNEL`
- `KRNL64UC`
- `KRNLUC`

---

## 🚀 Come utilizzare lo script

### **Opzione 1: Avvio tramite GUI**
1. **Esegui la GUI** dalla cartella in cui è presente:
   ```
   python GUI.py
   ```
2. **Nell'interfaccia grafica**:
   - Clicca su **"Esegui Confronto"** per avviare il processo.
   - Visualizza **l'output in tempo reale** nella finestra.
   - Osserva **la barra di avanzamento** mentre il confronto viene elaborato.
   - Una volta terminato, i file verranno salvati nella posizione scelta.

---

### **Opzione 2: Avvio da Terminale**
1. **Esegui il file Python** manualmente:
   ```
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

## 📈 Struttura del codice

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
- **`GUI.py`**:
  - Esegue il confronto attraverso un'interfaccia grafica con barra di avanzamento e output in tempo reale.

---

## 🔄 Esempio di utilizzo

```
📌 Controllo componente: SEM-BW | SPLevel: 21
🔴 Impattato: A1448 colorato di rosso
📌 Controllo componente: SAP_BW | SPLevel: 22
🔴 Impattato: A1494 colorato di rosso
🚀 Salvataggio completato: C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Note Extraction_Updated.xlsx e C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Impacted_Notes.xlsx
```

---

## 🛠️ Possibili miglioramenti

- **Ottimizzazione della gestione della memoria** per file molto grandi.
- **Miglioramento delle performance della GUI**.
- **Aggiunta di una barra di avanzamento più dettagliata con step progressivi**.

---

## 📚 Licenza

Questo progetto è rilasciato sotto la **licenza MIT**. Puoi modificarlo e distribuirlo liberamente.

---

## 📝 Contatti

Per domande o suggerimenti, contattami su **GitHub!** 😊

