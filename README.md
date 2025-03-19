# 📌 TOOL SECURITY NOTES - README

## 📖 Descrizione
Questo tool permette di **comparare e analizzare file Excel**, verificando le corrispondenze tra componenti e note estratte, applicando colori per evidenziare le note impattate e generando un file contenente solo queste ultime.  

Il programma è dotato di una **interfaccia grafica (GUI)** che permette di avviare il processo senza dover utilizzare la riga di comando.

---

## 🛠️ REQUISITI DI SISTEMA
Per eseguire il programma, **è necessario installare Python** e alcune librerie aggiuntive.

### 1️⃣ Installare Python (se non è installato)
Se Python **non è installato**, segui questi passaggi:

#### 🔹 Metodo 1: Installazione Automatica (Consigliato)
1. Apri il **Prompt dei comandi** (su Windows: premi `Win + R`, scrivi `cmd` e premi Invio).
2. Digita il seguente comando e premi Invio:
   ```cmd
   winget install -e --id Python.Python.3
   ```
3. **Riavvia il PC** dopo l'installazione.

#### 🔹 Metodo 2: Installazione Manuale
1. Scarica l'installer di Python da qui:  
   👉 [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)
2. Avvia il file `.exe` scaricato.
3. **IMPORTANTE**: **Seleziona "Add Python to PATH"** prima di cliccare "Install Now".
4. Attendi il completamento dell'installazione.
5. Apri il **Prompt dei comandi** e verifica l'installazione con:
   ```cmd
   python --version
   ```

Se tutto è corretto, vedrai una versione come questa:
```
Python 3.12.0
```

---

### 2️⃣ Installare le librerie necessarie
Dopo aver installato Python, è necessario installare i pacchetti richiesti.

1. Apri il **Prompt dei comandi**.
2. Digita e premi Invio:
   ```cmd
   pip install pandas openpyxl pyqt6
   ```

✅ Ora puoi eseguire il programma.

---

## 🚀 COME ESEGUIRE IL PROGRAMMA
Una volta installati **Python** e le librerie richieste, puoi avviare il programma.

### 🔹 Opzione 1: Eseguire il programma con GUI
1. Apri la cartella contenente il file `GUI.py`.
2. **Doppio clic su `GUI.py`** per avviare l'interfaccia grafica.
3. Premi il tasto `Esegui Confronto` per iniziare il processo.

### 🔹 Opzione 2: Eseguire il programma dal terminale
Se preferisci lanciare il programma manualmente:
1. Apri il **Prompt dei comandi**.
2. Vai nella cartella del programma:
   ```cmd
   cd "C:\Users\TuoNomeUtente\Documents\Script_pyton\script_note"
   ```
3. Avvia il programma con:
   ```cmd
   python GUI.py
   ```

---

## 📂 FILE GENERATI
Dopo l'esecuzione, il programma genera i seguenti file:

- **`Note Extraction Updated.xlsx`** → File aggiornato con i colori applicati.
- **`Impacted_Notes.xlsx`** → File contenente solo le note impattate, già pulito dai duplicati.

I file verranno salvati nella posizione scelta durante l'esecuzione.

---

## 📂 ORIGINE DEL FILE DELLE NOTE IMPATTATE
Il file delle note impattate si trova in:

📂 **Percorso file diviso per trimestri/mensilità:**
```bash
...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\Note 2025
```

📂 **Attualmente generato da uno script in:**
```bash
...EY\REMOTE-SERVICES - Documents\Remote\CLIENTI\Security Notes\script_python
```

---

## 📌 STRUTTURA DEL CODICE
### 🔹 Principali funzioni
- **`select_file(prompt)`** → Mostra una finestra per selezionare i file di input.
- **`select_save_location(default_name)`** → Permette di scegliere dove salvare i file.
- **`apply_color_to_note_number(components_df, notes_df, notes_file)`** → Confronta i dati e applica le modifiche.
- **`clean_impacted_notes(ws_red_notes)`** → Pulisce i duplicati dalla colonna "Note Number".
- **`run_comparison()`** → Avvia il processo e mostra l'avanzamento nella GUI.

---

## 🔄 ESEMPIO DI UTILIZZO
```
📌 Controllo componente: SEM-BW | SPLevel: 21
🔴 Impattato: A1448 colorato di rosso
📌 Controllo componente: SAP_BW | SPLevel: 22
🔴 Impattato: A1494 colorato di rosso
✅ Salvataggio completato: C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Note Extraction_Updated.xlsx e C:/Users/DD917MJ/OneDrive - EY/Documents/Script_pyton/Impacted_Notes.xlsx
```

---

## 🛠️ POSSIBILI MIGLIORAMENTI
- **Ottimizzazione della gestione della memoria** per file molto grandi.
- **Aggiunta di un'interfaccia grafica (GUI)** più avanzata con selezione dei file.

---

## 📜 LICENZA
Questo progetto è rilasciato sotto la licenza **MIT**. Puoi modificarlo e distribuirlo liberamente.

---

## 📧 CONTATTI
Per domande o suggerimenti, **contattami su GitHub!** 😊

