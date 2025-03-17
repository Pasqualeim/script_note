import sys
import os
import subprocess
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QProgressBar
from PyQt6.QtCore import QTimer

class FileSelectorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.process = None  # Processo che esegue `merge_and_clean_notes.py`
        self.timer = QTimer(self)  # Timer per aggiornare la GUI
        self.timer.timeout.connect(self.read_output)

    def initUI(self):
        self.setWindowTitle("Tool Security Notes")
        self.setGeometry(100, 100, 600, 450)  # Finestra più grande per mostrare output
        
        layout = QVBoxLayout()

        # Label per indicare lo stato
        self.label_status = QLabel("Estrai le note e confrontale con le componenti del sistema interessato")
        layout.addWidget(self.label_status)

        # Bottone per estrarre le note (da implementare)
        self.btn_extract = QPushButton("Estrai le note")
        self.btn_extract.clicked.connect(self.extract_notes)
        layout.addWidget(self.btn_extract)

        # Bottone per avviare il confronto
        self.btn_compare = QPushButton("Esegui Confronto")
        self.btn_compare.clicked.connect(self.run_comparison)
        layout.addWidget(self.btn_compare)

        # Barra di avanzamento
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)  # Inizialmente a 0%
        layout.addWidget(self.progress_bar)

        # Area di testo per mostrare l'output del terminale
        self.output_area = QTextEdit(self)
        self.output_area.setReadOnly(True)  # Non modificabile
        layout.addWidget(self.output_area)

        self.setLayout(layout)

    def extract_notes(self):
        """Funzione da implementare per estrarre le note."""
        self.label_status.setText("Estrazione delle note in corso... (Da implementare)")

    def run_comparison(self):
        """Avvia il confronto e controlla se `merge_and_clean_notes.py` viene eseguito correttamente"""
        self.label_status.setText("Elaborazione in corso... Attendere.")
        self.btn_compare.setEnabled(False)
        self.progress_bar.setValue(0)
        self.output_area.clear()

        python_exe = "python" if os.name == "nt" else "python3"

        # Avvia `merge_and_clean_notes.py` con output in tempo reale
        self.process = subprocess.Popen(
            [python_exe, "merge_and_clean_notes.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1  # Imposta il buffering a riga per ottenere output in tempo reale
        )

        self.timer.start(100)  # Controlla l'output ogni 100ms




    def read_output(self):
        """Legge l'output in tempo reale e aggiorna la barra di avanzamento."""
        if self.process is not None:
            output_line = self.process.stdout.readline()
            if output_line:
                self.output_area.append(output_line.strip())
                self.output_area.ensureCursorVisible()
                
                # Aggiorna la barra di avanzamento in base al numero di righe lette
                self.progress_bar.setValue(min(self.progress_bar.value() + 2, 100))

            return_code = self.process.poll()
            if return_code is not None:  # Il processo è terminato
                self.timer.stop()
                
                # Controlla eventuali errori
                errors = self.process.stderr.read().strip()
                if errors:
                    self.output_area.append(f"⚠ ERRORE: {errors}")
                    print(f"⚠ ERRORE: {errors}")  # Debug su terminale

                self.process.stdout.close()
                self.process = None
                self.label_status.setText("✅ Confronto completato!")
                self.progress_bar.setValue(100)
                self.btn_compare.setEnabled(True)




if __name__ == "__main__":
    os.environ["OBJC_DISABLE_INITIALIZE_FORK_SAFETY"] = "YES"  # Evita problemi su macOS
    app = QApplication(sys.argv)
    ex = FileSelectorApp()
    ex.show()
    sys.exit(app.exec())