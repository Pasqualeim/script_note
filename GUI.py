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
        """Esegue il confronto in un processo separato e aggiorna la GUI in tempo reale con una barra di avanzamento."""
        self.label_status.setText("Elaborazione in corso... Attendere.")
        self.btn_compare.setEnabled(False)
        self.progress_bar.setValue(0)  # Reset della barra di avanzamento

        # Avvia `merge_and_clean_notes.py` come un nuovo processo indipendente
        self.process = subprocess.Popen(
            ["python3", "merge_and_clean_notes.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1  # Imposta il buffering a riga per ottenere output in tempo reale
        )

        self.timer.start(100)  # Controlla l'output ogni 100ms

    def read_output(self):
        """Legge e aggiorna l'output e la barra di avanzamento in tempo reale senza bloccare la GUI."""
        if self.process is not None:
            lines = self.process.stdout.readlines()
            if lines:
                total_lines = len(lines)
                processed_lines = 0  # Contatore di righe lette
                
                for line in lines:
                    self.output_area.append(line.strip())
                    self.output_area.ensureCursorVisible()  # Scorre automaticamente l'output
                    processed_lines += 1

                    # Calcola la percentuale di avanzamento
                    progress = int((processed_lines / total_lines) * 100)
                    self.progress_bar.setValue(progress)

            return_code = self.process.poll()
            if return_code is not None:  # Il processo è terminato
                self.timer.stop()
                self.process.stdout.close()
                self.process = None
                self.label_status.setText("✅ Confronto completato!")
                self.progress_bar.setValue(100)  # Imposta la barra al 100% alla fine
                self.btn_compare.setEnabled(True)

if __name__ == "__main__":
    os.environ["OBJC_DISABLE_INITIALIZE_FORK_SAFETY"] = "YES"  # Evita problemi su macOS
    app = QApplication(sys.argv)
    ex = FileSelectorApp()
    ex.show()
    sys.exit(app.exec())