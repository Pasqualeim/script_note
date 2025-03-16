import sys
import multiprocessing
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
from merge_and_clean_notes import main  # Importa la funzione main dal file corretto

class FileSelectorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("Tool Security Notes")
        self.setGeometry(100, 100, 500, 250)  # Finestra leggermente più grande
        
        layout = QVBoxLayout()

        # Label per indicare lo stato
        self.label_status = QLabel("Estrai le note e confrontale con le componenti del sistema interessato")
        layout.addWidget(self.label_status)

        # Bottone per estrarre le note (deve essere separato)
        self.btn_extract = QPushButton("Estrai le note")
        self.btn_extract.clicked.connect(self.extract_notes)
        layout.addWidget(self.btn_extract)

        # Bottone per avviare il confronto
        self.btn_compare = QPushButton("Esegui Confronto")
        self.btn_compare.clicked.connect(self.run_comparison)
        layout.addWidget(self.btn_compare)
        
        self.setLayout(layout)

    def extract_notes(self):
        """Questa funzione dovrà essere implementata se vuoi estrarre le note."""
        self.label_status.setText("Estrazione delle note in corso... (Da implementare)")
        # Qui dovresti chiamare una funzione specifica per estrarre le note.

    def run_comparison(self):
        """Esegue il confronto delle componenti con le note estratte in un processo separato."""
        self.label_status.setText("Elaborazione in corso... Attendere.")
        self.btn_compare.setEnabled(False)  # Disabilita il pulsante durante l'elaborazione

        process = multiprocessing.Process(target=main)  # Avvia `main()` in un nuovo processo
        process.start()
        process.join()  # Aspetta la fine dell'elaborazione

        self.label_status.setText("✅ Confronto completato!")
        self.btn_compare.setEnabled(True)  # Riabilita il pulsante

if __name__ == "__main__":
    multiprocessing.freeze_support()  # Necessario per Windows
    app = QApplication(sys.argv)
    ex = FileSelectorApp()
    ex.show()
    sys.exit(app.exec())