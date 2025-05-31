from PyQt5.QtWidgets import QApplication, QLabel
from docx import Document
import sys

app = QApplication(sys.argv)
label = QLabel("¡Funciona, Fersa!")
label.show()
sys.exit(app.exec_())