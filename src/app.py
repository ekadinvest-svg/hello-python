import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("דמו PySide6")
        self.resize(600, 400)

        label = QLabel("שלום אלעד — PySide6 באוויר!")
        label.setAlignment(
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
        )
        self.setCentralWidget(label)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Main()
    w.show()
    sys.exit(app.exec())
