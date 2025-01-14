import sys
from src.gui import TimecardGUI
from PyQt6.QtWidgets import QApplication

def main():
    app = QApplication(sys.argv)
    gui = TimecardGUI()
    gui.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()