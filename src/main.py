import sys
from PyQt6.QtWidgets import QApplication
from .ui import MainWindow

def main():
    app = QApplication(sys.argv)
    
    # Check if instance already running (basic check using simple mutex or just single instance lib could be better but keeping simple)
    # Since this is a simple tool, we'll verify it behaves well.
    # To truly prevent multiple instances, we would need QLocalSocket/QLocalServer. 
    # For now, relying on User logic or System Tray to manage visibility.

    window = MainWindow()
    
    # If not started minimized (e.g. valid arguments or just default), show it.
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
