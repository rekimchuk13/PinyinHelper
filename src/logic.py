import time
from PyQt6.QtCore import QObject, pyqtSignal

try:
    from pynput import keyboard
except ImportError:
    print("Error: pynput not installed")

class GlobalHotKeyMonitor(QObject):
    activated = pyqtSignal()
    activated_replace = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.last_c_time = 0
        self.listener = None

    def start(self):
        try:
            self.listener = keyboard.GlobalHotKeys({
                '<ctrl>+c': self.on_ctrl_c,
                '<ctrl>+с': self.on_ctrl_c,
                '<ctrl>+x': self.on_ctrl_x,
                '<ctrl>+ч': self.on_ctrl_x,
            })
            self.listener.start()
        except Exception as e:
            print(f"Failed to start hotkey listener: {e}")

    def stop(self):
        if self.listener:
            try:
                self.listener.stop()
            except:
                pass

    def on_ctrl_c(self):
        current_time = time.time()
        if (current_time - self.last_c_time) < 0.6:
            self.activated.emit()
            self.last_c_time = 0
        else:
            self.last_c_time = current_time

    def on_ctrl_x(self):
        current_time = time.time()
        if (current_time - self.last_c_time) < 0.6:
            self.activated_replace.emit()
            self.last_c_time = 0
