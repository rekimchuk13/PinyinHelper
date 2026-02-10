import time
from PyQt6.QtCore import QObject, pyqtSignal

try:
    from pynput import keyboard
except ImportError:
    print("Error: pynput not installed")

class GlobalHotKeyMonitor(QObject):
    activated = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.last_press_time = 0
        self.listener = None

    def start(self):
        # Use GlobalHotKeys for reliable hotkey detection
        try:
            self.listener = keyboard.GlobalHotKeys({
                '<ctrl>+c': self.on_activate_c,
                '<ctrl>+с': self.on_activate_c # Russian 'с' just in case
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

    def on_activate_c(self):
        # Called when Ctrl+C is pressed
        current_time = time.time()
        # Check if less than 0.6s passed since last press
        if (current_time - self.last_press_time) < 0.6:
            self.activated.emit()
            self.last_press_time = 0 # Reset timer
        else:
            self.last_press_time = current_time
