import os
import sys
import json
import locale
from PyQt6.QtCore import QObject, pyqtSignal

class Utils:
    @staticmethod
    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    @staticmethod
    def load_translations(lang_code):
        """ Load translation for the given language code """
        try:
            # First try loading from internal assets (if bundled) or local folder
            path = Utils.resource_path(os.path.join("locales", f"{lang_code}.json"))
            if not os.path.exists(path):
                # Fallback to absolute path relative to src if running from source
                project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                path = os.path.join(project_root, "locales", f"{lang_code}.json")

            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading translation for {lang_code}: {e}")
            return {}

class ConfigManager:
    """ Simple JSON config manager """
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.config = {
            "language": "zh",
            "font_size_hanzi": 32,
            "font_size_pinyin": 18,
            "always_on_top": False,
            "favorite_fonts_hanzi": ["Microsoft YaHei", "KaiTi"],
            "favorite_fonts_pinyin": ["Arial"]
        }
        self.load()

    def load(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.config.update(data)
        except Exception as e:
            print(f"Error loading config: {e}")

    def save(self):
        # Only save specific keys to disk to enforce reset on restart
        save_data = {
            "language": self.config.get("language", "en"),
            "favorite_fonts_hanzi": self.config.get("favorite_fonts_hanzi", ["Microsoft YaHei", "KaiTi"]),
            "favorite_fonts_pinyin": self.config.get("favorite_fonts_pinyin", ["Arial"])
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(save_data, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

    def get(self, key, default=None):
        return self.config.get(key, default)

    def set(self, key, value):
        self.config[key] = value
        self.save()
