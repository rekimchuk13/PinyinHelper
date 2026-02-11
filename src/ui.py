import sys
import os
import re
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLineEdit, QLabel, QPushButton,
                             QScrollArea, QColorDialog, QFontComboBox,
                             QSpinBox, QMessageBox, QStyle, QFrame, QMenu, QComboBox,
                             QSystemTrayIcon, QFontDialog)
from PyQt6.QtCore import Qt, QBuffer, QIODevice, QByteArray, QMimeData, QTimer, QSize
from PyQt6.QtGui import QPainter, QColor, QFont, QPixmap, QFontMetrics, QImage, QAction, QIcon, QShortcut, QKeySequence, QFontDatabase, QCursor

from pypinyin import pinyin, Style

from .utils import Utils, ConfigManager
from .logic import GlobalHotKeyMonitor
from .updater import Updater

try:
    import win32clipboard
except ImportError:
    win32clipboard = None

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

HIGH_RES_SCALE = 6.0

class PairWidget(QWidget):
    def __init__(self, char, pinyin_text, index, parent_window):
        super().__init__()
        self.char = char
        self.pinyin = pinyin_text
        self.index = index
        self.main_window = parent_window
        self.color = parent_window.render_color

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(2)

        self.py_edit = QLineEdit(pinyin_text)
        self.py_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.py_edit.setStyleSheet(self.get_style(14))
        self.py_edit.textChanged.connect(self.on_text_changed)

        self.char_label = QLabel(char)
        self.char_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.char_label.setStyleSheet(f"color: {self.color.name()}; font-size: 24px; font-weight: bold;")

        self.layout.addWidget(self.py_edit)
        self.layout.addWidget(self.char_label)

    def get_style(self, font_size):
        return f"color: {self.color.name()}; background-color: #555; border: none; font-size: {font_size}px;"

    def on_text_changed(self, text):
        self.main_window.update_pair_text(self.index, text)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #444; color: white; border: 1px solid #666; }
            QMenu::item { padding: 5px 25px 5px 20px; }
            QMenu::item:selected { background-color: #0078d7; }
        """)

        tr = self.main_window.get_translation

        action_color = QAction(tr("ctx_color"), self)
        action_color.triggered.connect(self.change_pair_color)
        menu.addAction(action_color)

        menu.addSeparator()

        py_menu = menu.addMenu(tr("ctx_pinyin"))

        try:
            variations = pinyin(self.char, style=Style.TONE, heteronym=True)[0]
            unique_vars = sorted(list(set(variations)))

            if unique_vars:
                for py in unique_vars:
                    act = QAction(py, self)
                    if py == self.py_edit.text():
                        act.setCheckable(True)
                        act.setChecked(True)
                    act.triggered.connect(lambda checked, val=py: self.set_pinyin_text(val))
                    py_menu.addAction(act)
            else:
                empty_act = QAction(tr("ctx_no_variants"), self)
                empty_act.setEnabled(False)
                py_menu.addAction(empty_act)

        except Exception:
            pass

        menu.exec(event.globalPos())

    def change_pair_color(self):
        c = QColorDialog.getColor(self.color)
        if c.isValid():
            self.color = c
            self.py_edit.setStyleSheet(self.get_style(14))
            self.char_label.setStyleSheet(f"color: {self.color.name()}; font-size: 24px; font-weight: bold;")
            self.main_window.update_pair_color(self.index, c)

    def set_pinyin_text(self, text):
        self.py_edit.setText(text)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Load Config
        self.config = ConfigManager()
        self.current_lang = self.config.get("language", "en")
        self.translations = Utils.load_translations(self.current_lang)
        
        self.resize(1000, 800)
        self.setup_styles()

        self.pairs = []
        self.render_color = QColor(0, 0, 0)
        self.shortcuts = []
        
        # Font Favorites
        self.fav_fonts_h = self.config.get("favorite_fonts_hanzi", ["Microsoft YaHei", "KaiTi"])
        self.fav_fonts_p = self.config.get("favorite_fonts_pinyin", ["Arial"])
        self.remove_mode_h = False
        self.remove_mode_p = False
        self.auto_copy_font = False
        self._cached_com_info = None

        # --- TRAY ICON ---
        self.tray_icon = QSystemTrayIcon(self)
        # Fallback icon if file missing
        icon_path = Utils.resource_path(os.path.join("assets", "app_icon.png"))
        if os.path.exists(icon_path):
             self.tray_icon.setIcon(QIcon(icon_path))
        else:
             self.tray_icon.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
             
        self.tray_menu = QMenu()
        self.action_show = QAction("Show", self)
        self.action_show.triggered.connect(self.show_window)
        self.action_check_update = QAction("Check for Updates", self)
        self.action_check_update.triggered.connect(lambda: self.updater.check_for_updates(silent=False))
        self.action_quit = QAction("Quit", self)
        self.action_quit.triggered.connect(self.quit_app)
        self.tray_menu.addAction(self.action_show)
        self.tray_menu.addAction(self.action_check_update)
        self.tray_menu.addAction(self.action_quit)
        self.tray_icon.setContextMenu(self.tray_menu)

        self.tray_icon.activated.connect(self.on_tray_click)
        self.tray_icon.show()

        # --- HOTKEY MONITOR ---
        self.key_monitor = GlobalHotKeyMonitor()
        self.key_monitor.activated.connect(self.activate_from_clipboard)
        self.key_monitor.activated_replace.connect(self.quick_replace_from_clipboard)
        self.key_monitor.ctrl_c_pressed.connect(self._cache_com_info)
        self.key_monitor.start()

        # --- INTERFACE ---
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)

        # === TOP BAR ===
        top_bar = QHBoxLayout()
        self.entry = QLineEdit()
        self.entry.setFont(QFont("Microsoft YaHei", 12))
        self.entry.returnPressed.connect(self.process)
        
        # Explicit Ctrl+A support
        self.add_select_all_shortcut(self.entry)
        
        top_bar.addWidget(self.entry, 1)

        self.btn_process = QPushButton()
        self.btn_process.setStyleSheet("background-color: #0078d7; font-weight: bold;")
        self.btn_process.clicked.connect(self.process)
        top_bar.addWidget(self.btn_process)

        self.combo_lang = QComboBox()
        self.combo_lang.setFixedWidth(100)
        self.combo_lang.addItems(["English", "Русский", "中文"])
        
        # Set initial index based on config
        lang_map = {"en": 0, "ru": 1, "zh": 2}
        self.combo_lang.setCurrentIndex(lang_map.get(self.current_lang, 0))
        
        self.combo_lang.currentIndexChanged.connect(self.change_language)
        top_bar.addWidget(self.combo_lang)
        
        # Always on top button
        self.btn_top = QPushButton("Top")
        self.btn_top.setCheckable(True)
        self.btn_top.setChecked(self.config.get("always_on_top", False))
        self.btn_top.setStyleSheet("background-color: #555;")
        self.btn_top.toggled.connect(self.toggle_always_on_top)
        top_bar.addWidget(self.btn_top)

        layout.addLayout(top_bar)

        # === EDITOR AREA ===
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFixedHeight(140)
        self.scroll.setStyleSheet("background-color: #333; border-radius: 5px;")

        self.area = QWidget()
        self.area.setStyleSheet("background-color: #333;")
        self.area_layout = QHBoxLayout(self.area)
        self.area_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.scroll.setWidget(self.area)
        layout.addWidget(self.scroll)

        self.label_hint = QLabel()
        self.label_hint.setStyleSheet("color: #aaa; font-style: italic;")
        layout.addWidget(self.label_hint)

        # === SETTINGS PANEL ===
        sets = QWidget()
        sets.setStyleSheet("background-color: #3a3a3a; border-radius: 8px;")
        h_layout = QHBoxLayout(sets)
        h_layout.setContentsMargins(10, 13, 10, 13)
        h_layout.setSpacing(12)

        self.label_hanzi = QLabel()
        self.label_hanzi.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.label_hanzi.customContextMenuRequested.connect(lambda pos: self.show_font_context_menu(pos, "hanzi"))
        self.label_hanzi.setToolTip(self.get_translation("tip_hint"))
        h_layout.addWidget(self.label_hanzi)

        self.font_cb_h = QComboBox()
        self.font_cb_h.setMinimumWidth(120)
        self.update_font_combo("hanzi")
        self.font_cb_h.currentTextChanged.connect(self.on_hanzi_font_changed) # Changed signal
        self.add_select_all_shortcut(self.font_cb_h)
        h_layout.addWidget(self.font_cb_h, 1)

        self.spin_h = QSpinBox()
        self.spin_h.setFixedWidth(80)
        self.spin_h.setRange(10, 500)
        self.spin_h.setValue(self.config.get("font_size_hanzi", 32))
        self.spin_h.valueChanged.connect(self.on_hanzi_size_changed)
        self.add_select_all_shortcut(self.spin_h)
        h_layout.addWidget(self.spin_h)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.VLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("background-color: #555; margin: 0 5px;")
        h_layout.addWidget(line)

        self.label_pinyin = QLabel()
        self.label_pinyin.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.label_pinyin.customContextMenuRequested.connect(lambda pos: self.show_font_context_menu(pos, "pinyin"))
        self.label_pinyin.setToolTip(self.get_translation("tip_hint"))
        h_layout.addWidget(self.label_pinyin)

        self.font_cb_p = QComboBox()
        self.font_cb_p.setMinimumWidth(120)
        self.update_font_combo("pinyin")
        self.font_cb_p.currentTextChanged.connect(self.on_pinyin_font_changed)
        self.add_select_all_shortcut(self.font_cb_p)
        h_layout.addWidget(self.font_cb_p, 1)

        self.spin_p = QSpinBox()
        self.spin_p.setFixedWidth(80)
        self.spin_p.setRange(5, 300)
        self.spin_p.setValue(self.config.get("font_size_pinyin", 18))
        self.spin_p.valueChanged.connect(self.preview)
        self.add_select_all_shortcut(self.spin_p)
        h_layout.addWidget(self.spin_p)

        self.btn_col = QPushButton()
        self.btn_col.setFixedWidth(80)
        self.btn_col.setStyleSheet(
            f"background-color: {self.render_color.name()}; border: 1px solid #777; font-weight: bold; border-radius: 4px;")
        self.btn_col.clicked.connect(self.set_color)
        h_layout.addWidget(self.btn_col)

        layout.addWidget(sets)

        # === PREVIEW ===
        self.label_prev_title = QLabel()
        layout.addWidget(self.label_prev_title)

        self.lbl_prev = QLabel()
        self.lbl_prev.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_prev.setMinimumHeight(150)
        self.lbl_prev.setStyleSheet("""
            QLabel {
                border: 2px dashed #666;
                background-color: #444;
                background-image: linear-gradient(45deg, #555 25%, transparent 25%), 
                                  linear-gradient(-45deg, #555 25%, transparent 25%), 
                                  linear-gradient(45deg, transparent 75%, #555 75%), 
                                  linear-gradient(-45deg, transparent 75%, #555 75%);
            }
        """)
        layout.addWidget(self.lbl_prev, 1)

        # === COPY BUTTONS ===
        btns_layout = QHBoxLayout()

        # Copy Text Button
        self.btn_copy_txt = QPushButton()
        self.btn_copy_txt.setFixedHeight(50)
        self.btn_copy_txt.setStyleSheet("""
            QPushButton { background-color: #0078d7; color: white; font-weight: bold; font-size: 16px; border-radius: 8px; }
            QPushButton:hover { background-color: #0063b1; }
        """)
        self.btn_copy_txt.clicked.connect(self.copy_as_text_html)
        btns_layout.addWidget(self.btn_copy_txt)

        # Copy Image Button
        self.btn_copy_img = QPushButton()
        self.btn_copy_img.setFixedHeight(50)
        self.btn_copy_img.setStyleSheet("""
            QPushButton { background-color: #28a745; color: white; font-weight: bold; font-size: 16px; border-radius: 8px; }
            QPushButton:hover { background-color: #218838; }
        """)
        self.btn_copy_img.clicked.connect(self.copy_to_clipboard_win32)
        btns_layout.addWidget(self.btn_copy_img)

        layout.addLayout(btns_layout)
        
        # Apply logic
        if self.btn_top.isChecked():
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        
        self.retranslate_ui()
        
        # --- AUTO UPDATE ---
        from . import __version__, __repo_name__
        self.updater = Updater(self, __version__, __repo_name__)
        # Check updates silently after 2 seconds to not block startup
        QTimer.singleShot(2000, lambda: self.updater.check_for_updates(silent=True))

    @staticmethod
    def _detect_selection_info_com():
        """Try to get font size and per-character colors from running app via COM."""
        if not win32com:
            return None
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        com_apps = ["KWPP.Application", "PowerPoint.Application"]
        for prog_id in com_apps:
            try:
                app = win32com.client.GetActiveObject(prog_id)
                sel = app.ActiveWindow.Selection
                if sel.Type != 3:  # ppSelectionText
                    continue
                text_range = sel.TextRange
                size = text_range.Font.Size
                if not size or not (8 <= size <= 300):
                    size = 32

                colors = []
                count = text_range.Characters().Count
                for i in range(1, count + 1):
                    try:
                        rgb_bgr = text_range.Characters(i, 1).Font.Color.RGB
                        r = rgb_bgr & 0xFF
                        g = (rgb_bgr >> 8) & 0xFF
                        b = (rgb_bgr >> 16) & 0xFF
                        colors.append(QColor(r, g, b))
                    except Exception:
                        colors.append(QColor(0, 0, 0))

                return {"size": int(size), "colors": colors, "font_name": text_range.Font.Name}
            except Exception:
                continue
        return None

    def _cache_com_info(self):
        self._cached_com_info = self._detect_selection_info_com()

    def activate_from_clipboard(self):
        """Called on double Ctrl+C"""
        time.sleep(0.1)

        try:
            clipboard = QApplication.clipboard()
            mime = clipboard.mimeData()

            if mime.hasText():
                raw_text = mime.text()
                clean_text = raw_text.replace(" ", "").replace("\n", "").replace("\r", "")
                
                if clean_text:
                    detected_size = 32
                    detected_colors = None

                    # Try COM
                    info = self._detect_selection_info_com()
                    if info:
                        detected_size = info["size"]
                        detected_colors = info["colors"]
                        if self.auto_copy_font and info.get("font_name"):
                            font_name = info["font_name"]
                            if font_name not in self.fav_fonts_h:
                                self.fav_fonts_h.append(font_name)
                                self.save_favorite_fonts()
                                self.update_font_combo("hanzi")
                            self.font_cb_h.setCurrentText(font_name)
                    # Fallback: parse clipboard HTML
                    elif mime.hasHtml():
                        try:
                            html = mime.html()
                            val = None
                            match = re.search(r'(?:mso-ansi-)?font-size:\s*(\d+(?:\.\d+)?)\s*(pt|px)', html)
                            if match:
                                val = float(match.group(1))
                                unit = match.group(2)
                                if unit == 'px':
                                    val = val * 0.75
                            if val is None:
                                match_font = re.search(r'<font[^>]+size=["\']?(\d+)["\']?', html, re.IGNORECASE)
                                if match_font:
                                    html_size = int(match_font.group(1))
                                    html_to_pt = {1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36}
                                    val = html_to_pt.get(html_size, 12)

                            if val is not None and 8 <= val <= 300:
                                detected_size = int(val)
                        except Exception:
                            pass

                    self.entry.setText(clean_text)
                    self.show_window()

                    self.spin_h.setValue(detected_size)
                    self.process()

                    if detected_colors:
                        for i in range(min(len(detected_colors), len(self.pairs))):
                            color = detected_colors[i]
                            self.pairs[i]['color'] = color
                            if i < self.area_layout.count():
                                w = self.area_layout.itemAt(i).widget()
                                if w:
                                    w.color = color
                                    w.py_edit.setStyleSheet(w.get_style(14))
                                    w.char_label.setStyleSheet(f"color: {color.name()}; font-size: 24px; font-weight: bold;")
                        self.preview()

                    self.auto_adjust_pinyin_size()
                    return

            self.show_window()
            
        except Exception as e:
            self.show_window()

    def quick_replace_from_clipboard(self):
        """Called on Ctrl+C then Ctrl+X — silent inline replace."""
        time.sleep(0.15)

        try:
            clipboard = QApplication.clipboard()
            mime = clipboard.mimeData()

            if not mime.hasText():
                return

            raw_text = mime.text()
            clean_text = raw_text.replace(" ", "").replace("\n", "").replace("\r", "")
            if not clean_text:
                return

            detected_size = 32
            detected_colors = None

            info = self._cached_com_info
            self._cached_com_info = None
            if not info:
                info = self._detect_selection_info_com()
            if info:
                detected_size = info["size"]
                detected_colors = info["colors"]
                if self.auto_copy_font and info.get("font_name"):
                    font_name = info["font_name"]
                    if font_name not in self.fav_fonts_h:
                        self.fav_fonts_h.append(font_name)
                        self.save_favorite_fonts()
                        self.update_font_combo("hanzi")
                    self.font_cb_h.setCurrentText(font_name)

            from pypinyin import pinyin as py_pinyin, Style as py_Style
            raw = py_pinyin(clean_text, style=py_Style.TONE)
            self.pairs = []
            for i, ch in enumerate(clean_text):
                py_text = raw[i][0]
                color = self.render_color
                if detected_colors and i < len(detected_colors):
                    color = detected_colors[i]
                self.pairs.append({'ch': ch, 'py': py_text, 'color': color})

            self.spin_h.blockSignals(True)
            self.spin_h.setValue(detected_size)
            self.spin_h.blockSignals(False)
            self.auto_adjust_pinyin_size()

            # Copy HTML to clipboard
            self.copy_as_text_html()

            # Simulate Ctrl+V in the source app
            time.sleep(0.1)
            try:
                from pynput.keyboard import Controller, Key
                kb = Controller()
                kb.press(Key.ctrl)
                kb.press('v')
                kb.release('v')
                kb.release(Key.ctrl)
            except Exception:
                pass

        except Exception:
            pass

    def on_hanzi_size_changed(self):
        self.auto_adjust_pinyin_size()
        self.preview()
        # Save config
        self.config.set("font_size_hanzi", self.spin_h.value())

    def auto_adjust_pinyin_size(self):
        if not self.pairs: return

        hanzi_size_pt = self.spin_h.value()
        font_h = QFont(self.font_cb_h.currentText())
        font_h.setPointSize(hanzi_size_pt)
        fm_h = QFontMetrics(font_h)

        font_p = QFont(self.font_cb_p.currentText())
        best_size = 8

        for candidate in range(hanzi_size_pt, 7, -1):
            font_p.setPointSize(candidate)
            fm_p = QFontMetrics(font_p)

            all_fit = True
            for item in self.pairs:
                w_h = fm_h.horizontalAdvance(item['ch'])
                w_p = fm_p.horizontalAdvance(item['py'])

                if w_p > w_h:
                    all_fit = False
                    break

            if all_fit:
                best_size = candidate
                break

        self.spin_p.blockSignals(True)
        self.spin_p.setValue(best_size)
        self.spin_p.blockSignals(False)
        self.config.set("font_size_pinyin", best_size)
        self.preview()

    def reset_copy_btn(self, btn, text, color_hex):
        btn.setText(text)
        btn.setStyleSheet(
            f"background-color: {color_hex}; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray_icon.showMessage("Pinyin Helper", self.get_translation("tray_tooltip"),
                                   QSystemTrayIcon.MessageIcon.Information, 2000)

    def show_window(self):
        if self.isMinimized():
            self.showNormal()

        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        # If not always on top, remove the flag, otherwise keep it
        if not self.btn_top.isChecked():
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
            self.show()

        self.activateWindow()
        self.raise_()

    def toggle_always_on_top(self, checked):
        self.config.set("always_on_top", checked)
        if checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
            self.btn_top.setStyleSheet("background-color: #0078d7;")
            self.btn_top.setText(self.get_translation("btn_top_active"))
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
            self.btn_top.setStyleSheet("background-color: #555;")
            self.btn_top.setText(self.get_translation("btn_top"))
        self.show()

    def on_tray_click(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.show_window()

    def quit_app(self):
        self.key_monitor.stop()
        QApplication.quit()

    def change_language(self, index):
        codes = ["en", "ru", "zh"]
        self.current_lang = codes[index]
        self.config.set("language", self.current_lang)
        self.translations = Utils.load_translations(self.current_lang)
        self.retranslate_ui()

    def get_translation(self, key):
        return self.translations.get(key, key)

    def retranslate_ui(self):
        tr = self.get_translation
        self.setWindowTitle(tr("window_title"))
        icon_path = Utils.resource_path(os.path.join("assets", "app_icon.png"))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.entry.setPlaceholderText(tr("input_placeholder"))
        self.btn_process.setText(tr("btn_process"))
        self.label_hanzi.setText(tr("label_hanzi"))
        self.label_pinyin.setText(tr("label_pinyin"))
        self.btn_col.setText(tr("btn_color"))
        self.label_prev_title.setText(tr("label_preview"))
        self.btn_copy_img.setText(tr("btn_copy_img"))
        self.btn_copy_txt.setText(tr("btn_copy_txt"))
        self.label_hint.setText(tr("tip_hint"))
        self.tray_icon.setToolTip(tr("tray_tooltip"))
        self.action_show.setText(tr("tray_show"))
        self.action_check_update.setText(tr("tray_check_update"))
        self.action_quit.setText(tr("tray_quit"))
        if self.btn_top.isChecked():
            self.btn_top.setText(tr("btn_top_active"))
        else:
            self.btn_top.setText(tr("btn_top"))

    def process(self):
        txt = self.entry.text()
        if not txt: return
        raw = pinyin(txt, style=Style.TONE)

        while self.area_layout.count():
            w = self.area_layout.takeAt(0).widget()
            if w: w.deleteLater()

        self.pairs = []

        for i, ch in enumerate(txt):
            py = raw[i][0]
            self.pairs.append({
                'ch': ch, 'py': py, 'color': self.render_color
            })
            widget = PairWidget(ch, py, i, self)
            self.area_layout.addWidget(widget)

        self.auto_adjust_pinyin_size()
        self.preview()

    def update_pair_text(self, index, new_text):
        self.pairs[index]['py'] = new_text
        self.auto_adjust_pinyin_size()
        self.preview()

    def update_pair_color(self, index, new_color):
        self.pairs[index]['color'] = new_color
        self.preview()

    def set_color(self):
        c = QColorDialog.getColor(self.render_color)
        if c.isValid():
            self.render_color = c
            self.btn_col.setStyleSheet(
                f"background-color: {c.name()}; border: 1px solid #777; font-weight: bold; border-radius: 4px;")
            for i, item in enumerate(self.pairs):
                item['color'] = c
                if i < self.area_layout.count():
                    w = self.area_layout.itemAt(i).widget()
                    if w:
                        w.color = c
                        w.py_edit.setStyleSheet(w.get_style(14))
                        w.char_label.setStyleSheet(f"color: {c.name()}; font-size: 24px; font-weight: bold;")
            self.preview()

    def preview(self):
        if not self.pairs: return
        pix = self.generate(1.0)
        self.lbl_prev.setPixmap(pix)

    def update_font_combo(self, font_type):
        if font_type == "hanzi":
            combo = self.font_cb_h
            favorites = self.fav_fonts_h
        else:
            combo = self.font_cb_p
            favorites = self.fav_fonts_p
        
        current = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(favorites)
        
        if current in favorites:
            combo.setCurrentText(current)
        elif favorites:
            combo.setCurrentIndex(0)
            
        combo.blockSignals(False)

    def show_font_context_menu(self, pos, font_type):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #444; color: white; border: 1px solid #666; }
            QMenu::item { padding: 5px 25px 5px 20px; }
            QMenu::item:selected { background-color: #0078d7; }
        """)
        
        # Add Font Action
        action_add = QAction(self.get_translation("btn_add_font"), self)
        action_add.triggered.connect(lambda: self.add_font_dialog(font_type))
        menu.addAction(action_add)
        
        # Remove Current Font Action
        combo = self.font_cb_h if font_type == "hanzi" else self.font_cb_p
        current_font = combo.currentText()
        favorites = self.fav_fonts_h if font_type == "hanzi" else self.fav_fonts_p
        
        if current_font and current_font in favorites:
            action_remove = QAction(f"{self.get_translation('btn_remove_font')} '{current_font}'", self)
            action_remove.triggered.connect(lambda: self.remove_font(font_type, current_font))
            menu.addAction(action_remove)

        # Auto-detect font toggle (only for hanzi)
        if font_type == "hanzi":
            menu.addSeparator()
            action_auto_font = QAction(self.get_translation("chk_auto_copy_font"), self)
            action_auto_font.setCheckable(True)
            action_auto_font.setChecked(self.auto_copy_font)
            action_auto_font.toggled.connect(lambda checked: setattr(self, 'auto_copy_font', checked))
            menu.addAction(action_auto_font)
            
        menu.exec(QCursor.pos())

    def on_hanzi_font_changed(self, text):
        self.preview()

    def on_pinyin_font_changed(self, text):
        self.preview()

    def remove_font(self, font_type, font_name):
        if not font_name: return
        
        if font_type == "hanzi":
            favorites = self.fav_fonts_h
            combo = self.font_cb_h
        else:
            favorites = self.fav_fonts_p
            combo = self.font_cb_p

        if font_name in favorites:
            idx = favorites.index(font_name)
            favorites.remove(font_name)
            self.save_favorite_fonts()
            
            new_selection = None
            if favorites:
                new_idx = idx if idx < len(favorites) else len(favorites) - 1
                new_selection = favorites[new_idx]
            
            self.update_font_combo(font_type)
            
            # Start selection
            if new_selection:
                combo.setCurrentText(new_selection)

    def add_font_dialog(self, font_type):
        font, ok = QFontDialog.getFont(QFont("Arial"), self, self.get_translation("btn_add_font"))
        if ok:
            family = font.family()
            if font_type == "hanzi":
                if family not in self.fav_fonts_h:
                    self.fav_fonts_h.append(family)
                    self.save_favorite_fonts()
                    self.update_font_combo("hanzi")
                self.font_cb_h.setCurrentText(family)
            else:
                if family not in self.fav_fonts_p:
                    self.fav_fonts_p.append(family)
                    self.save_favorite_fonts()
                    self.update_font_combo("pinyin")
                self.font_cb_p.setCurrentText(family)

    def show_all_fonts_dialog(self, font_type):
        pass 

    def save_favorite_fonts(self):
        self.config.set("favorite_fonts_hanzi", self.fav_fonts_h)
        self.config.set("favorite_fonts_pinyin", self.fav_fonts_p)

    def generate(self, scale):
        base_h = self.spin_h.value()
        base_p = self.spin_p.value()
        size_h = int(base_h * scale)
        size_p = int(base_p * scale)
        spacing = int(10 * scale)

        font_h = QFont(self.font_cb_h.currentText())
        font_h.setPixelSize(size_h)
        font_p = QFont(self.font_cb_p.currentText())
        font_p.setPixelSize(size_p)

        fm_h = QFontMetrics(font_h)
        fm_p = QFontMetrics(font_p)

        blocks = []
        total_w = 0
        for item in self.pairs:
            w_h = fm_h.horizontalAdvance(item['ch'])
            w_p = fm_p.horizontalAdvance(item['py'])
            bw = max(w_h, w_p)
            total_w += bw + spacing
            blocks.append((item, bw, w_h, w_p))

        h_h = fm_h.height()
        h_p = fm_p.height()
        total_h = h_h + h_p + int(10 * scale)

        img = QImage(total_w, total_h, QImage.Format.Format_ARGB32)
        img.fill(Qt.GlobalColor.transparent)

        p = QPainter(img)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        x = 0
        y_py = fm_p.ascent()
        y_hz = h_p + int(5 * scale) + fm_h.ascent()

        for (item, bw, wh, wp) in blocks:
            color_to_use = item.get('color', self.render_color)
            p.setPen(color_to_use)

            p.setFont(font_p)
            p.drawText(int(x + (bw - wp) / 2), int(y_py), item['py'])
            p.setFont(font_h)
            p.drawText(int(x + (bw - wh) / 2), int(y_hz), item['ch'])
            x += bw + spacing

        p.end()
        return QPixmap.fromImage(img)

    def copy_to_clipboard_win32(self):
        if not self.pairs: return

        pix = self.generate(HIGH_RES_SCALE)
        ba = QByteArray()
        buff = QBuffer(ba)
        buff.open(QIODevice.OpenModeFlag.WriteOnly)
        pix.save(buff, "PNG")
        png_bytes = ba.data()

        try:
            if win32clipboard:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                cf_png = win32clipboard.RegisterClipboardFormat("PNG")
                win32clipboard.SetClipboardData(cf_png, bytes(png_bytes))
                win32clipboard.CloseClipboard()
            else:
                clipboard = QApplication.clipboard()
                clipboard.setPixmap(pix)

            old_text = self.btn_copy_img.text()
            self.btn_copy_img.setText("✅ OK!")
            self.btn_copy_img.setStyleSheet(
                "background-color: #1e7e34; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")

            QTimer.singleShot(1000, lambda: self.reset_copy_btn(self.btn_copy_img, old_text, "#28a745"))

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {e}")

    def copy_as_text_html(self):
        if not self.pairs: return

        size_h_pt = self.spin_h.value()
        size_p_pt = self.spin_p.value()
        font_h = QFont(self.font_cb_h.currentText())
        font_h.setPointSize(size_h_pt)
        font_p = QFont(self.font_cb_p.currentText())
        font_p.setPointSize(size_p_pt)
        fm_h = QFontMetrics(font_h)
        fm_p = QFontMetrics(font_p)

        col_widths_pt = []
        for item in self.pairs:
            px_w_h = fm_h.horizontalAdvance(item["ch"])
            px_w_p = fm_p.horizontalAdvance(item["py"])
            max_px = max(px_w_h, px_w_p)
            width_pt = int(max_px * 0.9)
            col_widths_pt.append(width_pt)

        html = '<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border: none;"><tr>'

        for i, item in enumerate(self.pairs):
            width = col_widths_pt[i]
            color = item.get('color', self.render_color).name()
            td_style = f"width: {width}pt; min-width: {width}pt; text-align: center; vertical-align: bottom; padding: 0;"
            span_style = f"font-family: '{font_p.family()}'; font-size: {size_p_pt}pt; color: {color}; line-height: 100%;"
            html += f'<td width="{width}" style="{td_style}"><span style="{span_style}">{item["py"]}</span></td>'

        html += "</tr><tr>"

        for i, item in enumerate(self.pairs):
            width = col_widths_pt[i]
            color = item.get('color', self.render_color).name()
            td_style = f"width: {width}pt; min-width: {width}pt; text-align: center; vertical-align: top; padding: 0;"
            span_style = f"font-family: '{font_h.family()}'; font-size: {size_h_pt}pt; color: {color}; line-height: 100%;"
            html += f'<td width="{width}" style="{td_style}"><span style="{span_style}">{item["ch"]}</span></td>'

        html += "</tr></table>"

        plain_py = " ".join(item["py"] for item in self.pairs)
        plain_hz = "".join(item["ch"] for item in self.pairs)
        plain_text = f"{plain_py}\n{plain_hz}"

        try:
            if win32clipboard:
                # Build CF_HTML with required headers
                fragment = html
                html_doc = f"<html><body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>"
                
                prefix = "Version:1.0\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\nStartFragment:{:010d}\r\nEndFragment:{:010d}\r\n"
                dummy = prefix.format(0, 0, 0, 0)
                prefix_len = len(dummy.encode("utf-8"))
                
                html_bytes = html_doc.encode("utf-8")
                start_html = prefix_len
                end_html = prefix_len + len(html_bytes)
                start_frag = prefix_len + html_bytes.find(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
                end_frag = prefix_len + html_bytes.find(b"<!--EndFragment-->")
                
                cf_html_data = prefix.format(start_html, end_html, start_frag, end_frag).encode("utf-8") + html_bytes
                
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                # Set plain text (CF_UNICODETEXT)
                win32clipboard.SetClipboardData(13, plain_text)
                # Set HTML (CF_HTML)
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                win32clipboard.SetClipboardData(cf_html, cf_html_data)
                win32clipboard.CloseClipboard()
            else:
                mime = QMimeData()
                mime.setText(plain_text)
                mime.setHtml(html)
                clipboard = QApplication.clipboard()
                clipboard.setMimeData(mime)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {e}")
            return

        old_text = self.btn_copy_txt.text()
        self.btn_copy_txt.setText("✅ OK!")
        self.btn_copy_txt.setStyleSheet(
            "background-color: #005a9e; color: white; font-weight: bold; font-size: 16px; border-radius: 8px;")

        QTimer.singleShot(1000, lambda: self.reset_copy_btn(self.btn_copy_txt, old_text, "#0078d7"))

    def setup_styles(self):
        plus_path = Utils.resource_path(os.path.join("assets", "plus.svg")).replace("\\", "/")
        minus_path = Utils.resource_path(os.path.join("assets", "minus.svg")).replace("\\", "/")

        self.setStyleSheet(f"""
            QMainWindow, QWidget {{ 
                background-color: #2b2b2b; color: #ffffff; 
                font-family: "Segoe UI", "Microsoft YaHei"; 
                selection-background-color: #0078d7; selection-color: white;
            }}
            QLineEdit {{ 
                background-color: #3b3b3b; border: 1px solid #555; 
                padding: 5px; color: #fff; border-radius: 4px; selection-background-color: #0078d7;
            }}
            QSpinBox {{ 
                background-color: #3b3b3b; border: 1px solid #555; 
                color: #fff; padding-right: 20px; border-radius: 4px;
                min-height: 25px;
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                subcontrol-origin: border; width: 20px;
                background-color: #444; border-left: 1px solid #555;
            }}
            QSpinBox::up-button {{ 
                subcontrol-position: top right; border-top-right-radius: 4px;
                image: url({plus_path});
                width: 20px; height: 15px; /* Adjust height slightly */
            }}
            QSpinBox::down-button {{ 
                subcontrol-position: bottom right; border-bottom-right-radius: 4px;
                image: url({minus_path});
                width: 20px; height: 15px; /* Adjust height slightly */
            }}
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background-color: #666; }}
            
            QSpinBox::up-arrow, QSpinBox::down-arrow {{
                width: 0; height: 0; border: none; background: none; image: none;
            }}
        """)

    def add_select_all_shortcut(self, widget):
        # Helper to add Ctrl+A shortcut
        if hasattr(widget, 'selectAll'):
            sc = QShortcut(QKeySequence(QKeySequence.StandardKey.SelectAll), widget)
            sc.activated.connect(widget.selectAll)
            self.shortcuts.append(sc)
        elif hasattr(widget, 'lineEdit') and widget.lineEdit():
            sc = QShortcut(QKeySequence(QKeySequence.StandardKey.SelectAll), widget)
            sc.activated.connect(widget.lineEdit().selectAll)
            self.shortcuts.append(sc)