import sys
import os
import tempfile
import subprocess
import requests
from PyQt6.QtCore import QThread, pyqtSignal, QObject
from PyQt6.QtWidgets import QMessageBox, QProgressDialog

# --- 1. UPDATE CHECKER THREAD ---
class UpdateChecker(QThread):
    found = pyqtSignal(str, str)  # version, url
    not_found = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, current_version, repo_name):
        super().__init__()
        self.current_version = current_version.lstrip("v")
        self.repo_name = repo_name

    def run(self):
        try:
            url = f"https://api.github.com/repos/{self.repo_name}/releases/latest"
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                data = response.json()
                latest_tag = data.get("tag_name", "").lstrip("v")
                
                # Compare versions 
                if latest_tag != self.current_version:
                    # Find .exe in assets
                    exe_url = ""
                    for asset in data.get("assets", []):
                        if asset["name"].endswith(".exe"):
                            exe_url = asset["browser_download_url"]
                            break
                    
                    if exe_url:
                        self.found.emit(latest_tag, exe_url)
                    else:
                        self.not_found.emit()
                else:
                    self.not_found.emit()
            else:
                self.error.emit(f"GitHub API Error: {response.status_code}")
        except Exception as e:
            self.error.emit(str(e))

# --- 2. DOWNLOAD WORKER THREAD ---
class DownloadWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str) 
    error = pyqtSignal(str)

    def __init__(self, url):
        super().__init__()
        self.url = url

    def run(self):
        try:
            # Create path in Windows temp directory
            temp_dir = tempfile.gettempdir()
            # Get filename from URL 
            filename = self.url.split("/")[-1]
            save_path = os.path.join(temp_dir, filename)

            response = requests.get(self.url, stream=True, timeout=10)
            total_size = int(response.headers.get('content-length', 0))
            
            downloaded = 0
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            percent = int((downloaded / total_size) * 100)
                            self.progress.emit(percent)
            
            self.finished.emit(save_path)
            
        except Exception as e:
            self.error.emit(str(e))

# --- 3. UPDATER CONTROLLER ---
class Updater(QObject):
    def __init__(self, parent_window, current_version, repo_name):
        super().__init__(parent_window)
        self.parent = parent_window
        self.current_version = current_version
        self.repo_name = repo_name
        self.download_url = ""
        self.checker = None
        self.downloader = None
        self.progress_dialog = None

    def check_for_updates(self, silent=True):
        """Start update check"""
        self.silent_mode = silent
        self.checker = UpdateChecker(self.current_version, self.repo_name)
        self.checker.found.connect(self.on_update_found)
        
        # Helper for translation
        tr = getattr(self.parent, "get_translation", lambda k: k)
        
        if not silent:
            self.checker.not_found.connect(lambda: QMessageBox.information(
                self.parent, 
                tr("msg_no_updates_title"), 
                tr("msg_no_updates_content")
            ))
            self.checker.error.connect(lambda e: QMessageBox.warning(
                self.parent, 
                tr("msg_error_title"), 
                tr("msg_update_check_error").format(error=e)
            ))
        self.checker.start()

    def on_update_found(self, version, url):
        self.download_url = url
        tr = getattr(self.parent, "get_translation", lambda k: k)
        
        title = tr("msg_update_available_title")
        msg = tr("msg_update_available_content").format(version=version)
        
        reply = QMessageBox.question(self.parent, title, msg, 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.start_download()

    def start_download(self):
        tr = getattr(self.parent, "get_translation", lambda k: k)
        
        # Show progress dialog
        self.progress_dialog = QProgressDialog(tr("lbl_downloading_update"), tr("btn_cancel"), 0, 100, self.parent)
        self.progress_dialog.setWindowModality(2) # Block window
        self.progress_dialog.show()

        self.downloader = DownloadWorker(self.download_url)
        self.downloader.progress.connect(self.progress_dialog.setValue)
        self.downloader.finished.connect(self.install_and_restart)
        self.downloader.error.connect(lambda e: QMessageBox.critical(
            self.parent, 
            tr("msg_error_title"), 
            tr("msg_download_error").format(error=e)
        ))
        self.downloader.start()

    def install_and_restart(self, file_path):
        self.progress_dialog.close()
        tr = getattr(self.parent, "get_translation", lambda k: k)
        
        try:
            # Run installer
            subprocess.Popen([file_path], shell=True)
            
            # Close current app to allow installer to overwrite files
            sys.exit(0)
            
        except Exception as e:
            QMessageBox.critical(
                self.parent, 
                tr("msg_error_title"), 
                tr("msg_install_error").format(error=e)
            )
