# Pinyin Helper

A useful tool for generating Pinyin from Chinese characters with coloring and clipboard integration.

## Features
- **Smart Pinyin Generation**: Automatically converts Hanzi to Pinyin.
- **Individual Styling**: Right-click any character to change its Pinyin or Color individually.
- **Clipboard Integration**: Select text anywhere, press `Ctrl+C` twice to instantly analyze.
- **Quick Replace** (`Ctrl+C` → `Ctrl+X`): Silently replaces selected text with Pinyin+Hanzi inline.
- **Color & Font Detection**: Automatically detects per-character colors and font from WPS/PowerPoint via COM.
- **Export**: Copy as Image or HTML table (for Word/Excel/PowerPoint/WPS).
- **Internationalization**: Supports English, Russian, and Chinese.
- **Auto-Update**: Checks for new versions on GitHub.
- **Manual Update Check**: Right-click the tray icon to check for updates manually.

## Installation

### For Users
1. Download the latest installer from the [Releases](https://github.com/rekimchuk13/PinyinHelper/releases) page.
2. Run the installer.
3. The app can start automatically with Windows.

### For Developers

1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python run.py
   ```

## Building

### Create Executable
```bash
pyinstaller PinyinHelper.spec
```

### Create Installer
Open `setup_script.iss` with Inno Setup Compiler and build.

## License
MIT
