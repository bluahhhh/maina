import sys
import os
import json
import subprocess
import xml.etree.ElementTree as ET
from PyQt5 import QtWidgets, QtGui, QtCore, uic
from qt_material import apply_stylesheet
import keyboard

# Try to import crosshair modules, but handle gracefully if they're missing
try:
    from crosshair_overlay import CrosshairOverlay
    CROSSHAIR_OVERLAY_AVAILABLE = True
except ImportError:
    CrosshairOverlay = None
    CROSSHAIR_OVERLAY_AVAILABLE = False

try:
    from crosshair_picker import CrosshairPickerDialog
    CROSSHAIR_PICKER_AVAILABLE = True
except ImportError:
    CrosshairPickerDialog = None
    CROSSHAIR_PICKER_AVAILABLE = False

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MACRO_FOLDER = os.path.join(SCRIPT_DIR, "macros")
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "settings.json")
THEME_DIR = os.path.join(SCRIPT_DIR, "themes")
ICON_DIR = os.path.join(SCRIPT_DIR, "icons")
CROSSHAIR_FOLDER = os.path.join(SCRIPT_DIR, "crosshairs")
DEFAULT_SETTINGS = {
    "theme": "dark_teal.xml",
    "hide_hotkey": "F9",
    "play_hotkey": "F5",
    "pause_hotkey": "F6",
    "stop_hotkey": "F7",
    "reload_hotkey": "F8",
    "crosshair_toggle_hotkey": "F11",
    "window_geometry": None,
    "macro_hotkeys": {},
    "dialog_geometries": {},
    "crosshair_settings": None,
    "crosshair_enabled": False,
}

def ensure_macro_folder():
    if not os.path.exists(MACRO_FOLDER):
        os.makedirs(MACRO_FOLDER)
def ensure_crosshair_folder():
    if not os.path.exists(CROSSHAIR_FOLDER):
        os.makedirs(CROSSHAIR_FOLDER)

def open_macro_folder():
    folder = os.path.abspath(MACRO_FOLDER)
    if not os.path.exists(folder):
        os.makedirs(folder)
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])

def get_theme_accent(theme_path=None):
    if theme_path and os.path.exists(theme_path):
        try:
            tree = ET.parse(theme_path)
            root = tree.getroot()
            for item in root.findall(".//palette/item"):
                if item.get("name", "").lower() in ("accent", "highlight"):
                    return item.get("value")
            for item in root.findall(".//palette/item"):
                if "accent" in item.get("name", "").lower():
                    return item.get("value")
        except Exception:
            pass
    accent = "#ff5ea2"
    try:
        palette = QtWidgets.QApplication.instance().palette()
        accent = palette.color(QtGui.QPalette.Highlight).name()
    except Exception:
        pass
    return accent

class HotkeyLineEdit(QtWidgets.QLineEdit):
    key_captured = QtCore.pyqtSignal(str)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setPlaceholderText("Click and press a key...")
        self._listening = False

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        self._listening = True
        self.setText("")

    def keyPressEvent(self, event):
        if self._listening:
            key_seq = QtGui.QKeySequence(event.modifiers() | event.key())
            key_str = key_seq.toString(QtGui.QKeySequence.NativeText)
            if event.key() == QtCore.Qt.Key_Escape or key_str.lower() == "esc":
                QtWidgets.QToolTip.showText(self.mapToGlobal(QtCore.QPoint(0, self.height())),
                                            "ESC cannot be used as a hotkey.", self)
                self.setText("")
                self._listening = False
                return
            if key_str:
                self.setText(key_str)
                self.key_captured.emit(key_str)
            self._listening = False
        else:
            super().keyPressEvent(event)

# ... (other helper classes remain unchanged, except MacroPlayerGUI below) ...

class MacroPlayerGUI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        ui_file = os.path.join(SCRIPT_DIR, "main_window.ui")
        if os.path.exists(ui_file):
            try:
                uic.loadUi(ui_file, self)
            except Exception as e:
                print(f"UI file error: {e}")
                self.setWindowTitle("Macro Player (Material) - UI Error")
        else:
            self.setWindowTitle("Macro Player (Material) - No UI file")
            label = QtWidgets.QLabel("main_window.ui not found. Please place it in the script folder.")
            layout = QtWidgets.QVBoxLayout()
            layout.addWidget(label)
            central = QtWidgets.QWidget()
            central.setLayout(layout)
            self.setCentralWidget(central)
            return
        self.settings = self.load_settings()
        self.setWindowTitle("Division 2 Macro Player (Material)")
        self.loop_toggle = LoopToggleWidget(self)
        grid = self.findChild(QtWidgets.QGridLayout, "buttonGridLayout")
        if grid:
            old_loop = self.findChild(QtWidgets.QCheckBox, "loop_checkbox")
            if old_loop:
                grid.removeWidget(old_loop)
                old_loop.deleteLater()
            grid.addWidget(self.loop_toggle, 1, 0)
        self.loop_toggle.toggled.connect(self.update_loop_checkbox_style)
        self.loop_toggle.setChecked(False)
        self.settings_btn.clicked.connect(self.show_settings_popup)
        self.info_btn.clicked.connect(self.show_info_popup)
        self.reload_btn.clicked.connect(self.reload_macros)
        self.play_btn.clicked.connect(self.play_macro)
        self.pause_btn.clicked.connect(self.pause_macro)
        self.stop_btn.clicked.connect(self.stop_macro)
        self.crosshair_btn.clicked.connect(self.show_crosshair_picker)

        # --- FIXED MACRO PANE: always QTreeWidget ---
        macro_listbox = getattr(self, "macro_listbox", None)
        macro_tree = getattr(self, "macro_tree", None)
        if macro_listbox:
            parent_layout = macro_listbox.parent().layout()
            idx = parent_layout.indexOf(macro_listbox)
            macro_tree_fixed = QtWidgets.QTreeWidget(self)
            macro_tree_fixed.setObjectName("macro_tree")
            macro_tree_fixed.setHeaderHidden(True)
            macro_tree_fixed.setColumnCount(1)
            macro_tree_fixed.setIndentation(18)
            macro_tree_fixed.setAnimated(True)
            macro_tree_fixed.setMinimumWidth(210)
            macro_tree_fixed.setMaximumWidth(220)
            macro_tree_fixed.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            macro_tree_fixed.setStyleSheet("""
QTreeWidget {
    background: #232629;
    color: white;
    border: none;
}
QTreeWidget::item {
    min-height: 22px;
    font-size: 10pt;
    padding-left: 10px;
}
QTreeWidget::item:selected {
    background: #ff5ea2;
    color: #232629;
    border-radius: 4px;
}
QTreeWidget::branch:has-children {
    background: #232629;
}
""")
            parent_layout.removeWidget(macro_listbox)
            macro_listbox.hide()
            macro_listbox.setParent(None)
            macro_listbox.deleteLater()
            parent_layout.insertWidget(idx, macro_tree_fixed)
            parent_layout.setStretch(idx, 0)
            self.macro_tree = macro_tree_fixed
        elif macro_tree:
            self.macro_tree = macro_tree
        else:
            # fallback if neither present
            self.macro_tree = QtWidgets.QTreeWidget(self)
            self.macro_tree.setHeaderHidden(True)
            self.macro_tree.setColumnCount(1)
            self.macro_tree.setIndentation(18)
            self.macro_tree.setAnimated(True)
            self.macro_tree.setMinimumWidth(210)
            self.macro_tree.setMaximumWidth(220)
            self.macro_tree.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            self.macro_tree.setStyleSheet("""
QTreeWidget {
    background: #232629;
    color: white;
    border: none;
}
QTreeWidget::item {
    min-height: 22px;
    font-size: 10pt;
    padding-left: 10px;
}
QTreeWidget::item:selected {
    background: #ff5ea2;
    color: #232629;
    border-radius: 4px;
}
QTreeWidget::branch:has-children {
    background: #232629;
}
""")
        # --- End Macro Tree Widget FIX ---

        self.tooltip = TooltipHelper(self.macro_tree, self)
        # ... rest of your original class unchanged ...

    # ... all your methods unchanged ...
    # (leave everything else in your file as-is)

if __name__ == "__main__":
    ensure_macro_folder()
    ensure_crosshair_folder()
    app = QtWidgets.QApplication(sys.argv)
    settings = DEFAULT_SETTINGS.copy()
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                settings.update(loaded)
        except Exception:
            pass
    theme_file = settings.get("theme", "dark_teal.xml")
    theme_path = os.path.join(THEME_DIR, theme_file)
    if not os.path.exists(theme_path):
        theme_path = os.path.join(THEME_DIR, "dark_teal.xml")
    try:
        apply_stylesheet(app, theme=theme_path)
    except Exception as e:
        print(f"Theme application failed at startup: {e}")
    window = MacroPlayerGUI()
    window.show()
    sys.exit(app.exec_())