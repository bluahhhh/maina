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

class ThemePickerDialog(QtWidgets.QDialog):
    def __init__(self, parent, theme_names, name_to_file, current_theme):
        super().__init__(parent)
        self.setWindowTitle("Choose Theme")
        self.setMinimumWidth(350)
        self.setMaximumWidth(380)
        layout = QtWidgets.QVBoxLayout(self)

        grid_widget = QtWidgets.QWidget()
        grid = QtWidgets.QGridLayout(grid_widget)
        grid.setSpacing(8)
        grid.setContentsMargins(0, 0, 0, 0)

        icon_size = 24
        row_height = 38
        font = QtGui.QFont("Arial", 9)
        col_count = 2

        color_icons = {
            "Amber": "yellow.png", "Blue": "blue.png", "Cyan": "cyan.png",
            "Lightgreen": "green.png", "Pink": "pink.png", "Purple": "purple.png",
            "Red": "red.png", "Teal": "teal.png", "Yellow": "yellow.png", "Orange": "orange.png"
        }
        mode_icons = {"Dark": "darkmode.png", "Light": "lightmode.png"}
        def get_color_icon(theme_name):
            for color in color_icons:
                if color.lower() in theme_name.lower():
                    path = os.path.join(ICON_DIR, color_icons[color])
                    if os.path.exists(path): return path
            path = os.path.join(ICON_DIR, "light.png")
            return path if os.path.exists(path) else None
        def get_mode_icon(theme_name):
            if theme_name.lower().startswith("dark"):
                path = os.path.join(ICON_DIR, mode_icons["Dark"])
                if os.path.exists(path): return path
            if theme_name.lower().startswith("light"):
                path = os.path.join(ICON_DIR, mode_icons["Light"])
                if os.path.exists(path): return path
            return None

        self.selected_theme = None
        self.selected_theme_file = None
        self.theme_cards = []

        for i, pretty_name in enumerate(theme_names):
            theme_file = name_to_file[pretty_name]
            card = QtWidgets.QPushButton()
            card.setCheckable(True)
            card.setAutoExclusive(True)
            card.setMinimumHeight(row_height)
            card.setMaximumHeight(row_height)
            card.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            card.setFont(font)

            hbox = QtWidgets.QHBoxLayout(card)
            hbox.setContentsMargins(7, 0, 7, 0)
            hbox.setSpacing(7)
            mode_icon_path = get_mode_icon(pretty_name)
            mode_label = QtWidgets.QLabel()
            if mode_icon_path:
                mode_pix = QtGui.QPixmap(mode_icon_path).scaled(icon_size, icon_size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                mode_label.setPixmap(mode_pix)
            hbox.addWidget(mode_label)
            color_icon_path = get_color_icon(pretty_name)
            color_label = QtWidgets.QLabel()
            if color_icon_path:
                color_pix = QtGui.QPixmap(color_icon_path).scaled(icon_size, icon_size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                color_label.setPixmap(color_pix)
            hbox.addWidget(color_label)
            text_label = QtWidgets.QLabel(pretty_name)
            text_label.setFont(font)
            text_label.setStyleSheet("color: white;")
            hbox.addWidget(text_label)
            hbox.addStretch()
            card.clicked.connect(lambda checked, name=pretty_name, tf=theme_file: self.on_select(name, tf))
            self.theme_cards.append((card, theme_file))
            grid.addWidget(card, i // col_count, i % col_count)
            if theme_file == current_theme:
                card.setChecked(True)
                self.selected_theme = pretty_name
                self.selected_theme_file = theme_file

        grid_widget.setMinimumWidth(300)
        grid_widget.setMaximumWidth(340)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(grid_widget)
        scroll.setMinimumHeight(210)
        scroll.setMaximumHeight(220)
        layout.addWidget(scroll)

        btnbox = QtWidgets.QHBoxLayout()
        savebtn = QtWidgets.QPushButton("SAVE")
        savebtn.setMinimumWidth(80)
        savebtn.setMaximumWidth(110)
        savebtn.setDefault(True)
        savebtn.clicked.connect(self.accept)
        cancelbtn = QtWidgets.QPushButton("CANCEL")
        cancelbtn.setMinimumWidth(80)
        cancelbtn.setMaximumWidth(110)
        cancelbtn.clicked.connect(self.reject)
        btnbox.addStretch()
        btnbox.addWidget(savebtn)
        btnbox.addWidget(cancelbtn)
        layout.addLayout(btnbox)
        self.setLayout(layout)
        self.update_theme_card_highlights(current_theme)

    def update_theme_card_highlights(self, highlight_theme_file):
        accent = get_theme_accent(os.path.join(THEME_DIR, highlight_theme_file))
        for card, theme_file in self.theme_cards:
            if card.isChecked() or theme_file == highlight_theme_file:
                card.setStyleSheet(f"""
                    QPushButton {{
                        background: {accent};
                        color: #222;
                        border-radius: 10px;
                        text-align: left;
                        padding: 4px 10px;
                        border: 2px solid {accent};
                    }}
                    QPushButton:hover {{
                        background: {accent};
                        color: #222;
                    }}
                """)
            else:
                card.setStyleSheet("""
                    QPushButton {
                        background: #232629;
                        color: white;
                        border-radius: 10px;
                        text-align: left;
                        padding: 4px 10px;
                        border: 2px solid #777;
                    }
                    QPushButton:hover {
                        background: #32363b;
                    }
                """)

    def on_select(self, theme_name, theme_file):
        self.selected_theme = theme_name
        self.selected_theme_file = theme_file
        theme_path = os.path.join(THEME_DIR, theme_file)
        if os.path.exists(theme_path):
            try:
                apply_stylesheet(QtWidgets.QApplication.instance(), theme=theme_path)
                for widget in QtWidgets.QApplication.instance().allWidgets():
                    widget.style().unpolish(widget)
                    widget.style().polish(widget)
                    widget.update()
            except Exception as e:
                print(f"Theme application failed: {e}")
        self.update_theme_card_highlights(theme_file)

    def get_selected_theme(self):
        return self.selected_theme
    def get_selected_theme_file(self):
        return self.selected_theme_file

class MacroErrorNotifier(QtCore.QObject):
    error_signal = QtCore.pyqtSignal(str)
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.error_signal.connect(self.show_error)
    def show_error(self, message):
        self.parent.statusBar().showMessage(message, 5000)
        QtWidgets.QMessageBox.warning(self.parent, "Macro Error", message)

class LoopToggleWidget(QtWidgets.QWidget):
    toggled = QtCore.pyqtSignal(bool)
    def __init__(self, parent=None):
        super().__init__(parent)
        self._checked = False
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        hbox = QtWidgets.QHBoxLayout(self)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(8)
        self.box = QtWidgets.QLabel(self)
        self.box.setFixedSize(22, 22)
        self.label = QtWidgets.QLabel("Loop", self)
        self.label.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        hbox.addWidget(self.box)
        hbox.addWidget(self.label)
        hbox.addStretch()
        self.setLayout(hbox)
        self.update_ui()
    def mouseReleaseEvent(self, event):
        self.setChecked(not self._checked)
        self.toggled.emit(self._checked)
        self.update_ui()
    def setChecked(self, checked):
        self._checked = checked
        self.update_ui()
    def isChecked(self):
        return self._checked
    def update_theme_color(self):
        return get_theme_accent(os.path.join(THEME_DIR, DEFAULT_SETTINGS.get("theme", "dark_teal.xml")))
    def update_ui(self):
        accent = get_theme_accent(os.path.join(THEME_DIR, DEFAULT_SETTINGS.get("theme", "dark_teal.xml")))
        tick = "✔"
        if self._checked:
            self.box.setText(tick)
            self.box.setAlignment(QtCore.Qt.AlignCenter)
            self.box.setStyleSheet(
                f"border-radius: 4px; background: #222; color: {accent}; font-size: 17px; font-weight: bold;"
            )
            self.label.setStyleSheet(f"color: {accent}; font-weight: bold;")
            self.label.setText("Loop")
        else:
            self.box.setText("")
            self.box.setStyleSheet("border-radius: 4px; background: #444;")
            self.label.setStyleSheet("color: #888888; font-weight: normal;")
            self.label.setText("Loop")

class TooltipHelper(QtCore.QObject):
    def __init__(self, tree_widget, parent_gui):
        super().__init__()
        self.tree_widget = tree_widget
        self.parent_gui = parent_gui
        self.tree_widget.setMouseTracking(True)
        self.tree_widget.viewport().installEventFilter(self)
    def eventFilter(self, obj, event):
        if obj == self.tree_widget.viewport() and event.type() == QtCore.QEvent.MouseMove:
            item = self.tree_widget.itemAt(event.pos())
            if item:
                if item.childCount() == 0 and item.data(0, QtCore.Qt.UserRole) == "file":
                    macro_file = item.text(0)
                    category = self.parent_gui.get_item_category(item)
                    info_text = self.parent_gui.get_macro_info(category, macro_file)
                    if info_text and not macro_file.startswith("No macros found"):
                        QtWidgets.QToolTip.showText(event.globalPos(), info_text, self.tree_widget)
                    else:
                        QtWidgets.QToolTip.hideText()
                else:
                    QtWidgets.QToolTip.hideText()
            else:
                QtWidgets.QToolTip.hideText()
        if obj == self.tree_widget.viewport() and event.type() == QtCore.QEvent.Leave:
            QtWidgets.QToolTip.hideText()
        return super().eventFilter(obj, event)

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

        # Macro tree setup for left pane
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
        macro_listbox = getattr(self, "macro_listbox", None)
        if macro_listbox:
            parent_layout = macro_listbox.parent().layout()
            idx = parent_layout.indexOf(macro_listbox)
            macro_listbox.hide()
            parent_layout.removeWidget(macro_listbox)
            macro_listbox.setParent(None)
            macro_listbox.deleteLater()
            parent_layout.insertWidget(idx, self.macro_tree)
            parent_layout.setStretch(idx, 0)  # keep left pane fixed

        self.tooltip = TooltipHelper(self.macro_tree, self)
        self.current_macro_file = None
        self.current_macro_category = None
        self.macro_actions = []
        self.macro_index = 0
        self.macro_running = False
        self.macro_paused = False
        self.macro_loop = False
        self.notifier = MacroErrorNotifier(self)
        self.macro_timer = QtCore.QTimer()
        self.macro_timer.timeout.connect(self.execute_macro_step)
        self.held_keys = set()
        theme_file = self.settings.get("theme", "dark_teal.xml")
        theme_path = os.path.join(THEME_DIR, theme_file)
        if not os.path.exists(theme_path):
            theme_path = os.path.join(THEME_DIR, "dark_teal.xml")
        apply_stylesheet(QtWidgets.QApplication.instance(), theme=theme_path)
        self.restore_window_geometry()
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        self.statusBar().showMessage("Ready")
        self.show()
        self.installEventFilter(self)
        self.setup_hotkey_global()
        self.setup_action_hotkeys()
        self.reload_macros()
        ensure_crosshair_folder()
        self.crosshair_overlay = None
        self.crosshair_settings = self.settings.get("crosshair_settings", None)
        if self.crosshair_settings is None or not isinstance(self.crosshair_settings, dict):
            self.crosshair_settings = {}
        self.crosshair_enabled = bool(self.settings.get("crosshair_enabled", False))
        if self.crosshair_enabled and self.crosshair_settings:
            self.create_crosshair_overlay()

    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.Close or event.type() == QtCore.QEvent.Resize:
            self.save_window_geometry()
        return super().eventFilter(obj, event)

    def save_window_geometry(self):
        geometry = self.saveGeometry().toHex().data().decode()
        self.settings["window_geometry"] = geometry
        self.save_settings()

    def restore_window_geometry(self):
        geometry = self.settings.get("window_geometry", None)
        if geometry:
            try:
                self.restoreGeometry(QtCore.QByteArray.fromHex(bytes(geometry, 'utf-8')))
            except Exception:
                pass

    def setup_hotkey_global(self):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        hotkey = self.settings.get("hide_hotkey", "F9")
        def hotkey_callback():
            QtCore.QMetaObject.invokeMethod(self, "toggle_visibility", QtCore.Qt.QueuedConnection)
        keyboard.add_hotkey(hotkey, hotkey_callback, suppress=False)

    def setup_action_hotkeys(self):
        for key in [
            "play_hotkey", "pause_hotkey", "stop_hotkey", "reload_hotkey", "crosshair_toggle_hotkey"
        ]:
            hotkey = self.settings.get(key)
            if hotkey:
                try:
                    keyboard.remove_hotkey(hotkey)
                except Exception:
                    pass
        hotkey_map = {
            "play_hotkey": self.play_macro,
            "pause_hotkey": self.pause_macro,
            "stop_hotkey": self.stop_macro,
            "reload_hotkey": self.reload_macros,
            "crosshair_toggle_hotkey": self.toggle_crosshair,
        }
        for key, func in hotkey_map.items():
            hotkey = self.settings.get(key)
            if hotkey:
                keyboard.add_hotkey(
                    hotkey,
                    lambda f=func: QtCore.QMetaObject.invokeMethod(self, f.__name__, QtCore.Qt.QueuedConnection),
                    suppress=False
                )

    @QtCore.pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible():
            self.release_all_keys()
            self.hide()
            QtWidgets.QApplication.processEvents()
        else:
            self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
            self.showNormal()
            self.raise_()
            self.activateWindow()
            self.show()
            QtWidgets.QApplication.processEvents()

    @QtCore.pyqtSlot()
    def toggle_crosshair(self):
        try:
            self.crosshair_enabled = not self.crosshair_enabled
            self.settings["crosshair_enabled"] = self.crosshair_enabled
            self.save_settings()
            if self.crosshair_enabled and self.crosshair_settings:
                self.create_crosshair_overlay()
            else:
                self.hide_crosshair_overlay()
        except Exception as e:
            self.statusBar().showMessage(f"Error toggling crosshair: {str(e)}", 5000)
            print(f"Error toggling crosshair: {e}")

    @QtCore.pyqtSlot()
    def play_macro(self):
        selected_items = self.macro_tree.selectedItems()
        if self.macro_running and self.macro_paused:
            self.macro_paused = False
            self.start_macro_timer()
            self.statusBar().showMessage("Macro resumed", 2000)
            return
        if self.macro_running and not self.macro_paused:
            return
        if selected_items and selected_items[0].childCount() == 0 and selected_items[0].data(0, QtCore.Qt.UserRole) == "file":
            item = selected_items[0]
            macro_file = item.text(0)
            category = self.get_item_category(item)
            macro_path = os.path.join(MACRO_FOLDER, category, macro_file) if category else os.path.join(MACRO_FOLDER, macro_file)
            self.stop_macro()
            self.current_macro_file = macro_file
            self.current_macro_category = category
            self.macro_loop = self.loop_toggle.isChecked()
            try:
                with open(macro_path, "r", encoding="utf-8") as f:
                    actions = json.load(f)
            except Exception as e:
                self.notifier.error_signal.emit(f"Failed to load macro: {e}")
                return
            if not isinstance(actions, list):
                self.notifier.error_signal.emit("Macro file is not a list of actions.")
                return
            if not actions:
                self.notifier.error_signal.emit("No actions found in macro.")
                return
            self.macro_actions = actions
            self.macro_index = 0
            self.macro_running = True
            self.macro_paused = False
            self.statusBar().showMessage(f"Playing macro: {macro_file}", 2000)
            self.start_macro_timer()

    @QtCore.pyqtSlot()
    def pause_macro(self):
        if self.macro_running:
            self.macro_paused = True
            self.macro_timer.stop()
            self.release_all_keys()
            self.statusBar().showMessage("Macro paused", 2000)

    @QtCore.pyqtSlot()
    def stop_macro(self):
        self.macro_running = False
        self.macro_paused = False
        self.macro_timer.stop()
        self.macro_index = 0
        self.macro_actions = []
        self.current_macro_file = None
        self.current_macro_category = None
        self.release_all_keys()
        self.statusBar().showMessage("Macro stopped", 2000)

    @QtCore.pyqtSlot()
    def reload_macros(self):
        self.macro_tree.clear()
        ensure_macro_folder()
        folder_icon = QtGui.QIcon(os.path.join(ICON_DIR, 'folder.png')) if os.path.exists(os.path.join(ICON_DIR, 'folder.png')) else QtGui.QIcon()
        file_icon = QtGui.QIcon(os.path.join(ICON_DIR, 'macro.png')) if os.path.exists(os.path.join(ICON_DIR, 'macro.png')) else QtGui.QIcon()

        categories = {}
        top_level_macros = []
        for entry in sorted(os.listdir(MACRO_FOLDER), key=lambda x: x.lower()):
            full_path = os.path.join(MACRO_FOLDER, entry)
            if os.path.isdir(full_path):
                files = [f for f in sorted(os.listdir(full_path), key=lambda x: x.lower()) if f.endswith(".json")]
                if files:
                    categories[entry] = files
            elif entry.endswith(".json"):
                top_level_macros.append(entry)

        for f in top_level_macros:
            item = QtWidgets.QTreeWidgetItem([f])
            item.setIcon(0, file_icon)
            item.setData(0, QtCore.Qt.UserRole, "file")
            self.macro_tree.addTopLevelItem(item)

        for cat in sorted(categories.keys(), key=lambda x: x.lower()):
            parent = QtWidgets.QTreeWidgetItem([cat])
            parent.setIcon(0, folder_icon)
            font = parent.font(0)
            font.setBold(True)
            parent.setFont(0, font)
            parent.setData(0, QtCore.Qt.UserRole, "category")
            for f in categories[cat]:
                child = QtWidgets.QTreeWidgetItem([f])
                child.setIcon(0, file_icon)
                child.setData(0, QtCore.Qt.UserRole, "file")
                parent.addChild(child)
            self.macro_tree.addTopLevelItem(parent)

        self.macro_tree.expandAll()
        if self.macro_tree.topLevelItemCount() == 0:
            item = QtWidgets.QTreeWidgetItem(["No macros found. Put .json macro files in the 'macros' folder."])
            self.macro_tree.addTopLevelItem(item)
        self.statusBar().showMessage("Macros reloaded", 2000)

    def release_all_keys(self):
        for key in list(self.held_keys):
            try:
                keyboard.release(key)
            except Exception:
                pass
        self.held_keys.clear()

    def closeEvent(self, event):
        self.stop_macro()
        self.release_all_keys()
        self.hide_crosshair_overlay()
        event.accept()

    def show_info_popup(self):
        info_ui = os.path.join(SCRIPT_DIR, "info_dialog.ui")
        if os.path.exists(info_ui):
            try:
                popup = QtWidgets.QDialog(self)
                uic.loadUi(info_ui, popup)
                if hasattr(popup, 'close_btn'):
                    popup.close_btn.clicked.connect(popup.accept)
                geom_key = "info_dialog"
                geom_value = self.settings.get("dialog_geometries", {}).get(geom_key)
                if geom_value:
                    try:
                        popup.restoreGeometry(QtCore.QByteArray.fromHex(bytes(geom_value, "utf-8")))
                    except Exception: pass
                def save_dialog_geom():
                    dialog_geom = popup.saveGeometry().toHex().data().decode()
                    dg = self.settings.get("dialog_geometries", {})
                    dg[geom_key] = dialog_geom
                    self.settings["dialog_geometries"] = dg
                    self.save_settings()
                popup.finished.connect(save_dialog_geom)
                popup.exec_()
            except Exception:
                QtWidgets.QMessageBox.information(self, "Info",
                    "Division 2 Macro Player\nModern Material Theme.\nEdit info_dialog.ui to customize this popup.")
        else:
            QtWidgets.QMessageBox.information(self, "Info",
                "Division 2 Macro Player\nModern Material Theme.\ninfo_dialog.ui not found.")

    def show_settings_popup(self):
        settings_ui = os.path.join(SCRIPT_DIR, "settings_dialog.ui")
        if os.path.exists(settings_ui):
            try:
                popup = QtWidgets.QDialog(self)
                uic.loadUi(settings_ui, popup)
                geom_key = "settings_dialog"
                geom_value = self.settings.get("dialog_geometries", {}).get(geom_key)
                if geom_value:
                    try:
                        popup.restoreGeometry(QtCore.QByteArray.fromHex(bytes(geom_value, "utf-8")))
                    except Exception: pass
                def save_dialog_geom():
                    dialog_geom = popup.saveGeometry().toHex().data().decode()
                    dg = self.settings.get("dialog_geometries", {})
                    dg[geom_key] = dialog_geom
                    self.settings["dialog_geometries"] = dg
                    self.save_settings()
                popup.finished.connect(save_dialog_geom)

                hotkey_fields = [
                    ("hotkey_edit", "hide_hotkey"),
                    ("play_hotkey_edit", "play_hotkey"),
                    ("pause_hotkey_edit", "pause_hotkey"),
                    ("stop_hotkey_edit", "stop_hotkey"),
                    ("reload_hotkey_edit", "reload_hotkey"),
                    ("crosshair_toggle_hotkey_edit", "crosshair_toggle_hotkey"),
                ]
                for field_name, setting_name in hotkey_fields:
                    orig = getattr(popup, field_name, None)
                    if orig is not None:
                        layout = orig.parent().layout() if orig.parent() else popup.layout()
                        if isinstance(layout, QtWidgets.QGridLayout):
                            found = False
                            for row in range(layout.rowCount()):
                                for col in range(layout.columnCount()):
                                    item = layout.itemAtPosition(row, col)
                                    if item and item.widget() == orig:
                                        hotkey = HotkeyLineEdit(popup)
                                        hotkey.setText(self.settings.get(setting_name, ""))
                                        hotkey.key_captured.connect(lambda key, h=hotkey: h.setText(key))
                                        orig.hide()
                                        layout.removeWidget(orig)
                                        orig.setParent(None)
                                        orig.deleteLater()
                                        layout.addWidget(hotkey, row, col)
                                        setattr(popup, field_name, hotkey)
                                        found = True
                                        break
                                if found:
                                    break
                        else:
                            idx = layout.indexOf(orig)
                            hotkey = HotkeyLineEdit(popup)
                            hotkey.setText(self.settings.get(setting_name, ""))
                            hotkey.key_captured.connect(lambda key, h=hotkey: h.setText(key))
                            orig.hide()
                            layout.removeWidget(orig)
                            orig.setParent(None)
                            orig.deleteLater()
                            layout.insertWidget(idx, hotkey)
                            setattr(popup, field_name, hotkey)

                if hasattr(popup, 'choose_theme_btn'):
                    def choose_theme():
                        if not os.path.exists(THEME_DIR):
                            QtWidgets.QMessageBox.warning(popup, "Theme Error",
                                f"Theme folder '{THEME_DIR}' does not exist.\nPlease create the folder and add .xml theme files.")
                            return
                        themes = [f for f in os.listdir(THEME_DIR) if f.endswith(".xml")]
                        if not themes:
                            QtWidgets.QMessageBox.warning(popup, "Theme Error",
                                f"No themes found in {THEME_DIR}.\nAdd .xml theme files to this folder.")
                            return
                        display_names = []
                        name_to_file = {}
                        for f in themes:
                            try:
                                pretty = os.path.splitext(f)[0].replace("_500", "").replace("_", " ").title()
                                display_names.append(pretty)
                                name_to_file[pretty] = f
                            except Exception:
                                continue
                        if not display_names:
                            QtWidgets.QMessageBox.warning(popup, "Theme Error",
                                f"No valid theme files found in {THEME_DIR}.\nAdd .xml theme files to this folder.")
                            return
                        theme_file = self.settings.get("theme", "dark_teal.xml")
                        try:
                            picker = ThemePickerDialog(popup, display_names, name_to_file, theme_file)
                        except Exception as e:
                            QtWidgets.QMessageBox.warning(popup, "Theme Error",
                                f"Failed to open theme picker dialog: {e}")
                            return
                        result = picker.exec_()
                        if result == QtWidgets.QDialog.Accepted and picker.get_selected_theme_file():
                            selected_theme_file = picker.get_selected_theme_file()
                            self.settings["theme"] = selected_theme_file
                            theme_path = os.path.join(THEME_DIR, selected_theme_file)
                            if not os.path.exists(theme_path):
                                QtWidgets.QMessageBox.warning(popup, "Theme Error",
                                    f"Theme file not found: {theme_path}")
                                return
                            try:
                                apply_stylesheet(QtWidgets.QApplication.instance(), theme=theme_path)
                                for widget in QtWidgets.QApplication.instance().allWidgets():
                                    widget.style().unpolish(widget)
                                    widget.style().polish(widget)
                                    widget.update()
                                self.loop_toggle.update_ui()
                                self.save_settings()
                            except Exception as e:
                                print(f"Theme application failed: {e}")
                        else:
                            prev_theme_file = self.settings.get("theme", "dark_teal.xml")
                            prev_theme_path = os.path.join(THEME_DIR, prev_theme_file)
                            if os.path.exists(prev_theme_path):
                                try:
                                    apply_stylesheet(QtWidgets.QApplication.instance(), theme=prev_theme_path)
                                    for widget in QtWidgets.QApplication.instance().allWidgets():
                                        widget.style().unpolish(widget)
                                        widget.style().polish(widget)
                                        widget.update()
                                    self.loop_toggle.update_ui()
                                except Exception as e:
                                    print(f"Theme application failed: {e}")
                    popup.choose_theme_btn.clicked.connect(choose_theme)
                if hasattr(popup, 'open_macro_btn'):
                    popup.open_macro_btn.clicked.connect(open_macro_folder)
                def save_settings():
                    try:
                        for field_name, setting_name in hotkey_fields:
                            hotkey_field = getattr(popup, field_name, None)
                            if hotkey_field is not None:
                                self.settings[setting_name] = hotkey_field.text()
                        self.save_settings()
                        self.setup_hotkey_global()
                        self.setup_action_hotkeys()
                        popup.accept()
                    except Exception as e:
                        QtWidgets.QMessageBox.warning(popup, "Settings Error", f"Failed to save settings: {e}")
                        print("Settings save error:", e)
                if hasattr(popup, 'save_btn'):
                    popup.save_btn.clicked.connect(save_settings)
                popup.exec_()
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Settings Error", f"Could not open settings dialog: {e}")
        else:
            QtWidgets.QMessageBox.information(self, "Settings",
                "settings_dialog.ui not found. Basic settings unavailable.")

    def show_crosshair_picker(self):
        # Check if crosshair picker module is available
        if not CROSSHAIR_PICKER_AVAILABLE:
            QtWidgets.QMessageBox.warning(
                self, 
                "Crosshair Picker Unavailable", 
                "The crosshair picker feature is not available.\n\n"
                "Required module 'crosshair_picker.py' is missing.\n"
                "Please ensure all application files are present."
            )
            self.statusBar().showMessage("Crosshair picker unavailable - missing module", 5000)
            return
        
        try:
            live_overlay = self.crosshair_overlay if (self.crosshair_enabled and self.crosshair_settings) else None
            picker = CrosshairPickerDialog(
                CROSSHAIR_FOLDER,
                crosshair_settings=self.crosshair_settings if self.crosshair_settings else {},
                crosshair_enabled=self.crosshair_enabled,
                parent=self,
                live_overlay=live_overlay
            )
            picker.settings_changed.connect(self.on_picker_overlay_changed)
            result = picker.exec_()
            picker.settings_changed.disconnect(self.on_picker_overlay_changed)
            if result == QtWidgets.QDialog.Accepted:
                self.crosshair_settings = picker.get_crosshair_settings()
                self.crosshair_enabled = self.crosshair_settings.get("enabled", False)
                self.settings["crosshair_settings"] = self.crosshair_settings
                self.settings["crosshair_enabled"] = self.crosshair_enabled
                self.save_settings()
                self.create_crosshair_overlay()
            else:
                if self.crosshair_overlay:
                    self.crosshair_overlay.close()
                    self.crosshair_overlay = None
                if self.crosshair_enabled and self.crosshair_settings:
                    self.create_crosshair_overlay()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Crosshair Picker Error",
                f"An error occurred while opening the crosshair picker:\n\n{str(e)}\n\n"
                "The crosshair picker feature may not be functioning correctly."
            )
            self.statusBar().showMessage(f"Crosshair picker error: {str(e)}", 5000)
            print(f"Crosshair picker error: {e}")

    def on_picker_overlay_changed(self, settings):
        # Check if crosshair overlay module is available
        if not CROSSHAIR_OVERLAY_AVAILABLE:
            self.statusBar().showMessage("Crosshair overlay unavailable - missing module", 3000)
            return
        
        try:
            if self.crosshair_overlay:
                if settings.get("enabled", False):
                    self.crosshair_overlay.update_settings(settings)
                    self.crosshair_overlay.show()
                else:
                    self.crosshair_overlay.hide()
            else:
                if settings.get("enabled", False):
                    self.crosshair_overlay = CrosshairOverlay(
                        style=settings.get("style", "lines"),
                        color=settings.get("color", "#ff5ea2"),
                        size=settings.get("size", 40),
                        opacity=settings.get("opacity", 1.0),
                        image_path=settings.get("image_path"),
                        thickness=settings.get("thickness", 3),
                        gap=settings.get("gap", 0),
                        outline_enabled=settings.get("outline_enabled", False),
                        outline_color=settings.get("outline_color", "#000000"),
                        outline_thickness=settings.get("outline_thickness", 2),
                        dot_enabled=settings.get("dot_enabled", False),
                        dot_color=settings.get("dot_color", settings.get("color", "#ff5ea2")),
                        dot_size=settings.get("dot_size", 6),
                    )
                    self.crosshair_overlay.show()
        except Exception as e:
            self.statusBar().showMessage(f"Crosshair overlay error: {str(e)}", 5000)
            print(f"Crosshair overlay error: {e}")

    def create_crosshair_overlay(self):
        # Check if crosshair overlay module is available
        if not CROSSHAIR_OVERLAY_AVAILABLE:
            self.statusBar().showMessage("Crosshair overlay unavailable - missing module", 3000)
            return
        
        try:
            if self.crosshair_overlay:
                self.crosshair_overlay.close()
                self.crosshair_overlay = None
            if self.crosshair_enabled and self.crosshair_settings:
                s = self.crosshair_settings
                self.crosshair_overlay = CrosshairOverlay(
                    style=s.get("style", "lines"),
                    color=s.get("color", "#ff5ea2"),
                    size=s.get("size", 40),
                    opacity=s.get("opacity", 1.0),
                    image_path=s.get("image_path"),
                    thickness=s.get("thickness", 3),
                    gap=s.get("gap", 0),
                    outline_enabled=s.get("outline_enabled", False),
                    outline_color=s.get("outline_color", "#000000"),
                    outline_thickness=s.get("outline_thickness", 2),
                    dot_enabled=s.get("dot_enabled", False),
                    dot_color=s.get("dot_color", s.get("color", "#ff5ea2")),
                    dot_size=s.get("dot_size", 6),
                )
                self.crosshair_overlay.show()
        except Exception as e:
            self.statusBar().showMessage(f"Crosshair overlay creation error: {str(e)}", 5000)
            print(f"Crosshair overlay creation error: {e}")
            self.crosshair_overlay = None

    def hide_crosshair_overlay(self):
        try:
            if self.crosshair_overlay:
                self.crosshair_overlay.close()
                self.crosshair_overlay = None
        except Exception as e:
            self.statusBar().showMessage(f"Error hiding crosshair overlay: {str(e)}", 3000)
            print(f"Error hiding crosshair overlay: {e}")
            self.crosshair_overlay = None

    def get_item_category(self, item):
        parent = item.parent()
        if parent and parent.data(0, QtCore.Qt.UserRole) == "category":
            return parent.text(0)
        return None

    def get_macro_info(self, category, macro_file):
        if macro_file.endswith(".json"):
            if category:
                macro_path = os.path.join(MACRO_FOLDER, category, macro_file)
            else:
                macro_path = os.path.join(MACRO_FOLDER, macro_file)
            if os.path.exists(macro_path):
                try:
                    with open(macro_path, "r", encoding="utf-8") as f:
                        actions = json.load(f)
                    if isinstance(actions, list):
                        for entry in actions:
                            if isinstance(entry, dict) and "info" in entry:
                                return entry["info"]
                except Exception:
                    return ""
        return ""

    def start_macro_timer(self):
        if self.macro_actions and self.macro_running and not self.macro_paused:
            try:
                delay = float(self.macro_actions[self.macro_index].get("delay", 0.1)) * 1000
            except Exception:
                delay = 100
            self.macro_timer.start(int(delay))

    def execute_macro_step(self):
        while self.macro_running and self.macro_index < len(self.macro_actions):
            action = self.macro_actions[self.macro_index]
            if "info" in action:
                self.macro_index += 1
                continue
            act = str(action.get("action", "send")).lower()
            key = str(action.get("key", "")).strip().lower()
            try:
                delay = float(action.get("delay", 0.1)) * 1000
            except Exception:
                delay = 100
            if not key:
                self.notifier.error_signal.emit(f"Macro step error: Step {self.macro_index+1} has an empty or missing key, skipping...")
                self.macro_index += 1
                continue
            try:
                if act in ["press"]:
                    keyboard.press(key)
                    self.held_keys.add(key)
                elif act in ["release"]:
                    keyboard.release(key)
                    self.held_keys.discard(key)
                elif act in ["send"]:
                    keyboard.press_and_release(key)
            except Exception as e:
                self.notifier.error_signal.emit(f"Macro step error: {e}")
            self.macro_index += 1
            self.macro_timer.start(int(delay))
            return
        if self.macro_index >= len(self.macro_actions):
            if self.loop_toggle.isChecked():
                self.macro_index = 0
                self.statusBar().showMessage("Looping macro", 1500)
                self.start_macro_timer()
            else:
                self.stop_macro()

    def update_loop_checkbox_style(self, checked=None):
        self.loop_toggle.update_ui()

    def load_settings(self):
        settings = DEFAULT_SETTINGS.copy()
        if os.path.exists(SETTINGS_FILE):
            try:
                loaded = json.load(open(SETTINGS_FILE, "r", encoding="utf-8"))
                settings.update(loaded)
            except Exception:
                pass
        if "macro_hotkeys" not in settings or not isinstance(settings["macro_hotkeys"], dict):
            settings["macro_hotkeys"] = {}
        if "dialog_geometries" not in settings or not isinstance(settings["dialog_geometries"], dict):
            settings["dialog_geometries"] = {}
        return settings

    def save_settings(self):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            pass

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
