# Updated Main.py

# This fix ensures the macro browser always uses QTreeWidget for the macro tree pane, replacing any other widget.

from PyQt5.QtWidgets import QTreeWidget

class MacroBrowser:
    def __init__(self):
        self.macro_tree_pane = QTreeWidget()  # Ensuring QTreeWidget is used
        # Additional initialization code...

    # Other methods of the MacroBrowser class...

# Rest of the Main.py code...