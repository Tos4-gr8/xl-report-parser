import os
import sys

# ==============================================================================
# GUI Settings
# ==============================================================================

APPEARANCE_MODE = "dark"
DEFAULT_COLOR_THEME = "blue"
WINDOW_GEOMETRY = "510x430"
WINDOW_TITLE = "Excel calc v1.1"


def get_resource_path(relative_path):
    """Returns the absolute path to a resource, supporting PyInstaller builds."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# Path to the application icon
ICON_PATH = get_resource_path("icon.ico")


# ==============================================================================
# Rates & Multipliers
# ==============================================================================

# Personal income tax rate (%) for net salary calculations
TAX_PERCENT = 13

# Report 1: RT Operations
RATES_REPORT_1 = {
    "pos_movement": 5.30,  # Items picked (Movements)
    "box_movement": 2.00,  # Boxes picked (Movements)
    "pos_retail": 6.80,    # Items placed (Retail)
    "box_retail": 3.80,    # Boxes placed (Retail)
    "pos_wholesale": 4.00, # Items picked (Wholesale picking)
    "box_wholesale": 2.20, # Boxes picked (Wholesale picking)
    "internal_move": 18.00 # Internal transfers
}

# Report 2: Pallet Splitting & Receiving
RATES_REPORT_2 = {
    "wholesale_box": 1.50,
    "wholesale_pos": 7.00,
    "retail_box": 3.80,
    "retail_pos": 6.80,
    "receipt_box": 2.25,
    "receipt_pos": 1.90,
    "ssci_click": 1.20      # Includes scanned SSCI (Data Matrix codes)
}

# Report 3: Shipping & Dispatch
RATES_REPORT_3 = {
    "base_rate": 4.90       # Rate per single dispatch row
}

# Report 4: Inventory Transfers (Room Cleanup)
RATES_REPORT_4 = {
    "retail_sector": 55.00, # Retail assembly area
    "opt_position": 12.10,  # Wholesale positions
    "opt_box": 5.80         # Wholesale maximums
}
