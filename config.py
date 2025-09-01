# =============================================================================
# CONFIGURATION MODULE
# =============================================================================
# Centralized configuration for all regex patterns and system settings

# LC Number extraction pattern (modify if your LC numbers have different format)
LC_PATTERN = r'\b(?:L/C|LC)[-\s]?\d+[/\s]?\d*\b'

# OLD PO Pattern (commented out - specific format matching)
# PO_PATTERN = r'[A-Z0-9]{2,4}/PO/\d{4}/\d{1,2}/\d{3,6}'

# NEW PO Pattern - Dynamic approach using /PO/ as anchor
# Finds PO blocks that are continuous text with /PO/ in them
# More flexible boundaries to catch PO numbers at sentence edges
# Examples: CIL/C//PO//11/2024, CCEL/Reno//PO///2024/9/191024, G24/PO/2024/9/29505
PO_PATTERN = r'(?:^|\s)([A-Z0-9/]+/PO/[A-Z0-9/]+)(?:\s|$|[,\.])'

# Amount matching tolerance (for rounding differences)
# AMOUNT_TOLERANCE = 0.01  # ❌ UNUSED - removed since all matching uses exact amounts

# File paths and output settings
# INPUT_FILE1_PATH = "Input Files/Interunit Steel.xlsx"
# INPUT_FILE2_PATH = "Input Files/Interunit GeoTex.xlsx"

INPUT_FILE1_PATH = "Input Files/Pole Book STEEL.xlsx"
INPUT_FILE2_PATH = "Input Files/Steel Book POLE.xlsx"

# INPUT_FILE1_PATH = "Input Files/Steel book Trans.xlsx"
# INPUT_FILE2_PATH = "Input Files/Trans book Steel.xlsx"

OUTPUT_FOLDER = "Output"
OUTPUT_SUFFIX = "_MATCHED.xlsx"
SIMPLE_SUFFIX = "_SIMPLE.xlsx"
CREATE_SIMPLE_FILES = False
# CREATE_ALT_FILES = False  # ❌ UNUSED - commenting out
VERBOSE_DEBUG = True

def print_configuration():
    """Print current configuration settings."""
    print("=" * 60)
    print("CURRENT CONFIGURATION")
    print("=" * 60)
    print(f"Input File 1: {INPUT_FILE1_PATH}")
    print(f"Input File 2: {INPUT_FILE2_PATH}")
    print(f"Output Folder: {OUTPUT_FOLDER}")
    print(f"Output Suffix: {OUTPUT_SUFFIX}")
    print(f"Simple Files: {'Yes' if CREATE_SIMPLE_FILES else 'No'}")
    # print(f"Alternative Files: {'Yes' if CREATE_ALT_FILES else 'No'}")  # ❌ UNUSED - commenting out
    print(f"Verbose Debug: {'Yes' if VERBOSE_DEBUG else 'No'}")
    print(f"LC Pattern: {LC_PATTERN}")
    print(f"PO Pattern: {PO_PATTERN}")
    # print(f"Amount Tolerance: {AMOUNT_TOLERANCE}")  # ❌ UNUSED - removed
    print("=" * 60)

def update_configuration():
    """Interactive configuration update (for future use)."""
    print("To update configuration, modify the variables in config.py:")
    print("1. INPUT_FILE1_PATH - Path to your first Excel file")
    print("2. INPUT_FILE2_PATH - Path to your second Excel file")
    print("3. OUTPUT_FOLDER - Where to save output files")
    print("4. OUTPUT_SUFFIX - Suffix for matched files")
    print("5. CREATE_SIMPLE_FILES - Whether to create simple test files")
    # print("6. CREATE_ALT_FILES - Whether to create alternative files")  # ❌ UNUSED - commenting out
    print("7. VERBOSE_DEBUG - Whether to show detailed debug output")
    print("8. LC_PATTERN - Regex pattern for LC number extraction")
    print("9. PO_PATTERN - Regex pattern for PO number extraction")
    # print("10. AMOUNT_TOLERANCE - Tolerance for amount matching (0 for exact)")  # ❌ UNUSED - removed
