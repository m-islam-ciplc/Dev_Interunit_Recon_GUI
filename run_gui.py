#!/usr/bin/env python3
"""
GUI Launcher for Interunit Loan Matcher
Simple launcher script to run the GUI application
"""

import sys
import os

# Add current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from main_gui import main
    if __name__ == "__main__":
        main()
except ImportError as e:
    print(f"Error importing GUI modules: {e}")
    print("Please install required dependencies:")
    print("pip install -r requirements_gui.txt")
    sys.exit(1)
except Exception as e:
    print(f"Error starting GUI application: {e}")
    sys.exit(1)
