"""
Launcher Entry Point delegating to Tkinter GUI Application (gui_app.py)
"""
import sys
import os

try:
    from gui_app import main
except ImportError:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from gui_app import main

if __name__ == "__main__":
    main()
