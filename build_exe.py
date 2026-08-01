"""
PyInstaller Standalone Executable Build Script (Tkinter GUI Version)
=====================================================================
Compiles 'Interactive_PDF_Code/gui_app.py' into a clean, standalone Windows GUI executable:
'Interactive_PDF_Code/Interactive_PDF_Generator_Suite.exe'

Key Configuration:
------------------
- Windowed GUI Mode: Uses --noconsole flag so no CMD window pops up.
- Included Dependencies: tkinter, openpyxl, reportlab, pypdf, PIL (Pillow - required by ReportLab).
- Excluded Heavy Unused Modules: matplotlib, numpy, pandas, scipy, pytest.

To Rebuild the Exe Anytime in the Future:
------------------------------------------
Run command:
    python Interactive_PDF_Code/build_exe.py
"""

import os
import sys
import subprocess

def build_executable():
    print("=" * 76)
    print("       BUILDING STANDALONE GUI EXECUTABLE (Interactive_PDF_Generator_Suite.exe)")
    print("=" * 76)

    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    gui_app_path = os.path.join(root_dir, "Interactive_PDF_Code", "gui_app.py")

    if not os.path.exists(gui_app_path):
        print(f"[ERROR] gui_app.py not found at: '{gui_app_path}'")
        return

    # PyInstaller CLI arguments for Tkinter GUI (Include PIL for ReportLab)
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--noconsole",  # Native GUI application (No CMD prompt window)
        "--clean",
        "--name", "Interactive_PDF_Generator_Suite",
        "--distpath", os.path.join(root_dir, "Interactive_PDF_Code"),
        "--workpath", os.path.join(root_dir, "build"),
        "--specpath", os.path.join(root_dir, "build"),
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        # Exclude heavy unused libraries to optimize exe size (PIL retained for ReportLab!)
        "--exclude-module", "matplotlib",
        "--exclude-module", "numpy",
        "--exclude-module", "pandas",
        "--exclude-module", "scipy",
        "--exclude-module", "pytest",
        "--exclude-module", "IPython",
        "--exclude-module", "notebook",
        gui_app_path
    ]

    print("\n[INFO] Running PyInstaller compilation command...")
    print(" ".join(cmd))
    print("\nProcessing... Please wait...")

    res = subprocess.run(cmd, cwd=root_dir)

    if res.returncode == 0:
        exe_path = os.path.join(root_dir, "Interactive_PDF_Code", "Interactive_PDF_Generator_Suite.exe")
        print("\n" + "=" * 76)
        print("SUCCESS! STANDALONE GUI EXECUTABLE CREATED AT:")
        print(f"   [EXE] {exe_path}")
        print("=" * 76)
    else:
        print("\n[ERROR] PyInstaller build failed. Check output logs above.")

if __name__ == "__main__":
    build_executable()
