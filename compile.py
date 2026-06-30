import PyInstaller.__main__
import os
import shutil
import glob
from datetime import datetime


try:
    version = os.popen("git rev-parse --short HEAD").read().strip()
    if not version:
        version = "unknown"
except Exception:
    version = "unknown"

# Define the script to compile
script_name = "main.py"
current_date = datetime.now().strftime("%d.%m.%Y")
app_name = f"Конвертер ПФУ ver.{version} ({current_date})"

# PyInstaller arguments
args = [
    script_name,
    f"--name={app_name}",
    "--onefile",                    # Create a single executable file
    "--noconsole",                  # Don't show the terminal window
    # "--windowed",                 # Same as --noconsole
    "--clean",                      # Clean cache
    "--add-data=Шаблон ПФУ.xlsx:.",   # Include template file
    "--add-data=icon.ico:.",        # Include ico for GUI
    "--icon=icon.ico",              # Exe icon
]

print(f"Starting compilation of {script_name}...")
PyInstaller.__main__.run(args)

# Cleanup and renaming
try:
    if os.path.exists("build"):
        shutil.rmtree("build")
    
    # Remove spec files
    for spec_file in glob.glob("*.spec"):
        os.remove(spec_file)
    
    print(f"Compilation finished. Executable created: dist/{app_name}.exe")
except Exception as e:
    print(f"An error occurred during cleanup or renaming: {e}")
