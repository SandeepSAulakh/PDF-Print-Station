import os
import sys
import venv
import subprocess
import shutil
from pathlib import Path

def setup_project():
    print("Setting up PDF Print Station...")
    
    # Create project directory
    project_dir = Path("pdf-print-station")
    project_dir.mkdir(exist_ok=True)
    os.chdir(project_dir)
    
    # Create virtual environment
    print("\nCreating virtual environment...")
    venv.create("venv", with_pip=True)
    
    # Determine python executable path
    if sys.platform == "win32":
        python_path = "venv\\Scripts\\python.exe"
        pip_path = "venv\\Scripts\\pip.exe"
    else:
        python_path = "venv/bin/python"
        pip_path = "venv/bin/pip"
    
    # Create required directories
    print("\nCreating project structure...")
    Path("assets").mkdir(exist_ok=True)
    Path("collections").mkdir(exist_ok=True)
    Path("temp_previews").mkdir(exist_ok=True)
    
    # Copy application files
    print("\nCopying application files...")
    files_to_copy = {
        "pdf_printer_app.py": "Main application file",
        "requirements.txt": "Dependencies file",
        ".gitignore": "Git ignore file",
        "LICENSE": "License file",
        "README.md": "Documentation",
        "assets/app_icon.svg": "Application icon"
    }
    
    for file, description in files_to_copy.items():
        if Path(file).exists():
            dest = project_dir / file
            dest.parent.mkdir(exist_ok=True)
            shutil.copy2(file, dest)
            print(f"Copied {description}")
    
    # Install dependencies
    print("\nInstalling dependencies...")
    subprocess.run([pip_path, "install", "-r", "requirements.txt"])
    
    print("\nSetup complete! To run the application:")
    if sys.platform == "win32":
        print("1. cd pdf-print-station")
        print("2. venv\\Scripts\\activate")
        print("3. python pdf_printer_app.py")
    else:
        print("1. cd pdf-print-station")
        print("2. source venv/bin/activate")
        print("3. python pdf_printer_app.py")

if __name__ == "__main__":
    try:
        setup_project()
    except Exception as e:
        print(f"\nError during setup: {e}")
        sys.exit(1) 