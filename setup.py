import os
import sys
import venv
import subprocess
from pathlib import Path

def create_virtual_environment():
    print("Creating virtual environment...")
    venv.create("venv", with_pip=True)

def get_python_executable():
    if sys.platform == "win32":
        return os.path.join("venv", "Scripts", "python.exe")
    return os.path.join("venv", "bin", "python")

def install_requirements():
    python_exe = get_python_executable()
    print("Installing requirements...")
    subprocess.check_call([python_exe, "-m", "pip", "install", "-r", "requirements.txt"])

def create_directories():
    print("Creating necessary directories...")
    directories = ['collections', 'temp_previews', 'assets']
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)

def main():
    try:
        print("Starting PDF Print Station setup...")
        
        # Create virtual environment
        create_virtual_environment()
        
        # Install requirements
        install_requirements()
        
        # Create necessary directories
        create_directories()
        
        print("\nSetup completed successfully!")
        print("\nTo run the application:")
        if sys.platform == "win32":
            print("1. Run: venv\\Scripts\\activate")
        else:
            print("1. Run: source venv/bin/activate")
        print("2. Run: python pdf_printer_app.py")
        
    except Exception as e:
        print(f"Error during setup: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 