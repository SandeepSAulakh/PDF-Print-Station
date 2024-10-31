import os
import sys
import venv
import subprocess
from pathlib import Path

def create_virtual_environment():
    print("Creating virtual environment...")
    venv.create("venv", with_pip=True)

def get_python_executable():
    # Handle both Windows and Unix systems
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

def make_run_script_executable():
    # Make run.sh executable on Unix systems
    if sys.platform != "win32":
        run_script = Path("run.sh")
        if run_script.exists():
            try:
                os.chmod(run_script, 0o755)  # rwxr-xr-x
                print("Made run.sh executable")
            except Exception as e:
                print(f"Warning: Could not make run.sh executable: {e}")

def main():
    try:
        print("Starting PDF Print Station setup...")
        
        # Create virtual environment
        create_virtual_environment()
        
        # Install requirements
        install_requirements()
        
        # Create necessary directories
        create_directories()
        
        # Make run script executable on Unix
        make_run_script_executable()
        
        print("\nSetup completed successfully!")
        print("\nTo run the application:")
        if sys.platform == "win32":
            print("Option 1: Double-click run.bat")
            print("Option 2: Run these commands:")
            print("1. venv\\Scripts\\activate")
            print("2. python pdf_printer_app.py")
        else:
            print("Option 1: Run ./run.sh")
            print("Option 2: Run these commands:")
            print("1. source venv/bin/activate")
            print("2. python3 pdf_printer_app.py")
        
    except Exception as e:
        print(f"Error during setup: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())