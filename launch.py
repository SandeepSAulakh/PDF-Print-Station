#!/usr/bin/env python3
import os
import sys
import subprocess
import venv
import shutil
from pathlib import Path

def is_venv_setup():
    """Check if virtual environment is set up"""
    if sys.platform == "win32":
        return os.path.exists("venv/Scripts/python.exe")
    return os.path.exists("venv/bin/python")

def get_python_path():
    """Get the correct Python executable path"""
    if sys.platform == "win32":
        return "venv\\Scripts\\python.exe"
    return "venv/bin/python"

def setup_and_run():
    """Set up environment and run the application"""
    try:
        # Get the script's directory
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        # Create virtual environment if it doesn't exist
        if not is_venv_setup():
            print("Setting up virtual environment...")
            venv.create("venv", with_pip=True)
            
            # Install dependencies
            python_path = get_python_path()
            subprocess.run([python_path, "-m", "pip", "install", "-r", "requirements.txt"])
        
        # Run the application
        python_path = get_python_path()
        subprocess.run([python_path, "pdf_printer_app.py"])
        
    except Exception as e:
        print(f"Error: {e}")
        input("Press Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    setup_and_run() 