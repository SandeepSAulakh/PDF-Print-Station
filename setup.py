import os
import sys
import venv
import subprocess
from pathlib import Path
import shutil

def find_project_root():
    """Find the project root directory containing setup.py"""
    # Get the directory containing setup.py
    setup_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check if required files exist
    required_files = ['requirements.txt', 'pdf_printer_app.py']
    if all(os.path.exists(os.path.join(setup_dir, f)) for f in required_files):
        return setup_dir
    
    print("Error: Cannot find project files!")
    print(f"Please run this script from the PDF Print Station directory")
    print(f"Current directory: {os.getcwd()}")
    print(f"Script directory: {setup_dir}")
    print("Files found:", os.listdir(setup_dir))
    return None

def create_virtual_environment(project_root):
    """Create virtual environment in the project directory"""
    print("Creating virtual environment...")
    venv_path = os.path.join(project_root, "venv")
    venv.create(venv_path, with_pip=True)
    return venv_path

def get_python_executable(venv_path):
    """Get the correct Python executable path for the platform"""
    if sys.platform == "win32":
        return os.path.join(venv_path, "Scripts", "python.exe")
    return os.path.join(venv_path, "bin", "python")

def install_requirements(project_root, python_exe):
    """Install required packages"""
    print("Installing requirements...")
    requirements_path = os.path.join(project_root, 'requirements.txt')
    
    if not os.path.exists(requirements_path):
        print(f"Error: requirements.txt not found in {project_root}")
        return False
    
    try:
        subprocess.check_call([python_exe, "-m", "pip", "install", "-r", requirements_path])
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error installing requirements: {e}")
        return False

def create_directories(project_root):
    """Create necessary application directories"""
    print("Creating necessary directories...")
    directories = ['collections', 'temp_previews', 'assets']
    for directory in directories:
        dir_path = os.path.join(project_root, directory)
        Path(dir_path).mkdir(exist_ok=True)

def make_run_script_executable(project_root):
    """Make run scripts executable on Unix systems"""
    if sys.platform == "win32":
        # Check for run.bat
        run_script = os.path.join(project_root, "run.bat")
        if not os.path.exists(run_script):
            # Create run.bat if it doesn't exist
            with open(run_script, 'w') as f:
                f.write('@echo off\ncall venv\\Scripts\\activate\npython pdf_printer_app.py\npause')
    else:
        # Check for run.sh
        run_script = os.path.join(project_root, "run.sh")
        if not os.path.exists(run_script):
            # Create run.sh if it doesn't exist
            with open(run_script, 'w') as f:
                f.write('#!/bin/bash\nsource venv/bin/activate\npython3 pdf_printer_app.py')
        try:
            os.chmod(run_script, 0o755)  # rwxr-xr-x
            print("Made run script executable")
        except Exception as e:
            print(f"Warning: Could not make run script executable: {e}")

def copy_assets(project_root):
    """Copy necessary assets if they don't exist"""
    assets_dir = os.path.join(project_root, 'assets')
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)
    
    # Copy icon files if they don't exist
    icon_files = ['app_icon.png', 'app_icon.svg']
    for icon in icon_files:
        src = os.path.join(project_root, 'assets', icon)
        dst = os.path.join(assets_dir, icon)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)

def main():
    try:
        print("Starting PDF Print Station setup...")
        
        # Find project root directory
        project_root = find_project_root()
        if not project_root:
            return 1
        
        # Change to project directory
        os.chdir(project_root)
        print(f"Working directory: {os.getcwd()}")
        
        # Create virtual environment
        venv_path = create_virtual_environment(project_root)
        python_exe = get_python_executable(venv_path)
        
        # Install requirements
        if not install_requirements(project_root, python_exe):
            return 1
        
        # Create directories and copy assets
        create_directories(project_root)
        copy_assets(project_root)
        
        # Make run script executable
        make_run_script_executable(project_root)
        
        print("\nSetup completed successfully!")
        print("\nTo run the application:")
        if sys.platform == "win32":
            print(f"Double-click 'run.bat' in the installation folder")
        else:
            print(f"Run './run.sh' in the installation folder")
        
    except Exception as e:
        print(f"Error during setup: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())