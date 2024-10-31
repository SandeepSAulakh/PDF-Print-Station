import PyInstaller.__main__
import os
import shutil

def build_app():
    print("Building PDF Print Station...")
    
    # Clean previous builds
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")
        
    # Create necessary directories
    os.makedirs("dist/PDF Print Station/assets", exist_ok=True)
    os.makedirs("dist/PDF Print Station/collections", exist_ok=True)
    os.makedirs("dist/PDF Print Station/temp_previews", exist_ok=True)
    
    # PyInstaller options
    PyInstaller.__main__.run([
        'pdf_printer_app.py',
        '--onefile',
        '--windowed',
        '--name=PDF Print Station',
        '--icon=assets/app_icon.ico',
        '--add-data=assets;assets',
        '--noconsole',
    ])
    
    # Copy additional files
    shutil.copy2("assets/app_icon.svg", "dist/PDF Print Station/assets/")
    
    print("Build complete! Check the 'dist' folder.")

if __name__ == "__main__":
    build_app() 