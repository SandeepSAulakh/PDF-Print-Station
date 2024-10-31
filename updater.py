import os
import sys
import json
import requests
from packaging import version
from PyQt5.QtWidgets import QMessageBox, QProgressDialog
from PyQt5.QtCore import Qt

CURRENT_VERSION = "0.0.1"  # Current app version
GITHUB_API_URL = "https://api.github.com/repos/SandeepSAulakh/PDF-Print-Station/releases/latest"

def check_for_updates(parent=None):
    try:
        # Show checking progress
        checking_dialog = QProgressDialog("Checking for updates...", None, 0, 0, parent)
        checking_dialog.setWindowModality(Qt.WindowModal)
        checking_dialog.setWindowTitle("Update Check")
        checking_dialog.setMinimumDuration(0)
        checking_dialog.setValue(0)
        
        # Get latest release info from GitHub
        response = requests.get(GITHUB_API_URL)
        latest_release = response.json()
        latest_version = latest_release['tag_name'].lstrip('v')
        
        checking_dialog.close()
        
        # Compare versions
        if version.parse(latest_version) > version.parse(CURRENT_VERSION):
            # Ask user if they want to update
            reply = QMessageBox.question(
                parent,
                "Update Available",
                f"A new version ({latest_version}) is available.\n\n" +
                f"Current version: {CURRENT_VERSION}\n" +
                f"Latest version: {latest_version}\n\n" +
                "Would you like to update now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # Download and install update
                download_url = latest_release['assets'][0]['browser_download_url']
                download_and_install_update(download_url, parent)
                return True
        else:
            if parent:  # Only show if called from GUI
                QMessageBox.information(parent, "No Updates", 
                    "You are running the latest version.")
                
    except Exception as e:
        print(f"Error checking for updates: {e}")
        if parent:
            QMessageBox.warning(parent, "Update Check Failed", 
                f"Failed to check for updates:\n{str(e)}")
    
    return False

def download_and_install_update(url, parent=None):
    try:
        # Create progress dialog
        progress = QProgressDialog("Downloading update...", "Cancel", 0, 100, parent)
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowTitle("Updating")
        progress.setMinimumDuration(0)
        
        # Download new version
        response = requests.get(url, stream=True)
        total_size = int(response.headers.get('content-length', 0))
        block_size = 8192
        temp_file = "update.zip"
        
        with open(temp_file, 'wb') as f:
            downloaded = 0
            for chunk in response.iter_content(chunk_size=block_size):
                if progress.wasCanceled():
                    os.remove(temp_file)
                    return
                    
                downloaded += len(chunk)
                f.write(chunk)
                
                # Update progress
                if total_size:
                    percent = int((downloaded / total_size) * 100)
                    progress.setValue(percent)
                    progress.setLabelText(f"Downloading update... {percent}%")
        
        # Extract and install
        progress.setLabelText("Installing update...")
        progress.setValue(0)
        
        import zipfile
        with zipfile.ZipFile(temp_file, 'r') as zip_ref:
            zip_ref.extractall("update")
        
        # Replace current files
        import shutil
        shutil.copytree("update", ".", dirs_exist_ok=True)
        
        # Clean up
        os.remove(temp_file)
        shutil.rmtree("update")
        
        QMessageBox.information(parent, "Update Complete", 
            "Update installed successfully.\nPlease restart the application.")
            
    except Exception as e:
        QMessageBox.critical(parent, "Update Error", 
            f"Failed to install update:\n{str(e)}")
    finally:
        progress.close()