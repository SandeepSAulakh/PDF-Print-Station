# Installation Guide for PDF Print Station

## One-Button Install

1. Download and extract the project files
2. Run the setup script:

On Windows:
```bash
python setup.py
```

On macOS/Linux:
```bash
python3 setup.py
```

The setup script will:
- Create necessary directories
- Set up virtual environment
- Install required dependencies
- Configure the application

## Running the Application

After installation, you have two options:

### Option 1: Using Run Scripts
On Windows:
- Double-click `run.bat`

On macOS/Linux:
- Run `./run.sh` in terminal

### Option 2: Manual Activation
On Windows:
```bash
venv\Scripts\activate
python pdf_printer_app.py
```

On macOS/Linux:
```bash
source venv/bin/activate
python3 pdf_printer_app.py
```

## Requirements

- Python 3.9 or higher
- PyQt5 5.15.9 or higher
- PyMuPDF (fitz) 1.24.12 or higher
- Operating System: Windows, macOS, or Linux

## Troubleshooting

1. Virtual Environment Issues:
   - Make sure Python is in your system PATH
   - Use `python3` instead of `python` on macOS/Linux
   - On Windows, you might need to run: `Set-ExecutionPolicy RemoteSigned -Scope Process`

2. Dependency Issues:
   - Try upgrading pip: `python -m pip install --upgrade pip`
   - Install dependencies one by one if batch install fails

3. Permission Issues:
   - Run terminal as administrator on Windows
   - On macOS/Linux:
     - Make run script executable: `chmod +x run.sh`
     - Use `sudo` if needed for installations

## Need Help?

If you encounter any issues:
1. Check the error message
2. Verify your Python version
3. Create an issue on GitHub with:
   - Your operating system
   - Python version
   - Error message
   - Steps to reproduce