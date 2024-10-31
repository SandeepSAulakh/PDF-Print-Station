# Installation Guide for PDF Print Station

## One-Button Install

1. Download and extract the project files
2. Run the setup script:

On Windows:
```bash
python setup.py
```

The setup script will:
- Create necessary directories
- Set up virtual environment
- Install required dependencies
- Configure the application

## Manual Installation

1. Clone the repository:
```bash
git clone https://github.com/SandeepSAulakh/PDF-Print-Station.git
cd PDF-Print-Station
```

2. Create and activate virtual environment:

On Windows:
```bash
python -m venv venv
venv\Scripts\activate
```

On macOS/Linux:
```bash
python -m venv venv
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Requirements

- Python 3.9 or higher
- PyQt5 5.15.9 or higher
- PyMuPDF (fitz) 1.24.12 or higher
- Operating System: Windows, macOS, or Linux

## Running the Application

1. After installation, run:
```bash
python pdf_printer_app.py
```

## Troubleshooting

1. Virtual Environment Issues:
   - Make sure Python is in your system PATH
   - Use `python3` instead of `python` on some systems
   - On Windows, you might need to run: `Set-ExecutionPolicy RemoteSigned -Scope Process`

2. Dependency Issues:
   - Try upgrading pip: `python -m pip install --upgrade pip`
   - Install dependencies one by one if batch install fails

3. Permission Issues:
   - Run terminal as administrator on Windows
   - Use `sudo` on Linux/macOS if needed

## Need Help?

If you encounter any issues:
1. Check the error message
2. Verify your Python version
3. Create an issue on GitHub with:
   - Your operating system
   - Python version
   - Error message
   - Steps to reproduce 