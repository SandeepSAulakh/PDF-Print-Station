#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"
export PYTHONPATH="${PYTHONPATH}:${DIR}"
source venv/bin/activate
export QT_MAC_WANTS_LAYER=1
export DISPLAY_NAME="PDF Print Station"
export PYAPP_DISPLAY_NAME="PDF Print Station"
exec pythonw pdf_printer_app.py 