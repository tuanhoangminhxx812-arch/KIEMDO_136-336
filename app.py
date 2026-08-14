import os
import sys
import runpy

# Ensure 'dữ liệu' folder is in Python path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(CURRENT_DIR, "dữ liệu")

if DATA_DIR not in sys.path:
    sys.path.insert(0, DATA_DIR)

# Run main app script inside 'dữ liệu/app.py'
target_app = os.path.join(DATA_DIR, "app.py")
runpy.run_path(target_app, run_name="__main__")
