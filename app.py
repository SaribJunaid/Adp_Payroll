import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from payroll_app.ui import run_app

if __name__ == "__main__":
    run_app()