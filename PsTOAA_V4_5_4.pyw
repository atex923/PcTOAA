"""Launch the PsTOAA GUI without opening a console window on Windows."""

from pathlib import Path
import sys


APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PsTOAA_V4_5_4 import main


if __name__ == "__main__":
    main()
