#!/usr/bin/env python3
"""
ReportIQ - Vulnerability Report Generator
Main application entry point
"""

import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.gui.main_window import MainWindow


def main():
    """Main entry point"""
    print("Starting ReportIQ...")
    
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
