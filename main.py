#!/usr/bin/env python3
"""
BVC Automator - Professional TMS Data Processing Suite
Main Entry Point

This is the primary entry point for the application, following industry best practices
for project structure and maintainability.
"""

import sys
import os
from pathlib import Path

# Add project root to Python path for imports
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def main():
    """Main entry point for BVC Automator"""
    try:
        # Import and run the TMS Processor GUI
        from tms_processor import main as run_gui
        print("Starting BVC Automator - TMS Data Processor Pro")
        run_gui()
        
    except ImportError as e:
        print(f"Failed to import required modules: {e}")
        print("Please ensure all dependencies are installed:")
        print("pip install pandas openpyxl numpy")
        input("Press Enter to exit...")
        sys.exit(1)
        
    except Exception as e:
        print(f"Application error: {e}")
        input("Press Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main()