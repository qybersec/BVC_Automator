#!/usr/bin/env python3
"""
Script to clean tms_processor.py by removing duplicate processor classes
"""
import os

def clean_tms_processor():
    """Remove duplicate processor classes that are now in separate modules"""

    # Read the original file
    with open('tms_processor.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Find the lines to remove
    start_utcfs = None
    start_transco = None
    end_transco = None
    start_gui = None

    for i, line in enumerate(lines):
        if line.strip() == 'class UTCFSProcessor:':
            start_utcfs = i
        elif line.strip() == 'class TranscoProcessor:':
            start_transco = i
        elif line.strip() == 'class ModernTMSProcessorGUI:':
            start_gui = i
            break

    print(f"Found UTCFSProcessor at line {start_utcfs}")
    print(f"Found TranscoProcessor at line {start_transco}")
    print(f"Found GUI class at line {start_gui}")

    if start_utcfs is not None and start_gui is not None:
        # Keep everything before UTCFSProcessor and everything from GUI onwards
        new_lines = lines[:start_utcfs] + lines[start_gui:]

        # Write the cleaned file
        with open('tms_processor_clean.py', 'w', encoding='utf-8') as f:
            f.writelines(new_lines)

        print(f"Removed {start_gui - start_utcfs} lines of duplicate processor classes")
        print("Created tms_processor_clean.py")
    else:
        print("Could not find the required class markers")

if __name__ == '__main__':
    clean_tms_processor()