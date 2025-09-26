#!/usr/bin/env python3
"""
Improved script to clean tms_processor.py by removing duplicate processor classes
"""
import os

def clean_tms_processor():
    """Remove duplicate processor classes that are now in separate modules"""

    # Read the original file
    with open('tms_processor.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Find the classes to remove
    markers = {}
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped == 'class UTCMainProcessor:':
            markers['UTCMainProcessor'] = i
        elif stripped == 'class UTCFSProcessor:':
            markers['UTCFSProcessor'] = i
        elif stripped == 'class TranscoProcessor:':
            markers['TranscoProcessor'] = i
        elif stripped == 'class ModernTMSProcessorGUI:':
            markers['GUI'] = i
            break

    print(f"Found markers: {markers}")

    if 'UTCFSProcessor' in markers and 'GUI' in markers:
        # Keep everything before the first processor class and everything from GUI onwards
        first_duplicate = markers['UTCFSProcessor']
        gui_start = markers['GUI']

        # Build the new file
        new_lines = []
        new_lines.extend(lines[:first_duplicate])  # Everything before duplicates
        new_lines.extend(lines[gui_start:])        # GUI class and onwards

        # Write the cleaned file
        with open('tms_processor_clean_v2.py', 'w', encoding='utf-8') as f:
            f.writelines(new_lines)

        removed_lines = gui_start - first_duplicate
        print(f"Removed {removed_lines} lines of duplicate processor classes")
        print("Created tms_processor_clean_v2.py")

        # Show line counts
        print(f"Original: {len(lines)} lines")
        print(f"Cleaned: {len(new_lines)} lines")

    else:
        print("Could not find the required class markers")
        print(f"Available markers: {list(markers.keys())}")

if __name__ == '__main__':
    clean_tms_processor()