#!/usr/bin/env python3
"""
Final script to clean tms_processor.py by removing duplicate processor classes
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

    # Find the first duplicate class
    first_duplicate = None
    for class_name in ['UTCMainProcessor', 'UTCFSProcessor', 'TranscoProcessor']:
        if class_name in markers:
            if first_duplicate is None or markers[class_name] < first_duplicate:
                first_duplicate = markers[class_name]

    if first_duplicate is not None and 'GUI' in markers:
        # Keep everything before the first duplicate class and everything from GUI onwards
        gui_start = markers['GUI']

        # Build the new file
        new_lines = []
        new_lines.extend(lines[:first_duplicate])  # Everything before duplicates
        new_lines.extend(lines[gui_start:])        # GUI class and onwards

        # Write the cleaned file
        with open('tms_processor_final_clean.py', 'w', encoding='utf-8') as f:
            f.writelines(new_lines)

        removed_lines = gui_start - first_duplicate
        print(f"Removed {removed_lines} lines of duplicate processor classes")
        print("Created tms_processor_final_clean.py")

        # Show line counts
        print(f"Original: {len(lines)} lines")
        print(f"Cleaned: {len(new_lines)} lines")
        print(f"Reduction: {len(lines) - len(new_lines)} lines")

        return True

    else:
        print("Could not find the required class markers")
        print(f"Available markers: {list(markers.keys())}")
        return False

if __name__ == '__main__':
    success = clean_tms_processor()
    if success:
        print("\nCleaning completed successfully!")
    else:
        print("\nCleaning failed!")