#!/usr/bin/env python3
"""
Test script for TranscoProcessor
"""
import sys
import os
sys.path.append(os.path.dirname(__file__))

try:
    # Import the TranscoProcessor from tms_processor
    from tms_processor import TranscoProcessor

    print("=== Testing TranscoProcessor ===")

    # Create processor instance
    processor = TranscoProcessor()
    print("SUCCESS: TranscoProcessor created successfully")
    print(f"  - HEADER_ROW: {processor.HEADER_ROW}")
    print(f"  - DATA_START_ROW: {processor.DATA_START_ROW}")

    # Test file path
    test_file = "1_05.05.xlsx"
    if not os.path.exists(test_file):
        print(f"ERROR: Test file {test_file} not found")
        sys.exit(1)

    print(f"SUCCESS: Test file found: {test_file}")

    # Try to process the file
    print("\n=== Processing File ===")
    try:
        processed_data = processor.process_excel_file(test_file)
        print("SUCCESS: File processed successfully!")
        print(f"  - Rows: {len(processed_data)}")
        print(f"  - Columns: {len(processed_data.columns)}")
        print(f"  - Title info: {processor.title_info}")
        print(f"  - Summary stats: {processor.summary_stats}")

        # Show first few rows
        if len(processed_data) > 0:
            print("\n=== First 3 rows ===")
            print(processed_data.head(3).to_string())

    except Exception as e:
        print(f"ERROR: Error processing file: {e}")
        import traceback
        traceback.print_exc()

except Exception as e:
    print(f"ERROR: Error importing or creating TranscoProcessor: {e}")
    import traceback
    traceback.print_exc()