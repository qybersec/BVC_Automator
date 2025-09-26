#!/usr/bin/env python3
"""
Test script for updated TranscoProcessor with complete Basic logic
"""
import sys
import os
sys.path.append(os.path.dirname(__file__))

try:
    from tms_processor import TranscoProcessor

    print("=== Testing Updated TranscoProcessor ===")

    # Create processor instance
    processor = TranscoProcessor()
    print("SUCCESS: TranscoProcessor created successfully")

    # Test file path
    test_file = "1_05.05.xlsx"
    if not os.path.exists(test_file):
        print(f"ERROR: Test file {test_file} not found")
        sys.exit(1)

    print(f"SUCCESS: Test file found: {test_file}")

    # Try to process the file
    print("\n=== Processing File with Complete Basic Logic ===")
    try:
        processed_data = processor.process_excel_file(test_file)
        print("SUCCESS: File processed successfully with complete Basic logic!")
        print(f"  - Rows: {len(processed_data)}")
        print(f"  - Columns: {len(processed_data.columns)}")

        # Check for key columns that should be properly mapped
        expected_columns = ['Selected Carrier', 'Least Cost Carrier', 'Selected Total Cost', 'Least Cost Total Cost', 'Potential Savings']
        found_columns = [col for col in expected_columns if col in processed_data.columns]
        print(f"  - Key columns found: {found_columns}")

        print(f"  - Title info: {processor.title_info}")
        print(f"  - Summary stats: {processor.summary_stats}")

        # Check if we have valid cost data now
        if 'Selected Total Cost' in processed_data.columns:
            selected_total = processed_data['Selected Total Cost'].sum()
            print(f"  - Total Selected Cost: ${selected_total:,.2f}")

        if 'Least Cost Total Cost' in processed_data.columns:
            least_total = processed_data['Least Cost Total Cost'].sum()
            print(f"  - Total Least Cost: ${least_total:,.2f}")

    except Exception as e:
        print(f"ERROR: Error processing file: {e}")
        import traceback
        traceback.print_exc()

except Exception as e:
    print(f"ERROR: Error importing or creating TranscoProcessor: {e}")
    import traceback
    traceback.print_exc()