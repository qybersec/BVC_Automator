#!/usr/bin/env python3
"""
Test the new modular structure
"""
import sys
import os
sys.path.append(os.path.dirname(__file__))

try:
    # Test importing the basic processor
    from basic_processor import BasicTMSProcessor
    print("SUCCESS: basic_processor imported")

    # Test importing city processors
    from city_processors import UTCMainProcessor, UTCFSProcessor, TranscoProcessor
    print("SUCCESS: city_processors imported")

    # Test creating instances
    basic = BasicTMSProcessor()
    utc_main = UTCMainProcessor()
    utc_fs = UTCFSProcessor()
    transco = TranscoProcessor()
    print("SUCCESS: All processors created")

    # Test that they inherit correctly
    print(f"UTCMainProcessor has load_data method: {hasattr(utc_main, 'load_data')}")
    print(f"TranscoProcessor has apply_basic_business_rules method: {hasattr(transco, 'apply_basic_business_rules')}")

    # Quick test processing
    test_file = "1_05.05.xlsx"
    if os.path.exists(test_file):
        print(f"\n=== Testing Transco with modular structure ===")
        processed_data = transco.process_excel_file(test_file)
        print(f"SUCCESS: Processed {len(processed_data)} rows with modular Transco")
        print(f"Total Selected Cost: ${transco.summary_stats['total_selected_cost']:,.2f}")
    else:
        print("Test file not found, skipping file processing test")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()