#!/usr/bin/env python3
"""
Test script to validate TMS processor improvements
"""

import pandas as pd
import os
from tms_processor import ModernTMSProcessor

def test_data_structure_detection():
    """Test the intelligent data structure detection"""
    print("🧪 Testing data structure detection...")
    
    processor = ModernTMSProcessor()
    
    # Test with sample files
    test_files = [
        "Ferro_raw_08.18.xlsx",
        "FERROWORKS BVC 08.18.25 - 08.22.25.xlsx"
    ]
    
    for file_name in test_files:
        if os.path.exists(file_name):
            print(f"📁 Testing file: {file_name}")
            try:
                # Read raw file
                df_raw = pd.read_excel(file_name, sheet_name=0, header=None)
                
                # Test structure detection
                header_row, data_start_row = processor._detect_data_structure(df_raw)
                print(f"   📊 Detected header row: {header_row}")
                print(f"   📈 Detected data start row: {data_start_row}")
                
                # Test title extraction
                title_info = processor._extract_title_info(df_raw)
                print(f"   🏢 Company: {title_info.get('company_name', 'Not found')}")
                print(f"   📅 Date range: {title_info.get('date_range', 'Not found')}")
                
            except Exception as e:
                print(f"   ❌ Error testing {file_name}: {e}")
        else:
            print(f"   ⚠️  File not found: {file_name}")
    
    print("✅ Data structure detection test completed\n")

def test_performance_improvements():
    """Test performance improvements"""
    print("🚀 Testing performance improvements...")
    
    processor = ModernTMSProcessor()
    
    # Test with available files
    test_files = [f for f in ["Ferro_raw_08.18.xlsx", "FERROWORKS BVC 08.18.25 - 08.22.25.xlsx"] 
                  if os.path.exists(f)]
    
    if not test_files:
        print("   ⚠️  No test files available")
        return
    
    for file_name in test_files:
        print(f"📁 Testing performance with: {file_name}")
        try:
            import time
            start_time = time.time()
            
            # Process the file
            processed_data = processor.clean_and_process_data(file_name)
            
            end_time = time.time()
            processing_time = end_time - start_time
            
            print(f"   ⏱️  Processing time: {processing_time:.2f} seconds")
            print(f"   📊 Processed records: {len(processed_data)}")
            print(f"   💰 Total selected cost: ${processor.summary_stats['total_selected_cost']:,.2f}")
            print(f"   💵 Total potential savings: ${processor.summary_stats['total_potential_savings']:,.2f}")
            
        except Exception as e:
            print(f"   ❌ Error processing {file_name}: {e}")
    
    print("✅ Performance test completed\n")

def test_error_handling():
    """Test improved error handling"""
    print("🛡️  Testing error handling...")
    
    processor = ModernTMSProcessor()
    
    # Test with non-existent file
    try:
        processor.clean_and_process_data("non_existent_file.xlsx")
    except FileNotFoundError as e:
        print(f"   ✅ Correctly caught FileNotFoundError: {e}")
    except Exception as e:
        print(f"   ⚠️  Unexpected exception type: {type(e).__name__}: {e}")
    
    # Test with invalid file (if exists)
    try:
        # Create a dummy invalid file for testing
        with open("test_invalid.txt", "w") as f:
            f.write("This is not an Excel file")
        
        processor.clean_and_process_data("test_invalid.txt")
    except (ValueError, RuntimeError) as e:
        print(f"   ✅ Correctly caught file format error: {type(e).__name__}: {e}")
        # Clean up
        if os.path.exists("test_invalid.txt"):
            os.remove("test_invalid.txt")
    except Exception as e:
        print(f"   ⚠️  Unexpected exception: {type(e).__name__}: {e}")
        # Clean up
        if os.path.exists("test_invalid.txt"):
            os.remove("test_invalid.txt")
    
    print("✅ Error handling test completed\n")

def main():
    """Run all tests"""
    print("🧪 TMS Processor Improvement Tests")
    print("=" * 50)
    
    test_data_structure_detection()
    test_performance_improvements()
    test_error_handling()
    
    print("🎉 All tests completed!")

if __name__ == "__main__":
    main()