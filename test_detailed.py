#!/usr/bin/env python3
"""
Simple test for detailed processor
"""

from tms_detailed_processor import TMSDetailedDataProcessor
import os
import sys

def test_detailed_processor():
    """Test the detailed processor"""
    print("=== Testing TMSDetailedDataProcessor ===")
    
    try:
        processor = TMSDetailedDataProcessor()
        
        # Test with the raw detailed report
        input_file = 'CN_raw_08.25.xlsx'
        output_file = 'test_detailed_output.xlsx'
        
        if not os.path.exists(input_file):
            print(f"ERROR: Input file {input_file} not found")
            return False
            
        print(f"Processing {input_file}...")
        
        # Process the data
        processed_data = processor.clean_and_process_data(input_file)
        
        print(f"SUCCESS: Processing completed!")
        print(f"Records processed: {len(processed_data)}")
        print(f"Columns: {len(processed_data.columns)}")
        print(f"Column names: {list(processed_data.columns)}")
        
        # Display summary statistics
        print("\n=== Summary Statistics ===")
        stats = processor.summary_stats
        print(f"Total Loads: {stats['total_loads']}")
        print(f"Total Selected Cost: ${stats['total_selected_cost']:,.2f}")
        print(f"Total Potential Savings: ${stats['total_potential_savings']:,.2f}")
        print(f"Percentage Savings: {stats['percentage_savings']:.2f}%")
        print(f"Loads with Savings: {stats['loads_with_savings']}")
        print(f"Report Type: {stats['report_type']}")
        
        # Save the processed data
        processor.save_processed_data(output_file)
        print(f"\nSUCCESS: Saved to {output_file}")
        
        return True
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_detailed_processor()
    if success:
        print("\n✓ Detailed processor test PASSED")
        sys.exit(0)
    else:
        print("\n✗ Detailed processor test FAILED")
        sys.exit(1)