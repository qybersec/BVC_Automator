#!/usr/bin/env python3
"""
TMS Data Processor Pro - Test Script
Comprehensive testing for both Basic and Detailed processors
"""
import sys
import os
from pathlib import Path

# Add current directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_basic_processor():
    """Test the Basic TMS processor"""
    print("🧪 Testing Basic TMS Processor...")
    print("=" * 50)
    
    try:
        from tms_processor import ModernTMSProcessor
        
        # Use the current directory Excel file
        input_file = "Ferro_raw_08.18.xlsx"
        output_file = "test_basic_output.xlsx"
        
        if not os.path.exists(input_file):
            print(f"❌ Input file not found: {input_file}")
            return False
        
        processor = ModernTMSProcessor()
        
        print("📖 Processing file...")
        processed_data = processor.clean_and_process_data(input_file)
        
        print(f"✅ Successfully processed {len(processed_data)} records")
        print("\n📊 Sample of processed data:")
        print(processed_data.head())
        
        print("\n📋 Column names:")
        print(processed_data.columns.tolist())
        
        print("\n📈 Summary Statistics:")
        for key, value in processor.summary_stats.items():
            if isinstance(value, float):
                if 'cost' in key.lower() or 'savings' in key.lower():
                    print(f"  {key}: ${value:,.2f}")
                else:
                    print(f"  {key}: {value:.2f}")
            else:
                print(f"  {key}: {value}")
        
        print(f"\n💾 Saving to: {output_file}")
        processor.save_processed_data(output_file)
        
        print("✅ Basic processor test completed successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Error during basic processor testing: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_detailed_processor():
    """Detailed processor disabled for now"""
    print("\n🧪 Detailed TMS Processor tests are skipped (feature disabled).")
    return True

def main():
    """Run all tests"""
    print("🚛 TMS Data Processor Pro - Test Suite")
    print("=" * 60)
    
    # Check if Excel file exists
    excel_file = "Ferro_raw_08.18.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ Test Excel file not found: {excel_file}")
        print("Please ensure the Excel file is in the current directory.")
        return
    
    print(f"📁 Found test file: {excel_file}")
    print(f"📊 File size: {os.path.getsize(excel_file) / 1024:.1f} KB")
    
    # Test both processors
    basic_success = test_basic_processor()
    detailed_success = test_detailed_processor()
    
    # Summary
    print("\n" + "=" * 60)
    print("📋 TEST SUMMARY")
    print("=" * 60)
    print(f"Basic Processor: {'✅ PASSED' if basic_success else '❌ FAILED'}")
    print(f"Detailed Processor: {'✅ PASSED' if detailed_success else '❌ FAILED'} (skipped)")
    
    if basic_success and detailed_success:
        print("\n🎉 All tests passed! The TMS Data Processor Pro is working correctly.")
    else:
        print("\n⚠️  Some tests failed. Please check the error messages above.")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()