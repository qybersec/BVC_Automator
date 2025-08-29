"""
Comprehensive testing suite for the enhanced TMS processor
"""
import unittest
import pandas as pd
import numpy as np
import tempfile
import os
from pathlib import Path
import time
import sys

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from tms_processor import ModernTMSProcessor
    from config import tms_config
    from logger_config import main_logger
    from validators import tms_validator
except ImportError as e:
    print(f"Warning: Could not import enhanced modules: {e}")
    print("Some tests may be skipped.")
    ModernTMSProcessor = None

class TestEnhancedTMSProcessor(unittest.TestCase):
    """Test suite for enhanced TMS processor"""
    
    @classmethod
    def setUpClass(cls):
        """Set up test class"""
        cls.test_data_dir = Path(tempfile.mkdtemp(prefix="tms_test_"))
        cls.sample_data = cls._create_sample_data()
        cls.sample_file = cls._create_sample_excel_file()
    
    @classmethod
    def tearDownClass(cls):
        """Clean up test files"""
        import shutil
        if cls.test_data_dir.exists():
            shutil.rmtree(cls.test_data_dir)
    
    def setUp(self):
        """Set up each test"""
        if ModernTMSProcessor is None:
            self.skipTest("Enhanced TMS processor not available")
        self.processor = ModernTMSProcessor()
    
    @staticmethod
    def _create_sample_data():
        """Create sample TMS data for testing"""
        return {
            'title_rows': [
                ['', 'TMS Transportation Report', '', '', ''],
                ['', '', '', '', ''],
                ['', 'Sample Company Inc.', '', '', ''],
                ['', '', '', '', ''],
                ['', 'Date Range: 01/01/2024 - 01/31/2024', '', '', ''],
                ['', '', '', '', ''],
                ['', '', '', '', ''],
                ['', '', '', '', ''],
            ],
            'headers': ['', '', 'Load No.', 'Ship Date', 'Origin City', 'Origin State', 'Origin Postal',
                       'Destination City', 'Destination State', 'Destination Postal',
                       'Selected Carrier', 'Selected Service Type', 'Selected Transit Days', 
                       'Selected Freight Cost', 'Selected Accessorial Cost', 'Selected Total Cost',
                       'Least Cost Carrier', 'Least Cost Service Type', 'Least Cost Transit Days',
                       'Least Cost Freight Cost', 'Least Cost Accessorial Cost', 'Least Cost Total Cost',
                       'Potential Savings'],
            'data_rows': [
                ['', '', 'A12345', '2024-01-15', 'New York', 'NY', '10001',
                 'Los Angeles', 'CA', '90210',
                 'Carrier A', 'Standard', '5', '1000', '100', '1100',
                 'Carrier B', 'Express', '3', '950', '120', '1070',
                 '30'],
                ['', '', 'A12346', '2024-01-16', 'Chicago', 'IL', '60601',
                 'Miami', 'FL', '33101',
                 'Carrier C', 'Standard', '4', '800', '80', '880',
                 'Carrier C', 'Standard', '4', '800', '80', '880',
                 '0'],
                ['', '', 'A12347', '2024-01-17', 'Houston', 'TX', '77001',
                 'Seattle', 'WA', '98101',
                 'Carrier D', 'Express', '2', '1200', '150', '1350',
                 '', '', '', '', '', '',
                 ''],
                ['', '', 'A12348', '2024-01-18', 'Boston', 'MA', '02101',
                 'Denver', 'CO', '80201',
                 'Carrier E', 'Standard', '6', '900', '90', '990',
                 'Carrier F', 'Standard', '5', '1000', '100', '1100',
                 '-110'],
            ]
        }
    
    @classmethod
    def _create_sample_excel_file(cls):
        """Create a sample Excel file for testing"""
        file_path = cls.test_data_dir / "sample_tms_data.xlsx"
        
        # Create DataFrame with all data
        all_data = []
        all_data.extend(cls.sample_data['title_rows'])
        all_data.append(cls.sample_data['headers'])
        all_data.append(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])  # Empty row
        all_data.extend(cls.sample_data['data_rows'])
        
        df = pd.DataFrame(all_data)
        df.to_excel(file_path, index=False, header=False)
        
        return str(file_path)
    
    def test_enhanced_validation(self):
        """Test enhanced validation capabilities"""
        if not hasattr(tms_validator, 'run_full_validation'):
            self.skipTest("Enhanced validation not available")
        
        # Test file validation
        validation_results = tms_validator.run_full_validation(self.sample_file)
        
        self.assertIsInstance(validation_results, dict)
        self.assertIn('overall_valid', validation_results)
        self.assertIn('validation_steps', validation_results)
        
        # Should pass validation
        if not validation_results['overall_valid']:
            print(f"Validation failed: {validation_results}")
        
        self.assertTrue(validation_results.get('overall_valid', False))
    
    def test_enhanced_processing_with_logging(self):
        """Test processing with enhanced logging"""
        start_time = time.time()
        
        try:
            processed_data = self.processor.clean_and_process_data(self.sample_file)
            processing_time = time.time() - start_time
            
            # Verify processing results
            self.assertIsNotNone(processed_data)
            self.assertIsInstance(processed_data, pd.DataFrame)
            self.assertGreater(len(processed_data), 0)
            
            # Check if processing stats are available
            if hasattr(self.processor, 'processing_stats'):
                stats = self.processor.processing_stats
                self.assertIn('total_time', stats)
                self.assertIn('records_processed', stats)
                self.assertGreater(stats['total_time'], 0)
                self.assertEqual(stats['records_processed'], len(processed_data))
            
            print(f"Processing completed in {processing_time:.2f} seconds")
            print(f"Processed {len(processed_data)} records")
            
        except Exception as e:
            self.fail(f"Enhanced processing failed: {str(e)}")
    
    def test_business_logic_application(self):
        """Test enhanced business logic application"""
        processed_data = self.processor.clean_and_process_data(self.sample_file)
        
        # Test same carrier rule (row 2: Carrier C -> Carrier C, should have 0 savings)
        same_carrier_rows = processed_data[
            (processed_data['Selected Carrier'] == processed_data['Least Cost Carrier']) &
            (processed_data['Selected Carrier'].notna()) &
            (processed_data['Least Cost Carrier'].notna())
        ]
        
        if len(same_carrier_rows) > 0 and 'Potential Savings' in processed_data.columns:
            for idx, row in same_carrier_rows.iterrows():
                self.assertEqual(
                    float(row['Potential Savings']), 
                    0.0, 
                    f"Same carrier rule failed for row {idx}"
                )
        
        # Test empty data rule (row 3: empty least cost data)
        empty_least_cost_rows = processed_data[
            processed_data['Least Cost Carrier'].isna() | 
            (processed_data['Least Cost Carrier'].astype(str) == '') |
            (processed_data['Least Cost Carrier'].astype(str) == 'nan')
        ]
        
        if len(empty_least_cost_rows) > 0:
            for idx, row in empty_least_cost_rows.iterrows():
                # Should have copied selected data to least cost
                self.assertEqual(
                    str(row['Least Cost Carrier']), 
                    str(row['Selected Carrier']),
                    f"Empty data rule failed for row {idx}"
                )
                if 'Potential Savings' in processed_data.columns:
                    self.assertEqual(
                        float(row['Potential Savings']), 
                        0.0,
                        f"Empty data savings rule failed for row {idx}"
                    )
    
    def test_configuration_integration(self):
        """Test configuration system integration"""
        if not hasattr(tms_config, 'get'):
            self.skipTest("Configuration system not available")
        
        # Test configuration access
        header_row = tms_config.get('data_structure.default_header_row', 8)
        self.assertIsInstance(header_row, int)
        self.assertGreaterEqual(header_row, 0)
        
        min_columns = tms_config.get('data_structure.min_data_columns', 5)
        self.assertIsInstance(min_columns, int)
        self.assertGreater(min_columns, 0)
        
        # Test that processor uses configuration
        self.assertIsNotNone(self.processor.config)
    
    def test_error_handling_robustness(self):
        """Test enhanced error handling"""
        # Test with non-existent file
        with self.assertRaises((FileNotFoundError, ValueError)):
            self.processor.clean_and_process_data("non_existent_file.xlsx")
        
        # Test with invalid file (create a text file with .xlsx extension)
        invalid_file = self.test_data_dir / "invalid.xlsx"
        with open(invalid_file, 'w') as f:
            f.write("This is not an Excel file")
        
        with self.assertRaises((ValueError, RuntimeError)):
            self.processor.clean_and_process_data(str(invalid_file))
    
    def test_summary_statistics_accuracy(self):
        """Test summary statistics calculation"""
        processed_data = self.processor.clean_and_process_data(self.sample_file)
        
        # Verify summary stats were calculated
        self.assertIsNotNone(self.processor.summary_stats)
        stats = self.processor.summary_stats
        
        # Check required statistics
        required_stats = [
            'total_loads', 'total_selected_cost', 'total_least_cost',
            'total_potential_savings', 'average_savings_per_load',
            'percentage_savings', 'loads_with_savings'
        ]
        
        for stat in required_stats:
            self.assertIn(stat, stats, f"Missing statistic: {stat}")
            self.assertIsInstance(stats[stat], (int, float, np.integer, np.floating), f"Invalid type for {stat}")
        
        # Verify accuracy
        self.assertEqual(stats['total_loads'], len(processed_data))
        
        if 'Potential Savings' in processed_data.columns:
            expected_total_savings = processed_data['Potential Savings'].sum()
            self.assertAlmostEqual(
                stats['total_potential_savings'], 
                expected_total_savings, 
                places=2,
                msg="Total potential savings calculation incorrect"
            )
    
    def test_performance_benchmarks(self):
        """Test processing performance"""
        # Create larger dataset for performance testing
        large_data = self.sample_data.copy()
        
        # Duplicate data rows to create larger dataset
        original_rows = large_data['data_rows'].copy()
        for i in range(10):  # Create 40 total rows (4 original * 10)
            for row in original_rows:
                new_row = row.copy()
                new_row[2] = f"A{12349 + i * 100 + len(large_data['data_rows'])}"  # Unique Load No.
                large_data['data_rows'].append(new_row)
        
        # Create larger test file
        large_file = self.test_data_dir / "large_tms_data.xlsx"
        all_data = []
        all_data.extend(large_data['title_rows'])
        all_data.append(large_data['headers'])
        all_data.append([''] * len(large_data['headers']))
        all_data.extend(large_data['data_rows'])
        
        df = pd.DataFrame(all_data)
        df.to_excel(large_file, index=False, header=False)
        
        # Benchmark processing
        start_time = time.time()
        processed_data = self.processor.clean_and_process_data(str(large_file))
        processing_time = time.time() - start_time
        
        # Performance assertions
        records_per_second = len(processed_data) / processing_time if processing_time > 0 else 0
        
        print(f"Performance benchmark:")
        print(f"  Records processed: {len(processed_data)}")
        print(f"  Processing time: {processing_time:.3f}s")
        print(f"  Records per second: {records_per_second:.0f}")
        
        # Should process at reasonable speed (adjust threshold as needed)
        self.assertGreater(records_per_second, 10, "Processing speed too slow")
        self.assertLess(processing_time, 30, "Processing time too long")
    
    def test_data_integrity(self):
        """Test data integrity throughout processing"""
        processed_data = self.processor.clean_and_process_data(self.sample_file)
        
        # Check that essential columns exist
        essential_columns = ['Load No.', 'Selected Carrier', 'Least Cost Carrier']
        for col in essential_columns:
            if col in processed_data.columns:
                # Should not have completely empty Load No. column
                if col == 'Load No.':
                    non_empty_loads = processed_data[col].notna().sum()
                    self.assertGreater(non_empty_loads, 0, "All Load Numbers are empty")
        
        # Check data types
        if 'Potential Savings' in processed_data.columns:
            # Should be numeric
            numeric_savings = pd.to_numeric(processed_data['Potential Savings'], errors='coerce')
            non_numeric_count = numeric_savings.isna().sum()
            original_na_count = processed_data['Potential Savings'].isna().sum()
            conversion_failures = non_numeric_count - original_na_count
            
            self.assertEqual(conversion_failures, 0, "Potential Savings should be numeric")
        
        # Check for data consistency
        for idx, row in processed_data.iterrows():
            # If selected and least cost are the same, savings should be 0
            if (row.get('Selected Carrier') == row.get('Least Cost Carrier') and
                pd.notna(row.get('Selected Carrier')) and
                pd.notna(row.get('Least Cost Carrier'))):
                
                if 'Potential Savings' in processed_data.columns:
                    savings = pd.to_numeric(row['Potential Savings'], errors='coerce')
                    if pd.notna(savings):
                        self.assertEqual(
                            savings, 0.0,
                            f"Same carrier should have 0 savings at row {idx}"
                        )


class TestConfigurationSystem(unittest.TestCase):
    """Test configuration system"""
    
    def test_config_loading(self):
        """Test configuration loading"""
        if not hasattr(tms_config, 'get'):
            self.skipTest("Configuration system not available")
        
        # Test default values
        default_header_row = tms_config.get('data_structure.default_header_row')
        self.assertIsNotNone(default_header_row)
        
        # Test non-existent key with default
        non_existent = tms_config.get('non.existent.key', 'default_value')
        self.assertEqual(non_existent, 'default_value')
    
    def test_config_validation(self):
        """Test configuration validation"""
        if not hasattr(tms_config, 'validate_config'):
            self.skipTest("Configuration validation not available")
        
        issues = tms_config.validate_config()
        self.assertIsInstance(issues, list)
        
        # Should have no critical validation issues with default config
        critical_issues = [issue for issue in issues if 'missing required' in issue.lower()]
        self.assertEqual(len(critical_issues), 0, f"Critical config issues: {critical_issues}")


def run_tests():
    """Run all tests with detailed output"""
    # Set up test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add test classes
    suite.addTests(loader.loadTestsFromTestCase(TestEnhancedTMSProcessor))
    suite.addTests(loader.loadTestsFromTestCase(TestConfigurationSystem))
    
    # Run tests with detailed output
    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "="*50)
    print("TEST SUMMARY")
    print("="*50)
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    print(f"Skipped: {len(result.skipped) if hasattr(result, 'skipped') else 0}")
    
    if result.failures:
        print(f"\nFAILURES:")
        for test, traceback in result.failures:
            print(f"  {test}: {traceback}")
    
    if result.errors:
        print(f"\nERRORS:")
        for test, traceback in result.errors:
            print(f"  {test}: {traceback}")
    
    # Return success status
    return len(result.failures) == 0 and len(result.errors) == 0


if __name__ == '__main__':
    print("Enhanced TMS Processor Test Suite")
    print("="*40)
    
    success = run_tests()
    
    if success:
        print("\nAll tests passed!")
        sys.exit(0)
    else:
        print("\nSome tests failed!")
        sys.exit(1)