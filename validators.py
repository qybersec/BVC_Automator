"""
Input validation and data verification for TMS Processor
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from datetime import datetime
import re
from logger_config import data_logger
from config import tms_config

class TMSDataValidator:
    """Comprehensive data validation for TMS files"""
    
    def __init__(self):
        self.logger = data_logger
        self.config = tms_config
        self.validation_results = {}
    
    def validate_file_path(self, file_path: str) -> Tuple[bool, str]:
        """Validate file path and accessibility"""
        try:
            path = Path(file_path)
            
            # Check if file exists
            if not path.exists():
                return False, f"File does not exist: {file_path}"
            
            # Check if it's a file (not directory)
            if not path.is_file():
                return False, f"Path is not a file: {file_path}"
            
            # Check file extension
            allowed_extensions = self.config.get('validation.allowed_extensions', ['.xlsx', '.xls'])
            if path.suffix.lower() not in allowed_extensions:
                return False, f"Unsupported file format. Allowed: {allowed_extensions}"
            
            # Check file size
            max_size_mb = self.config.get('validation.max_file_size_mb', 100)
            file_size_mb = path.stat().st_size / (1024 * 1024)
            if file_size_mb > max_size_mb:
                return False, f"File too large: {file_size_mb:.1f}MB (max: {max_size_mb}MB)"
            
            # Check read permissions (Windows compatible)
            try:
                with open(path, 'rb') as test_file:
                    test_file.read(1)
            except (PermissionError, IOError):
                return False, f"File is not readable: {file_path}"
            
            self.logger.log_file_operation("validation", str(path), True, path.stat().st_size)
            return True, "File validation passed"
            
        except Exception as e:
            self.logger.error("File validation failed", exception=e, file_path=file_path)
            return False, f"File validation error: {str(e)}"
    
    def validate_excel_structure(self, file_path: str) -> Tuple[bool, str, Optional[pd.DataFrame]]:
        """Validate Excel file structure and content"""
        try:
            # Try to read the Excel file
            df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
            
            # Check minimum dimensions
            min_rows = self.config.get('validation.min_rows', 10)
            max_rows = self.config.get('validation.max_rows', 50000)
            
            if len(df_raw) < min_rows:
                return False, f"File has too few rows: {len(df_raw)} (minimum: {min_rows})", None
            
            if len(df_raw) > max_rows:
                return False, f"File has too many rows: {len(df_raw)} (maximum: {max_rows})", None
            
            # Check for minimum columns
            min_cols = self.config.get('data_structure.min_data_columns', 5)
            if len(df_raw.columns) < min_cols:
                return False, f"File has too few columns: {len(df_raw.columns)} (minimum: {min_cols})", None
            
            # Check if file is completely empty
            if df_raw.empty or df_raw.dropna(how='all').empty:
                return False, "File appears to be empty or contains no data", None
            
            self.logger.log_data_stats({
                'rows': len(df_raw),
                'columns': len(df_raw.columns),
                'non_empty_cells': df_raw.count().sum()
            }, "EXCEL_STRUCTURE")
            
            return True, "Excel structure validation passed", df_raw
            
        except pd.errors.EmptyDataError:
            return False, "Excel file is empty", None
        except pd.errors.ParserError as e:
            return False, f"Excel file parsing error: {str(e)}", None
        except Exception as e:
            self.logger.error("Excel structure validation failed", exception=e)
            return False, f"Excel validation error: {str(e)}", None
    
    def detect_and_validate_headers(self, df_raw: pd.DataFrame) -> Tuple[bool, str, Dict[str, Any]]:
        """Detect and validate TMS headers"""
        try:
            header_indicators = self.config.get('data_structure.header_indicators', [])
            detection_results = {
                'header_row': None,
                'data_start_row': None,
                'detected_headers': [],
                'missing_indicators': [],
                'confidence_score': 0
            }
            
            best_match = {'row': -1, 'score': 0, 'headers': []}
            
            # Search for header row within reasonable range
            search_range = min(20, len(df_raw))
            for row_idx in range(search_range):
                try:
                    row_data = df_raw.iloc[row_idx].dropna().astype(str).tolist()
                    row_str = ' '.join(row_data).lower()
                    
                    # Count matches with header indicators
                    matches = 0
                    found_indicators = []
                    
                    for indicator in header_indicators:
                        if indicator.lower() in row_str:
                            matches += 1
                            found_indicators.append(indicator)
                    
                    # Calculate confidence score
                    confidence = matches / len(header_indicators) if header_indicators else 0
                    
                    if confidence > best_match['score']:
                        best_match = {
                            'row': row_idx,
                            'score': confidence,
                            'headers': row_data,
                            'found_indicators': found_indicators
                        }
                
                except Exception as e:
                    self.logger.debug(f"Error checking row {row_idx} for headers: {e}")
                    continue
            
            # Validate best match
            min_confidence = 0.5  # At least 50% of indicators must be found
            if best_match['score'] >= min_confidence:
                detection_results.update({
                    'header_row': best_match['row'],
                    'data_start_row': best_match['row'] + 2,  # Skip potential blank row
                    'detected_headers': best_match['headers'],
                    'confidence_score': best_match['score'],
                    'missing_indicators': [ind for ind in header_indicators 
                                         if ind not in best_match.get('found_indicators', [])]
                })
                
                self.logger.log_validation_result(
                    "header_detection", 
                    True, 
                    f"Headers found at row {best_match['row']} with {best_match['score']:.1%} confidence"
                )
                return True, "Header detection successful", detection_results
            else:
                self.logger.log_validation_result(
                    "header_detection", 
                    False, 
                    f"Low confidence: {best_match['score']:.1%}"
                )
                return False, f"Could not reliably detect headers. Best confidence: {best_match['score']:.1%}", detection_results
        
        except Exception as e:
            self.logger.error("Header detection failed", exception=e)
            return False, f"Header detection error: {str(e)}", {}
    
    def validate_data_quality(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Validate data quality and identify issues"""
        issues = []
        stats = {}
        
        try:
            # Basic statistics
            stats['total_rows'] = len(df)
            stats['total_columns'] = len(df.columns)
            stats['empty_rows'] = df.isnull().all(axis=1).sum()
            stats['empty_columns'] = df.isnull().all(axis=0).sum()
            stats['duplicate_rows'] = df.duplicated().sum()
            
            # Check for required columns
            required_cols = self.config.get('data_structure.required_columns', [])
            missing_required = [col for col in required_cols if col not in df.columns]
            if missing_required:
                issues.append(f"Missing required columns: {missing_required}")
            
            # Check for Load No. column specifically
            load_no_col = None
            for col in df.columns:
                if 'load' in str(col).lower() and ('no' in str(col).lower() or 'number' in str(col).lower()):
                    load_no_col = col
                    break
            
            if load_no_col:
                empty_load_nos = df[load_no_col].isnull().sum()
                invalid_load_nos = df[df[load_no_col].astype(str).str.strip().isin(['', 'nan'])].shape[0]
                stats['empty_load_numbers'] = empty_load_nos
                stats['invalid_load_numbers'] = invalid_load_nos
                
                if empty_load_nos > len(df) * 0.1:  # More than 10% empty
                    issues.append(f"High number of empty Load Numbers: {empty_load_nos}")
            else:
                issues.append("Could not identify Load Number column")
            
            # Check numeric columns
            numeric_indicators = ['cost', 'price', 'amount', 'savings', 'days']
            numeric_columns = [col for col in df.columns 
                             if any(ind in str(col).lower() for ind in numeric_indicators)]
            
            for col in numeric_columns:
                try:
                    numeric_values = pd.to_numeric(df[col], errors='coerce')
                    null_count = numeric_values.isnull().sum()
                    if null_count > 0:
                        stats[f'{col}_conversion_failures'] = null_count
                        
                    # Check for reasonable ranges
                    if 'cost' in str(col).lower() or 'amount' in str(col).lower():
                        negative_values = (numeric_values < 0).sum()
                        if negative_values > 0:
                            stats[f'{col}_negative_values'] = negative_values
                            
                        extreme_values = (numeric_values > 100000).sum()  # Over $100k
                        if extreme_values > 0:
                            stats[f'{col}_extreme_values'] = extreme_values
                            
                except Exception as e:
                    issues.append(f"Could not validate numeric column {col}: {e}")
            
            # Check date columns
            date_indicators = ['date', 'time', 'ship', 'delivery']
            date_columns = [col for col in df.columns 
                          if any(ind in str(col).lower() for ind in date_indicators)]
            
            for col in date_columns:
                try:
                    date_values = pd.to_datetime(df[col], errors='coerce')
                    null_count = date_values.isnull().sum()
                    if null_count > 0:
                        stats[f'{col}_invalid_dates'] = null_count
                        
                    # Check for reasonable date ranges
                    current_year = datetime.now().year
                    future_dates = (date_values.dt.year > current_year + 2).sum()
                    old_dates = (date_values.dt.year < current_year - 10).sum()
                    
                    if future_dates > 0:
                        stats[f'{col}_future_dates'] = future_dates
                    if old_dates > 0:
                        stats[f'{col}_old_dates'] = old_dates
                        
                except Exception as e:
                    issues.append(f"Could not validate date column {col}: {e}")
            
            self.logger.log_data_stats(stats, "DATA_QUALITY")
            
            return {
                'issues': issues,
                'stats': stats,
                'is_valid': len(issues) == 0,
                'quality_score': max(0, 100 - len(issues) * 10)  # Simple scoring
            }
            
        except Exception as e:
            self.logger.error("Data quality validation failed", exception=e)
            return {
                'issues': [f"Validation error: {str(e)}"],
                'stats': stats,
                'is_valid': False,
                'quality_score': 0
            }
    
    def validate_business_logic_requirements(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Validate that data meets business logic requirements"""
        try:
            validation_results = {
                'can_apply_same_carrier_rule': False,
                'can_apply_empty_data_rule': False,
                'can_apply_negative_savings_rule': False,
                'carrier_columns_found': [],
                'cost_columns_found': [],
                'savings_column_found': None,
                'issues': []
            }
            
            # Look for carrier columns
            carrier_patterns = [
                r'selected.*carrier',
                r'least.*cost.*carrier',
                r'carrier.*selected',
                r'carrier.*least'
            ]
            
            carrier_columns = []
            for col in df.columns:
                col_str = str(col).lower()
                for pattern in carrier_patterns:
                    if re.search(pattern, col_str):
                        carrier_columns.append(col)
                        break
            
            validation_results['carrier_columns_found'] = carrier_columns
            
            if len(carrier_columns) >= 2:
                validation_results['can_apply_same_carrier_rule'] = True
                validation_results['can_apply_empty_data_rule'] = True
            else:
                validation_results['issues'].append(
                    f"Need at least 2 carrier columns for business rules, found: {len(carrier_columns)}"
                )
            
            # Look for cost columns
            cost_patterns = [
                r'.*cost.*',
                r'.*price.*',
                r'.*amount.*',
                r'.*freight.*'
            ]
            
            cost_columns = []
            for col in df.columns:
                col_str = str(col).lower()
                for pattern in cost_patterns:
                    if re.search(pattern, col_str):
                        cost_columns.append(col)
                        break
            
            validation_results['cost_columns_found'] = cost_columns
            
            # Look for savings column
            savings_patterns = [
                r'.*savings.*',
                r'.*potential.*',
                r'.*difference.*'
            ]
            
            savings_column = None
            for col in df.columns:
                col_str = str(col).lower()
                for pattern in savings_patterns:
                    if re.search(pattern, col_str):
                        savings_column = col
                        break
            
            validation_results['savings_column_found'] = savings_column
            
            if savings_column:
                validation_results['can_apply_negative_savings_rule'] = True
            else:
                validation_results['issues'].append("No savings column found for negative savings rule")
            
            # Log validation results
            self.logger.log_validation_result(
                "business_logic_requirements",
                len(validation_results['issues']) == 0,
                f"Found {len(carrier_columns)} carrier cols, {len(cost_columns)} cost cols, savings: {bool(savings_column)}"
            )
            
            return validation_results
            
        except Exception as e:
            self.logger.error("Business logic validation failed", exception=e)
            return {
                'can_apply_same_carrier_rule': False,
                'can_apply_empty_data_rule': False,
                'can_apply_negative_savings_rule': False,
                'issues': [f"Business logic validation error: {str(e)}"]
            }
    
    def run_full_validation(self, file_path: str) -> Dict[str, Any]:
        """Run complete validation suite"""
        self.logger.info(f"Starting full validation for: {Path(file_path).name}")
        
        full_results = {
            'file_path': file_path,
            'timestamp': datetime.now().isoformat(),
            'overall_valid': False,
            'validation_steps': {}
        }
        
        # Step 1: File validation
        file_valid, file_msg = self.validate_file_path(file_path)
        full_results['validation_steps']['file_path'] = {
            'valid': file_valid,
            'message': file_msg
        }
        
        if not file_valid:
            self.logger.error(f"File validation failed: {file_msg}")
            return full_results
        
        # Step 2: Excel structure validation
        excel_valid, excel_msg, df_raw = self.validate_excel_structure(file_path)
        full_results['validation_steps']['excel_structure'] = {
            'valid': excel_valid,
            'message': excel_msg
        }
        
        if not excel_valid:
            self.logger.error(f"Excel validation failed: {excel_msg}")
            return full_results
        
        # Step 3: Header detection
        header_valid, header_msg, header_info = self.detect_and_validate_headers(df_raw)
        full_results['validation_steps']['header_detection'] = {
            'valid': header_valid,
            'message': header_msg,
            'details': header_info
        }
        
        # Step 4: Data quality validation (can proceed even with header issues)
        if header_valid and header_info.get('data_start_row'):
            try:
                # Extract data portion for validation
                data_df = df_raw.iloc[header_info['data_start_row']:].copy()
                data_df = data_df.dropna(how='all')
                
                if not data_df.empty:
                    # Set column names for validation
                    if len(header_info['detected_headers']) <= len(data_df.columns):
                        data_df.columns = header_info['detected_headers'] + [f'Col_{i}' for i in range(len(header_info['detected_headers']), len(data_df.columns))]
                    
                    quality_results = self.validate_data_quality(data_df)
                    full_results['validation_steps']['data_quality'] = quality_results
                    
                    # Step 5: Business logic validation
                    business_results = self.validate_business_logic_requirements(data_df)
                    full_results['validation_steps']['business_logic'] = business_results
                    
            except Exception as e:
                self.logger.error("Data validation failed", exception=e)
                full_results['validation_steps']['data_quality'] = {
                    'is_valid': False,
                    'issues': [f"Data validation error: {str(e)}"]
                }
        
        # Determine overall validity
        critical_steps = ['file_path', 'excel_structure', 'header_detection']
        full_results['overall_valid'] = all(
            full_results['validation_steps'].get(step, {}).get('valid', False)
            for step in critical_steps
        )
        
        # Log final results
        if full_results['overall_valid']:
            self.logger.info(f"Full validation PASSED for {Path(file_path).name}")
        else:
            failed_steps = [step for step in critical_steps 
                           if not full_results['validation_steps'].get(step, {}).get('valid', False)]
            self.logger.warning(f"Full validation FAILED for {Path(file_path).name}. Failed steps: {failed_steps}")
        
        return full_results


class TMSDataCleaner:
    """Data cleaning utilities with validation"""
    
    def __init__(self):
        self.logger = data_logger
        self.config = tms_config
    
    def clean_and_validate_types(self, df: pd.DataFrame, column_mapping: Dict[str, str]) -> pd.DataFrame:
        """Clean and validate data types with extensive error checking"""
        df_clean = df.copy()
        cleaning_stats = {}
        
        try:
            # Clean numeric columns
            numeric_columns = [col for col, dtype in column_mapping.items() 
                             if dtype == 'numeric' and col in df_clean.columns]
            
            for col in numeric_columns:
                original_nulls = df_clean[col].isnull().sum()
                
                # Convert to numeric with error tracking
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')
                
                new_nulls = df_clean[col].isnull().sum()
                conversion_failures = new_nulls - original_nulls
                
                if conversion_failures > 0:
                    cleaning_stats[f'{col}_conversion_failures'] = conversion_failures
                    self.logger.warning(f"Failed to convert {conversion_failures} values in {col} to numeric")
                
                # Fill nulls with 0 for numeric columns
                df_clean[col] = df_clean[col].fillna(0)
            
            # Clean date columns
            date_columns = [col for col, dtype in column_mapping.items() 
                           if dtype == 'date' and col in df_clean.columns]
            
            for col in date_columns:
                original_nulls = df_clean[col].isnull().sum()
                
                # Convert to datetime
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
                
                new_nulls = df_clean[col].isnull().sum()
                conversion_failures = new_nulls - original_nulls
                
                if conversion_failures > 0:
                    cleaning_stats[f'{col}_date_failures'] = conversion_failures
                    self.logger.warning(f"Failed to convert {conversion_failures} values in {col} to date")
                
                # Format dates
                date_format = self.config.get('formatting.date_format', '%m/%d/%y')
                df_clean[col] = df_clean[col].dt.strftime(date_format)
            
            # Clean string columns
            string_columns = [col for col, dtype in column_mapping.items() 
                            if dtype == 'string' and col in df_clean.columns]
            
            for col in string_columns:
                # Convert to string and clean
                df_clean[col] = df_clean[col].astype(str).str.strip()
                df_clean[col] = df_clean[col].replace('nan', '')
                
                # Count empty values
                empty_count = (df_clean[col] == '').sum()
                if empty_count > 0:
                    cleaning_stats[f'{col}_empty_values'] = empty_count
            
            self.logger.log_data_stats(cleaning_stats, "CLEANING_STATS")
            return df_clean
            
        except Exception as e:
            self.logger.error("Data cleaning failed", exception=e)
            raise ValueError(f"Data cleaning error: {str(e)}")


# Global validator instance
tms_validator = TMSDataValidator()
tms_cleaner = TMSDataCleaner()