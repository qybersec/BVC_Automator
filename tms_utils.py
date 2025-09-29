#!/usr/bin/env python3
"""
TMS Processing Utilities
Shared utility functions for TMS data processing
"""

import pandas as pd
from typing import List, Tuple, Dict, Any


class CityRuleProcessor:
    """Utility class for applying city-based exclusion rules"""

    @staticmethod
    def apply_city_exclusion_rule(df: pd.DataFrame, city_name: str, rule_description: str = None) -> pd.DataFrame:
        """
        Generic city exclusion rule - Zero out PS for any origin or destination matching city_name

        Args:
            df: DataFrame to process
            city_name: Name of city to exclude (case-insensitive)
            rule_description: Optional description for logging

        Returns:
            Modified DataFrame with Potential Savings zeroed for matching cities
        """
        if rule_description is None:
            rule_description = f"{city_name} City Rule"

        city_count = 0

        # Check if city columns exist
        origin_col_found = 'Origin City' in df.columns
        dest_col_found = 'Destination City' in df.columns

        if not (origin_col_found or dest_col_found):
            print(f"{rule_description}: No city columns found, skipping rule")
            return df

        # Create mask for matching cities
        city_mask = pd.Series([False] * len(df), index=df.index)

        # Check Origin City
        if origin_col_found:
            origin_matches = df['Origin City'].astype(str).str.upper().str.contains(
                city_name.upper(), na=False
            )
            city_mask |= origin_matches
            origin_count = origin_matches.sum()
            if origin_count > 0:
                print(f"Found {origin_count} rows with {city_name} in Origin City")

        # Check Destination City
        if dest_col_found:
            dest_matches = df['Destination City'].astype(str).str.upper().str.contains(
                city_name.upper(), na=False
            )
            city_mask |= dest_matches
            dest_count = dest_matches.sum()
            if dest_count > 0:
                print(f"Found {dest_count} rows with {city_name} in Destination City")

        city_count = city_mask.sum()

        if city_count > 0 and 'Potential Savings' in df.columns:
            # Zero out Potential Savings for matching rows
            df.loc[city_mask, 'Potential Savings'] = 0
            print(f"{rule_description}: Zeroed out PS for {city_count} rows")
        elif city_count > 0:
            print(
                f"{rule_description}: Found {city_count} {city_name} rows "
                f"but no Potential Savings column"
            )
        else:
            print(f"{rule_description}: No {city_name} cities found")

        return df


class DataValidationUtils:
    """Utility functions for data validation and cleaning"""

    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str]]:
        """
        Validate that required columns exist in DataFrame

        Args:
            df: DataFrame to validate
            required_columns: List of required column names

        Returns:
            Tuple of (is_valid, missing_columns)
        """
        missing = [col for col in required_columns if col not in df.columns]
        return len(missing) == 0, missing

    @staticmethod
    def clean_numeric_column(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
        """
        Clean and convert column to numeric, handling common issues

        Args:
            df: DataFrame to process
            column_name: Name of column to clean

        Returns:
            DataFrame with cleaned numeric column
        """
        if column_name not in df.columns:
            return df

        # Convert to string first to handle mixed types
        df[column_name] = df[column_name].astype(str)

        # Remove common non-numeric characters
        df[column_name] = df[column_name].str.replace('$', '', regex=False)
        df[column_name] = df[column_name].str.replace(',', '', regex=False)
        df[column_name] = df[column_name].str.replace('(', '-', regex=False)
        df[column_name] = df[column_name].str.replace(')', '', regex=False)

        # Convert to numeric, setting errors to NaN
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce')

        return df

    @staticmethod
    def get_column_statistics(df: pd.DataFrame, column_name: str) -> Dict[str, Any]:
        """
        Get basic statistics for a numeric column

        Args:
            df: DataFrame to analyze
            column_name: Name of column to analyze

        Returns:
            Dictionary with statistics
        """
        if column_name not in df.columns:
            return {'error': f'Column {column_name} not found'}

        series = df[column_name]
        if not pd.api.types.is_numeric_dtype(series):
            return {'error': f'Column {column_name} is not numeric'}

        return {
            'count': len(series),
            'non_null_count': series.count(),
            'null_count': series.isnull().sum(),
            'sum': series.sum(),
            'mean': series.mean(),
            'min': series.min(),
            'max': series.max(),
            'positive_count': (series > 0).sum(),
            'negative_count': (series < 0).sum(),
            'zero_count': (series == 0).sum()
        }


class ExcelUtils:
    """Utility functions for Excel file operations"""

    @staticmethod
    def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
        """
        Standardize column names for consistent processing

        Args:
            df: DataFrame with potentially inconsistent column names

        Returns:
            DataFrame with standardized column names
        """
        # Common column name mappings
        column_mappings = {
            'selected cost': 'Selected Cost',
            'selectedcost': 'Selected Cost',
            'least cost': 'Least Cost',
            'leastcost': 'Least Cost',
            'potential savings': 'Potential Savings',
            'potentialsavings': 'Potential Savings',
            'savings': 'Potential Savings',
            'origin city': 'Origin City',
            'origincity': 'Origin City',
            'destination city': 'Destination City',
            'destinationcity': 'Destination City',
            'carrier': 'Carrier',
            'mode': 'Mode'
        }

        # Create a copy to avoid modifying original
        df_copy = df.copy()

        # Rename columns using mapping (case-insensitive)
        for old_name, new_name in column_mappings.items():
            for col in df_copy.columns:
                if col.lower().strip() == old_name.lower():
                    df_copy = df_copy.rename(columns={col: new_name})
                    break

        return df_copy

    @staticmethod
    def detect_processor_type_from_file(file_path: str) -> str:
        """
        Detect the appropriate processor type based on file name or content

        Args:
            file_path: Path to the Excel file

        Returns:
            Suggested processor type ('basic', 'utc_main', 'utc_fs', 'transco', 'detailed')
        """
        import os
        filename = os.path.basename(file_path).lower()

        # File name based detection
        if 'utc_main' in filename or 'evanston' in filename:
            return 'utc_main'
        elif 'utc_fs' in filename or 'green_river' in filename:
            return 'utc_fs'
        elif 'transco' in filename or 'miles_city' in filename:
            return 'transco'
        elif 'cast_nylon' in filename or 'detailed' in filename or 'cn_' in filename:
            return 'detailed'
        else:
            return 'basic'


def create_summary_report(processing_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Create a summary report from multiple processing results

    Args:
        processing_results: List of processing result dictionaries

    Returns:
        Summary report dictionary
    """
    if not processing_results:
        return {'error': 'No processing results provided'}

    total_files = len(processing_results)
    successful_files = sum(1 for result in processing_results if result.get('success', False))
    failed_files = total_files - successful_files

    total_savings = sum(
        result.get('stats', {}).get('total_potential_savings', 0)
        for result in processing_results
        if result.get('success', False)
    )

    total_loads = sum(
        result.get('stats', {}).get('total_loads', 0)
        for result in processing_results
        if result.get('success', False)
    )

    return {
        'summary': {
            'total_files_processed': total_files,
            'successful_files': successful_files,
            'failed_files': failed_files,
            'success_rate': (successful_files / total_files * 100) if total_files > 0 else 0,
            'total_potential_savings': total_savings,
            'total_loads': total_loads,
            'average_savings_per_load': (total_savings / total_loads) if total_loads > 0 else 0
        },
        'details': processing_results
    }