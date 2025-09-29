#!/usr/bin/env python3
"""
City-Specific TMS Processors
Each processor inherits from BasicTMSProcessor and adds its own city-specific rule
"""

import os
import pandas as pd
from basic_processor import BasicTMSProcessor
from tms_utils import CityRuleProcessor


class UTCMainProcessor(BasicTMSProcessor):
    """UTC Main processor with Evanston city rule"""

    def process_excel_file(self, file_path: str) -> pd.DataFrame:
        """Complete processing pipeline for UTC Main"""
        print(f"Processing UTC Main report: {os.path.basename(file_path)}")

        # Step 1: Load data (inherited from Basic)
        df = self.load_data(file_path)

        # Step 2: Apply basic business rules (inherited from Basic)
        df = self.apply_basic_business_rules(df)

        # Step 3: Apply UTC Main specific rule (Evanston)
        df = self._apply_evanston_rule(df)

        # Step 4: Calculate summary stats (inherited from Basic)
        self.calculate_summary_stats(df)

        # Step 5: Sort by destination city
        if 'Destination City' in df.columns:
            df = df.sort_values('Destination City', na_position='last')
            print(f"Sorted by Destination City - {len(df)} rows")

        self.processed_data = df
        return df

    def clean_and_process_data(self, file_path: str) -> pd.DataFrame:
        """Compatibility method for GUI - calls process_excel_file"""
        return self.process_excel_file(file_path)


    def _apply_evanston_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Evanston City Rule - Zero out PS for any Evanston origin or destination"""
        return CityRuleProcessor.apply_city_exclusion_rule(df, "Evanston", "Evanston City Rule")


class UTCFSProcessor(BasicTMSProcessor):
    """UTC FS processor with Green River city rule"""

    def process_excel_file(self, file_path: str) -> pd.DataFrame:
        """Complete processing pipeline for UTC FS"""
        print(f"Processing UTC FS report: {os.path.basename(file_path)}")

        # Step 1: Load data (inherited from Basic)
        df = self.load_data(file_path)

        # Step 2: Apply basic business rules (inherited from Basic)
        df = self.apply_basic_business_rules(df)

        # Step 3: Apply UTC FS specific rule (Green River)
        df = self._apply_green_river_rule(df)

        # Step 4: Calculate summary stats (inherited from Basic)
        self.calculate_summary_stats(df)

        # Step 5: Sort by destination city
        if 'Destination City' in df.columns:
            df = df.sort_values('Destination City', na_position='last')
            print(f"Sorted by Destination City - {len(df)} rows")

        self.processed_data = df
        return df

    def clean_and_process_data(self, file_path: str) -> pd.DataFrame:
        """Compatibility method for GUI - calls process_excel_file"""
        return self.process_excel_file(file_path)


    def _apply_green_river_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Green River City Rule - Zero out PS for any Green River origin or destination"""
        return CityRuleProcessor.apply_city_exclusion_rule(df, "Green River", "Green River Rule")


class TranscoProcessor(BasicTMSProcessor):
    """Transco processor with Miles City rule"""

    def process_excel_file(self, file_path: str) -> pd.DataFrame:
        """Complete processing pipeline for Transco"""
        print(f"Processing Transco report: {os.path.basename(file_path)}")

        # Step 1: Load data (inherited from Basic)
        df = self.load_data(file_path)

        # Step 2: Apply basic business rules (inherited from Basic)
        df = self.apply_basic_business_rules(df)

        # Step 3: Apply Transco specific rule (Miles City)
        df = self._apply_miles_city_rule(df)

        # Step 4: Calculate summary stats (inherited from Basic)
        self.calculate_summary_stats(df)

        # Step 5: Sort by destination city
        if 'Destination City' in df.columns:
            df = df.sort_values('Destination City', na_position='last')
            print(f"Sorted by Destination City - {len(df)} rows")

        self.processed_data = df
        return df

    def clean_and_process_data(self, file_path: str) -> pd.DataFrame:
        """Compatibility method for GUI - calls process_excel_file"""
        return self.process_excel_file(file_path)


    def _apply_miles_city_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Miles City Rule - Zero out PS for any Miles City origin or destination"""
        return CityRuleProcessor.apply_city_exclusion_rule(df, "Miles City", "Miles City Rule")