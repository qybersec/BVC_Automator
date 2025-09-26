#!/usr/bin/env python3
"""
City-Specific TMS Processors
Each processor inherits from BasicTMSProcessor and adds its own city-specific rule
"""

import os
import pandas as pd
from basic_processor import BasicTMSProcessor


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

    def save_processed_data(self, output_file: str) -> None:
        """Save processed data to Excel file with formatting"""
        if self.processed_data is None:
            raise ValueError("No processed data available to save")

        # Simple Excel save for now - just save the DataFrame
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.processed_data.to_excel(writer, sheet_name='Processed Data', index=False)

    def _apply_evanston_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Evanston City Rule - Zero out PS for any Evanston origin or destination"""
        evanston_count = 0

        # Check if city columns exist
        origin_col_found = 'Origin City' in df.columns
        dest_col_found = 'Destination City' in df.columns

        if not (origin_col_found or dest_col_found):
            print("Evanston Rule: No city columns found, skipping rule")
            return df

        # Create mask for Evanston cities
        evanston_mask = pd.Series([False] * len(df), index=df.index)

        # Check Origin City for "Evanston"
        if origin_col_found:
            origin_evanston = df['Origin City'].astype(str).str.upper().str.contains(
                'EVANSTON', na=False
            )
            evanston_mask |= origin_evanston
            origin_count = origin_evanston.sum()
            if origin_count > 0:
                print(f"Found {origin_count} rows with Evanston in Origin City")

        # Check Destination City for "Evanston"
        if dest_col_found:
            dest_evanston = df['Destination City'].astype(str).str.upper().str.contains(
                'EVANSTON', na=False
            )
            evanston_mask |= dest_evanston
            dest_count = dest_evanston.sum()
            if dest_count > 0:
                print(f"Found {dest_count} rows with Evanston in Destination City")

        evanston_count = evanston_mask.sum()

        if evanston_count > 0 and 'Potential Savings' in df.columns:
            # Zero out Potential Savings for Evanston rows
            df.loc[evanston_mask, 'Potential Savings'] = 0
            print(f"Evanston City Rule: Zeroed out PS for {evanston_count} rows")
        elif evanston_count > 0:
            print(
                f"Evanston City Rule: Found {evanston_count} Evanston rows "
                f"but no Potential Savings column"
            )
        else:
            print("Evanston City Rule: No Evanston cities found")

        return df


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

    def save_processed_data(self, output_file: str) -> None:
        """Save processed data to Excel file with formatting"""
        if self.processed_data is None:
            raise ValueError("No processed data available to save")

        # Simple Excel save for now - just save the DataFrame
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.processed_data.to_excel(writer, sheet_name='Processed Data', index=False)

    def _apply_green_river_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Green River City Rule - Zero out PS for any Green River origin or destination"""
        green_river_count = 0

        # Check if city columns exist
        origin_col_found = 'Origin City' in df.columns
        dest_col_found = 'Destination City' in df.columns

        if not (origin_col_found or dest_col_found):
            print("Green River Rule: No city columns found, skipping rule")
            return df

        # Create mask for Green River cities
        green_river_mask = pd.Series([False] * len(df), index=df.index)

        # Check Origin City for "Green River"
        if origin_col_found:
            origin_green_river = df['Origin City'].astype(str).str.upper().str.contains(
                'GREEN RIVER', na=False
            )
            green_river_mask |= origin_green_river
            origin_count = origin_green_river.sum()
            if origin_count > 0:
                print(f"Found {origin_count} rows with Green River in Origin City")

        # Check Destination City for "Green River"
        if dest_col_found:
            dest_green_river = df['Destination City'].astype(str).str.upper().str.contains(
                'GREEN RIVER', na=False
            )
            green_river_mask |= dest_green_river
            dest_count = dest_green_river.sum()
            if dest_count > 0:
                print(f"Found {dest_count} rows with Green River in Destination City")

        green_river_count = green_river_mask.sum()

        if green_river_count > 0 and 'Potential Savings' in df.columns:
            # Zero out Potential Savings for Green River rows
            df.loc[green_river_mask, 'Potential Savings'] = 0
            print(f"Green River Rule: Zeroed out PS for {green_river_count} rows")
        elif green_river_count > 0:
            print(
                f"Green River Rule: Found {green_river_count} Green River rows "
                f"but no Potential Savings column"
            )
        else:
            print("Green River Rule: No Green River cities found")

        return df


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

    def save_processed_data(self, output_file: str) -> None:
        """Save processed data to Excel file with formatting"""
        if self.processed_data is None:
            raise ValueError("No processed data available to save")

        # Simple Excel save for now - just save the DataFrame
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.processed_data.to_excel(writer, sheet_name='Processed Data', index=False)

    def _apply_miles_city_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Miles City Rule - Zero out PS for any Miles City origin or destination"""
        miles_city_count = 0

        # Check if city columns exist
        origin_col_found = 'Origin City' in df.columns
        dest_col_found = 'Destination City' in df.columns

        if not (origin_col_found or dest_col_found):
            print("Miles City Rule: No city columns found, skipping rule")
            return df

        # Create mask for Miles City cities
        miles_city_mask = pd.Series([False] * len(df), index=df.index)

        # Check Origin City for "Miles City"
        if origin_col_found:
            origin_miles_city = df['Origin City'].astype(str).str.upper().str.contains(
                'MILES CITY', na=False
            )
            miles_city_mask |= origin_miles_city
            origin_count = origin_miles_city.sum()
            if origin_count > 0:
                print(f"Found {origin_count} rows with Miles City in Origin City")

        # Check Destination City for "Miles City"
        if dest_col_found:
            dest_miles_city = df['Destination City'].astype(str).str.upper().str.contains(
                'MILES CITY', na=False
            )
            miles_city_mask |= dest_miles_city
            dest_count = dest_miles_city.sum()
            if dest_count > 0:
                print(f"Found {dest_count} rows with Miles City in Destination City")

        miles_city_count = miles_city_mask.sum()

        if miles_city_count > 0 and 'Potential Savings' in df.columns:
            # Zero out Potential Savings for Miles City rows
            df.loc[miles_city_mask, 'Potential Savings'] = 0
            print(f"Miles City Rule: Zeroed out PS for {miles_city_count} rows")
        elif miles_city_count > 0:
            print(
                f"Miles City Rule: Found {miles_city_count} Miles City rows "
                f"but no Potential Savings column"
            )
        else:
            print("Miles City Rule: No Miles City cities found")

        return df