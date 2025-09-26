#!/usr/bin/env python3
"""
Basic TMS Processor - Core processing logic shared by all processors
This contains the fundamental TMS business rules that apply to all clients
"""

import pandas as pd
import os
from typing import Dict, Any


class BasicTMSProcessor:
    """
    Core TMS processing logic shared by UTC Main, UTC FS, and Transco
    Contains all the fundamental business rules without city-specific customizations
    """

    def __init__(self):
        # File structure settings
        self.HEADER_ROW = 8          # Row 9 in Excel (0-indexed) - actual header row
        self.DATA_START_ROW = 10     # Row 11 in Excel (0-indexed) - first data row

        # TL carriers requiring special processing
        self.TL_CARRIERS = {
            'LANDSTAR RANGER INC',
            'SMARTWAY TRANSPORTATION INC',
            'ONX LOGISTICS INC'
        }

        # Results storage
        self.processed_data = None
        self.summary_stats = {}
        self.title_info = {}

    def load_data(self, file_path: str) -> pd.DataFrame:
        """Load and clean Excel data using Basic_Processor logic"""
        # Read Excel file
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # Extract title information
        self.title_info = self._extract_title_info(df_raw)

        # Get headers and data
        headers = df_raw.iloc[self.HEADER_ROW].fillna('').astype(str).tolist()
        data_df = df_raw.iloc[self.DATA_START_ROW:].copy()
        data_df.columns = headers

        # Map column names to expected format
        data_df = self._map_columns(data_df)

        # Handle PS -> Potential Savings column mapping
        if 'PS' in data_df.columns:
            data_df = data_df.rename(columns={'PS': 'Potential Savings'})

        # Clean the data
        data_df = self._clean_data(data_df)

        # Calculate Potential Savings if missing
        data_df = self._calculate_potential_savings(data_df)

        return data_df

    def apply_basic_business_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply the 5 core TMS business rules (excluding city-specific rules)"""
        df = df.copy()
        print("Applying Basic TMS business rules...")

        # Rule 1: Same Carrier Rule
        df = self._apply_same_carrier_rule(df)

        # Rule 2: Empty Least Cost Data Rule
        df = self._apply_empty_data_rule(df)

        # Rule 3: Negative Savings Rule
        df = self._apply_negative_savings_rule(df)

        # Rule 4: TL Carriers Rule
        df = self._apply_tl_carriers_rule(df)

        # Rule 5: DALKO DEFENDER INSURANCE Rule
        df = self._apply_dalko_rule(df)

        return df

    def _apply_same_carrier_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rule 1: Set PS to 0 when selected carrier = least cost carrier"""
        if 'Selected Carrier' in df.columns and 'Least Cost Carrier' in df.columns:
            mask = (
                (df['Selected Carrier'].astype(str) == df['Least Cost Carrier'].astype(str)) &
                (df['Selected Carrier'].notna()) &
                (df['Least Cost Carrier'].notna()) &
                (df['Selected Carrier'].astype(str) != '') &
                (df['Least Cost Carrier'].astype(str) != '')
            )

            count = mask.sum()
            if count > 0 and 'Potential Savings' in df.columns:
                df.loc[mask, 'Potential Savings'] = 0
                print(f"Same Carrier Rule: {count} rows")

        return df

    def _apply_empty_data_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rule 2: Copy selected data when least cost data is missing"""
        if 'Least Cost Carrier' in df.columns:
            mask = (
                df['Least Cost Carrier'].isna() |
                (df['Least Cost Carrier'].astype(str) == '') |
                (df['Least Cost Carrier'].astype(str) == 'nan')
            )

            count = mask.sum()
            if count > 0:
                self._copy_selected_to_least_cost(df, mask)
                if 'Potential Savings' in df.columns:
                    df.loc[mask, 'Potential Savings'] = 0
                print(f"Empty Data Rule: {count} rows")

        return df

    def _apply_negative_savings_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rule 3: Fix negative potential savings"""
        if 'Potential Savings' in df.columns:
            ps_numeric = pd.to_numeric(df['Potential Savings'], errors='coerce').fillna(0)
            mask = ps_numeric < 0
            count = mask.sum()

            if count > 0:
                self._copy_selected_to_least_cost(df, mask)
                df.loc[mask, 'Potential Savings'] = 0
                print(f"Negative Savings Rule: {count} rows")

        return df

    def _apply_tl_carriers_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rule 4: Special handling for TL carriers"""
        if 'Selected Carrier' in df.columns and 'Least Cost Carrier' in df.columns:
            mask = (
                df['Selected Carrier'].astype(str).str.upper().isin(
                    [c.upper() for c in self.TL_CARRIERS]
                ) |
                df['Least Cost Carrier'].astype(str).str.upper().isin(
                    [c.upper() for c in self.TL_CARRIERS]
                )
            )

            count = mask.sum()
            if count > 0:
                self._copy_selected_to_least_cost(df, mask)
                if 'Potential Savings' in df.columns:
                    df.loc[mask, 'Potential Savings'] = 0
                print(f"TL Carriers Rule: {count} rows")

        return df

    def _apply_dalko_rule(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rule 5: DALKO DEFENDER INSURANCE pattern matching"""
        if 'Selected Carrier' in df.columns and 'Least Cost Carrier' in df.columns:
            dalko_matches = []

            for idx, row in df.iterrows():
                selected = str(row['Selected Carrier']).strip().upper()
                least_cost = str(row['Least Cost Carrier']).strip().upper()

                # Skip empty or nan values
                if selected in ['', 'NAN', 'NONE'] or least_cost in ['', 'NAN', 'NONE']:
                    continue

                # Check for DALKO DEFENDER INSURANCE patterns
                # Pattern 1: "DALKO DEFENDER INSURANCE/XXX" where least cost is "XXX"
                if 'DALKO DEFENDER INSURANCE/' in selected:
                    carrier_after_slash = selected.split('DALKO DEFENDER INSURANCE/')[-1].strip()
                    if carrier_after_slash == least_cost:
                        dalko_matches.append(idx)

                # Pattern 2: "XXX/DALKO DEFENDER INSURANCE" where least cost is "XXX"
                elif '/DALKO DEFENDER INSURANCE' in selected:
                    carrier_before_slash = selected.split('/DALKO DEFENDER INSURANCE')[0].strip()
                    if carrier_before_slash == least_cost:
                        dalko_matches.append(idx)

            count = len(dalko_matches)
            if count > 0:
                dalko_mask = df.index.isin(dalko_matches)
                self._copy_selected_to_least_cost(df, dalko_mask)
                if 'Potential Savings' in df.columns:
                    df.loc[dalko_mask, 'Potential Savings'] = 0
                print(f"DALKO DEFENDER INSURANCE Rule: {count} rows")

        return df

    def _copy_selected_to_least_cost(self, df: pd.DataFrame, mask: pd.Series) -> None:
        """Helper: Copy selected carrier data to least cost columns"""
        column_pairs = [
            ('Selected Carrier', 'Least Cost Carrier'),
            ('Selected Service Type', 'Least Cost Service Type'),
            ('Selected Transit Days', 'Least Cost Transit Days'),
            ('Selected Freight Cost', 'Least Cost Freight Cost'),
            ('Selected Accessorial Cost', 'Least Cost Accessorial Cost'),
            ('Selected Total Cost', 'Least Cost Total Cost')
        ]

        for selected_col, least_cost_col in column_pairs:
            if selected_col in df.columns and least_cost_col in df.columns:
                df.loc[mask, least_cost_col] = df.loc[mask, selected_col]

    def _map_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Map column names to standardized format"""
        df = df.copy()

        # Find carrier columns - there should be two "Carrier" columns
        new_columns = list(df.columns)
        carrier_positions = []

        for i, col in enumerate(new_columns):
            if col == 'Carrier':
                carrier_positions.append(i)

        # Rename the two "Carrier" columns
        if len(carrier_positions) >= 2:
            new_columns[carrier_positions[0]] = 'Selected Carrier'
            new_columns[carrier_positions[1]] = 'Least Cost Carrier'

        # Map other column names to standardized format
        column_mapping = {
            'Service Type': 'Selected Service Type',  # First occurrence
            'Transit\nDays': 'Selected Transit Days',  # First occurrence
            'Freight + Fuel': 'Selected Freight Cost',  # First occurrence
            'Total Acc.': 'Selected Accessorial Cost',  # First occurrence
            'Total Cost ': 'Selected Total Cost'  # First occurrence (note the trailing space)
        }

        # Apply mapping to first occurrence of each column
        for old_name, new_name in column_mapping.items():
            if old_name in new_columns:
                # Find first occurrence and rename it
                idx = new_columns.index(old_name)
                new_columns[idx] = new_name

        # Now handle the second set (Least Cost columns)
        least_cost_mapping = {
            'Service Type': 'Least Cost Service Type',
            'Transit\nDays': 'Least Cost Transit Days',
            'Freight + Fuel': 'Least Cost Freight Cost',
            'Total Acc.': 'Least Cost Accessorial Cost',
            'Total Cost ': 'Least Cost Total Cost'
        }

        # Find and rename second occurrences
        for old_name, new_name in least_cost_mapping.items():
            # Count occurrences and rename the second one
            occurrences = [i for i, col in enumerate(new_columns) if col == old_name]
            if len(occurrences) >= 2:
                new_columns[occurrences[1]] = new_name

        df.columns = new_columns
        return df

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and validate data types"""
        df = df.copy()

        # Remove rows with missing Load No.
        if 'Load No.' in df.columns:
            df = df.dropna(subset=['Load No.'])
            df = df[df['Load No.'].astype(str).str.strip() != '']
            df = df[df['Load No.'].astype(str) != 'nan']

        # Clean numeric columns
        numeric_columns = [
            'Selected Transit Days', 'Selected Freight Cost',
            'Selected Accessorial Cost', 'Selected Total Cost',
            'Least Cost Transit Days', 'Least Cost Freight Cost',
            'Least Cost Accessorial Cost', 'Least Cost Total Cost',
            'Potential Savings'
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Clean string columns
        string_columns = [
            'Selected Carrier', 'Selected Service Type',
            'Least Cost Carrier', 'Least Cost Service Type'
        ]

        for col in string_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('nan', '')

        # Remove mostly empty rows
        df = df.dropna(thresh=5).reset_index(drop=True)

        return df

    def _calculate_potential_savings(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate Potential Savings if missing"""
        df = df.copy()

        # If PS already exists and has valid data, keep it
        if 'Potential Savings' in df.columns:
            ps_numeric = pd.to_numeric(df['Potential Savings'], errors='coerce')
            if ps_numeric.notna().sum() > 0:
                return df

        # Calculate PS from cost difference
        if 'Selected Total Cost' in df.columns and 'Least Cost Total Cost' in df.columns:
            selected_cost = pd.to_numeric(df['Selected Total Cost'], errors='coerce').fillna(0)
            least_cost = pd.to_numeric(df['Least Cost Total Cost'], errors='coerce').fillna(0)
            df['Potential Savings'] = selected_cost - least_cost

            # Handle cases where least cost is 0
            mask_zero_least = least_cost == 0
            df.loc[mask_zero_least, 'Potential Savings'] = 0
        else:
            df['Potential Savings'] = 0

        return df

    def _extract_title_info(self, df_raw: pd.DataFrame) -> Dict[str, Any]:
        """Extract title and company information from Excel header"""
        title_info = {}

        try:
            # Company name (row 4, column B)
            if len(df_raw) > 3 and len(df_raw.columns) > 1:
                company = df_raw.iloc[3, 1]
                if pd.notna(company):
                    title_info['company_name'] = str(company)

            # Date range (row 6, column B)
            if len(df_raw) > 5 and len(df_raw.columns) > 1:
                date_range = df_raw.iloc[5, 1]
                if pd.notna(date_range):
                    title_info['date_range'] = str(date_range)
        except Exception:
            pass

        return title_info

    def calculate_summary_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate comprehensive summary statistics"""
        if df.empty:
            stats = {
                'total_loads': 0,
                'total_potential_savings': 0,
                'average_savings_per_load': 0,
                'loads_with_savings': 0,
                'total_selected_cost': 0,
                'total_least_cost': 0,
                'percentage_savings': 0
            }
            self.summary_stats = stats
            return stats

        # Calculate all stats
        total_loads = len(df)

        if 'Selected Total Cost' in df.columns:
            total_selected_cost = pd.to_numeric(
                df['Selected Total Cost'], errors='coerce').fillna(0).sum()
        else:
            total_selected_cost = 0

        if 'Least Cost Total Cost' in df.columns:
            total_least_cost = pd.to_numeric(
                df['Least Cost Total Cost'], errors='coerce').fillna(0).sum()
        else:
            total_least_cost = 0

        if 'Potential Savings' in df.columns:
            ps_series = pd.to_numeric(df['Potential Savings'], errors='coerce').fillna(0)
            total_potential_savings = ps_series.sum()
            ps_numeric = ps_series
        else:
            total_potential_savings = 0
            ps_numeric = pd.Series([0] * len(df))

        loads_with_savings = (ps_numeric > 0).sum()

        percentage_savings = (
            (total_potential_savings / total_selected_cost * 100)
            if total_selected_cost > 0 else 0
        )
        average_savings_per_load = total_potential_savings / total_loads if total_loads > 0 else 0

        stats = {
            'total_loads': total_loads,
            'total_potential_savings': total_potential_savings,
            'average_savings_per_load': average_savings_per_load,
            'loads_with_savings': loads_with_savings,
            'total_selected_cost': total_selected_cost,
            'total_least_cost': total_least_cost,
            'percentage_savings': percentage_savings
        }

        self.summary_stats = stats
        return stats