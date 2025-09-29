#!/usr/bin/env python3
"""
Abstract base class for TMS processors
Provides a consistent interface for all processor types
"""

from abc import ABC, abstractmethod
import pandas as pd
from typing import Dict, Any, Optional


class TMSProcessorInterface(ABC):
    """Abstract base class defining the interface for all TMS processors"""

    def __init__(self):
        self.processed_data: Optional[pd.DataFrame] = None
        self.stats: Dict[str, Any] = {}

    @abstractmethod
    def process_excel_file(self, file_path: str) -> pd.DataFrame:
        """
        Process an Excel file and return the processed DataFrame

        Args:
            file_path: Path to the Excel file to process

        Returns:
            pd.DataFrame: Processed data
        """
        pass

    @abstractmethod
    def clean_and_process_data(self, file_path: str) -> pd.DataFrame:
        """
        Compatibility method for GUI integration

        Args:
            file_path: Path to the Excel file to process

        Returns:
            pd.DataFrame: Processed data
        """
        pass

    @abstractmethod
    def save_processed_data(self, output_file: str) -> None:
        """
        Save processed data to an Excel file

        Args:
            output_file: Path where to save the processed data
        """
        pass

    @abstractmethod
    def calculate_summary_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Calculate summary statistics for the processed data

        Args:
            df: Processed DataFrame

        Returns:
            Dict containing summary statistics
        """
        pass

    def get_stats(self) -> Dict[str, Any]:
        """Get the current statistics"""
        return self.stats

    def get_processed_data(self) -> Optional[pd.DataFrame]:
        """Get the processed data"""
        return self.processed_data


class ProcessorFactory:
    """Factory class to create processor instances"""

    @staticmethod
    def create_processor(processor_type: str) -> TMSProcessorInterface:
        """
        Create a processor instance based on type

        Args:
            processor_type: Type of processor to create
                          ('basic', 'utc_main', 'utc_fs', 'transco', 'detailed')

        Returns:
            TMSProcessorInterface: Processor instance

        Raises:
            ValueError: If processor type is not recognized
        """
        processor_type = processor_type.lower()

        if processor_type == 'basic':
            from basic_processor import BasicTMSProcessor
            return BasicTMSProcessor()
        elif processor_type == 'utc_main':
            from city_processors import UTCMainProcessor
            return UTCMainProcessor()
        elif processor_type == 'utc_fs':
            from city_processors import UTCFSProcessor
            return UTCFSProcessor()
        elif processor_type == 'transco':
            from city_processors import TranscoProcessor
            return TranscoProcessor()
        elif processor_type == 'detailed':
            # For Cast Nylon processing - would need to implement DetailedProcessor
            # For now, fall back to BasicTMSProcessor
            from basic_processor import BasicTMSProcessor
            return BasicTMSProcessor()
        else:
            raise ValueError(f"Unknown processor type: {processor_type}")

    @staticmethod
    def get_available_processors():
        """Get list of available processor types"""
        return ['basic', 'utc_main', 'utc_fs', 'transco', 'detailed']


def process_file_headless(file_path: str, processor_type: str, output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Automation-ready function to process a file without GUI

    Args:
        file_path: Path to the Excel file to process
        processor_type: Type of processor to use
        output_path: Optional path to save processed data

    Returns:
        Dict containing processing results and statistics
    """
    try:
        # Create processor
        processor = ProcessorFactory.create_processor(processor_type)

        # Process the file
        processed_df = processor.process_excel_file(file_path)

        # Get statistics
        stats = processor.get_stats()

        # Save if output path provided
        if output_path:
            processor.save_processed_data(output_path)

        return {
            'success': True,
            'stats': stats,
            'row_count': len(processed_df),
            'processor_type': processor_type,
            'file_path': file_path,
            'output_path': output_path
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'processor_type': processor_type,
            'file_path': file_path
        }