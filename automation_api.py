#!/usr/bin/env python3
"""
BVC Automator - Python API
Simple programmatic interface for TMS processing
"""

from typing import List, Dict, Any, Union, Optional
from pathlib import Path

from processor_interface import ProcessorFactory, process_file_headless
from tms_utils import create_summary_report, ExcelUtils


class TMSAutomator:
    """
    Main class for automated TMS processing
    Provides a simple Python API for processing Excel files
    """

    def __init__(self):
        """Initialize the TMS Automator"""
        self.results_history: List[Dict[str, Any]] = []

    def process_file(self, file_path: Union[str, Path], processor_type: str = 'auto',
                    output_path: Optional[Union[str, Path]] = None) -> Dict[str, Any]:
        """
        Process a single Excel file

        Args:
            file_path: Path to the Excel file to process
            processor_type: Type of processor ('auto', 'basic', 'utc_main', 'utc_fs', 'transco', 'detailed')
            output_path: Optional path to save processed file

        Returns:
            Dictionary with processing results and statistics

        Example:
            >>> automator = TMSAutomator()
            >>> result = automator.process_file('data.xlsx', 'utc_main')
            >>> print(f"Savings: ${result['stats']['total_potential_savings']}")
        """
        file_path = str(Path(file_path).absolute())

        # Auto-detect processor type if needed
        if processor_type == 'auto':
            processor_type = ExcelUtils.detect_processor_type_from_file(file_path)

        # Convert output_path to string if provided
        output_path_str = str(output_path) if output_path else None

        # Process the file
        result = process_file_headless(file_path, processor_type, output_path_str)

        # Add to history
        self.results_history.append(result)

        return result

    def process_files(self, file_paths: List[Union[str, Path]], processor_type: str = 'auto',
                     output_dir: Optional[Union[str, Path]] = None) -> Dict[str, Any]:
        """
        Process multiple Excel files

        Args:
            file_paths: List of file paths to process
            processor_type: Type of processor ('auto', 'basic', 'utc_main', 'utc_fs', 'transco', 'detailed')
            output_dir: Optional directory to save processed files

        Returns:
            Dictionary with batch processing summary

        Example:
            >>> automator = TMSAutomator()
            >>> files = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']
            >>> summary = automator.process_files(files, 'basic', './output')
            >>> print(f"Total savings: ${summary['summary']['total_potential_savings']}")
        """
        results = []

        # Create output directory if specified
        if output_dir:
            output_dir_path = Path(output_dir)
            output_dir_path.mkdir(parents=True, exist_ok=True)

        for file_path in file_paths:
            file_path = Path(file_path)

            # Generate output path if output_dir specified
            output_path = None
            if output_dir:
                output_path = Path(output_dir) / f"{file_path.stem}_processed.xlsx"

            # Process the file
            result = self.process_file(file_path, processor_type, output_path)
            results.append(result)

        # Create and return summary
        summary = create_summary_report(results)
        return summary

    def get_available_processors(self) -> List[str]:
        """
        Get list of available processor types

        Returns:
            List of processor type names

        Example:
            >>> automator = TMSAutomator()
            >>> types = automator.get_available_processors()
            >>> print(types)  # ['basic', 'utc_main', 'utc_fs', 'transco', 'detailed']
        """
        return ProcessorFactory.get_available_processors()

    def detect_processor_type(self, file_path: Union[str, Path]) -> str:
        """
        Auto-detect the appropriate processor type for a file

        Args:
            file_path: Path to the Excel file

        Returns:
            Suggested processor type

        Example:
            >>> automator = TMSAutomator()
            >>> ptype = automator.detect_processor_type('CN_data.xlsx')
            >>> print(ptype)  # 'detailed'
        """
        return ExcelUtils.detect_processor_type_from_file(str(file_path))

    def get_results_history(self) -> List[Dict[str, Any]]:
        """
        Get history of all processing results from this session

        Returns:
            List of processing result dictionaries

        Example:
            >>> automator = TMSAutomator()
            >>> automator.process_file('file1.xlsx')
            >>> automator.process_file('file2.xlsx')
            >>> history = automator.get_results_history()
            >>> print(f"Processed {len(history)} files this session")
        """
        return self.results_history.copy()

    def get_session_summary(self) -> Dict[str, Any]:
        """
        Get summary of all files processed in this session

        Returns:
            Summary dictionary with totals

        Example:
            >>> automator = TMSAutomator()
            >>> # ... process some files ...
            >>> summary = automator.get_session_summary()
            >>> print(f"Session total: ${summary['summary']['total_potential_savings']}")
        """
        if not self.results_history:
            return {
                'summary': {
                    'total_files_processed': 0,
                    'successful_files': 0,
                    'failed_files': 0,
                    'success_rate': 0.0,
                    'total_potential_savings': 0.0,
                    'total_loads': 0,
                    'average_savings_per_load': 0.0
                },
                'details': []
            }

        return create_summary_report(self.results_history)

    def clear_history(self) -> None:
        """
        Clear the results history

        Example:
            >>> automator = TMSAutomator()
            >>> # ... process some files ...
            >>> automator.clear_history()
        """
        self.results_history.clear()


# Convenience functions for quick usage
def quick_process(file_path: Union[str, Path], processor_type: str = 'auto') -> Dict[str, Any]:
    """
    Quick function to process a single file without creating an automator instance

    Args:
        file_path: Path to Excel file
        processor_type: Processor type to use

    Returns:
        Processing result dictionary

    Example:
        >>> from automation_api import quick_process
        >>> result = quick_process('data.xlsx', 'basic')
        >>> print(f"Savings: ${result['stats']['total_potential_savings']}")
    """
    automator = TMSAutomator()
    return automator.process_file(file_path, processor_type)


def quick_batch(file_paths: List[Union[str, Path]], processor_type: str = 'auto') -> Dict[str, Any]:
    """
    Quick function to process multiple files without creating an automator instance

    Args:
        file_paths: List of file paths
        processor_type: Processor type to use

    Returns:
        Batch processing summary

    Example:
        >>> from automation_api import quick_batch
        >>> files = ['file1.xlsx', 'file2.xlsx']
        >>> summary = quick_batch(files, 'utc_main')
        >>> print(f"Total: ${summary['summary']['total_potential_savings']}")
    """
    automator = TMSAutomator()
    return automator.process_files(file_paths, processor_type)


if __name__ == '__main__':
    # Example usage
    print("BVC Automator API - Example Usage")
    print("="*40)

    # Create automator instance
    automator = TMSAutomator()

    # Show available processors
    print("Available processors:", automator.get_available_processors())

    # Example file detection (would need actual file)
    # print("Detected type for 'CN_data.xlsx':", automator.detect_processor_type('CN_data.xlsx'))

    print("\nAPI ready for use. Import this module in your Python scripts:")