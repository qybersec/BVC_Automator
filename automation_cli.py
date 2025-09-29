#!/usr/bin/env python3
"""
BVC Automator - Headless CLI Interface
Automation-ready TMS processing without GUI dependencies
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

from processor_interface import ProcessorFactory, process_file_headless
from tms_utils import create_summary_report, ExcelUtils


def process_single_file(file_path: str, processor_type: str, output_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Process a single Excel file with the specified processor type

    Args:
        file_path: Path to the Excel file
        processor_type: Type of processor to use
        output_dir: Optional output directory for processed files

    Returns:
        Dictionary with processing results
    """
    try:
        # Auto-detect processor type if not specified
        if processor_type == 'auto':
            processor_type = ExcelUtils.detect_processor_type_from_file(file_path)
            print(f"Auto-detected processor type: {processor_type}")

        # Generate output path if output_dir provided
        output_path = None
        if output_dir:
            output_dir_path = Path(output_dir)
            output_dir_path.mkdir(parents=True, exist_ok=True)

            file_stem = Path(file_path).stem
            output_path = str(output_dir_path / f"{file_stem}_processed.xlsx")

        # Process the file
        result = process_file_headless(file_path, processor_type, output_path)

        # Add file-specific information
        result['input_file'] = file_path
        result['output_file'] = output_path if output_path else 'Not saved'

        return result

    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'input_file': file_path,
            'processor_type': processor_type
        }


def process_batch_files(file_paths: List[str], processor_type: str,
                       output_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Process multiple Excel files in batch

    Args:
        file_paths: List of file paths to process
        processor_type: Type of processor to use (or 'auto' for detection)
        output_dir: Optional output directory for processed files

    Returns:
        Dictionary with batch processing results
    """
    results = []

    print(f"Processing {len(file_paths)} files...")

    for i, file_path in enumerate(file_paths, 1):
        print(f"\nProcessing file {i}/{len(file_paths)}: {Path(file_path).name}")

        result = process_single_file(file_path, processor_type, output_dir)
        results.append(result)

        if result['success']:
            stats = result.get('stats', {})
            savings = stats.get('total_potential_savings', 0)
            loads = stats.get('total_loads', 0)
            print(f"  ✓ Success: ${savings:.2f} savings from {loads} loads")
        else:
            print(f"  ✗ Failed: {result.get('error', 'Unknown error')}")

    # Create summary report
    summary = create_summary_report(results)

    return summary


def export_results(results: Dict[str, Any], export_path: str) -> None:
    """Export processing results to JSON file"""
    try:
        with open(export_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, default=str)
        print(f"\nResults exported to: {export_path}")
    except Exception as e:
        print(f"Failed to export results: {e}")


def print_summary(results: Dict[str, Any]) -> None:
    """Print a formatted summary of processing results"""
    summary = results.get('summary', {})

    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)

    print(f"Total files processed: {summary.get('total_files_processed', 0)}")
    print(f"Successful files: {summary.get('successful_files', 0)}")
    print(f"Failed files: {summary.get('failed_files', 0)}")
    print(f"Success rate: {summary.get('success_rate', 0):.1f}%")
    print(f"Total potential savings: ${summary.get('total_potential_savings', 0):.2f}")
    print(f"Total loads processed: {summary.get('total_loads', 0)}")

    avg_savings = summary.get('average_savings_per_load', 0)
    if avg_savings > 0:
        print(f"Average savings per load: ${avg_savings:.2f}")

    # Show failed files if any
    failed_files = [r for r in results.get('details', []) if not r.get('success', False)]
    if failed_files:
        print(f"\nFailed files:")
        for result in failed_files:
            print(f"  - {Path(result.get('input_file', '')).name}: {result.get('error', 'Unknown error')}")


def main():
    """Main CLI interface"""
    parser = argparse.ArgumentParser(
        description="BVC Automator - Headless TMS Processing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process single file with auto-detection
  python automation_cli.py file.xlsx

  # Process with specific processor type
  python automation_cli.py file.xlsx --type utc_main

  # Process multiple files in batch
  python automation_cli.py file1.xlsx file2.xlsx file3.xlsx --type basic

  # Process with output directory
  python automation_cli.py *.xlsx --type auto --output-dir ./processed

  # Export results to JSON
  python automation_cli.py files/*.xlsx --export results.json

Available processor types: basic, utc_main, utc_fs, transco, detailed, auto
        """
    )

    parser.add_argument(
        'files',
        nargs='+',
        help='Excel file(s) to process'
    )

    parser.add_argument(
        '--type', '-t',
        choices=['basic', 'utc_main', 'utc_fs', 'transco', 'detailed', 'auto'],
        default='auto',
        help='Processor type to use (default: auto-detect)'
    )

    parser.add_argument(
        '--output-dir', '-o',
        help='Output directory for processed files'
    )

    parser.add_argument(
        '--export', '-e',
        help='Export results to JSON file'
    )

    parser.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='Suppress detailed output'
    )

    parser.add_argument(
        '--list-types',
        action='store_true',
        help='List available processor types and exit'
    )

    args = parser.parse_args()

    # Handle list types command
    if args.list_types:
        print("Available processor types:")
        for ptype in ProcessorFactory.get_available_processors():
            print(f"  {ptype}")
        return 0

    # Validate files exist
    valid_files = []
    for file_path in args.files:
        path = Path(file_path)
        if path.exists() and path.is_file():
            valid_files.append(str(path.absolute()))
        else:
            print(f"Warning: File not found: {file_path}")

    if not valid_files:
        print("Error: No valid files to process")
        return 1

    # Process files
    try:
        if len(valid_files) == 1:
            # Single file processing
            if not args.quiet:
                print(f"Processing single file: {Path(valid_files[0]).name}")

            result = process_single_file(valid_files[0], args.type, args.output_dir)

            if result['success']:
                if not args.quiet:
                    stats = result.get('stats', {})
                    savings = stats.get('total_potential_savings', 0)
                    loads = stats.get('total_loads', 0)
                    print(f"Success: ${savings:.2f} savings from {loads} loads")
                    if args.output_dir:
                        print(f"Output saved to: {result.get('output_file')}")
            else:
                print(f"Error: {result.get('error')}")
                return 1

            # Wrap single result for consistent handling
            results = {
                'summary': {
                    'total_files_processed': 1,
                    'successful_files': 1 if result['success'] else 0,
                    'failed_files': 0 if result['success'] else 1,
                    'success_rate': 100.0 if result['success'] else 0.0,
                    'total_potential_savings': result.get('stats', {}).get('total_potential_savings', 0),
                    'total_loads': result.get('stats', {}).get('total_loads', 0),
                    'average_savings_per_load': result.get('stats', {}).get('average_savings_per_load', 0)
                },
                'details': [result]
            }

        else:
            # Batch processing
            results = process_batch_files(valid_files, args.type, args.output_dir)

        # Print summary unless quiet mode
        if not args.quiet:
            print_summary(results)

        # Export results if requested
        if args.export:
            export_results(results, args.export)

        # Return appropriate exit code
        summary = results.get('summary', {})
        failed_files = summary.get('failed_files', 0)
        return 1 if failed_files > 0 else 0

    except KeyboardInterrupt:
        print("\nProcessing interrupted by user")
        return 1
    except Exception as e:
        print(f"Unexpected error: {e}")
        return 1


if __name__ == '__main__':
    sys.exit(main())