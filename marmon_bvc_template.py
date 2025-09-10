"""
Marmon BVC Excel Template Generator
Integrated with TMS Data Processor Pro
"""
import os
from openpyxl import Workbook
from pathlib import Path
from datetime import datetime
import logging

class MarmonBVCTemplateGenerator:
    """Generate Excel templates for Marmon BVC reporting"""
    
    def __init__(self):
        self.output_folder = "output_reports"
        self.logger = logging.getLogger(__name__)
        
        # Sheet names in the required order
        self.sheet_names = [
            "UTC MAIN",
            "UTC FS", 
            "TRANSCO",
            "PROCOR",
            "UTXLA",
            "MCKENZIE",
            "PENN MACHINE",
            "RAILSERVE",
            "TRACKMOBILE",
            "MARMON RR",
            "BVC"
        ]
    
    def generate_template(self, date_range: str, output_file: str = None) -> str:
        """
        Generate Marmon BVC template with specified date range
        
        Args:
            date_range: Date range string (e.g., "08.04.25 - 08.08.25")
            output_file: Optional custom output file path
            
        Returns:
            str: Path to the generated file
        """
        try:
            # Validate date range
            if not date_range or not date_range.strip():
                raise ValueError("Date range cannot be empty")
            
            date_range = date_range.strip()
            
            # Ensure output folder exists
            os.makedirs(self.output_folder, exist_ok=True)
            
            # Generate filename
            if output_file:
                file_path = output_file
            else:
                file_name = f"MARMON BVC {date_range}.xlsx"
                file_path = os.path.join(self.output_folder, file_name)
            
            # Create workbook and sheets
            wb = Workbook()
            
            # Rename the first default sheet
            ws = wb.active
            ws.title = self.sheet_names[0]
            
            # Add remaining sheets
            for name in self.sheet_names[1:]:
                wb.create_sheet(title=name)
            
            # Save the workbook
            wb.save(file_path)
            wb.close()
            
            self.logger.info(f"Marmon BVC template generated: {file_path}")
            return file_path
            
        except Exception as e:
            self.logger.error(f"Failed to generate Marmon BVC template: {e}")
            raise RuntimeError(f"Template generation error: {str(e)}")
    
    def validate_date_format(self, date_range: str) -> bool:
        """
        Validate the date range format
        
        Args:
            date_range: Date range string to validate
            
        Returns:
            bool: True if format appears valid
        """
        try:
            # Basic validation - check if it contains expected patterns
            if not date_range:
                return False
                
            # Should contain a dash or similar separator
            separators = ['-', 'to', 'TO', '–', '—']
            has_separator = any(sep in date_range for sep in separators)
            
            # Should have some digits (for dates)
            has_digits = any(c.isdigit() for c in date_range)
            
            return has_separator and has_digits
            
        except Exception:
            return False

def main():
    """Command line interface for template generation"""
    generator = MarmonBVCTemplateGenerator()
    
    # Ask user for date range
    date_range = input("Enter the date range (e.g., 08.04.25 - 08.08.25): ").strip()
    
    if not date_range:
        print("❌ No date range provided. Exiting.")
        return
    
    try:
        file_path = generator.generate_template(date_range)
        print(f"✅ File created: {file_path}")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
