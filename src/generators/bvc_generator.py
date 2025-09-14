"""
BVC Template Generator - Following SOLID principles and industry best practices
"""
from .base_generator import BaseTemplateGenerator
from typing import List
import re


class BVCTemplateGenerator(BaseTemplateGenerator):
    """
    BVC-specific template generator
    
    Inherits from BaseTemplateGenerator and implements BVC-specific logic.
    Follows Single Responsibility Principle - only handles BVC templates.
    """
    
    # BVC-specific sheet configuration
    BVC_SHEETS = [
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
    
    def get_template_name(self) -> str:
        """Return the name of this template type"""
        return "MARMON BVC"
    
    def get_default_sheets(self) -> List[str]:
        """Return the default sheet names for BVC templates"""
        return self.BVC_SHEETS.copy()
    
    def validate_input(self, date_range: str) -> bool:
        """
        Validate BVC date range format
        
        Expected formats:
        - MM.DD.YY - MM.DD.YY
        - MM/DD/YY - MM/DD/YY
        - Similar variations with 'to', 'TO', etc.
        """
        if not date_range or not date_range.strip():
            return False
        
        date_range = date_range.strip()
        
        # Check for separator patterns
        separators = ['-', 'to', 'TO', '–', '—', 'through', 'THROUGH']
        has_separator = any(sep in date_range for sep in separators)
        
        # Check for date-like patterns (digits and date separators)
        has_dates = bool(re.search(r'\d+[./]\d+[./]\d+', date_range))
        
        return has_separator and has_dates
    
    def generate_filename(self, date_range: str) -> str:
        """Generate BVC-specific filename"""
        clean_date = date_range.strip()
        return f"MARMON BVC {clean_date}.xlsx"
    
    def get_supported_date_formats(self) -> List[str]:
        """Return examples of supported date formats"""
        return [
            "08.04.25 - 08.08.25",
            "08/04/25 - 08/08/25", 
            "08.04.25 to 08.08.25",
            "Aug 4, 2025 - Aug 8, 2025"
        ]
    
    def get_validation_help_text(self) -> str:
        """Return help text for date format validation"""
        formats = self.get_supported_date_formats()
        examples = "\n".join([f"  • {fmt}" for fmt in formats])
        return f"Please enter a date range in one of these formats:\n{examples}"


# Factory function for easy instantiation
def create_bvc_generator(output_folder: str = "output_reports") -> BVCTemplateGenerator:
    """
    Factory function to create BVC template generator
    
    This follows the Factory pattern for consistent object creation
    """
    return BVCTemplateGenerator(output_folder)