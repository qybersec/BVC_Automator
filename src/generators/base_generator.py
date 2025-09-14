"""
Abstract base class for all template generators
Following industry best practices with SOLID principles
"""
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional
import logging
from pathlib import Path


class BaseTemplateGenerator(ABC):
    """
    Abstract base class for template generators
    
    This class defines the contract that all template generators must follow.
    Following the Template Method pattern for consistent behavior across generators.
    """
    
    def __init__(self, output_folder: str = "output_reports"):
        self.output_folder = Path(output_folder)
        self.logger = logging.getLogger(self.__class__.__name__)
        self._setup_output_directory()
    
    def _setup_output_directory(self) -> None:
        """Ensure output directory exists"""
        self.output_folder.mkdir(parents=True, exist_ok=True)
    
    @abstractmethod
    def get_template_name(self) -> str:
        """Return the name of this template type"""
        pass
    
    @abstractmethod
    def get_default_sheets(self) -> list:
        """Return the default sheet names for this template"""
        pass
    
    @abstractmethod
    def validate_input(self, user_input: str) -> bool:
        """Validate user input specific to this template type"""
        pass
    
    @abstractmethod
    def generate_filename(self, user_input: str) -> str:
        """Generate the output filename based on user input"""
        pass
    
    def generate_template(self, user_input: str, output_file: Optional[str] = None) -> str:
        """
        Main template generation method - Template Method pattern
        
        This method defines the algorithm structure while allowing subclasses
        to customize specific steps.
        """
        try:
            # Step 1: Validate input
            if not self.validate_input(user_input):
                raise ValueError(f"Invalid input for {self.get_template_name()}: {user_input}")
            
            # Step 2: Determine output file
            if output_file is None:
                filename = self.generate_filename(user_input)
                output_file = self.output_folder / filename
            else:
                output_file = Path(output_file)
            
            # Step 3: Create the template
            self._create_template_file(output_file, user_input)
            
            # Step 4: Log success
            self.logger.info(f"{self.get_template_name()} template generated: {output_file}")
            
            return str(output_file)
            
        except Exception as e:
            self.logger.error(f"Failed to generate {self.get_template_name()} template: {e}")
            raise RuntimeError(f"Template generation error: {str(e)}")
    
    def _create_template_file(self, output_file: Path, user_input: str) -> None:
        """Create the actual template file - can be overridden by subclasses"""
        from openpyxl import Workbook
        
        wb = Workbook()
        sheets = self.get_default_sheets()
        
        # Rename first sheet
        wb.active.title = sheets[0]
        
        # Add remaining sheets
        for sheet_name in sheets[1:]:
            wb.create_sheet(title=sheet_name)
        
        # Save and close
        wb.save(output_file)
        wb.close()
    
    def get_info(self) -> Dict[str, Any]:
        """Return information about this generator"""
        return {
            'name': self.get_template_name(),
            'output_folder': str(self.output_folder),
            'default_sheets': self.get_default_sheets()
        }