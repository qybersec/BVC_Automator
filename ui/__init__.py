"""
UI Components Package

Modular UI components for the TMS Data Processor application.
Following React-like component architecture for better maintainability.
"""

from .styles import COLORS, FONTS, SPACING, DIMENSIONS, BUTTON_STYLES, WIDGET_CONFIGS
from .navigation import NavigationComponent
from .file_input import FileInputComponent
from .date_input import DateInputComponent

__all__ = [
    'COLORS',
    'FONTS', 
    'SPACING',
    'DIMENSIONS',
    'BUTTON_STYLES',
    'WIDGET_CONFIGS',
    'NavigationComponent',
    'FileInputComponent',
    'DateInputComponent'
]

__version__ = "1.0.0"