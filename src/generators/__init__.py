"""
Template Generators Package

This package contains all template generators following SOLID principles.
Each generator inherits from BaseTemplateGenerator for consistent behavior.
"""
from .base_generator import BaseTemplateGenerator
from .bvc_generator import BVCTemplateGenerator, create_bvc_generator

# Export public API
__all__ = [
    'BaseTemplateGenerator',
    'BVCTemplateGenerator', 
    'create_bvc_generator'
]

# Version info
__version__ = "2.0.0"