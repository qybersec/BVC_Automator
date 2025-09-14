"""
UI Styles and Constants
Centralized styling configuration for consistent theming across the application.
"""

# Color Palette
COLORS = {
    # Primary colors
    'PRIMARY_BLUE': '#4299e1',
    'PRIMARY_BLUE_HOVER': '#3182ce',
    'PRIMARY_BLUE_PRESSED': '#2c5aa0',
    
    # Background colors
    'BACKGROUND_WHITE': '#ffffff',
    'BACKGROUND_GRAY': '#f8f9fa',
    'BACKGROUND_LIGHT': '#f7fafc',
    'BACKGROUND_BORDER': '#e2e8f0',
    
    # Text colors
    'TEXT_PRIMARY': '#2d3748',
    'TEXT_SECONDARY': '#4a5568',
    'TEXT_MUTED': '#718096',
    'TEXT_DISABLED': '#a0aec0',
    
    # Status colors
    'SUCCESS_GREEN': '#38a169',
    'SUCCESS_LIGHT': '#e6fffa',
    'WARNING_YELLOW': '#ffd700',
    'ERROR_RED': '#e53e3e',
    
    # Navigation colors
    'NAV_INACTIVE': '#ffffff',
    'NAV_ACTIVE': '#4299e1',
}

# Typography
FONTS = {
    'HEADER': ('Segoe UI', 11, 'bold'),
    'BODY': ('Segoe UI', 10),
    'BODY_BOLD': ('Segoe UI', 10, 'bold'),
    'SMALL': ('Segoe UI', 9),
    'NAV_PRIMARY': ('Segoe UI', 11, 'bold'),
    'NAV_SECONDARY': ('Segoe UI', 10),
    'NAV_SECONDARY_BOLD': ('Segoe UI', 10, 'bold'),
}

# Spacing
SPACING = {
    'SMALL': 4,
    'MEDIUM': 8,
    'LARGE': 15,
    'XLARGE': 20,
}

# Component Dimensions
DIMENSIONS = {
    'FILE_DISPLAY_HEIGHT': 6,
    'FILE_DISPLAY_WIDTH': 50,
    'CALENDAR_HEADER_HEIGHT': 35,
    'PROCESS_BUTTON_WIDTH': 200,
}

# Button Styles Configuration
BUTTON_STYLES = {
    'nav_primary_active': {
        'bg': COLORS['PRIMARY_BLUE'],
        'fg': COLORS['BACKGROUND_WHITE'],
        'font': FONTS['NAV_PRIMARY']
    },
    'nav_primary_inactive': {
        'bg': COLORS['BACKGROUND_LIGHT'],
        'fg': COLORS['TEXT_SECONDARY'],
        'font': FONTS['NAV_PRIMARY']
    },
    'nav_secondary_active': {
        'bg': COLORS['PRIMARY_BLUE'],
        'fg': COLORS['BACKGROUND_WHITE'],
        'font': FONTS['NAV_SECONDARY_BOLD']
    },
    'nav_secondary_inactive': {
        'bg': COLORS['NAV_INACTIVE'],
        'fg': COLORS['TEXT_SECONDARY'],
        'font': FONTS['NAV_SECONDARY']
    }
}

# Common Widget Configurations
WIDGET_CONFIGS = {
    'frame_default': {
        'bg': COLORS['BACKGROUND_GRAY']
    },
    'frame_white': {
        'bg': COLORS['BACKGROUND_WHITE'],
        'relief': 'flat',
        'bd': 0
    },
    'entry_default': {
        'font': FONTS['BODY'],
        'bg': COLORS['BACKGROUND_WHITE'],
        'fg': COLORS['TEXT_PRIMARY']
    }
}