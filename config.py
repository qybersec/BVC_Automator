"""
Configuration management for TMS Processor
"""
import json
from pathlib import Path
from typing import Dict, Any, List
import logging

class TMSConfig:
    """Configuration management for TMS Processor"""
    
    # Default configuration
    DEFAULT_CONFIG = {
        'data_structure': {
            'default_header_row': 8,
            'default_data_start_row': 11,
            'min_data_columns': 5,
            'expected_columns': 21,
            'header_indicators': ['Load No.', 'Carrier', 'Service Type', 'Ship Date'],
            'required_columns': ['Load No.', 'Ship Date', 'Selected Carrier', 'Least Cost Carrier']
        },
        'business_rules': {
            'same_carrier_savings': 0.0,
            'negative_savings_action': 'copy_selected',
            'min_non_empty_values': 5,
            'empty_load_no_action': 'remove'
        },
        'formatting': {
            'currency_format': '"$"#,##0.00',
            'date_format': '%m/%d/%y',
            'decimal_places': 2,
            'auto_fit_columns': True,
            'max_column_width': 25,
            'min_column_width': 6
        },
        'colors': {
            'selected_carrier': {
                'header': '87CEEB',
                'data_even': 'E6F3FF',
                'data_odd': 'F0F8FF'
            },
            'least_cost_carrier': {
                'header': 'FFB366',
                'data_even': 'FFF2E6',
                'data_odd': 'FFF8F0'
            },
            'potential_savings': {
                'positive': 'C6EFCE',
                'negative': 'FFCDD2',
                'zero': 'F8F9FA'
            },
            'headers': {
                'main': '1F4E79',
                'section': '4A90E2'
            }
        },
        'performance': {
            'chunk_size': 1000,
            'memory_limit_mb': 512,
            'enable_threading': True,
            'max_threads': 4
        },
        'validation': {
            'max_file_size_mb': 100,
            'allowed_extensions': ['.xlsx', '.xls'],
            'min_rows': 10,
            'max_rows': 50000
        }
    }
    
    def __init__(self, config_file: str = None):
        self.config_file = config_file or 'tms_config.json'
        self.config = self.load_config()
        
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from file or create default"""
        config_path = Path(self.config_file)
        
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    config = json.load(f)
                # Merge with defaults to ensure all keys exist
                return self._merge_configs(self.DEFAULT_CONFIG, config)
            except (json.JSONDecodeError, IOError) as e:
                logging.warning(f"Failed to load config from {config_path}: {e}. Using defaults.")
                return self.DEFAULT_CONFIG.copy()
        else:
            # Create default config file
            self.save_config(self.DEFAULT_CONFIG)
            return self.DEFAULT_CONFIG.copy()
    
    def save_config(self, config: Dict[str, Any] = None) -> None:
        """Save configuration to file"""
        config = config or self.config
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except IOError as e:
            logging.error(f"Failed to save config to {self.config_file}: {e}")
    
    def get(self, key_path: str, default=None):
        """Get configuration value using dot notation (e.g., 'data_structure.header_row')"""
        keys = key_path.split('.')
        value = self.config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any) -> None:
        """Set configuration value using dot notation"""
        keys = key_path.split('.')
        config = self.config
        
        # Navigate to the parent of the target key
        for key in keys[:-1]:
            if key not in config:
                config[key] = {}
            config = config[key]
        
        # Set the final value
        config[keys[-1]] = value
    
    def _merge_configs(self, default: Dict, custom: Dict) -> Dict:
        """Recursively merge custom config with defaults"""
        result = default.copy()
        
        for key, value in custom.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_configs(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def validate_config(self) -> List[str]:
        """Validate configuration and return list of issues"""
        issues = []
        
        # Check required sections
        required_sections = ['data_structure', 'business_rules', 'formatting', 'colors']
        for section in required_sections:
            if section not in self.config:
                issues.append(f"Missing required section: {section}")
        
        # Validate numeric ranges
        numeric_checks = [
            ('data_structure.min_data_columns', 1, 100),
            ('data_structure.expected_columns', 1, 100),
            ('performance.chunk_size', 100, 10000),
            ('performance.max_threads', 1, 16),
            ('validation.max_file_size_mb', 1, 1000),
            ('validation.min_rows', 1, 1000000),
            ('validation.max_rows', 10, 1000000)
        ]
        
        for key_path, min_val, max_val in numeric_checks:
            value = self.get(key_path)
            if value is not None and not (min_val <= value <= max_val):
                issues.append(f"{key_path} must be between {min_val} and {max_val}, got {value}")
        
        return issues
    
    def reset_to_defaults(self) -> None:
        """Reset configuration to defaults"""
        self.config = self.DEFAULT_CONFIG.copy()
        self.save_config()

# Global configuration instance
tms_config = TMSConfig()