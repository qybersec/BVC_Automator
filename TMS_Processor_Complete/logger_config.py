"""
Comprehensive logging system for TMS Processor
"""
import logging
import logging.handlers
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional
import traceback

class TMSLogger:
    """Enhanced logging system with multiple handlers and formatters"""
    
    def __init__(self, name: str = "TMSProcessor", log_dir: str = "logs"):
        self.name = name
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Create logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        
        # Prevent duplicate handlers
        if not self.logger.handlers:
            self._setup_handlers()
    
    def _setup_handlers(self):
        """Setup different logging handlers"""
        
        # Console handler with color support
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_formatter = ColoredFormatter(
            '%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(console_formatter)
        
        # File handler for all logs
        log_file = self.log_dir / f"{self.name}_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = logging.handlers.RotatingFileHandler(
            log_file, maxBytes=10*1024*1024, backupCount=5
        )
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(name)s | %(funcName)s:%(lineno)d | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(file_formatter)
        
        # Error handler for critical issues
        error_file = self.log_dir / f"{self.name}_errors_{datetime.now().strftime('%Y%m%d')}.log"
        error_handler = logging.handlers.RotatingFileHandler(
            error_file, maxBytes=5*1024*1024, backupCount=3
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(file_formatter)
        
        # Add handlers
        self.logger.addHandler(console_handler)
        self.logger.addHandler(file_handler)
        self.logger.addHandler(error_handler)
    
    def debug(self, message: str, **kwargs):
        """Log debug message with optional context"""
        self.logger.debug(self._format_message(message, kwargs))
    
    def info(self, message: str, **kwargs):
        """Log info message with optional context"""
        self.logger.info(self._format_message(message, kwargs))
    
    def warning(self, message: str, **kwargs):
        """Log warning message with optional context"""
        self.logger.warning(self._format_message(message, kwargs))
    
    def error(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """Log error message with optional exception and context"""
        msg = self._format_message(message, kwargs)
        if exception:
            msg += f" | Exception: {str(exception)}"
            self.logger.error(msg, exc_info=True)
        else:
            self.logger.error(msg)
    
    def critical(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """Log critical message with optional exception and context"""
        msg = self._format_message(message, kwargs)
        if exception:
            msg += f" | Exception: {str(exception)}"
            self.logger.critical(msg, exc_info=True)
        else:
            self.logger.critical(msg)
    
    def log_processing_step(self, step: str, details: dict = None):
        """Log processing steps with consistent formatting"""
        details_str = ""
        if details:
            details_str = " | ".join([f"{k}: {v}" for k, v in details.items()])
        
        message = f"PROCESSING: {step}"
        if details_str:
            message += f" | {details_str}"
        
        self.info(message)
    
    def log_data_stats(self, stats: dict, prefix: str = "DATA_STATS"):
        """Log data statistics in a structured format"""
        stats_str = " | ".join([f"{k}: {v}" for k, v in stats.items()])
        self.info(f"{prefix}: {stats_str}")
    
    def log_performance(self, operation: str, duration: float, records: int = None):
        """Log performance metrics"""
        msg = f"PERFORMANCE: {operation} took {duration:.2f}s"
        if records:
            rate = records / duration if duration > 0 else 0
            msg += f" | {records} records | {rate:.0f} records/sec"
        self.info(msg)
    
    def log_file_operation(self, operation: str, file_path: str, success: bool = True, size: int = None):
        """Log file operations"""
        status = "SUCCESS" if success else "FAILED"
        msg = f"FILE_{operation.upper()}: {status} | {file_path}"
        if size:
            msg += f" | Size: {self._format_bytes(size)}"
        
        if success:
            self.info(msg)
        else:
            self.error(msg)
    
    def log_validation_result(self, item: str, passed: bool, message: str = ""):
        """Log validation results"""
        status = "PASS" if passed else "FAIL"
        msg = f"VALIDATION: {item} | {status}"
        if message:
            msg += f" | {message}"
        
        if passed:
            self.info(msg)
        else:
            self.warning(msg)
    
    def _format_message(self, message: str, context: dict) -> str:
        """Format message with context"""
        if not context:
            return message
        
        context_str = " | ".join([f"{k}: {v}" for k, v in context.items()])
        return f"{message} | {context_str}"
    
    def _format_bytes(self, size: int) -> str:
        """Format byte size to human readable"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f}{unit}"
            size /= 1024
        return f"{size:.1f}TB"


class ColoredFormatter(logging.Formatter):
    """Colored console formatter"""
    
    COLORS = {
        'DEBUG': '\033[36m',     # Cyan
        'INFO': '\033[32m',      # Green
        'WARNING': '\033[33m',   # Yellow
        'ERROR': '\033[31m',     # Red
        'CRITICAL': '\033[35m',  # Magenta
        'RESET': '\033[0m'       # Reset
    }
    
    def format(self, record):
        # Add color to level name
        if record.levelname in self.COLORS:
            record.levelname = (
                f"{self.COLORS[record.levelname]}{record.levelname}"
                f"{self.COLORS['RESET']}"
            )
        return super().format(record)


class ProgressLogger:
    """Progress logging with percentage tracking"""
    
    def __init__(self, logger: TMSLogger, total: int, operation: str):
        self.logger = logger
        self.total = total
        self.operation = operation
        self.current = 0
        self.last_logged_percent = -1
    
    def update(self, increment: int = 1):
        """Update progress and log if percentage changed"""
        self.current += increment
        percent = int((self.current / self.total) * 100) if self.total > 0 else 0
        
        # Log every 10% or at completion
        if percent != self.last_logged_percent and (percent % 10 == 0 or percent == 100):
            self.logger.info(f"{self.operation}: {percent}% ({self.current}/{self.total})")
            self.last_logged_percent = percent
    
    def complete(self):
        """Mark operation as complete"""
        self.logger.info(f"{self.operation}: COMPLETED ({self.total}/{self.total})")


def setup_exception_logging():
    """Setup global exception logging"""
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        logger = TMSLogger("UNCAUGHT")
        logger.critical(
            "Uncaught exception",
            exception=exc_value
        )
    
    sys.excepthook = handle_exception


# Global logger instances
main_logger = TMSLogger("TMS_MAIN")
data_logger = TMSLogger("TMS_DATA")
gui_logger = TMSLogger("TMS_GUI")

# Setup global exception handling
setup_exception_logging()