"""
Logging configuration for TextTopo package.
Provides consistent logging across all modules.
"""

import logging
import sys
from typing import Optional

from .config import ConversionConfig


def setup_logging(config: Optional[ConversionConfig] = None, 
                 log_file: Optional[str] = None) -> logging.Logger:
    """
    Set up logging configuration for TextTopo.
    
    Args:
        config: Configuration object with log level settings
        log_file: Optional file to write logs to (in addition to console)
    
    Returns:
        Configured logger instance
    """
    if config is None:
        from .config import default_config
        config = default_config
    
    # Create logger
    logger = logging.getLogger("texttopo")
    logger.setLevel(getattr(logging, config.log_level, logging.INFO))
    
    # Clear any existing handlers
    logger.handlers.clear()
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, config.log_level, logging.INFO))
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler (if specified)
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(getattr(logging, config.log_level, logging.INFO))
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str) -> logging.Logger:
    """Get a logger instance for a specific module."""
    return logging.getLogger(f"texttopo.{name}")


# Default logger instance
logger = get_logger("main")
