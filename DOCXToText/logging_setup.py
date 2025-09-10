"""
Logging configuration for TextTopo package.
Provides consistent logging across all modules.
"""

import logging
import os
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
        
    Raises:
        ValueError: If log level is invalid
    """
    if config is None:
        from .config import default_config
        config = default_config
    
    # Validate log level
    valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
    if config.log_level not in valid_levels:
        raise ValueError(f"Invalid log level: {config.log_level}. Must be one of: {', '.join(valid_levels)}")
    
    # Create logger
    logger = logging.getLogger("texttopo")
    log_level = getattr(logging, config.log_level)
    logger.setLevel(log_level)
    
    # Clear any existing handlers to avoid duplicates
    logger.handlers.clear()
    
    # Prevent propagation to root logger to avoid duplicate messages
    logger.propagate = False
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler (if specified)
    if log_file:
        try:
            # Ensure directory exists
            log_dir = os.path.dirname(log_file)
            if log_dir:
                os.makedirs(log_dir, exist_ok=True)
            
            file_handler = logging.FileHandler(log_file)
            file_handler.setLevel(log_level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except (OSError, IOError) as e:
            logger.warning(f"Could not create log file {log_file}: {e}")
    
    return logger


def get_logger(name: str) -> logging.Logger:
    """Get a logger instance for a specific module."""
    return logging.getLogger(f"texttopo.{name}")


# Default logger instance
logger = get_logger("main")
