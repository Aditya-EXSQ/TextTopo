"""
Configuration management for TextTopo package.
Handles environment variables and default settings.
"""

import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class ConversionConfig:
    """Configuration class for DOCX to text conversion."""
    
    # Processing settings
    concurrency_limit: int = 4
    temp_dir_name: str = "texttopo_temp"
    
    # Output settings
    output_extension: str = ".txt"
    overwrite_existing: bool = False
    
    # Logging settings
    log_level: str = "INFO"
    
    @classmethod
    def from_env(cls) -> 'ConversionConfig':
        """Create configuration from environment variables with validation."""
        try:
            # Validate and convert integer values
            concurrency = int(os.getenv("CONCURRENCY_LIMIT", "4"))
            if concurrency <= 0:
                raise ValueError("CONCURRENCY_LIMIT must be positive")
            
            # Validate log level
            log_level = os.getenv("LOG_LEVEL", "INFO").upper()
            valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
            if log_level not in valid_levels:
                raise ValueError(f"LOG_LEVEL must be one of: {', '.join(valid_levels)}")
            
            # Validate output extension
            ext = os.getenv("OUTPUT_EXTENSION", ".txt")
            if not ext.startswith('.'):
                ext = f".{ext}"
            
            return cls(
                concurrency_limit=concurrency,
                temp_dir_name=os.getenv("TEMP_DIR_NAME", "texttopo_temp"),
                output_extension=ext,
                overwrite_existing=os.getenv("OVERWRITE_EXISTING", "false").lower() == "true",
                log_level=log_level
            )
        except (ValueError, TypeError) as e:
            raise ValueError(f"Invalid configuration from environment: {e}")
            
    def validate(self) -> None:
        """Validate configuration values."""
        if self.concurrency_limit <= 0:
            raise ValueError("concurrency_limit must be positive")
        if not self.output_extension.startswith('.'):
            raise ValueError("output_extension must start with '.'")
        if not self.temp_dir_name.strip():
            raise ValueError("temp_dir_name cannot be empty")
    
    def get_temp_dir_path(self, base_dir: str = ".") -> str:
        """Get the full path to the temporary directory in the project."""
        # Ensure we're always working from the project directory
        if not os.path.isabs(base_dir):
            # Make relative paths absolute from current working directory
            base_dir = os.path.abspath(base_dir)
        return os.path.join(base_dir, self.temp_dir_name)


# Default configuration instance
default_config = ConversionConfig.from_env()
