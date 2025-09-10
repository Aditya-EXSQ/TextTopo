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
    
    # LibreOffice settings
    soffice_path: Optional[str] = None
    conversion_timeout: int = 60  # seconds
    enable_libreoffice: bool = False  # Disabled by default due to common installation issues
    
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
        """Create configuration from environment variables."""
        return cls(
            soffice_path=os.getenv("SOFFICE_PATH"),
            conversion_timeout=int(os.getenv("CONVERSION_TIMEOUT", "60")),
            enable_libreoffice=os.getenv("ENABLE_LIBREOFFICE", "false").lower() == "true",
            concurrency_limit=int(os.getenv("CONCURRENCY_LIMIT", "4")),
            temp_dir_name=os.getenv("TEMP_DIR_NAME", "texttopo_temp"),
            output_extension=os.getenv("OUTPUT_EXTENSION", ".txt"),
            overwrite_existing=os.getenv("OVERWRITE_EXISTING", "false").lower() == "true",
            log_level=os.getenv("LOG_LEVEL", "INFO").upper()
        )
    
    def get_temp_dir_path(self, base_dir: str = ".") -> str:
        """Get the full path to the temporary directory."""
        return os.path.join(base_dir, self.temp_dir_name)


# Default configuration instance
default_config = ConversionConfig.from_env()
