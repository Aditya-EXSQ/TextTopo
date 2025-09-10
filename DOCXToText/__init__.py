"""
TextTopo: A Python package for comprehensive text extraction from DOCX files.

This package provides robust text extraction from DOCX files by:
1. Using advanced XML parsing to extract complete document content
2. Extracting headers, main content, footers, tables, and placeholders
3. Supporting both single file and batch folder processing with async processing
4. Providing comprehensive error handling for corrupted DOCX files
5. Working without external dependencies like LibreOffice
"""

__version__ = "1.0.0"
__author__ = "TextTopo"
__email__ = "support@texttopo.com"

# Core functionality
from .config import ConversionConfig
from .logging_setup import get_logger, setup_logging
from .Pipeline.Batch import process_file, process_files_in_parallel
from .Extractors.DOCXExtractor import extract_content

__all__ = [
    "ConversionConfig",
    "get_logger",
    "setup_logging", 
    "process_file", 
    "process_files_in_parallel",
    "extract_content"
]
