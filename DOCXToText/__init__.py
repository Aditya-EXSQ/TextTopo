"""
TextTopo: A Python package for extracting text from DOCX files via LibreOffice conversion.

This package provides robust text extraction from DOCX files by:
1. Converting DOCX to DOC and back to DOCX using LibreOffice for normalization
2. Extracting text content including tables and placeholders using python-docx
3. Supporting both single file and batch folder processing
4. Providing async processing for better performance
"""

__version__ = "1.0.0"
__author__ = "TextTopo"
__email__ = ""

from .config import ConversionConfig
from .Pipeline.Batch import process_file, process_files_in_parallel

__all__ = [
    "ConversionConfig",
    "process_file", 
    "process_files_in_parallel"
]
