"""
Pipeline package for TextTopo.
Contains batch processing utilities.
"""

from .Batch import process_file, process_files_in_parallel

__all__ = ["process_file", "process_files_in_parallel"]
