"""
Batch processing pipeline for TextTopo.
Handles single file and parallel multi-file processing.
"""

import os
import asyncio
import shutil
from typing import Dict, List, Optional
from pathlib import Path

from ..config import ConversionConfig
from ..logging_setup import get_logger
from ..Converters.LibreOffice import convert_docx_via_libreoffice
from ..Extractors.DOCXExtractor import extract_content

logger = get_logger("pipeline.batch")


async def process_file(
    input_path: str, 
    output_dir: Optional[str] = None,
    config: Optional[ConversionConfig] = None
) -> str:
    """
    Process a single DOCX file: Convert via LibreOffice, then extract text.
    
    Args:
        input_path: Path to input DOCX file
        output_dir: Directory to save extracted text (if None, returns text only)
        config: Configuration object
        
    Returns:
        Extracted text content
        
    Raises:
        Exception: If processing fails
    """
    if config is None:
        from ..config import default_config
        config = default_config
    
    logger.info(f"Processing file: {input_path}")
    
    # Validate input
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
    
    if not input_path.lower().endswith('.docx'):
        raise ValueError(f"File must be a .docx file: {input_path}")
    
    # Create temporary working directory in CWD
    temp_dir = config.get_temp_dir_path()
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # Generate paths
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        converted_docx_path = os.path.join(temp_dir, f"{base_name}_converted.docx")
        
        # Step 1: Try LibreOffice conversion (if enabled)
        if config.enable_libreoffice:
            logger.debug("Attempting LibreOffice conversion...")
            conversion_success = await convert_docx_via_libreoffice(
                input_path=input_path,
                output_path=converted_docx_path,
                config=config
            )
            
            # Determine which file to extract from
            if conversion_success and os.path.exists(converted_docx_path):
                extraction_path = converted_docx_path
                logger.info("Using converted file for extraction")
            else:
                extraction_path = input_path
                logger.warning("Conversion failed, extracting from original file")
        else:
            logger.debug("LibreOffice conversion disabled, extracting from original file")
            extraction_path = input_path
        
        # Step 2: Extract text content
        logger.debug("Extracting text content...")
        extracted_text = extract_content(extraction_path)
        
        if not extracted_text.strip():
            logger.warning(f"No text content extracted from {input_path}")
        
        # Step 3: Save to output file if output_dir specified
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, f"{base_name}{config.output_extension}")
            
            # Check if output file exists and handle overwrite settings
            if os.path.exists(output_file) and not config.overwrite_existing:
                counter = 1
                while os.path.exists(output_file):
                    output_file = os.path.join(
                        output_dir, 
                        f"{base_name}_{counter}{config.output_extension}"
                    )
                    counter += 1
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(extracted_text)
            
            logger.info(f"Text saved to: {output_file}")
        
        return extracted_text
        
    finally:
        # Clean up converted file
        if os.path.exists(converted_docx_path):
            try:
                os.remove(converted_docx_path)
                logger.debug("Cleaned up converted file")
            except Exception as e:
                logger.warning(f"Failed to cleanup converted file: {e}")


async def process_files_in_parallel(
    files: List[str], 
    output_dir: Optional[str] = None,
    config: Optional[ConversionConfig] = None
) -> Dict[str, str]:
    """
    Process multiple DOCX files in parallel with concurrency control.
    
    Args:
        files: List of paths to DOCX files
        output_dir: Directory to save extracted text files
        config: Configuration object
        
    Returns:
        Dictionary mapping input file paths to extracted text content
        
    Raises:
        Exception: If processing fails for all files
    """
    if config is None:
        from ..config import default_config
        config = default_config
    
    if not files:
        logger.warning("No files provided for processing")
        return {}
    
    logger.info(f"Processing {len(files)} files with concurrency limit: {config.concurrency_limit}")
    
    # Create semaphore to limit concurrency
    semaphore = asyncio.Semaphore(config.concurrency_limit)
    results = {}
    errors = {}
    
    async def process_with_semaphore(file_path: str) -> None:
        """Process a single file with semaphore control."""
        async with semaphore:
            try:
                logger.debug(f"Starting processing: {file_path}")
                extracted_text = await process_file(
                    input_path=file_path,
                    output_dir=output_dir,
                    config=config
                )
                results[file_path] = extracted_text
                logger.info(f"Completed processing: {file_path}")
            except Exception as e:
                error_msg = f"Failed to process {file_path}: {e}"
                logger.error(error_msg)
                errors[file_path] = str(e)
    
    # Create tasks for all files
    tasks = [process_with_semaphore(file_path) for file_path in files]
    
    # Execute all tasks
    await asyncio.gather(*tasks, return_exceptions=True)
    
    # Log summary
    success_count = len(results)
    error_count = len(errors)
    
    logger.info(f"Processing completed: {success_count} successful, {error_count} failed")
    
    if errors:
        logger.error("Errors encountered:")
        for file_path, error in errors.items():
            logger.error(f"  {file_path}: {error}")
    
    if not results:
        raise Exception("All files failed to process")
    
    return results


def find_docx_files(directory: str, recursive: bool = True) -> List[str]:
    """
    Find all DOCX files in a directory.
    
    Args:
        directory: Directory to search
        recursive: Whether to search subdirectories
        
    Returns:
        List of DOCX file paths
    """
    docx_files = []
    
    if not os.path.exists(directory):
        logger.error(f"Directory not found: {directory}")
        return docx_files
    
    if recursive:
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.docx') and not file.startswith('~$'):
                    docx_files.append(os.path.join(root, file))
    else:
        for file in os.listdir(directory):
            if file.lower().endswith('.docx') and not file.startswith('~$'):
                docx_files.append(os.path.join(directory, file))
    
    logger.info(f"Found {len(docx_files)} DOCX files in {directory}")
    return sorted(docx_files)
