"""
LibreOffice conversion utilities for TextTopo.
Handles DOCX to DOC to DOCX conversion pipeline using LibreOffice.
"""

import os
import asyncio
import subprocess
import tempfile
import shutil
from typing import Optional, List

from ..config import ConversionConfig
from ..logging_setup import get_logger

logger = get_logger("converters.libreoffice")


def discover_libreoffice(config: Optional[ConversionConfig] = None) -> Optional[str]:
    """
    Discover LibreOffice executable path.
    
    Args:
        config: Configuration object with potential soffice_path
        
    Returns:
        Path to LibreOffice executable, or None if not found
    """
    if config and config.soffice_path:
        # Test the configured path
        if test_libreoffice_executable(config.soffice_path):
            return config.soffice_path
    
    # Try common LibreOffice installation paths
    possible_paths: List[str] = [
        r"D:\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",  # If it's in PATH
        "libreoffice"  # Alternative command name
    ]
    
    for path in possible_paths:
        if test_libreoffice_executable(path):
            logger.info(f"Found LibreOffice at: {path}")
            return path
    
    logger.error("LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH environment variable")
    return None


def test_libreoffice_executable(path: str) -> bool:
    """
    Test if a LibreOffice executable is working.
    
    Args:
        path: Path to test
        
    Returns:
        True if executable works, False otherwise
    """
    try:
        result = subprocess.run(
            [path, "--version"], 
            capture_output=True, 
            text=True, 
            timeout=10
        )
        return result.returncode == 0
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return False


async def convert_docx_via_libreoffice(
    input_path: str, 
    output_path: str,
    config: Optional[ConversionConfig] = None
) -> bool:
    """
    Convert DOCX to DOC and back to DOCX using LibreOffice for normalization.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path where converted DOCX should be saved
        config: Configuration object
        
    Returns:
        True if conversion successful, False otherwise
    """
    if config is None:
        from ..config import default_config
        config = default_config
    
    # Discover LibreOffice
    soffice_path = discover_libreoffice(config)
    if not soffice_path:
        return False
    
    # Validate input file
    if not os.path.exists(input_path):
        logger.error(f"Input file does not exist: {input_path}")
        return False
    
    try:
        # Create temporary directory with unique profile for concurrency safety
        temp_dir = tempfile.mkdtemp(prefix="texttopo_conversion_")
        temp_profile_dir = tempfile.mkdtemp(prefix="libreoffice_profile_")
        
        try:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            doc_path = os.path.join(temp_dir, f"{base_name}.doc")
            converted_docx_path = os.path.join(temp_dir, f"{base_name}.docx")
            
            # Step 1: Convert DOCX to DOC
            logger.debug(f"Converting {input_path} to DOC format...")
            success = await _run_libreoffice_conversion(
                soffice_path=soffice_path,
                input_file=input_path,
                output_format="doc",
                output_dir=temp_dir,
                profile_dir=temp_profile_dir,
                timeout=config.conversion_timeout
            )
            
            if not success or not os.path.exists(doc_path):
                logger.error(f"Failed to convert {input_path} to DOC format")
                return False
            
            logger.debug("Successfully converted to DOC format")
            
            # Step 2: Convert DOC back to DOCX
            logger.debug(f"Converting DOC back to DOCX format...")
            success = await _run_libreoffice_conversion(
                soffice_path=soffice_path,
                input_file=doc_path,
                output_format="docx",
                output_dir=temp_dir,
                profile_dir=temp_profile_dir,
                timeout=config.conversion_timeout
            )
            
            if not success or not os.path.exists(converted_docx_path):
                logger.error("Failed to convert DOC back to DOCX format")
                return False
            
            logger.debug("Successfully converted back to DOCX format")
            
            # Step 3: Remove intermediate DOC file
            if os.path.exists(doc_path):
                os.remove(doc_path)
                logger.debug("Removed intermediate DOC file")
            
            # Step 4: Copy converted file to final destination
            shutil.copy2(converted_docx_path, output_path)
            logger.info(f"Successfully converted and saved to: {output_path}")
            
            return True
            
        finally:
            # Clean up temporary directories
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
                shutil.rmtree(temp_profile_dir, ignore_errors=True)
                logger.debug("Cleaned up temporary directories")
            except Exception as cleanup_error:
                logger.warning(f"Failed to cleanup temporary directories: {cleanup_error}")
                
    except Exception as e:
        logger.error(f"Error during conversion: {e}")
        return False


async def _run_libreoffice_conversion(
    soffice_path: str,
    input_file: str,
    output_format: str,
    output_dir: str,
    profile_dir: str,
    timeout: int = 60
) -> bool:
    """
    Run a single LibreOffice conversion command asynchronously.
    
    Args:
        soffice_path: Path to LibreOffice executable
        input_file: Input file path
        output_format: Target format (doc, docx, etc.)
        output_dir: Output directory
        profile_dir: Temporary profile directory for this conversion
        timeout: Conversion timeout in seconds
        
    Returns:
        True if successful, False otherwise
    """
    # Convert path to proper URI format for LibreOffice
    # Based on: https://stackoverflow.com/questions/67208383/
    if os.name == 'nt':
        # Convert Windows path to POSIX format and create proper file:/// URI
        from pathlib import Path
        posix_path = Path(profile_dir).as_posix()
        profile_path = f"file:///{posix_path}"
    else:
        profile_path = f"file://{profile_dir}"
    
    cmd = [
        soffice_path,
        f"-env:UserInstallation={profile_path}",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nolockcheck",
        "--nologo",
        "--norestore",
        "--convert-to", output_format,
        "--outdir", output_dir,
        input_file
    ]
    
    try:
        # Set environment to avoid interactive prompts
        env = os.environ.copy()
        env['DISPLAY'] = ''  # For Linux systems
        
        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            stdin=asyncio.subprocess.PIPE,
            env=env
        )
        
        # Close stdin immediately to prevent hanging on prompts
        if process.stdin:
            process.stdin.close()
        
        stdout, stderr = await asyncio.wait_for(
            process.communicate(),
            timeout=timeout
        )
        
        if process.returncode != 0:
            error_msg = stderr.decode().strip() or "Unknown conversion error"
            logger.error(f"LibreOffice conversion failed: {error_msg}")
            logger.error(f"LibreOffice stdout: {stdout.decode().strip()}")
            return False
            
        logger.debug(f"LibreOffice conversion successful")
        return True
        
    except asyncio.TimeoutError:
        logger.error(f"LibreOffice conversion timed out after {timeout} seconds")
        try:
            process.kill()
            await process.wait()
        except:
            pass
        return False
    except Exception as e:
        logger.error(f"Error running LibreOffice: {e}")
        return False
