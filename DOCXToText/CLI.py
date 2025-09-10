"""
Command Line Interface for TextTopo.
Provides CLI functionality for single file and batch folder processing.
"""

import argparse
import asyncio
import os
import sys
from pathlib import Path
from typing import List, Optional

from .config import ConversionConfig
from .logging_setup import setup_logging, get_logger
from .Pipeline.Batch import process_file, process_files_in_parallel, find_docx_files

logger = get_logger("cli")


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser."""
    parser = argparse.ArgumentParser(
        description="TextTopo: Extract text from DOCX files via LibreOffice conversion",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process single file
  python -m DOCXToText.CLI --input document.docx --output ./extracted/

  # Process entire folder
  python -m DOCXToText.CLI --input ./documents/ --output ./extracted/

  # Process with custom settings
  python -m DOCXToText.CLI --input ./docs/ --output ./text/ --concurrency 8 --log-level DEBUG

  # Process single file and print to stdout
  python -m DOCXToText.CLI --input document.docx --stdout
        """
    )
    
    # Input/Output arguments
    parser.add_argument(
        "--input", "-i",
        type=str,
        required=True,
        help="Input DOCX file or directory containing DOCX files"
    )
    
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Output directory for extracted text files (required unless --stdout is used)"
    )
    
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print extracted text to stdout instead of saving to files (single file only)"
    )
    
    # Processing options
    parser.add_argument(
        "--concurrency", "-c",
        type=int,
        default=4,
        help="Maximum number of concurrent conversions (default: 4)"
    )
    
    
    parser.add_argument(
        "--no-recursive", "-nr",
        action="store_true",
        help="Don't search subdirectories when input is a folder"
    )
    
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output files"
    )
    
    
    # Logging options
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Set the logging level (default: INFO)"
    )
    
    parser.add_argument(
        "--log-file",
        type=str,
        help="Write logs to file in addition to console"
    )
    
    # Temporary directory
    parser.add_argument(
        "--temp-dir",
        type=str,
        default="texttopo_temp",
        help="Name of temporary directory in current working directory (default: texttopo_temp)"
    )
    
    return parser


def validate_arguments(args: argparse.Namespace) -> None:
    """Validate command line arguments."""
    # Check input exists
    if not os.path.exists(args.input):
        logger.error(f"Input path does not exist: {args.input}")
        sys.exit(1)
    
    # Check output requirements
    if not args.stdout and not args.output:
        logger.error("Either --output or --stdout must be specified")
        sys.exit(1)
    
    if args.stdout and args.output:
        logger.error("Cannot use both --output and --stdout")
        sys.exit(1)
    
    # Check stdout is only for single files
    if args.stdout and os.path.isdir(args.input):
        logger.error("--stdout can only be used with single file input")
        sys.exit(1)
    
    # Validate concurrency
    if args.concurrency < 1:
        logger.error("Concurrency must be at least 1")
        sys.exit(1)
    


async def main_async(args: argparse.Namespace) -> None:
    """Main async function for processing files."""
    # Create configuration from arguments
    config = ConversionConfig(
        concurrency_limit=args.concurrency,
        temp_dir_name=args.temp_dir,
        overwrite_existing=args.overwrite,
        log_level=args.log_level
    )
    
    # Determine input files
    if os.path.isfile(args.input):
        if not args.input.lower().endswith('.docx'):
            logger.error("Input file must be a .docx file")
            sys.exit(1)
        files = [args.input]
    else:
        # Directory input
        recursive = not args.no_recursive
        files = find_docx_files(args.input, recursive=recursive)
        
        if not files:
            logger.error(f"No DOCX files found in {args.input}")
            sys.exit(1)
        
        logger.info(f"Found {len(files)} DOCX files to process")
    
    try:
        if len(files) == 1 and args.stdout:
            # Single file to stdout
            logger.info(f"Processing file: {files[0]}")
            extracted_text = await process_file(
                input_path=files[0],
                output_dir=None,
                config=config
            )
            print(extracted_text)
        else:
            # Process files to output directory
            if args.output:
                os.makedirs(args.output, exist_ok=True)
            
            if len(files) == 1:
                # Single file
                logger.info(f"Processing file: {files[0]}")
                await process_file(
                    input_path=files[0],
                    output_dir=args.output,
                    config=config
                )
            else:
                # Multiple files
                logger.info(f"Processing {len(files)} files in parallel")
                await process_files_in_parallel(
                    files=files,
                    output_dir=args.output,
                    config=config
                )
            
        logger.info("Processing completed successfully")
        
    except KeyboardInterrupt:
        logger.warning("Processing interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Processing failed: {e}")
        sys.exit(1)
    finally:
        # Clean up temporary directory
        temp_dir = config.get_temp_dir_path()
        if os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
                logger.debug("Cleaned up temporary directory")
            except Exception as e:
                logger.warning(f"Failed to cleanup temporary directory: {e}")


def main() -> None:
    """Main CLI entry point."""
    # Parse arguments
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup logging
    config = ConversionConfig(log_level=args.log_level)
    setup_logging(config, args.log_file)
    
    # Validate arguments
    validate_arguments(args)
    
    # Show banner
    logger.info("TextTopo - DOCX Text Extraction Tool")
    logger.info(f"Input: {args.input}")
    if args.output:
        logger.info(f"Output: {args.output}")
    logger.info(f"Concurrency: {args.concurrency}")
    
    # Run async main
    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        logger.warning("Interrupted by user")
        sys.exit(1)


if __name__ == "__main__":
    main()
