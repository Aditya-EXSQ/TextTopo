import logging
import os
import tempfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import List, Optional, Tuple

from DOCXToText.config import ConversionConfig
from DOCXToText.Converters.LibreOffice import convert_docx_via_libreoffice
from DOCXToText.Extractors.DOCXExtractor import extract_content_with_python_docx

LOGGER = logging.getLogger(__name__)


def extract_docx_to_text_file(input_docx_path: str, output_txt_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Extract DOCX document content using the clean two-step process:
	1. Try LibreOffice conversion for document normalization (optional)
	2. Extract text using python-docx
	"""
	cfg = cfg or ConversionConfig()
	
	LOGGER.info("=== DOCUMENT CONTENT EXTRACTION WITH LIBREOFFICE CONVERSION ===")
	
	# Check if input file exists
	if not os.path.isfile(input_docx_path):
		LOGGER.error("Input file not found: %s", input_docx_path)
		return False
	
	# Prepare temp path for converted docx
	temp_converted_path = None
	conversion_success = False
	converted_file = input_docx_path
	
	try:
		# Step 1: Convert DOCX via LibreOffice
		LOGGER.info("Step 1: Converting document via LibreOffice...")
		try:
			temp_converted_path = tempfile.mktemp(suffix=".docx")
			conversion_success = convert_docx_via_libreoffice(input_docx_path, temp_converted_path, cfg=cfg)
			if conversion_success:
				LOGGER.info("✅ Conversion successful!")
				converted_file = temp_converted_path
			else:
				LOGGER.info("❌ Conversion failed, trying to extract from original file...")
				converted_file = input_docx_path
		except Exception as e:
			LOGGER.error("❌ Conversion error: %s", e)
			LOGGER.info("Falling back to original file...")
			converted_file = input_docx_path
		
		# Step 2: Extract content using python-docx
		LOGGER.info("Step 2: Extracting content...")
		try:
			content = extract_content_with_python_docx(converted_file)
			if content.strip():
				LOGGER.info("✅ Content extraction successful!")
			else:
				LOGGER.warning("⚠️ Content extraction completed but no text was found")
				LOGGER.info("The document might be empty or contain only images/formatted content")
			
		except Exception as e:
			LOGGER.error("❌ Error extracting content: %s", e)
			LOGGER.error("This might be due to:")
			LOGGER.error("  - Corrupted document file")
			LOGGER.error("  - Unsupported document format")
			LOGGER.error("  - Missing python-docx library")
			raise
		
		# Save extracted content
		os.makedirs(os.path.dirname(output_txt_path) or ".", exist_ok=True)
		with open(output_txt_path, "w", encoding="utf-8") as fh:
			fh.write(content or "")
		
		LOGGER.info("📁 Extracted text saved to: %s", output_txt_path)
		
		# Note about LibreOffice
		if not conversion_success:
			LOGGER.info("💡 Tip: Install LibreOffice to enable document conversion for better text extraction")
		
		return True
		
	except Exception as exc:
		LOGGER.exception("Failed to extract '%s' -> '%s': %s", input_docx_path, output_txt_path, exc)
		try:
			os.makedirs(os.path.dirname(output_txt_path) or ".", exist_ok=True)
			with open(output_txt_path, "w", encoding="utf-8") as fh:
				fh.write("")
		except Exception:
			pass
		return False
	
	finally:
		# Clean up temporary converted file
		if conversion_success and temp_converted_path and os.path.exists(temp_converted_path):
			try:
				os.remove(temp_converted_path)
				LOGGER.debug("Deleted temp converted file: %s", temp_converted_path)
			except OSError:
				pass


def _process_single_file(args: Tuple[str, str, ConversionConfig]) -> Tuple[str, bool]:
	input_path, output_txt, cfg = args
	ok = extract_docx_to_text_file(input_path, output_txt, cfg=cfg)
	return (output_txt, ok)


def process_docx_folder(input_folder: str, output_folder: Optional[str] = None, cfg: Optional[ConversionConfig] = None, workers: int = 1) -> int:
	"""Process all .docx files in a folder with optional multiprocessing."""
	if not os.path.isdir(input_folder):
		raise ValueError(f"Input folder not found: {input_folder}")

	cfg = cfg or ConversionConfig()
	output_dir = output_folder or input_folder
	os.makedirs(output_dir, exist_ok=True)

	tasks: List[Tuple[str, str, ConversionConfig]] = []
	for name in os.listdir(input_folder):
		if not name.lower().endswith(".docx"):
			continue
		if name.startswith("~$"):
			continue
		input_path = os.path.join(input_folder, name)
		base_name = os.path.splitext(name)[0]
		output_txt = os.path.join(output_dir, f"{base_name}.txt")
		tasks.append((input_path, output_txt, cfg))

	if not tasks:
		LOGGER.info("No DOCX files found in '%s'.", input_folder)
		return 0

	processed = 0
	workers = max(1, int(workers or 1))
	if workers == 1:
		for t in tasks:
			_, ok = _process_single_file(t)
			if ok:
				processed += 1
	else:
		LOGGER.info("Processing %d files with %d workers...", len(tasks), workers)
		with ProcessPoolExecutor(max_workers=workers) as executor:
			future_to_task = {executor.submit(_process_single_file, t): t for t in tasks}
			for future in as_completed(future_to_task):
				try:
					_, ok = future.result()
					if ok:
						processed += 1
				except Exception as exc:
					task = future_to_task[future]
					LOGGER.exception("Error processing file '%s': %s", task[0], exc)

	LOGGER.info("Processed %d/%d DOCX file(s) from '%s' into '%s'.", processed, len(tasks), input_folder, output_dir)
	return processed


