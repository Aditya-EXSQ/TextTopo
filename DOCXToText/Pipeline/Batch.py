import logging
import os
import tempfile
from typing import List, Optional, Tuple

from DOCXToText.config import ConversionConfig
from DOCXToText.Converters.LibreOffice import convert_docx_via_libreoffice
from DOCXToText.Extractors.DOCXExtractor import extract_content_with_python_docx

LOGGER = logging.getLogger(__name__)


def extract_docx_to_text_file(input_docx_path: str, output_txt_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Extract DOCX document content using the clean two-step process:
	1. Try LibreOffice conversion for document normalization (optional, ALWAYS single-threaded)
	2. Extract text using python-docx
	
	Args:
		input_docx_path: Path to input DOCX file
		output_txt_path: Path where extracted text will be saved
		cfg: Configuration object
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
		LOGGER.debug("Step 1: Converting document via LibreOffice...")
		try:
			# Create temp file in current working directory for easier debugging
			temp_converted_path = tempfile.mktemp(suffix=".docx", dir=".")
			conversion_success = convert_docx_via_libreoffice(input_docx_path, temp_converted_path, cfg=cfg)
			if conversion_success:
				LOGGER.debug("✅ Conversion successful!")
				converted_file = temp_converted_path
			else:
				LOGGER.info("❌ Conversion failed, trying to extract from original file...")
				converted_file = input_docx_path
		except Exception as e:
			error_msg = str(e).lower()
			if "bootstrap.ini" in error_msg or "corrupt" in error_msg:
				LOGGER.error("❌ LibreOffice installation is corrupted (bootstrap.ini issue)")
				LOGGER.error("💡 Please repair LibreOffice: Settings → Apps → LibreOffice → Repair")
				LOGGER.info("📄 Continuing with original file extraction...")
			else:
				LOGGER.warning("❌ Conversion error: %s", e)
			LOGGER.debug("Falling back to original file...")
			converted_file = input_docx_path
		
		# Step 2: Extract content using python-docx
		LOGGER.debug("Step 2: Extracting content from: %s", converted_file)
		try:
			content = extract_content_with_python_docx(converted_file)
			if content and content.strip():
				LOGGER.debug("✅ Content extraction successful! Length: %d characters", len(content))
			else:
				LOGGER.warning("⚠️ Content extraction completed but no text was found in file: %s", os.path.basename(input_docx_path))
				LOGGER.debug("The document might be empty or contain only images/formatted content")
				content = ""
			
		except Exception as e:
			LOGGER.error("❌ Error extracting content from %s: %s", os.path.basename(input_docx_path), e)
			LOGGER.debug("This might be due to:")
			LOGGER.debug("  - Corrupted document file")
			LOGGER.debug("  - Unsupported document format")
			LOGGER.debug("  - Missing python-docx library")
			raise
		
		# Save extracted content
		os.makedirs(os.path.dirname(output_txt_path) or ".", exist_ok=True)
		with open(output_txt_path, "w", encoding="utf-8") as fh:
			fh.write(content or "")
		
		LOGGER.info("✅ Successfully extracted: %s -> %s (%d chars)", 
				   os.path.basename(input_docx_path), 
				   os.path.basename(output_txt_path),
				   len(content))
		
		# Note about LibreOffice
		if not conversion_success:
			LOGGER.debug("💡 Tip: Install LibreOffice to enable document conversion for better text extraction")
		
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


def process_docx_folder(input_folder: str, output_folder: Optional[str] = None, cfg: Optional[ConversionConfig] = None) -> int:
	"""Process all .docx files in a folder sequentially."""
	if not os.path.isdir(input_folder):
		raise ValueError(f"Input folder not found: {input_folder}")

	cfg = cfg or ConversionConfig()
	output_dir = output_folder or input_folder
	os.makedirs(output_dir, exist_ok=True)

	tasks: List[Tuple[str, str]] = []
	for name in os.listdir(input_folder):
		if not name.lower().endswith(".docx"):
			continue
		if name.startswith("~$"):
			continue
		input_path = os.path.join(input_folder, name)
		base_name = os.path.splitext(name)[0]
		output_txt = os.path.join(output_dir, f"{base_name}.txt")
		tasks.append((input_path, output_txt))

	if not tasks:
		LOGGER.info("No DOCX files found in '%s'.", input_folder)
		return 0

	processed = 0
	for input_path, output_txt in tasks:
		ok = extract_docx_to_text_file(input_path, output_txt, cfg=cfg)
		if ok:
			processed += 1

	LOGGER.info("Processed %d/%d DOCX file(s) from '%s' into '%s'.", processed, len(tasks), input_folder, output_dir)
	return processed


