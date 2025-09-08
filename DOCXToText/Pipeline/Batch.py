import logging
import os
import tempfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import List, Optional, Tuple

from DOCXToText.config import ConversionConfig
from DOCXToText.Converters.LibreOffice import convert_docx_via_libreoffice, convert_via_fake_doc_roundtrip
from DOCXToText.Converters.Repair import repair_docx_strip_customxml
from DOCXToText.Extractors.DOCXExtractor import extract_content_with_python_docx

LOGGER = logging.getLogger(__name__)


def extract_docx_to_text_file(input_docx_path: str, output_txt_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""Normalize via LibreOffice (if available) and extract text to .txt file."""
	cfg = cfg or ConversionConfig()

	# Prepare temp path for converted docx
	temp_converted_fd = None
	temp_converted_path: Optional[str] = None
	try:
		temp_converted_fd, temp_converted_path = tempfile.mkstemp(suffix=".docx")
		os.close(temp_converted_fd)
	except Exception:
		temp_converted_fd = None
		temp_converted_path = os.path.join(tempfile.gettempdir(), f"converted_{os.path.basename(input_docx_path)}")

	conversion_success = False
	converted_used_path = input_docx_path
	temp_repaired_path: Optional[str] = None

	try:
		conversion_success = convert_docx_via_libreoffice(input_docx_path, temp_converted_path, cfg=cfg)
		if conversion_success:
			converted_used_path = temp_converted_path
		else:
			LOGGER.info("Proceeding without conversion for: %s", input_docx_path)

		try:
			content = extract_content_with_python_docx(converted_used_path)
		except Exception as first_exc:
			# Fallback 1: strip customXml parts and retry
			try:
				fd, temp_repaired_path = tempfile.mkstemp(suffix=".docx")
				os.close(fd)
			except Exception:
				temp_repaired_path = os.path.join(tempfile.gettempdir(), f"repaired_{os.path.basename(input_docx_path)}")

			repaired = repair_docx_strip_customxml(converted_used_path, temp_repaired_path)
			if repaired:
				LOGGER.info("Retrying extraction after repair for: %s", input_docx_path)
				try:
					content = extract_content_with_python_docx(temp_repaired_path)
				except Exception:
					# Fallback 2: rename-trick roundtrip via LibreOffice
					fd2 = None
					try:
						fd2, temp_converted_path2 = tempfile.mkstemp(suffix=".docx")
						os.close(fd2)
					except Exception:
						temp_converted_path2 = os.path.join(tempfile.gettempdir(), f"rename_trick_{os.path.basename(input_docx_path)}")
					if convert_via_fake_doc_roundtrip(converted_used_path, temp_converted_path2, cfg=cfg):
						LOGGER.info("Retrying extraction after rename-trick for: %s", input_docx_path)
						content = extract_content_with_python_docx(temp_converted_path2)
						try:
							os.remove(temp_converted_path2)
						except Exception:
							pass
					else:
						raise first_exc
			else:
				# If repair failed entirely, try rename-trick before giving up
				fd2 = None
				try:
					fd2, temp_converted_path2 = tempfile.mkstemp(suffix=".docx")
					os.close(fd2)
				except Exception:
					temp_converted_path2 = os.path.join(tempfile.gettempdir(), f"rename_trick_{os.path.basename(input_docx_path)}")
				if convert_via_fake_doc_roundtrip(converted_used_path, temp_converted_path2, cfg=cfg):
					LOGGER.info("Retrying extraction after rename-trick for: %s", input_docx_path)
					content = extract_content_with_python_docx(temp_converted_path2)
					try:
						os.remove(temp_converted_path2)
					except Exception:
						pass
				else:
					raise first_exc

		os.makedirs(os.path.dirname(output_txt_path) or ".", exist_ok=True)
		with open(output_txt_path, "w", encoding="utf-8") as fh:
			fh.write(content or "")

		LOGGER.info("Extracted %s -> %s", input_docx_path, output_txt_path)
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
		if conversion_success and temp_converted_path and os.path.exists(temp_converted_path):
			try:
				os.remove(temp_converted_path)
				LOGGER.debug("Deleted temp converted file: %s", temp_converted_path)
			except OSError:
				pass
		if temp_repaired_path and os.path.exists(temp_repaired_path):
			try:
				os.remove(temp_repaired_path)
				LOGGER.debug("Deleted temp repaired file: %s", temp_repaired_path)
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


