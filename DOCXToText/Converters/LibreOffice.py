import logging
import os
import shutil
import subprocess
import tempfile
from typing import List, Optional

from DOCXToText.config import ConversionConfig

LOGGER = logging.getLogger(__name__)


def _windows_startupinfo():
	"""Return a STARTUPINFO that hides the window on Windows."""
	if os.name != "nt":
		return None
	startupinfo = subprocess.STARTUPINFO()
	startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
	startupinfo.wShowWindow = 0  # SW_HIDE
	return startupinfo




def _find_soffice_executable(cfg: ConversionConfig) -> Optional[str]:
	"""Discover LibreOffice `soffice` executable."""
	candidates: List[str] = []

	# Windows: try `where soffice` first
	if os.name == "nt":
		try:
			creationflags = subprocess.CREATE_NO_WINDOW
			res = subprocess.run(
				["where", "soffice"],
				stdin=subprocess.DEVNULL,
				capture_output=True,
				text=True,
				timeout=cfg.version_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo(),
			)
			if res.returncode == 0 and res.stdout:
				for line in res.stdout.splitlines():
					path = line.strip().strip('"')
					if path and os.path.isfile(path):
						candidates.append(path)
		except Exception:
			pass
	if cfg.soffice_path:
		# Accept direct path or directory
		if os.path.isdir(cfg.soffice_path):
			candidates.extend([
				os.path.join(cfg.soffice_path, "soffice.exe"),
				os.path.join(cfg.soffice_path, "soffice.com"),
				os.path.join(cfg.soffice_path, "soffice.bin"),
			])
		else:
			candidates.append(cfg.soffice_path)
	
	# Add common LibreOffice installation paths
	candidates.extend([
		r"C:\Program Files\LibreOffice\program\soffice.exe",
		r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
		r"C:\Program Files\LibreOffice\program\soffice.com",
		r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
		r"C:\Program Files\LibreOffice\program\soffice.bin",
		r"C:\Program Files (x86)\LibreOffice\program\soffice.bin",
		"soffice",
		"libreoffice",
	])

	for path in candidates:
		try:
			LOGGER.debug("Probing LibreOffice executable: %s", path)
			creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
			test = subprocess.run(
				[path, "--version"],
				stdin=subprocess.DEVNULL,
				capture_output=True,
				text=True,
				timeout=cfg.version_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo(),
			)
			if test.returncode == 0:
				LOGGER.info("Detected LibreOffice at: %s", path)
				return path
			LOGGER.debug("Probe failed for %s: rc=%s", path, test.returncode)
		except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
			continue

	return None


def convert_docx_via_libreoffice(input_docx_path: str, output_docx_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Convert DOCX to DOC and back to DOCX using LibreOffice CLI to normalize the format.
	This helps extract content that might not be accessible in the original format.
	"""
	cfg = cfg or ConversionConfig()
	soffice = _find_soffice_executable(cfg)
	if not soffice:
		LOGGER.warning("LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH in .env file")
		return False

	try:
		# Create temporary directory for conversion
		with tempfile.TemporaryDirectory() as temp_dir:
			# Get the base name of the input file without extension
			base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
			
			# Step 1: Convert DOCX to DOC
			doc_path = os.path.join(temp_dir, f"{base_name}.doc")
			cmd1 = [
				soffice,
				"--headless",
				"--convert-to", "doc",
				"--outdir", temp_dir,
				input_docx_path
			]
			
			LOGGER.debug("Converting DOCX to DOC with command: %s", ' '.join(cmd1))
			creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
			result1 = subprocess.run(
				cmd1, 
				capture_output=True, 
				text=True, 
				timeout=cfg.convert_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo()
			)
			
			LOGGER.debug("DOCX->DOC conversion result: returncode=%s, stdout='%s', stderr='%s'", 
						result1.returncode, result1.stdout.strip(), result1.stderr.strip())
			
			if result1.returncode != 0:
				LOGGER.error("Error converting to DOC: %s", result1.stderr)
				return False
			
			# Check if DOC file was created
			if not os.path.exists(doc_path):
				LOGGER.error("DOC file not created: %s", doc_path)
				# List files in temp directory for debugging
				LOGGER.debug("Files in temp directory after step 1: %s", os.listdir(temp_dir))
				return False
			
			# Step 2: Convert DOC back to DOCX
			cmd2 = [
				soffice,
				"--headless", 
				"--convert-to", "docx",
				"--outdir", temp_dir,
				doc_path
			]
			
			LOGGER.debug("Converting DOC back to DOCX with command: %s", ' '.join(cmd2))
			result2 = subprocess.run(
				cmd2, 
				capture_output=True, 
				text=True, 
				timeout=cfg.convert_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo()
			)
			
			LOGGER.debug("DOC->DOCX conversion result: returncode=%s, stdout='%s', stderr='%s'", 
						result2.returncode, result2.stdout.strip(), result2.stderr.strip())
			
			if result2.returncode != 0:
				LOGGER.error("Error converting back to DOCX: %s", result2.stderr)
				return False
			
			# Step 3: Delete the intermediate DOC file
			if os.path.exists(doc_path):
				os.remove(doc_path)
				LOGGER.debug("Deleted intermediate DOC file")
			
			# Step 4: Copy the converted file to output location
			converted_docx = os.path.join(temp_dir, f"{base_name}.docx")
			if os.path.exists(converted_docx):
				shutil.copy2(converted_docx, output_docx_path)
				LOGGER.debug("Successfully converted and saved to: %s", output_docx_path)
				return True
			else:
				LOGGER.error("Converted file not found: %s", converted_docx)
				# List files in temp directory for debugging
				LOGGER.debug("Files in temp directory after step 2: %s", os.listdir(temp_dir))
				return False
				
	except subprocess.TimeoutExpired:
		LOGGER.error("Conversion timed out after %s seconds", cfg.convert_timeout_sec)
		return False
	except Exception as e:
		LOGGER.exception("Error during conversion: %s", e)
		return False




