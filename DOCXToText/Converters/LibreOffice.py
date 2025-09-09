import logging
import os
import shutil
import subprocess
import tempfile
import asyncio
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
		# Try portable LibreOffice first (more reliable)
		r"C:\LibreOfficePortable\App\LibreOffice\program\soffice.exe",
		r".\LibreOffice\program\soffice.exe",
		r"LibreOffice\program\soffice.exe",
		# Standard installations
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


async def convert_docx_via_libreoffice_async(input_docx_path: str, output_docx_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Convert DOCX to DOC and back to DOCX using LibreOffice CLI with separate user profiles.
	This approach uses asyncio subprocess and unique LibreOffice profiles to avoid conflicts.
	"""
	cfg = cfg or ConversionConfig()
	soffice = _find_soffice_executable(cfg)
	if not soffice:
		LOGGER.warning("LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH in .env file")
		return False

	# Create a temporary profile directory for this LibreOffice instance
	temp_profile_dir = tempfile.mkdtemp(prefix="textopo_lo_profile_")
	
	try:
		# Create temporary directory for conversion
		with tempfile.TemporaryDirectory() as temp_dir:
			# Get the base name of the input file without extension
			base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
			
			# Step 1: Convert DOCX to DOC
			doc_path = os.path.join(temp_dir, f"{base_name}.doc")
			
			LOGGER.debug("Converting DOCX to DOC with separate profile: %s", temp_profile_dir)
			process1 = await asyncio.create_subprocess_exec(
				soffice,
				f"-env:UserInstallation=file://{temp_profile_dir}",
				"--headless",
				"--convert-to", "doc",
				"--outdir", temp_dir,
				input_docx_path,
				stdout=asyncio.subprocess.PIPE,
				stderr=asyncio.subprocess.PIPE
			)
			
			stdout1, stderr1 = await asyncio.wait_for(
				process1.communicate(), 
				timeout=cfg.convert_timeout_sec
			)
			
			LOGGER.debug("DOCX->DOC conversion result: returncode=%s, stdout='%s', stderr='%s'", 
						process1.returncode, stdout1.decode().strip(), stderr1.decode().strip())
			
			if process1.returncode != 0:
				LOGGER.error("Error converting to DOC: %s", stderr1.decode().strip())
				return False
			
			# Check if DOC file was created
			if not os.path.exists(doc_path):
				LOGGER.error("DOC file not created: %s", doc_path)
				LOGGER.debug("Files in temp directory after step 1: %s", os.listdir(temp_dir))
				return False
			
			# Step 2: Convert DOC back to DOCX
			LOGGER.debug("Converting DOC back to DOCX with same profile")
			process2 = await asyncio.create_subprocess_exec(
				soffice,
				f"-env:UserInstallation=file://{temp_profile_dir}",
				"--headless", 
				"--convert-to", "docx",
				"--outdir", temp_dir,
				doc_path,
				stdout=asyncio.subprocess.PIPE,
				stderr=asyncio.subprocess.PIPE
			)
			
			stdout2, stderr2 = await asyncio.wait_for(
				process2.communicate(),
				timeout=cfg.convert_timeout_sec
			)
			
			LOGGER.debug("DOC->DOCX conversion result: returncode=%s, stdout='%s', stderr='%s'", 
						process2.returncode, stdout2.decode().strip(), stderr2.decode().strip())
			
			if process2.returncode != 0:
				LOGGER.error("Error converting back to DOCX: %s", stderr2.decode().strip())
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
				LOGGER.debug("Files in temp directory after step 2: %s", os.listdir(temp_dir))
				return False
				
	except asyncio.TimeoutError:
		LOGGER.error("Conversion timed out after %s seconds", cfg.convert_timeout_sec)
		return False
	except Exception as e:
		LOGGER.exception("Error during conversion: %s", e)
		return False
	finally:
		# Clean up the temporary profile directory
		try:
			shutil.rmtree(temp_profile_dir)
			LOGGER.debug("Cleaned up LibreOffice profile directory: %s", temp_profile_dir)
		except Exception as cleanup_error:
			LOGGER.warning("Failed to cleanup temp profile directory %s: %s", temp_profile_dir, cleanup_error)

def convert_docx_via_libreoffice(input_docx_path: str, output_docx_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Synchronous wrapper for async LibreOffice conversion.
	Creates an event loop to run the async conversion.
	"""
	try:
		# Try to get the current event loop
		loop = asyncio.get_event_loop()
		if loop.is_running():
			# If we're in an async context, we need to run in a thread
			import concurrent.futures
			with concurrent.futures.ThreadPoolExecutor() as executor:
				future = executor.submit(
					lambda: asyncio.run(convert_docx_via_libreoffice_async(input_docx_path, output_docx_path, cfg))
				)
				return future.result()
		else:
			# We can run directly
			return asyncio.run(convert_docx_via_libreoffice_async(input_docx_path, output_docx_path, cfg))
	except RuntimeError:
		# No event loop exists, create one
		return asyncio.run(convert_docx_via_libreoffice_async(input_docx_path, output_docx_path, cfg))




