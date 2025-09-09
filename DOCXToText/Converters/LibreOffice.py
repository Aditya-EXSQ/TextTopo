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


# Replace the conversion functions with simplified synchronous version
def convert_docx_via_libreoffice(input_docx_path: str, output_docx_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
    cfg = cfg or ConversionConfig()
    soffice = _find_soffice_executable(cfg)
    if not soffice:
        LOGGER.warning("LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH in .env file")
        return False

    try:
        with tempfile.TemporaryDirectory(prefix="textopo_conversion_", dir=".") as temp_dir:
            base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
            doc_path = os.path.join(temp_dir, f"{base_name}.doc")

            cmd_args = [
                soffice,
                "--headless",
                "--convert-to", "doc",
                "--outdir", temp_dir,
                input_docx_path
            ]

            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

            process1 = subprocess.run(
                cmd_args,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                timeout=cfg.convert_timeout_sec,
                creationflags=creationflags,
                startupinfo=_windows_startupinfo(),
            )

            if process1.returncode != 0:
                LOGGER.error("Error converting to DOC: %s", process1.stderr)
                return False

            if not os.path.exists(doc_path):
                LOGGER.error("DOC file not created: %s", doc_path)
                return False

            cmd_args2 = [
                soffice,
                "--headless",
                "--convert-to", "docx",
                "--outdir", temp_dir,
                doc_path
            ]

            process2 = subprocess.run(
                cmd_args2,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                timeout=cfg.convert_timeout_sec,
                creationflags=creationflags,
                startupinfo=_windows_startupinfo(),
            )

            if process2.returncode != 0:
                LOGGER.error("Error converting back to DOCX: %s", process2.stderr)
                return False

            if os.path.exists(doc_path):
                os.remove(doc_path)
                LOGGER.debug("Deleted intermediate DOC file")

            converted_docx = os.path.join(temp_dir, f"{base_name}.docx")
            if os.path.exists(converted_docx):
                shutil.copy2(converted_docx, output_docx_path)
                LOGGER.debug("Successfully converted and saved to: %s", output_docx_path)
                return True
            else:
                LOGGER.error("Converted file not found: %s", converted_docx)
                return False

    except subprocess.TimeoutExpired:
        LOGGER.error("Conversion timed out after %s seconds", cfg.convert_timeout_sec)
        return False
    except Exception as e:
        LOGGER.warning("Error during conversion: %s", e)
        return False




