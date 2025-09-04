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
	candidates.extend([
		r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
		r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
		r"C:\\Program Files\\LibreOffice\\program\\soffice.com",
		r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
		r"C:\\Program Files\\LibreOffice\\program\\soffice.bin",
		r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.bin",
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
	"""Convert DOCX->DOC->DOCX via LibreOffice to normalize shape."""
	cfg = cfg or ConversionConfig()
	soffice = _find_soffice_executable(cfg)
	if not soffice:
		LOGGER.warning("LibreOffice not found. Install it or set SOFFICE_PATH.")
		return False

	try:
		with tempfile.TemporaryDirectory() as temp_dir:
			base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
			doc_path = os.path.join(temp_dir, f"{base_name}.doc")
			# Use a temp user profile to avoid any first-run dialogs or locks
			lo_profile = os.path.join(temp_dir, "lo_profile")
			os.makedirs(lo_profile, exist_ok=True)
			lo_profile_uri = "file:///" + lo_profile.replace("\\", "/")

			common_flags = ["--headless", "--norestore", f"-env:UserInstallation={lo_profile_uri}"]
			cmd1 = [soffice, *common_flags, "--convert-to", "doc", "--outdir", temp_dir, input_docx_path]
			LOGGER.info("Converting DOCX to DOC using LibreOffice...")
			creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
			res1 = subprocess.run(
				cmd1,
				stdin=subprocess.DEVNULL,
				capture_output=True,
				text=True,
				timeout=cfg.convert_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo(),
			)
			if res1.returncode != 0:
				LOGGER.error("DOCX->DOC conversion failed: %s", res1.stderr.strip())
				return False

			if not os.path.exists(doc_path):
				LOGGER.error("DOC not created at: %s", doc_path)
				return False

			cmd2 = [soffice, *common_flags, "--convert-to", "docx", "--outdir", temp_dir, doc_path]
			LOGGER.info("Converting DOC back to DOCX using LibreOffice...")
			res2 = subprocess.run(
				cmd2,
				stdin=subprocess.DEVNULL,
				capture_output=True,
				text=True,
				timeout=cfg.convert_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo(),
			)
			if res2.returncode != 0:
				LOGGER.error("DOC->DOCX conversion failed: %s", res2.stderr.strip())
				return False

			try:
				if os.path.exists(doc_path):
					os.remove(doc_path)
					LOGGER.debug("Deleted intermediate DOC: %s", doc_path)
			except OSError:
				pass

			converted_docx = os.path.join(temp_dir, f"{base_name}.docx")
			if not os.path.exists(converted_docx):
				LOGGER.error("Converted DOCX not found: %s", converted_docx)
				return False

			shutil.copy2(converted_docx, output_docx_path)
			LOGGER.debug("Converted file saved to: %s", output_docx_path)
			return True

	except subprocess.TimeoutExpired:
		LOGGER.error("LibreOffice conversion timed out")
		return False
	except Exception as exc:
		LOGGER.exception("Unexpected error during LibreOffice conversion: %s", exc)
		return False


