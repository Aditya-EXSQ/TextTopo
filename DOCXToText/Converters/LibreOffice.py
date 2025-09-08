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


_LO_PROFILE_DIR: Optional[str] = None


def _get_lo_profile_dir() -> str:
	"""Return a per-process LibreOffice user profile directory, reused across calls."""
	global _LO_PROFILE_DIR
	if _LO_PROFILE_DIR and os.path.isdir(_LO_PROFILE_DIR):
		return _LO_PROFILE_DIR
	# Create a stable path under system temp using PID to avoid cross-process clashes
	pid = os.getpid()
	base_dir = os.path.join(tempfile.gettempdir(), f"textopo_lo_profile_{pid}")
	os.makedirs(base_dir, exist_ok=True)
	_LO_PROFILE_DIR = base_dir
	return _LO_PROFILE_DIR


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
			# Use a persistent per-process user profile for speed
			lo_profile = _get_lo_profile_dir()
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


def convert_via_fake_doc_roundtrip(input_docx_path: str, output_docx_path: str, cfg: Optional[ConversionConfig] = None) -> bool:
	"""
	Rename trick fallback: copy the DOCX bytes to a temporary path with a .doc extension
	and ask LibreOffice to convert that .doc directly to .docx.

	This mimics the manual rename you performed and sometimes bypasses broken
	customXml references encountered by python-docx.
	"""
	cfg = cfg or ConversionConfig()
	soffice = _find_soffice_executable(cfg)
	if not soffice:
		LOGGER.warning("LibreOffice not found. Install it or set SOFFICE_PATH.")
		return False

	try:
		with tempfile.TemporaryDirectory() as temp_dir:
			base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
			fake_doc_path = os.path.join(temp_dir, f"{base_name}.doc")
			# Copy bytes verbatim, only the extension changes
			shutil.copy2(input_docx_path, fake_doc_path)

			# Use a persistent per-process user profile for speed
			lo_profile = _get_lo_profile_dir()
			lo_profile_uri = "file:///" + lo_profile.replace("\\", "/")

			common_flags = ["--headless", "--norestore", f"-env:UserInstallation={lo_profile_uri}"]
			creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
			cmd = [soffice, *common_flags, "--convert-to", "docx", "--outdir", temp_dir, fake_doc_path]
			LOGGER.info("Converting fake .doc to .docx using LibreOffice (rename trick)...")
			res = subprocess.run(
				cmd,
				stdin=subprocess.DEVNULL,
				capture_output=True,
				text=True,
				timeout=cfg.convert_timeout_sec,
				creationflags=creationflags,
				startupinfo=_windows_startupinfo(),
			)
			if res.returncode != 0:
				LOGGER.error("Fake .doc -> .docx conversion failed: %s", res.stderr.strip())
				return False

			converted_docx = os.path.join(temp_dir, f"{base_name}.docx")
			if not os.path.exists(converted_docx):
				LOGGER.error("Converted DOCX not found after rename trick: %s", converted_docx)
				return False

			shutil.copy2(converted_docx, output_docx_path)
			LOGGER.debug("Rename-trick converted file saved to: %s", output_docx_path)
			return True

	except subprocess.TimeoutExpired:
		LOGGER.error("LibreOffice rename-trick conversion timed out")
		return False
	except Exception as exc:
		LOGGER.exception("Unexpected error during rename-trick conversion: %s", exc)
		return False


