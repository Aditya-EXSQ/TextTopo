import os
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class ConversionConfig:
	"""Runtime configuration for LibreOffice conversion and probing."""
	soffice_path: Optional[str] = os.getenv("SOFFICE_PATH") or None
	convert_timeout_sec: int = int(os.getenv("CONVERT_TIMEOUT_SEC", "60"))
	version_timeout_sec: int = int(os.getenv("VERSION_TIMEOUT_SEC", "10"))


