import io
import logging
import os
import tempfile
import zipfile
from typing import Optional


LOGGER = logging.getLogger(__name__)


def _should_skip_member(name: str) -> bool:
	"""Return True if the zip member should be skipped in repaired docx."""
	# Drop any customXml parts entirely
	if name.lower().startswith("customxml/"):
		return True
	return False


def _filter_relationships(xml_bytes: bytes) -> bytes:
	"""Remove relationships that target customXml/* to avoid broken refs."""
	try:
		text = xml_bytes.decode("utf-8", errors="ignore")
		# Remove any Relationship whose Target contains customXml/ (with optional ../) using either quote style
		import re
		patterns = [
			re.compile(r"<Relationship[^>]*Target=\"(?:\.\./)?customXml/[^\"]*\"[^>]*/>\s*", re.IGNORECASE),
			re.compile(r"<Relationship[^>]*Target='(?:\.\./)?customXml/[^']*'[^>]*/>\s*", re.IGNORECASE),
		]
		for pat in patterns:
			text = pat.sub("", text)
		return text.encode("utf-8")
	except Exception:
		return xml_bytes


def _filter_content_types(xml_bytes: bytes) -> bytes:
	"""Remove [Content_Types].xml overrides for customXml parts."""
	try:
		text = xml_bytes.decode("utf-8", errors="ignore")
		import re
		# Remove any <Override PartName="/customXml/..." .../>
		pattern = re.compile(r"<Override[^>]*PartName=\"/customXml/[^\"]*\"[^>]*/>\s*", re.IGNORECASE)
		text = pattern.sub("", text)
		return text.encode("utf-8")
	except Exception:
		return xml_bytes


def repair_docx_strip_customxml(input_docx_path: str, output_docx_path: str) -> bool:
	"""
	Repair a DOCX by removing customXml parts and relationships that reference them.

	This addresses errors like: "There is no item named 'customXML/itemX.xml' in the archive".
	"""
	if not os.path.isfile(input_docx_path):
		LOGGER.error("Input DOCX not found for repair: %s", input_docx_path)
		return False

	try:
		with zipfile.ZipFile(input_docx_path, "r") as zin:
			# Write to a temp zip first
			with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
				temp_out_path = tmp_out.name
				pass
			with zipfile.ZipFile(temp_out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
				for info in zin.infolist():
					name = info.filename
					if _should_skip_member(name):
						LOGGER.debug("Stripping member from docx: %s", name)
						continue
					data = zin.read(name)
					# Filter relationship parts that may point to customXml
					if name.endswith(".rels") and ("/_rels/" in name or name.endswith("_rels/.rels") or name == "_rels/.rels"):
						data = _filter_relationships(data)
					# Also filter [Content_Types].xml overrides
					if name == "[Content_Types].xml":
						data = _filter_content_types(data)
					zout.writestr(name, data)

			# Move temp to requested output path
			os.replace(temp_out_path, output_docx_path)
			LOGGER.info("Repaired DOCX saved to: %s", output_docx_path)
			return True

	except Exception as exc:
		LOGGER.exception("Failed to repair DOCX '%s': %s", input_docx_path, exc)
		# Best effort cleanup
		try:
			if 'temp_out_path' in locals() and os.path.exists(temp_out_path):
				os.remove(temp_out_path)
		except Exception:
			pass
		return False


