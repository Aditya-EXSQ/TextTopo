import argparse
import logging
import os
import sys
from typing import Optional

from DOCXToText.config import ConversionConfig
from DOCXToText.logging_setup import configure_logging
from DOCXToText.Pipeline.Batch import process_docx_folder, extract_docx_to_text_file

LOGGER = logging.getLogger(__name__)


def parse_args(argv: Optional[list] = None) -> argparse.Namespace:
	parser = argparse.ArgumentParser(description="Normalize DOCX via LibreOffice and extract text to .txt.")
	parser.add_argument("--input-file", type=str, help="Path to a single .docx file to extract.")
	parser.add_argument("--input-folder", type=str, help="Path to a folder containing .docx files to extract.")
	parser.add_argument("--output-folder", type=str, help="Destination folder for .txt outputs (defaults to input location).")
	parser.add_argument("--soffice", type=str, default=os.getenv("SOFFICE_PATH"), help="Path to LibreOffice soffice executable.")
	parser.add_argument("--convert-timeout-sec", type=int, default=int(os.getenv("CONVERT_TIMEOUT_SEC", "60")), help="Timeout for each conversion step.")
	parser.add_argument("--version-timeout-sec", type=int, default=int(os.getenv("VERSION_TIMEOUT_SEC", "10")), help="Timeout when probing LibreOffice version.")
	parser.add_argument("-v", "--verbose", action="count", default=0, help="Increase verbosity (repeat for more detail, e.g. -vv).")

	args = parser.parse_args(argv)
	if not args.input_file and not args.input_folder:
		parser.error("Provide either --input-file or --input-folder")
	return args


def main(argv: Optional[list] = None) -> int:
	args = parse_args(argv)
	configure_logging(args.verbose)

	cfg = ConversionConfig(
		soffice_path=args.soffice,
		convert_timeout_sec=args.convert_timeout_sec,
		version_timeout_sec=args.version_timeout_sec,
	)

	try:
		if args.input_file:
			if not os.path.isfile(args.input_file):
				LOGGER.error("Input file does not exist: %s", args.input_file)
				return 2
			out_dir = args.output_folder or os.path.dirname(os.path.abspath(args.input_file)) or "."
			base = os.path.splitext(os.path.basename(args.input_file))[0]
			out_path = os.path.join(out_dir, f"{base}.txt")
			os.makedirs(out_dir, exist_ok=True)
			ok = extract_docx_to_text_file(args.input_file, out_path, cfg=cfg)
			return 0 if ok else 1

		if args.input_folder:
			count = process_docx_folder(args.input_folder, args.output_folder, cfg=cfg)
			return 0 if count > 0 else 1

	except ValueError as ve:
		LOGGER.error(str(ve))
		return 2
	except Exception as exc:
		LOGGER.exception("Unexpected error: %s", exc)
		return 3

	return 0


if __name__ == "__main__":
	sys.exit(main())


