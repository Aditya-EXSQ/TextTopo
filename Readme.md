## TextTopo: DOCX to Text Extraction Pipeline

TextTopo is a modular Python toolkit to normalize and extract text from Microsoft Word documents (.docx), with optional normalization via LibreOffice. It supports batch processing with multiprocessing and provides a clean CLI for single-file or folder-wide extraction. This project is the first step in a larger pipeline for embeddings and clustering.

### Key features
- Normalize DOCX via LibreOffice (DOCX → DOC → DOCX) to improve text extraction reliability
- Extract text from paragraphs and tables (cells joined by tabs; rows as new lines)
- Batch processing with optional multiprocessing (`--workers`)
- Structured logging with configurable verbosity (-v / -vv)
- Clear, modular codebase ready for downstream NLP (embeddings, clustering)

### Directory structure
```
TextTopo/
  Data/                         # Source .docx files (input)
  DOCXToText/                   # Core library (modular package)
    __init__.py
    CLI.py                      # CLI wiring and argument parsing
    config.py                   # ConversionConfig and defaults (env-aware)
    logging_setup.py            # Logging configuration
    Converters/
      LibreOffice.py            # LibreOffice discovery and conversion logic
    Extractors/
      DOCXExtractor.py          # python-docx traversal and text extraction
    Pipeline/
      Batch.py                  # Single-file and folder processing, multiprocessing
  Scripts/
    Extract.py                  # Script entrypoint (thin wrapper over CLI)
  Readme.md                     # This file
  Journey.txt                   # Notes / development log (optional)
```

### Installation
1) Python 3.9+ recommended.
2) Install dependencies:
```bash
pip install python-docx python-dotenv
```
3) Optional: Install LibreOffice for better normalization and extraction:
  - Windows: Install LibreOffice and ensure `soffice.exe` is available. You can set `SOFFICE_PATH` if not in PATH.

### Environment variables (optional)
- `SOFFICE_PATH`: full path to `soffice` executable. Example: `C:\Program Files\LibreOffice\program\soffice.exe`
- `CONVERT_TIMEOUT_SEC`: timeout per conversion step (default 60)
- `VERSION_TIMEOUT_SEC`: timeout when probing LibreOffice version (default 10)
- `WORKERS`: default number of worker processes for folder processing (default 1)

### Usage
Run via script entrypoint:
```bash
python Scripts/Extract.py --input-file "Data\\MHCA MP Approval Letter.docx" -v
```

Process an entire folder with multiprocessing:
```bash
python Scripts/Extract.py --input-folder Data --output-folder TXT --workers 4 -v
```

CLI options:
- `--input-file`: path to a single `.docx` file to extract
- `--input-folder`: path to a folder containing `.docx` files
- `--output-folder`: destination folder for `.txt` outputs (defaults to input location)
- `--soffice`: path to LibreOffice `soffice` executable (overrides `SOFFICE_PATH`)
- `--convert-timeout-sec`: seconds per conversion step (default 60)
- `--version-timeout-sec`: seconds for probing LibreOffice (default 10)
- `--workers`: number of parallel worker processes (default 1)
- `-v` / `-vv`: increase logging verbosity (INFO/DEBUG)

### How it works
1) For each `.docx`, the pipeline optionally normalizes the file via LibreOffice (DOCX → DOC → DOCX). This can surface text that is otherwise harder to read by libraries.
2) The extractor walks the document in block order, extracting:
   - Paragraph text
   - Tables (cells joined with tabs, each row on a new line)
3) The output is written to a `.txt` file with the same base name as the source document.

Notes:
- Temporary files created during LibreOffice conversion live in a temp directory and are cleaned up. Messages like "Converted file saved to: ..." are logged at DEBUG level only.

### Extending for embeddings and clustering
This repository is structured to evolve into a full NLP pipeline. Suggested next modules:
- `NLP/Vectorize.py`: embedding model loader and batched inference (e.g., SentenceTransformers)
- `NLP/Index.py`: ANN index (FAISS) build/load/save
- `NLP/Cluster.py`: clustering algorithms (k-means/HDBSCAN) and metrics
- `IO/Dataset.py`: manifest and metadata management for processed text files

Example future commands:
```bash
python Scripts/BuildEmbeddings.py --input TXT --out data/embeddings --model all-MiniLM-L6-v2 --batch-size 64
python Scripts/ClusterEmbeddings.py --embeddings data/embeddings --algo kmeans --k 50
```

### Troubleshooting
- LibreOffice not found:
  - Set `SOFFICE_PATH` or add LibreOffice to PATH
  - Run with `-vv` to see probe attempts
- Conversion timeouts:
  - Increase `CONVERT_TIMEOUT_SEC`
- Empty output files:
  - Ensure documents contain extractable text (images/graphics are not extracted)
  - Try enabling LibreOffice normalization if it wasn’t used

### License
Internal/Private. Update as appropriate for your organization.


