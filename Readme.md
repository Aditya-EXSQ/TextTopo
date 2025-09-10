# TextTopo

TextTopo is a Python package for robust text extraction from DOCX files. It uses LibreOffice conversion to normalize documents before extracting text content with python-docx, ensuring maximum compatibility and text extraction quality.

## Features

- **Robust DOCX Processing**: Converts DOCX → DOC → DOCX using LibreOffice for normalization
- **Async Processing**: Fully asynchronous conversions with configurable concurrency
- **Batch Processing**: Process single files or entire directories
- **Placeholder Handling**: Preserves and normalizes MERGEFIELD placeholders
- **Table Support**: Extracts text from tables with proper formatting
- **Configurable**: Environment-aware configuration with sensible defaults
- **CLI Interface**: Easy-to-use command line interface
- **Concurrent Safety**: Uses temporary profile directories to avoid LibreOffice conflicts

## Installation

### Prerequisites

1. **LibreOffice**: Required for document conversion
   - Windows: Download from [LibreOffice website](https://www.libreoffice.org/)
   - The package will auto-detect LibreOffice in common installation paths

2. **Python Dependencies**: Install required packages
   ```bash
   pip install python-docx python-dotenv
   ```

### Package Setup

1. Clone or download the TextTopo package
2. Navigate to the TextTopo directory
3. The package is ready to use!

## Usage

### Command Line Interface

#### Basic Usage

```bash
# Process single file
python Scripts/Extract.py --input document.docx --output ./extracted/

# Process entire folder
python Scripts/Extract.py --input ./documents/ --output ./extracted/

# Print single file content to stdout
python Scripts/Extract.py --input document.docx --stdout
```

#### Advanced Options

```bash
# Process with custom concurrency and timeout
python Scripts/Extract.py \
    --input ./documents/ \
    --output ./extracted/ \
    --concurrency 8 \
    --timeout 120

# Enable debug logging with log file
python Scripts/Extract.py \
    --input ./documents/ \
    --output ./extracted/ \
    --log-level DEBUG \
    --log-file extraction.log

# Use custom LibreOffice path
python Scripts/Extract.py \
    --input document.docx \
    --output ./extracted/ \
    --soffice-path "C:\Custom\LibreOffice\program\soffice.exe"

# Process only current directory (no subdirectories)
python Scripts/Extract.py \
    --input ./documents/ \
    --output ./extracted/ \
    --no-recursive
```

#### CLI Options

| Option | Description | Default |
|--------|-------------|---------|
| `--input`, `-i` | Input DOCX file or directory | Required |
| `--output`, `-o` | Output directory for text files | Required (unless --stdout) |
| `--stdout` | Print to stdout (single file only) | False |
| `--concurrency`, `-c` | Max concurrent conversions | 4 |
| `--timeout`, `-t` | Conversion timeout (seconds) | 60 |
| `--no-recursive`, `-nr` | Don't search subdirectories | False |
| `--overwrite` | Overwrite existing files | False |
| `--soffice-path` | Path to LibreOffice executable | Auto-detect |
| `--log-level` | Logging level (DEBUG/INFO/WARNING/ERROR) | INFO |
| `--log-file` | Log file path | Console only |
| `--temp-dir` | Temporary directory name | texttopo_temp |

### Python API

#### Single File Processing

```python
import asyncio
from DOCXToText import process_file, ConversionConfig

async def extract_single_file():
    # Basic usage
    text = await process_file("document.docx", output_dir="./extracted/")
    print(f"Extracted {len(text)} characters")
    
    # With custom configuration
    config = ConversionConfig(
        concurrency_limit=8,
        conversion_timeout=120,
        overwrite_existing=True
    )
    text = await process_file("document.docx", output_dir="./extracted/", config=config)

# Run the async function
asyncio.run(extract_single_file())
```

#### Batch Processing

```python
import asyncio
from DOCXToText import process_files_in_parallel
from DOCXToText.Pipeline.Batch import find_docx_files

async def extract_multiple_files():
    # Find all DOCX files in a directory
    files = find_docx_files("./documents/", recursive=True)
    
    # Process all files in parallel
    results = await process_files_in_parallel(
        files=files,
        output_dir="./extracted/",
        config=None  # Uses default configuration
    )
    
    print(f"Successfully processed {len(results)} files")
    for file_path, text in results.items():
        print(f"{file_path}: {len(text)} characters extracted")

asyncio.run(extract_multiple_files())
```

#### Custom Configuration

```python
from DOCXToText import ConversionConfig

# Create configuration from environment variables
config = ConversionConfig.from_env()

# Create custom configuration
config = ConversionConfig(
    soffice_path="C:\\Custom\\LibreOffice\\program\\soffice.exe",
    concurrency_limit=8,
    conversion_timeout=120,
    temp_dir_name="my_temp_dir",
    overwrite_existing=True,
    log_level="DEBUG"
)
```

## Configuration

### Environment Variables

TextTopo can be configured using environment variables:

```bash
# LibreOffice settings
export SOFFICE_PATH="/path/to/soffice"
export CONVERSION_TIMEOUT="60"

# Processing settings  
export CONCURRENCY_LIMIT="4"
export TEMP_DIR_NAME="texttopo_temp"

# Output settings
export OUTPUT_EXTENSION=".txt"
export OVERWRITE_EXISTING="false"

# Logging
export LOG_LEVEL="INFO"
```

### Configuration File

You can also use a `.env` file in your project directory:

```env
# .env file
SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
CONCURRENCY_LIMIT=8
CONVERSION_TIMEOUT=120
LOG_LEVEL=DEBUG
```

## How It Works

1. **LibreOffice Conversion**: 
   - Creates a temporary directory in the current working directory
   - Converts DOCX to DOC format using LibreOffice
   - Converts DOC back to DOCX for normalization
   - Uses unique temporary profile directories for concurrency safety

2. **Text Extraction**:
   - Uses python-docx to traverse the normalized document
   - Extracts text from paragraphs and tables
   - Normalizes MERGEFIELD placeholders (`{PLACEHOLDER}` → `PLACEHOLDER`)
   - Preserves table structure with tab separators

3. **Fallback Processing**:
   - If LibreOffice conversion fails, attempts extraction from original file
   - Provides detailed logging for troubleshooting

## Troubleshooting

### LibreOffice Not Found

```
Error: LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH
```

**Solution**: Install LibreOffice or set the `SOFFICE_PATH` environment variable:
```bash
export SOFFICE_PATH="/path/to/libreoffice/program/soffice"
```

### LibreOffice Interactive Popup

```
LibreOffice 25.8.1.1 ... Press Enter to continue...
```

**Problem**: Some LibreOffice installations show interactive prompts despite headless mode.

**Solutions**:
1. **Recommended**: Use without LibreOffice (still provides excellent text extraction)
   ```bash
   python Scripts/Extract.py --input docs/ --output text/
   ```

2. **Single File Processing**: LibreOffice works better for individual files:
   ```bash
   # Process single files with LibreOffice (requires pressing Enter for popup)
   python Scripts/Extract.py --input document.docx --stdout --enable-libreoffice
   ```

3. **Batch Processing**: LibreOffice popups interfere with batch processing, so use without:
   ```bash
   # Batch processing works reliably without LibreOffice
   python Scripts/Extract.py --input docs/ --output text/ --concurrency 6
   ```

4. **Environment Variable**: Disable LibreOffice by default:
   ```bash
   set ENABLE_LIBREOFFICE=false
   ```

### Conversion Timeout

```
Error: LibreOffice conversion timed out after 60 seconds
```

**Solution**: Increase timeout or reduce concurrency:
```bash
python Scripts/Extract.py --input docs/ --output text/ --timeout 120 --concurrency 2
```

### Permission Errors

**Solution**: Ensure you have write permissions to the output directory and current working directory.

### Memory Issues

**Solution**: Reduce concurrency limit:
```bash
python Scripts/Extract.py --input docs/ --output text/ --concurrency 2
```

## Examples

### Extract from Sample Documents

```bash
# Process the sample documents in Data/DOCX/
python Scripts/Extract.py \
    --input Data/DOCX/ \
    --output Data/TXT/ \
    --concurrency 6 \
    --log-level INFO
```

### Extract Single Document with Debug Info

```bash
python Scripts/Extract.py \
    --input "Data/DOCX/MHCA MP Approval Letter.docx" \
    --output Data/TXT/ \
    --log-level DEBUG \
    --log-file debug.log
```

### Extract and View Content

```bash
python Scripts/Extract.py \
    --input "Data/DOCX/MHCA MP Approval Letter.docx" \
    --stdout
```

## Package Structure

```
TextTopo/
├── Data/                         # Sample DOCX files
├── DOCXToText/                   # Core package
│   ├── __init__.py
│   ├── CLI.py                    # Command line interface
│   ├── config.py                 # Configuration management
│   ├── logging_setup.py          # Logging configuration
│   ├── Converters/
│   │   ├── __init__.py
│   │   └── LibreOffice.py        # LibreOffice conversion logic
│   ├── Extractors/
│   │   ├── __init__.py
│   │   └── DOCXExtractor.py      # python-docx text extraction
│   └── Pipeline/
│       ├── __init__.py
│       └── Batch.py              # Batch processing pipeline
├── Scripts/
│   └── Extract.py                # Entry point script
└── Readme.md                     # This file
```

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is provided as-is for educational and development purposes.