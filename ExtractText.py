import subprocess
import os
import tempfile
import shutil
from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from dotenv import load_dotenv
import re

load_dotenv()


def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, DocumentType):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_paragraph_text_with_fields(paragraph):
    """
    Extract text from a paragraph, including MERGEFIELD placeholders.
    """
    text = ""
    for run in paragraph.runs:
        text += run.text
    return text


def convert_docx_via_libreoffice(input_docx_path, output_docx_path):
    """
    Convert DOCX to DOC and back to DOCX using LibreOffice CLI to normalize the format.
    This helps extract content that might not be accessible in the original format.
    """
    # LibreOffice executable path - try common locations if not in env
    SOFFICE_PATH = os.getenv("SOFFICE_PATH")
    
    if not SOFFICE_PATH:
        # Try common LibreOffice installation paths on Windows
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "soffice",  # If it's in PATH
            "libreoffice"  # Alternative command name
        ]
        
        for path in possible_paths:
            try:
                # Test if the executable exists and works
                test_result = subprocess.run([path, "--version"], 
                                           capture_output=True, text=True, timeout=10)
                if test_result.returncode == 0:
                    SOFFICE_PATH = path
                    print(f"Found LibreOffice at: {path}")
                    break
            except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
                continue
    
    if not SOFFICE_PATH:
        print("LibreOffice not found. Please install LibreOffice or set SOFFICE_PATH in .env file")
        return False
    
    try:
        # Create temporary directory for conversion
        with tempfile.TemporaryDirectory() as temp_dir:
            # Get the base name of the input file without extension
            base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
            
            # Step 1: Convert DOCX to DOC
            doc_path = os.path.join(temp_dir, f"{base_name}.doc")
            cmd1 = [
                SOFFICE_PATH,
                "--headless",
                "--convert-to", "doc",
                "--outdir", temp_dir,
                input_docx_path
            ]
            
            print("Converting DOCX to DOC...")
            result1 = subprocess.run(cmd1, capture_output=True, text=True, timeout=60)
            if result1.returncode != 0:
                print(f"Error converting to DOC: {result1.stderr}")
                return False
            
            # Check if DOC file was created
            if not os.path.exists(doc_path):
                print(f"DOC file not created: {doc_path}")
                return False
            
            # Step 2: Convert DOC back to DOCX
            cmd2 = [
                SOFFICE_PATH,
                "--headless", 
                "--convert-to", "docx",
                "--outdir", temp_dir,
                doc_path
            ]
            
            print("Converting DOC back to DOCX...")
            result2 = subprocess.run(cmd2, capture_output=True, text=True, timeout=60)
            if result2.returncode != 0:
                print(f"Error converting back to DOCX: {result2.stderr}")
                return False
            
            # Step 3: Delete the intermediate DOC file
            if os.path.exists(doc_path):
                os.remove(doc_path)
                print("Deleted intermediate DOC file")
            
            # Step 4: Copy the converted file to output location
            converted_docx = os.path.join(temp_dir, f"{base_name}.docx")
            if os.path.exists(converted_docx):
                shutil.copy2(converted_docx, output_docx_path)
                print(f"Successfully converted and saved to: {output_docx_path}")
                return True
            else:
                print(f"Converted file not found: {converted_docx}")
                # List files in temp directory for debugging
                print(f"Files in temp directory: {os.listdir(temp_dir)}")
                return False
                
    except subprocess.TimeoutExpired:
        print("Conversion timed out")
        return False
    except Exception as e:
        print(f"Error during conversion: {e}")
        return False

def extract_content_with_python_docx(docx_path):
    """
    Extract document content using python-docx library with proper table handling.
    """
    document = Document(docx_path)
    all_text = []
    placeholder_re = re.compile(r"\{\s*([^{}\s].*?)\s*\}")
    
    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            paragraph_text = get_paragraph_text_with_fields(block).strip()
            if paragraph_text:
                # Replace placeholder braces with just the placeholder name
                paragraph_text = placeholder_re.sub(r'\1', paragraph_text)
                all_text.append(paragraph_text)
        elif isinstance(block, Table):
            for row in block.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = ''
                    for paragraph in cell.paragraphs:
                        cell_text += get_paragraph_text_with_fields(paragraph)
                    # Replace placeholder braces with just the placeholder name
                    cell_text = placeholder_re.sub(r'\1', cell_text).strip()
                    row_text.append(cell_text)
                # Join cell text with a tab to represent table columns
                full_row_text = "\t".join(row_text).strip()
                if full_row_text:
                    all_text.append(full_row_text)
    
    return '\n'.join(all_text)

def main():
    input_file = "Master Approval Letter.docx"
    converted_file = "new.docx"
    
    print("=== DOCUMENT CONTENT EXTRACTION WITH LIBREOFFICE CONVERSION ===\n")
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found!")
        print("Available files in current directory:")
        for file in os.listdir("."):
            if file.endswith((".docx", ".doc")):
                print(f"  - {file}")
        return
    
    # Step 1: Convert DOCX via LibreOffice
    print("Step 1: Converting document via LibreOffice...")
    conversion_success = False
    try:
        conversion_success = convert_docx_via_libreoffice(input_file, converted_file)
        if conversion_success:
            print("✅ Conversion successful!")
        else:
            print("❌ Conversion failed, trying to extract from original file...")
            converted_file = input_file
    except Exception as e:
        print(f"❌ Conversion error: {e}")
        print("Falling back to original file...")
        converted_file = input_file
    
    # Step 2: Extract content using python-docx
    print("\nStep 2: Extracting content...")
    try:
        content = extract_content_with_python_docx(converted_file)
        if content.strip():
            print("✅ Content extraction successful!")
            print("\n" + "="*50)
            print("EXTRACTED CONTENT:")
            print("="*50)
            print(content)
        else:
            print("⚠️ Content extraction completed but no text was found")
            print("The document might be empty or contain only images/formatted content")
        
    except Exception as e:
        print(f"❌ Error extracting content: {e}")
        print("This might be due to:")
        print("  - Corrupted document file")
        print("  - Unsupported document format")
        print("  - Missing python-docx library")
    
    # Note: We keep the converted file as new.docx for user reference
    if conversion_success and converted_file == "new.docx" and os.path.exists(converted_file):
        print(f"\n📁 Converted file saved as: {converted_file}")
    elif not conversion_success:
        print(f"\n💡 Tip: Install LibreOffice to enable document conversion for better text extraction")

if __name__ == "__main__":
    main()