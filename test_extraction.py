#!/usr/bin/env python3
"""
Quick test script to debug DOCX extraction issues.
Usage: python test_extraction.py path/to/document.docx
"""
import os
import sys
import logging
import tempfile
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from DOCXToText.config import ConversionConfig
from DOCXToText.Converters.LibreOffice import convert_docx_via_libreoffice
from DOCXToText.Extractors.DOCXExtractor import extract_content_with_python_docx

def setup_debug_logging():
    """Setup detailed logging for debugging."""
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

def test_single_document(docx_path: str):
    """Test extraction on a single document with detailed logging."""
    if not os.path.exists(docx_path):
        print(f"❌ File not found: {docx_path}")
        return False
    
    print(f"🔍 Testing extraction for: {os.path.basename(docx_path)}")
    print(f"📁 Full path: {docx_path}")
    print(f"📊 File size: {os.path.getsize(docx_path):,} bytes")
    print("="*60)
    
    cfg = ConversionConfig()
    
    # Test 1: Try LibreOffice conversion
    print("\n🔄 STEP 1: Testing LibreOffice conversion...")
    temp_converted_path = None
    conversion_success = False
    
    try:
        temp_converted_path = tempfile.mktemp(suffix=".docx")
        conversion_success = convert_docx_via_libreoffice(docx_path, temp_converted_path, cfg=cfg, max_instances=1)
        
        if conversion_success:
            print(f"✅ LibreOffice conversion successful!")
            print(f"📁 Converted file: {temp_converted_path}")
            print(f"📊 Converted size: {os.path.getsize(temp_converted_path):,} bytes")
            converted_file = temp_converted_path
        else:
            print(f"❌ LibreOffice conversion failed")
            converted_file = docx_path
    except Exception as e:
        print(f"❌ LibreOffice conversion error: {e}")
        converted_file = docx_path
    
    # Test 2: Try python-docx extraction
    print(f"\n📄 STEP 2: Testing python-docx extraction on: {os.path.basename(converted_file)}")
    
    try:
        content = extract_content_with_python_docx(converted_file)
        
        if content and content.strip():
            print(f"✅ Extraction successful!")
            print(f"📊 Content length: {len(content):,} characters")
            print(f"📊 Line count: {len(content.split(chr(10))):,} lines")
            print(f"📊 Word count (approx): {len(content.split()):,} words")
            
            # Show first 500 characters
            preview = content[:500].replace('\n', '\\n').replace('\t', '\\t')
            print(f"\n📖 CONTENT PREVIEW (first 500 chars):")
            print(f"'{preview}{'...' if len(content) > 500 else ''}'")
            
            # Save to test file
            output_path = f"test_output_{os.path.splitext(os.path.basename(docx_path))[0]}.txt"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"\n💾 Full content saved to: {output_path}")
            
            return True
        else:
            print(f"⚠️ Extraction completed but no content found")
            print(f"📊 Content: {repr(content)}")
            return False
            
    except Exception as e:
        print(f"❌ Extraction failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # Cleanup
        if conversion_success and temp_converted_path and os.path.exists(temp_converted_path):
            try:
                os.remove(temp_converted_path)
                print(f"🗑️ Cleaned up temp file")
            except:
                pass

def main():
    if len(sys.argv) != 2:
        print("Usage: python test_extraction.py path/to/document.docx")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    setup_debug_logging()
    
    success = test_single_document(docx_path)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
