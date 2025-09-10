"""
DOCX text extraction utilities using python-docx.
Extracts text from paragraphs and tables with placeholder handling.
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Generator, Union

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from ..logging_setup import get_logger

logger = get_logger("extractors.docx")


def iter_block_items(parent: Union[DocumentType, _Cell]) -> Generator[Union[Paragraph, Table], None, None]:
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    
    Args:
        parent: Document or Cell object to iterate through
        
    Yields:
        Paragraph or Table objects in document order
    """
    if isinstance(parent, DocumentType):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Parent must be Document or _Cell instance")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_paragraph_text_with_fields(paragraph: Paragraph) -> str:
    """
    Extract text from a paragraph, including MERGEFIELD placeholders.
    
    Args:
        paragraph: Paragraph object from python-docx
        
    Returns:
        Complete text content of the paragraph
    """
    text = ""
    for run in paragraph.runs:
        text += run.text
    return text


def extract_content(docx_path: str) -> str:
    """
    Extract document content using python-docx library with proper table handling.
    
    Args:
        docx_path: Path to the DOCX file
        
    Returns:
        Extracted text content with placeholders normalized
        
    Raises:
        Exception: If document cannot be opened or processed
    """
    logger.debug(f"Starting text extraction from: {docx_path}")
    
    try:
        # Try to open the document with error handling for custom XML issues
        try:
            document = Document(docx_path)
        except Exception as e:
            error_msg = str(e)
            if "customXML" in error_msg and "archive" in error_msg:
                logger.warning(f"Document has corrupted custom XML parts, attempting alternative extraction: {docx_path}")
                # Try to extract using zipfile directly as a fallback
                return _extract_with_zipfile_fallback(docx_path)
            else:
                raise e
        all_text = []
        
        # Regular expression to match placeholder braces and normalize them
        # placeholder_re = re.compile(r"\{\s*([^{}\s].*?)\s*\}")
        
        for block in iter_block_items(document):
            if isinstance(block, Paragraph):
                paragraph_text = get_paragraph_text_with_fields(block).strip()
                if paragraph_text:
                    # Replace placeholder braces with just the placeholder name
                    # paragraph_text = placeholder_re.sub(r'\1', paragraph_text)
                    all_text.append(paragraph_text)
                    
            elif isinstance(block, Table):
                for row in block.rows:
                    row_text = []
                    for cell in row.cells:
                        cell_text = ""
                        for paragraph in cell.paragraphs:
                            cell_text += get_paragraph_text_with_fields(paragraph)
                        # Replace placeholder braces with just the placeholder name
                        # cell_text = placeholder_re.sub(r'\1', cell_text).strip()
                        row_text.append(cell_text)
                    
                    # Join cell text with a tab to represent table columns
                    full_row_text = "\t".join(row_text).strip()
                    if full_row_text:
                        all_text.append(full_row_text)
        
        extracted_text = '\n'.join(all_text)
        logger.info(f"Successfully extracted {len(extracted_text)} characters from {docx_path}")
        
        return extracted_text
        
    except Exception as e:
        logger.error(f"Failed to extract content from {docx_path}: {e}")
        raise


def _extract_with_zipfile_fallback(docx_path: str) -> str:
    """
    Fallback extraction method for DOCX files with corrupted custom XML.
    Uses direct XML parsing of the document.xml file.
    
    Args:
        docx_path: Path to the DOCX file
        
    Returns:
        Extracted text content
    """
    logger = get_logger(__name__)
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_file:
            # Try to read the main document content
            try:
                with zip_file.open('word/document.xml') as doc_xml:
                    tree = ET.parse(doc_xml)
                    root = tree.getroot()
                    
                    # Extract text preserving paragraph structure
                    paragraphs = []
                    
                    # DOCX uses the WordProcessingML namespace
                    namespaces = {
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                    }
                    
                    # Process each paragraph separately to preserve structure
                    for para_elem in root.findall('.//w:p', namespaces):
                        para_text_parts = []
                        
                        # Get all text runs within this paragraph
                        for text_elem in para_elem.findall('.//w:t', namespaces):
                            if text_elem.text:
                                para_text_parts.append(text_elem.text)
                        
                        # Join text runs within paragraph and add to paragraphs
                        if para_text_parts:
                            paragraph_text = ''.join(para_text_parts).strip()
                            if paragraph_text:  # Only add non-empty paragraphs
                                paragraphs.append(paragraph_text)
                    
                    # Also extract table content if present
                    for table_elem in root.findall('.//w:tbl', namespaces):
                        for row_elem in table_elem.findall('.//w:tr', namespaces):
                            cell_texts = []
                            for cell_elem in row_elem.findall('.//w:tc', namespaces):
                                cell_text_parts = []
                                for text_elem in cell_elem.findall('.//w:t', namespaces):
                                    if text_elem.text:
                                        cell_text_parts.append(text_elem.text)
                                if cell_text_parts:
                                    cell_texts.append(''.join(cell_text_parts).strip())
                            
                            if cell_texts:
                                # Join cells with tabs, add row to paragraphs
                                row_text = '\t'.join(cell_texts)
                                if row_text.strip():
                                    paragraphs.append(row_text)
                    
                    extracted_text = '\n'.join(paragraphs)
                    
                    if extracted_text:
                        logger.info(f"Successfully extracted {len(extracted_text)} characters using fallback method from {docx_path}")
                        return extracted_text
                    else:
                        logger.warning(f"No text content found in document.xml for {docx_path}")
                        return ""
                        
            except KeyError:
                logger.error(f"Could not find word/document.xml in {docx_path}")
                return ""
                
    except Exception as e:
        logger.error(f"Fallback extraction failed for {docx_path}: {e}")
        return ""
