import re
from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph


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


