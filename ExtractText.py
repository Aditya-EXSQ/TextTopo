from docx import Document
import zipfile
import xml.etree.ElementTree as ET
import re

def extract_placeholders(docx_path):
    """
    Extract placeholders from a .docx file, including:
      - MERGEFIELD names from field instructions
      - Content control metadata (tags/aliases) and any visible placeholder-like text
      - Double-curly placeholders like {{Name}}

    Returns a dict with keys: "merge_fields", "content_controls", "curly_brace_tokens".
    """
    merge_fields = set()
    content_controls = set()
    curly_tokens = set()
    data_binding_xpaths = set()
    data_binding_names = set()

    with zipfile.ZipFile(docx_path) as docx:
        # Look inside document.xml, headers, and footers
        parts = ["word/document.xml"]
        parts += [name for name in docx.namelist() if name.startswith("word/header")]
        parts += [name for name in docx.namelist() if name.startswith("word/footer")]

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        merge_re = re.compile(r"\\bMERGEFIELD\\b\s+\"?([^\"\\s]+)\"?", re.IGNORECASE)
        curly_re = re.compile(r"\{\{\s*([^{}\s].*?)\s*\}\}")

        for part in parts:
            try:
                xml_content = docx.read(part)
            except KeyError:
                continue
            root = ET.fromstring(xml_content)

            # 1) Field instructions: w:instrText and w:fldSimple/@w:instr
            #    Extract MERGEFIELD names
            concatenated_instr_text = []
            for instr in root.findall(".//w:instrText", ns):
                if instr.text:
                    text = instr.text
                    concatenated_instr_text.append(text)
                    for m in merge_re.finditer(text):
                        merge_fields.add(m.group(1).strip())

            for fld in root.findall(".//w:fldSimple", ns):
                instr = fld.attrib.get(f"{{{ns['w']}}}instr")
                if instr:
                    for m in merge_re.finditer(instr):
                        merge_fields.add(m.group(1).strip())

            # Also scan concatenated instruction text for split runs
            if concatenated_instr_text:
                full_instr = " ".join(concatenated_instr_text)
                for m in merge_re.finditer(full_instr):
                    merge_fields.add(m.group(1).strip())

            # 2) Content controls (w:sdt). Capture tag/alias as placeholders
            for sdt in root.findall(".//w:sdt", ns):
                sdt_pr = sdt.find(".//w:sdtPr", ns)
                if sdt_pr is not None:
                    tag = sdt_pr.find(".//w:tag", ns)
                    alias = sdt_pr.find(".//w:alias", ns)
                    databinding = sdt_pr.find(".//w:dataBinding", ns)
                    if tag is not None:
                        val = tag.attrib.get(f"{{{ns['w']}}}val")
                        if val:
                            content_controls.add(val.strip())
                    if alias is not None:
                        val = alias.attrib.get(f"{{{ns['w']}}}val")
                        if val:
                            content_controls.add(val.strip())
                    if databinding is not None:
                        xpath_val = databinding.attrib.get(f"{{{ns['w']}}}xpath")
                        if xpath_val:
                            xpath_val = xpath_val.strip()
                            data_binding_xpaths.add(xpath_val)
                            # Derive last step name(s) from XPath (strip predicates and namespaces)
                            # Example: /ns0:Root/ns0:Customer[1]/ns0:Name -> Name
                            try:
                                # Split on '/' and filter blanks
                                steps = [seg for seg in xpath_val.split('/') if seg]
                                if steps:
                                    last = steps[-1]
                                    # Remove predicates [ ... ]
                                    last = re.sub(r"\[.*?\]", "", last)
                                    # Drop namespace prefix if any
                                    last = last.split(":")[-1]
                                    if last:
                                        data_binding_names.add(last)
                            except Exception:
                                pass

                # Also look for visible placeholder-like text inside the sdt
                for t in sdt.findall(".//w:sdtContent//w:t", ns):
                    if t.text:
                        # {{ tokens inside content controls
                        for m in curly_re.finditer(t.text):
                            curly_tokens.add(m.group(1).strip())

            # 3) Double-curly placeholders anywhere in the part (scan all w:t)
            for t in root.findall(".//w:t", ns):
                if t.text:
                    for m in curly_re.finditer(t.text):
                        curly_tokens.add(m.group(1).strip())

    return {
        "merge_fields": sorted(merge_fields),
        "content_controls": sorted(content_controls),
        "curly_brace_tokens": sorted(curly_tokens),
        "data_binding_xpaths": sorted(data_binding_xpaths),
        "data_binding_names": sorted(data_binding_names),
    }

def extract_text_from_docx(file_path):
    # Load the document
    doc = Document(file_path)
    
    # Extract text from each paragraph
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    
    # Join with line breaks
    return "\n".join(text)

file_path = "Master Approval Letter.docx"

# Example usage (Python-DOCX)
# all_text = extract_text_from_docx(file_path)
# print(all_text)

# Example usage (Placeholder Content)
placeholders = extract_placeholders(file_path)
print("Extracted placeholders:")
print("- MERGEFIELD names:")
for name in placeholders["merge_fields"]:
    print(f"  {name}")
print("- Content control tags/aliases:")
for name in placeholders["content_controls"]:
    print(f"  {name}")
print("- Content control data-binding names (from XPath):")
for name in placeholders["data_binding_names"]:
    print(f"  {name}")
print("- Content control data-binding XPaths:")
for name in placeholders["data_binding_xpaths"]:
    print(f"  {name}")
print("- {{...}} tokens:")
for name in placeholders["curly_brace_tokens"]:
    print(f"  {name}")
