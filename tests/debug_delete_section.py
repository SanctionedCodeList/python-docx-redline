"""Debug script to see XML structure after delete_section."""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from docx_redline import Document

# Create minimal contract document
doc_path = Path(tempfile.mktemp(suffix=".docx"))

document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Services</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Company A will provide consulting services.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

with zipfile.ZipFile(doc_path, "w") as docx:
    docx.writestr("word/document.xml", document_xml)
    docx.writestr("[Content_Types].xml", content_types)
    docx.writestr("_rels/.rels", rels)

# Load document and delete section
doc = Document(doc_path)
doc.delete_section("Services", track=True)

# Print the resulting XML
print("=" * 80)
print("RESULTING XML AFTER delete_section():")
print("=" * 80)
print(etree.tostring(doc.xml_root, encoding="unicode", pretty_print=True))

# Cleanup
doc_path.unlink()
