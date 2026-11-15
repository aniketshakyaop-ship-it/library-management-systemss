import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc = Document()
s = doc.sections[0]
footer = s.footer
p_f = footer.paragraphs[0]
p_f.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = p_f.add_run("Test Header | Page ")

fld = parse_xml(r'<w:fldSimple %s w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>' % nsdecls('w'))
p_f._p.append(fld)

r2 = p_f.add_run(" of ")
fld2 = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"><w:r><w:t>2</w:t></w:r></w:fldSimple>' % nsdecls('w'))
p_f._p.append(fld2)

doc.add_paragraph("Hello Page 1")
doc.add_page_break()
doc.add_paragraph("Hello Page 2")

doc.save(r"C:\Users\anike\.gemini\antigravity\scratch\library-management-system\test_fld.docx")
print("Saved test_fld.docx")
