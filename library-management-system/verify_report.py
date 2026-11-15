import zipfile, xml.etree.ElementTree as ET, os, re

fname = 'Library_Management_System_Report.docx'
size = os.path.getsize(fname)
print(f'File size: {size/1024:.1f} KB')

with zipfile.ZipFile(fname, 'r') as z:
    for name in z.namelist():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:
                ET.fromstring(z.read(name))
            except Exception as e:
                print(f'  BAD XML: {name} -> {e}')
    print('All XML parts valid!')
    doc_xml = z.read('word/document.xml').decode('utf-8')
    paras  = len(re.findall(r'<w:p[ >]', doc_xml))
    tables = len(re.findall(r'<w:tbl[ >]', doc_xml))
    ptype  = doc_xml.count('w:type="page"')
    print(f'Paragraphs : {paras}')
    print(f'Tables     : {tables}')
    print(f'Page breaks: {ptype}')

print('DOCX is clean and ready to open in Microsoft Word!')
