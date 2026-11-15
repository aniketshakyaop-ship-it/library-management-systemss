import docx

doc = docx.Document(r"C:\Users\anike\.gemini\antigravity\scratch\library-management-system\Library_Management_System_50_Pages_Project_Report.docx")
xml = doc._element.xml
breaks = xml.count('w:type="page"')
print("Total Explicit Page Breaks:", breaks)
print("Total Target Pages:", breaks + 1)

# Check total paragraphs and tables
print("Total Paragraphs:", len(doc.paragraphs))
print("Total Tables:", len(doc.tables))
print("File Size (KB):", len(xml) // 1024)
