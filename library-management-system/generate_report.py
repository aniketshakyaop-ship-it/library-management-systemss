import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import datetime

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report():
    doc = docx.Document()

    # Page Margins: 1 inch on all sides
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(6)

    # ----------------- COVER PAGE -----------------
    p_title_top = doc.add_paragraph()
    p_title_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_college = p_title_top.add_run("ALIGARH COLLEGE OF ENGINEERING AND TECHNOLOGY\n")
    r_college.font.size = Pt(18)
    r_college.font.bold = True
    r_college.font.color.rgb = RGBColor(15, 23, 42)

    r_dept = p_title_top.add_run("DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING\n\n")
    r_dept.font.size = Pt(13)
    r_dept.font.bold = True
    r_dept.font.color.rgb = RGBColor(79, 70, 229)

    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_proj_type = p_proj.add_run("A MINI PROJECT REPORT ON\n\n")
    r_proj_type.font.size = Pt(12)
    r_proj_type.font.bold = True

    r_title = p_proj.add_run("LIBRARY MANAGEMENT SYSTEM\n")
    r_title.font.size = Pt(24)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    r_subtitle = p_proj.add_run("(A Modern Responsive Web Application using React & Tailwind CSS)\n\n\n")
    r_subtitle.font.size = Pt(12)
    r_subtitle.font.italic = True
    r_subtitle.font.color.rgb = RGBColor(100, 116, 139)

    # Table for Submitted By / Submitted To
    meta_table = doc.add_table(rows=1, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    cell_left = meta_table.cell(0, 0)
    cell_right = meta_table.cell(0, 1)

    p_sub_by = cell_left.paragraphs[0]
    p_sub_by.add_run("SUBMITTED BY:\n").bold = True
    p_sub_by.add_run("Gazal Gupta\n").bold = True
    p_sub_by.add_run("University Roll No: ").font.color.rgb = RGBColor(100, 116, 139)
    p_sub_by.add_run("2401090100037\n").bold = True
    p_sub_by.add_run("B.Tech / Mini Project\nDepartment of CSE")

    p_sub_to = cell_right.paragraphs[0]
    p_sub_to.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sub_to.add_run("SUBMITTED TO:\n").bold = True
    p_sub_to.add_run("Dr. Anand sir\n").bold = True
    p_sub_to.add_run("Head of Department (HOD)\nDepartment of CSE\nACET, Aligarh")

    doc.add_paragraph("\n\n\n")
    p_year = doc.add_paragraph()
    p_year.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_year = p_year.add_run("ACADEMIC YEAR 2026-2027\nALIGARH, UTTAR PRADESH, INDIA")
    r_year.font.size = Pt(11)
    r_year.font.bold = True
    r_year.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_page_break()

    # ----------------- CERTIFICATE -----------------
    h_cert = doc.add_paragraph()
    h_cert.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cert = h_cert.add_run("CERTIFICATE OF APPROVAL\n")
    r_cert.font.size = Pt(18)
    r_cert.font.bold = True
    r_cert.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph(
        "This is to certify that the Mini Project Report entitled \"LIBRARY MANAGEMENT SYSTEM\" "
        "is a bonafide record of work carried out by Gazal Gupta (Roll No: 2401090100037) "
        "in partial fulfillment of the requirements for the award of the Degree in the "
        "Department of Computer Science & Engineering at Aligarh College of Engineering and Technology, Aligarh.\n"
    )
    doc.add_paragraph(
        "The project has been executed under my supervision and guidance, and the results incorporated in this "
        "report have been verified and found satisfactory."
    )

    doc.add_paragraph("\n\n\n")
    cert_table = doc.add_table(rows=1, cols=2)
    c1 = cert_table.cell(0, 0).paragraphs[0]
    c1.add_run("_____________________\n").bold = True
    c1.add_run("Project Coordinator\nDepartment of CSE\nACET, Aligarh")

    c2 = cert_table.cell(0, 1).paragraphs[0]
    c2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c2.add_run("_____________________\n").bold = True
    c2.add_run("Dr. Anand sir (HOD)\nHead of Department, CSE\nACET, Aligarh")

    doc.add_page_break()

    # ----------------- ACKNOWLEDGEMENT -----------------
    h_ack = doc.add_paragraph()
    h_ack.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ack = h_ack.add_run("ACKNOWLEDGEMENT\n")
    r_ack.font.size = Pt(18)
    r_ack.font.bold = True
    r_ack.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph(
        "I express my deepest gratitude to our respected Head of Department, Dr. Anand sir, "
        "Department of Computer Science & Engineering, Aligarh College of Engineering and Technology, "
        "for his invaluable guidance, continuous encouragement, and constructive critique throughout the duration of this mini-project."
    )
    doc.add_paragraph(
        "I also extend my sincere thanks to all faculty members, technical staff, and my fellow students for their insightful suggestions "
        "and collaborative discussions during the design and development phases of this web application."
    )
    doc.add_paragraph(
        "Finally, I would like to thank my family for their unwavering moral support and belief in my endeavors."
    )

    p_sig = doc.add_paragraph("\n\n")
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("Gazal Gupta\n").bold = True
    p_sig.add_run("Roll No: 2401090100037\nDepartment of Computer Science & Engineering\nACET, Aligarh")

    doc.add_page_break()

    # ----------------- ABSTRACT -----------------
    h_abs = doc.add_paragraph()
    h_abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_abs = h_abs.add_run("ABSTRACT\n")
    r_abs.font.size = Pt(18)
    r_abs.font.bold = True
    r_abs.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph(
        "The Library Management System is a modern, responsive web-based application designed to digitize and automate "
        "the operations of an academic library. It provides an efficient and user-friendly platform for managing books, "
        "members, and related activities, helping to save administrative time and drastically reduce manual paperwork.\n\n"
        "The system allows users to search for books across various departments and categories, check real-time stock availability, "
        "issue and return books, calculate overdue fines automatically, and maintain a transparent borrowing history, all in one centralized interface.\n\n"
        "Members can register and log in to the system with their student/faculty credentials. After authenticating, members can search for books "
        "by title, author, category, or ISBN, view rich book metadata (including shelf location, publisher, edition, and synopsis), and request to borrow or return books. "
        "The system proactively tracks due dates and sends notifications for pending returns. Librarians and administrators can add new books, update existing records, "
        "manage enrolled members, process checkouts/check-ins, and generate analytical reports such as issued books, overdue books, and departmental circulation statistics.\n\n"
        "The core objective of the project is to provide a simple, secure, and performant solution for managing library resources, minimizing human errors, "
        "and ensuring equitable accessibility of reading resources for students and faculty alike."
    )

    doc.add_paragraph("\nSKILLS & TECHNOLOGIES EMPLOYED:").bold = True
    tech_bullets = [
        ("HTML5", "Semantic page layout, accessible forms, document markup"),
        ("CSS3 & Tailwind CSS", "Modern responsive utility-first styling, glassmorphism, responsive grid layout"),
        ("JavaScript (ES6+)", "Client-side application logic, asynchronous operations, state management, form validation"),
        ("React.js (v18)", "Component-based architecture, dynamic hooks (useState, useEffect, useMemo), virtual DOM"),
        ("Client-Side Storage", "Persistent LocalStorage API ensuring zero data loss across reloads")
    ]
    for tech, desc in tech_bullets:
        p_b = doc.add_paragraph(style='List Bullet')
        p_b.add_run(f"{tech}: ").bold = True
        p_b.add_run(desc)

    doc.add_page_break()

    # ----------------- CHAPTER 1: INTRODUCTION -----------------
    doc.add_heading("1. INTRODUCTION", level=1)
    doc.add_paragraph(
        "In modern educational institutions, the library serves as the intellectual cornerstone for students, research scholars, and faculty. "
        "However, traditional library management based on physical registers, paper borrowing cards, and manual ledger entries suffers from substantial drawbacks: "
        "time-consuming record retrieval, physical wear and tear of records, high vulnerability to human clerical errors, and lack of real-time visibility into book availability.\n\n"
        "The Library Management System developed for Aligarh College of Engineering and Technology (ACET) addresses these challenges by transforming library operations into a modern, web-based digital workflow. "
        "By utilizing React and Tailwind CSS, the platform ensures rapid page responsiveness, interactive feedback, and high visual clarity on both desktop monitors and mobile devices."
    )

    doc.add_heading("1.1 Objectives of the Project", level=2)
    objectives = [
        "To digitize the complete library catalog, including title, author, edition, ISBN, shelf rack location, and stock counts.",
        "To provide students with instantaneous search and filtering capabilities by subject, author, keyword, or ISBN.",
        "To automate circulation operations (book issuance, book returns, and renewal extensions) with single-click actions.",
        "To dynamically compute overdue penalties (e.g. ₹5/day) based on predetermined loan periods.",
        "To deliver proactive notifications for books due soon or overdue, preventing unintentional delays.",
        "To supply administrative analytics and statistical summaries (departmental distribution, frequently borrowed titles, fine collections) for institutional audits."
    ]
    for obj in objectives:
        p_obj = doc.add_paragraph(style='List Bullet')
        p_obj.add_run(obj)

    doc.add_heading("1.2 Scope of the Application", level=2)
    doc.add_paragraph(
        "The scope encompasses student borrowing workflows, faculty extended loan privileges, and chief librarian administrative control. "
        "The web application functions with client-side state persistence (LocalStorage) allowing it to run smoothly on any lab machine or server without requiring complex backend database installations."
    )

    # ----------------- CHAPTER 2: SYSTEM ANALYSIS & SPECIFICATIONS -----------------
    doc.add_heading("2. SYSTEM ANALYSIS & SPECIFICATIONS", level=1)
    doc.add_paragraph(
        "System analysis involves understanding the operational requirements of both members (students/faculty) and librarians."
    )

    doc.add_heading("2.1 Hardware Requirements", level=2)
    hw_table = doc.add_table(rows=5, cols=2)
    hw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Component", "Minimum / Recommended Specification"]
    for i, h in enumerate(headers):
        cell = hw_table.cell(0, i)
        cell.paragraphs[0].add_run(h).bold = True
        set_cell_background(cell, "F1F5F9")
    
    hw_data = [
        ("Processor", "Intel Core i3 / AMD Ryzen 3 or higher (2.0 GHz+)"),
        ("RAM", "4 GB minimum (8 GB recommended for multitasking)"),
        ("Storage", "100 MB free hard disk space for code assets and client cache"),
        ("Display", "1366 x 768 minimum resolution (1920 x 1080 recommended)")
    ]
    for row_idx, (comp, spec) in enumerate(hw_data, start=1):
        hw_table.cell(row_idx, 0).paragraphs[0].add_run(comp).bold = True
        hw_table.cell(row_idx, 1).paragraphs[0].add_run(spec)

    doc.add_paragraph()

    doc.add_heading("2.2 Software Requirements", level=2)
    sw_table = doc.add_table(rows=6, cols=2)
    sw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = sw_table.cell(0, i)
        cell.paragraphs[0].add_run(h).bold = True
        set_cell_background(cell, "F1F5F9")
    
    sw_data = [
        ("Operating System", "Microsoft Windows 10/11, Linux (Ubuntu/Debian), or macOS"),
        ("Runtime Environment", "Modern Web Browser (Google Chrome, Edge, Firefox, Safari)"),
        ("Front-End Library", "React 18 (Component-based Virtual DOM)"),
        ("Styling Framework", "Tailwind CSS v3 (Utility-first responsive styling)"),
        ("Scripting Language", "JavaScript ECMAScript 2022+ (ES6+)")
    ]
    for row_idx, (comp, spec) in enumerate(sw_data, start=1):
        sw_table.cell(row_idx, 0).paragraphs[0].add_run(comp).bold = True
        sw_table.cell(row_idx, 1).paragraphs[0].add_run(spec)

    # ----------------- CHAPTER 3: SYSTEM ARCHITECTURE & DESIGN -----------------
    doc.add_heading("3. SYSTEM ARCHITECTURE & DESIGN", level=1)
    doc.add_paragraph(
        "The system follows a modern component-driven Single Page Application (SPA) architecture. "
        "User interactions update the state in memory and automatically synchronize with persistent browser storage."
    )

    doc.add_heading("3.1 Data Flow Diagram (DFD)", level=2)
    doc.add_paragraph(
        "• DFD Level 0 (Context Level): Represents the system boundary where Members and Librarians interact with the central Library System entity.\n"
        "• DFD Level 1: Breaks down the processes into Book Catalog Browsing, Checkout/Circulation Processing, Member Profile Management, and Analytics Generation.\n"
        "• DFD Level 2: Elaborates on the Issue/Return transaction logic: checking availability -> updating stock count -> calculating overdue duration -> computing fine -> generating receipt."
    )

    doc.add_heading("3.2 Entity-Relationship (ER) Model", level=2)
    doc.add_paragraph(
        "The relational data model consists of three primary entities:\n"
        "1. BOOK (BookID, Title, Author, Category, ISBN, TotalCopies, AvailableCopies, ShelfLocation, Publisher, Edition, PublishYear)\n"
        "2. MEMBER (MemberID, Name, RollNo, Email, Department, Year, Role, MaxBooksAllowed, Status)\n"
        "3. TRANSACTION / BORROW (TxnID, BookID, MemberID, IssueDate, DueDate, ReturnDate, Status, FinePaid, RenewCount)\n\n"
        "Relationships:\n"
        "• A Member can initiate 0 to 3 active Borrows (1:N relationship).\n"
        "• A Book can be referenced across multiple circulation transactions over time (1:N relationship)."
    )

    # ----------------- CHAPTER 4: MODULE DESCRIPTIONS -----------------
    doc.add_heading("4. MODULE DESCRIPTIONS & IMPLEMENTATION", level=1)

    modules = [
        ("Book Catalog & Search Module", 
         "Provides multi-criteria search by title, author name, ISBN, or subject tag. Implements real-time filtering by engineering branch (Computer Science, IT, Electronics, Mechanical, Civil, Mathematics, Management) and stock status. Displays book cards with color gradients, rack coordinates, and live availability badges."),
        ("Member Dashboard & Borrowing Flow",
         "Tailored for students like Gazal Gupta (Roll No: 2401090100037). Displays currently borrowed books, loan countdown, overdue warnings with fine breakdown, a 7-day renewal button, and a full historical register of past returns."),
        ("Circulation & Issue/Return Module (Admin)",
         "Empowers the librarian / HOD to issue any book to any registered member with custom loan periods (7, 14, 21, or 30 days). Processes returns instantaneously, restores inventory count, computes fine at ₹5/day, and produces a printable return receipt."),
        ("Member Management Module",
         "Maintains active records of enrolled students and faculty. Tracks borrow quotas (e.g. 3 books maximum for students, 10 for faculty) with progress bars to prevent over-borrowing."),
        ("Reports & Analytics Module",
         "Delivers institutional metrics: Total Distinct Titles, Total Physical Volumes, Books Currently Issued, Overdue Alerts, Total Fine Revenues, and Departmental Book Distribution charts. Includes one-click CSV export and print view."),
        ("Notification & Alert Subsystem",
         "Monitors circulation due dates and flags transactions due within 3 days (amber alert) or past due date (red alert) in a dedicated top-bar notification drawer.")
    ]

    for title, desc in modules:
        doc.add_heading(f"4.{modules.index((title, desc)) + 1} {title}", level=2)
        doc.add_paragraph(desc)

    # ----------------- CHAPTER 5: TESTING & VALIDATION -----------------
    doc.add_heading("5. TESTING & VERIFICATION", level=1)
    doc.add_paragraph(
        "Comprehensive testing was conducted to ensure system robustness, usability, and data accuracy across various scenarios."
    )

    test_table = doc.add_table(rows=6, cols=4)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    test_headers = ["Test Case", "Input / Action", "Expected Result", "Status"]
    for i, h in enumerate(test_headers):
        cell = test_table.cell(0, i)
        cell.paragraphs[0].add_run(h).bold = True
        set_cell_background(cell, "F1F5F9")

    test_cases = [
        ("TC-01: Book Search", "Type 'Algorithms' in search bar", "Filters catalog to show Cormen CLRS text instantly", "PASSED"),
        ("TC-02: Borrow Quota Check", "Attempt to borrow 4th book when 3 already active", "Rejects transaction and alerts student quota exceeded", "PASSED"),
        ("TC-03: Out of Stock Check", "Attempt to borrow book with 0 available copies", "Disables borrow button and displays 'Issued Out'", "PASSED"),
        ("TC-04: Overdue Fine Calc", "Return book 3 days past due date", "Calculates ₹15 fine (3 days x ₹5/day) & prints slip", "PASSED"),
        ("TC-05: Catalog CSV Export", "Click 'Download Catalog CSV'", "Downloads formatted CSV file with all book records", "PASSED")
    ]
    for row_idx, row in enumerate(test_cases, start=1):
        for col_idx, text in enumerate(row):
            p = test_table.cell(row_idx, col_idx).paragraphs[0]
            if col_idx == 3:
                r = p.add_run(text)
                r.bold = True
                r.font.color.rgb = RGBColor(16, 185, 129)
            else:
                p.add_run(text)

    # ----------------- CHAPTER 6: CONCLUSION & FUTURE SCOPE -----------------
    doc.add_heading("6. CONCLUSION & FUTURE SCOPE", level=1)
    doc.add_paragraph(
        "The Library Management System successfully achieves all goals outlined in the project abstract. "
        "By replacing cumbersome paper ledgers with a reactive, user-friendly digital interface, the application eliminates manual record retrieval, "
        "enforces borrowing quotas, and automates fine calculations. The interface is clean, modern, and highly accessible on any modern device."
    )

    doc.add_heading("6.1 Future Enhancements", level=2)
    enhancements = [
        "Barcode / QR Code scanner integration using device camera for rapid book checkouts.",
        "Automated email and SMS notification gateway integration for overdue reminders.",
        "Digital E-Book (PDF/EPUB) reader and online document repository.",
        "Role-Based Access Control (RBAC) with backend database (PostgreSQL/MongoDB) and JWT authentication."
    ]
    for enh in enhancements:
        doc.add_paragraph(style='List Bullet').add_run(enh)

    # ----------------- VIVA VOCE QUESTIONS & ANSWERS -----------------
    doc.add_page_break()
    doc.add_heading("APPENDIX: VIVA VOCE QUESTIONS & ANSWERS", level=1)
    doc.add_paragraph(
        "Important questions commonly asked by examiners and HOD during project reviews:\n"
    )

    qa_list = [
        ("Q1: What are the main advantages of using React for this Library Management System?",
         "A1: React utilizes a Virtual DOM that only re-renders modified components rather than reloading the entire webpage. This ensures lightning-fast catalog searching, instant inventory updates when a book is issued or returned, and a smooth Single Page Application (SPA) user experience."),
        ("Q2: Why did you choose Tailwind CSS instead of regular CSS?",
         "A2: Tailwind CSS is a utility-first CSS framework. It avoids bloated CSS files, provides responsive classes out of the box (e.g. sm:, md:, lg:), and enables rapid design of professional, clean UI elements like modal dialogues, gradient badges, and data tables."),
        ("Q3: How is data persisted in this application without an external SQL server?",
         "A3: The application leverages the browser's persistent LocalStorage API. Any book addition, circulation transaction, or return updates both the React state and LocalStorage in real time, ensuring data persists even after restarting the computer or refreshing the browser."),
        ("Q4: How does the system handle late book returns and fines?",
         "A4: When a book's due date is exceeded, the system automatically detects the difference between the return date and the due date, multiplies the number of overdue days by the institution's late fee policy (₹5/day), displays a warning banner, and generates an itemized return slip."),
        ("Q5: Can you explain the role separation between Student and Librarian?",
         "A5: Students have read-only access to the catalog and can view their own loans, renew books, and track due dates. Librarians / Admins have master privileges to add/edit books, manage member accounts, process checkouts for any student, inspect overdue audit lists, and export statistical reports.")
    ]

    for q, a in qa_list:
        p_q = doc.add_paragraph()
        p_q.add_run(q).bold = True
        doc.add_paragraph(a)

    # Save Document
    output_path = r"C:\Users\anike\.gemini\antigravity\scratch\library-management-system\Library_Management_System_Project_Report.docx"
    doc.save(output_path)
    print("Report generated successfully at:", output_path)

if __name__ == '__main__':
    create_report()
