# 📚 Library Management System (ACET Mini Project)

**Institution:** Aligarh College of Engineering and Technology (ACET), Aligarh  
**Department:** Department of Computer Science & Engineering  
**Project Type:** Mini Project  
**Student Name:** Gazal Gupta  
**University Roll No:** `2401090100037`  
**Project Guide / HOD:** Dr. Anand sir  

---

## 🌟 Overview

The **Library Management System** is a responsive web application built with **React** and **Tailwind CSS** designed to digitize and automate the operations of an academic library. It replaces tedious manual paper registers and cards with a real-time, interactive dashboard for students, faculty, and administrators.

---

## 🚀 Key Features

### 1. 📖 Book Catalog & Multi-Criteria Search
* Real-time search across title, author, ISBN, and topic tags.
* Category filtering by engineering departments:
  * Computer Science
  * Information Technology
  * Electronics & Communication
  * Mechanical Engineering
  * Civil Engineering
  * Mathematics
  * Management & General
* Stock filter: In-Stock vs. Out-of-Stock.
* Dynamic sorting by Title (A-Z), Top Rated, Most Copies, and Publication Year.
* Detailed Book Modal showing synopsis, shelf rack coordinates (e.g., `Rack CS-101`), publisher, edition, and total vs. available copies.

### 2. 🎒 Student Member Portal (Gazal Gupta - 2401090100037)
* Instant borrowing flow respecting institutional quota limits (max 3 books for students).
* Real-time loan countdown and due date alerts.
* Late fee / fine calculation (₹5/day for overdue days).
* One-click loan renewal (+7 days extension).
* Complete historical register of returned books and past transactions.

### 3. 🛡️ Administrator / Librarian Portal (Dr. Anand sir - HOD)
* **Add & Edit Books:** Add new textbooks with title, author, ISBN, total copies, rack location, edition, and tags.
* **Circulation Desk:** Issue books to any student/faculty with custom loan periods (7, 14, 21, or 30 days).
* **Return Processing:** Instant return processing with automated fine computation and printable return slips.
* **Member Management:** Enrol new students/faculty, view active borrow counts, and monitor quota limits.
* **Audit & Reports:**
  * Key institutional metrics: Total Distinct Titles, Total Volumes, Books Issued, Overdue Alerts, and Fines Collected.
  * Category distribution breakdown charts.
  * Leaderboard of most frequently borrowed titles.
  * One-click CSV export of the entire library catalog.
  * Formatted print report mode.

### 4. 🔔 Smart Alerts & Notifications
* Floating notification bell showing real-time count of due soon (<= 3 days) and overdue books.
* Interactive toast alerts for every action (borrow, return, edit, error).

### 5. 💾 Offline-Ready & Persistent State
* Uses browser `LocalStorage` so books added, loans created, and returns processed are preserved across page reloads.
* Includes a "Reset Data" option to easily return to the initial demo state anytime.

---

## 📁 Project Structure

```
library-management-system/
│
├── index.html                      # Complete, standalone, zero-dependency web app
├── run_app.bat                     # Windows 1-click launcher
├── package.json                    # Modern Vite + React + Tailwind configuration
├── vite.config.js                  # Vite configuration file
├── tailwind.config.js              # Tailwind CSS theme & plugin setup
├── postcss.config.js               # PostCSS configuration
│
├── src/
│   ├── main.jsx                    # React application entrypoint
│   ├── App.jsx                     # Core application component with state & tabs
│   ├── index.css                   # Global Tailwind styles & scrollbars
│   └── data/
│       └── initialData.js          # Pre-loaded academic textbooks, members, & loans
│
├── Library_Management_System_Project_Report.docx  # Full academic project report (Word)
├── generate_report.py              # Script that generated the Word report
└── README.md                       # Project documentation & presentation guide
```

---

## 💻 How to Run the Application

### Option A: Instant Zero-Setup Run (Recommended)
1. Double-click the file **`run_app.bat`** (or simply double-click **`index.html`**).
2. It will open immediately in your default web browser (Google Chrome, Microsoft Edge, etc.).
3. **No Node.js or `npm install` is required** for this mode!

### Option B: Run via Vite / Node.js
If you wish to run it in a developer server environment:
```bash
# 1. Navigate to the project directory
cd library-management-system

# 2. Install dependencies (requires internet on first run)
npm install

# 3. Start development server
npm run dev
```

---

## 🎯 How to Demonstrate to HOD (Dr. Anand sir)

1. **Open the Application:** Double-click `run_app.bat`.
2. **Show the Top Banner:** Point out your name (**Gazal Gupta**), Roll Number (**2401090100037**), and HOD (**Dr. Anand sir**).
3. **Demonstrate Student View:**
   * Search for *"Algorithms"* or select *"Computer Science"* from the department filter.
   * Click **"View Details"** on *Introduction to Algorithms (CLRS)* to show shelf coordinates (`Rack CS-101`), ISBN, and synopsis.
   * Switch to **"🎒 My Borrowed Books"** to showcase:
     * Active loan countdown.
     * The overdue alert on *Operating System Concepts* (3 days late -> ₹15 fine at ₹5/day).
     * Click **"Return Book"** to show the generated return receipt slip.
4. **Switch to Admin View (Dr. Anand sir):**
   * Click the **"Admin"** button on the top right.
   * Show the **"Issue & Return"** transaction register.
   * Click **"⚡ Quick Issue Book"** to demonstrate issuing a book to a student with duration selection.
   * Go to **"📊 Reports & Analytics"** to display:
     * Total Inventory & Active Borrow metrics.
     * Category distribution progress charts.
     * Top 5 most borrowed titles leaderboard.
     * Click **"📥 Download Catalog CSV"** to demonstrate automated data export.
5. **Submit Project Report:**
   * Hand over or submit the generated `Library_Management_System_Project_Report.docx`.

---

## 📜 Academic Details
* **Student:** Gazal Gupta (2401090100037)
* **College:** Aligarh College of Engineering and Technology (ACET)
* **HOD / Guide:** Dr. Anand sir
* **Tech Stack:** HTML5, CSS3, JavaScript (ES6+), React 18, Tailwind CSS
