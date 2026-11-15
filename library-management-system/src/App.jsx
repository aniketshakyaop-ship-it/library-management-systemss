import React, { useState, useEffect, useMemo } from 'react';
import { INITIAL_BOOKS, INITIAL_MEMBERS, INITIAL_BORROWS, CATEGORIES } from './data/initialData';

// Helper: fine calculation
export const calculateFine = (dueDateStr, returnDateStr = null) => {
  const due = new Date(dueDateStr);
  const referenceDate = returnDateStr ? new Date(returnDateStr) : new Date();
  due.setHours(0, 0, 0, 0);
  referenceDate.setHours(0, 0, 0, 0);
  const diffTime = referenceDate - due;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  if (diffDays > 0) {
    return { overdueDays: diffDays, fine: diffDays * 5 }; // ₹5 per day
  }
  return { overdueDays: 0, fine: 0 };
};

// Helper: borrow status
export const getBorrowStatus = (borrow) => {
  if (borrow.returnDate) return 'Returned';
  const due = new Date(borrow.dueDate);
  const today = new Date();
  due.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((due - today) / (1000 * 60 * 60 * 24));
  if (diffDays < 0) return 'Overdue';
  if (diffDays <= 3) return 'Due Soon';
  return 'Active';
};

export default function App() {
  const [books, setBooks] = useState(() => {
    const saved = localStorage.getItem('acet_lms_books');
    return saved ? JSON.parse(saved) : INITIAL_BOOKS;
  });

  const [members, setMembers] = useState(() => {
    const saved = localStorage.getItem('acet_lms_members');
    return saved ? JSON.parse(saved) : INITIAL_MEMBERS;
  });

  const [borrows, setBorrows] = useState(() => {
    const saved = localStorage.getItem('acet_lms_borrows');
    return saved ? JSON.parse(saved) : INITIAL_BORROWS;
  });

  const [persona, setPersona] = useState('member'); // 'member' (Gazal Gupta) or 'admin' (Dr. Anand / Librarian)
  const currentMemberId = 'M001'; // Gazal Gupta

  const [activeTab, setActiveTab] = useState('catalog');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [availabilityFilter, setAvailabilityFilter] = useState('all');
  const [sortBy, setSortBy] = useState('title');

  // Modals
  const [selectedBookModal, setSelectedBookModal] = useState(null);
  const [showAddEditBookModal, setShowAddEditBookModal] = useState(false);
  const [editingBook, setEditingBook] = useState(null);
  const [showIssueModal, setShowIssueModal] = useState(false);
  const [preselectedBookForIssue, setPreselectedBookForIssue] = useState(null);
  const [returnReceiptModal, setReturnReceiptModal] = useState(null);
  const [showAddMemberModal, setShowAddMemberModal] = useState(false);
  const [showNotificationsDropdown, setShowNotificationsDropdown] = useState(false);

  // Toasts
  const [toasts, setToasts] = useState([]);
  const showToast = (message, type = 'success') => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 4000);
  };

  useEffect(() => {
    localStorage.setItem('acet_lms_books', JSON.stringify(books));
  }, [books]);

  useEffect(() => {
    localStorage.setItem('acet_lms_members', JSON.stringify(members));
  }, [members]);

  useEffect(() => {
    localStorage.setItem('acet_lms_borrows', JSON.stringify(borrows));
  }, [borrows]);

  const liveBorrows = useMemo(() => {
    return borrows.map(b => {
      const status = getBorrowStatus(b);
      const fineInfo = calculateFine(b.dueDate, b.returnDate);
      return {
        ...b,
        status,
        currentFine: b.returnDate ? b.finePaid : fineInfo.fine,
        overdueDays: fineInfo.overdueDays
      };
    });
  }, [borrows]);

  const pendingNotifications = useMemo(() => {
    if (persona === 'member') {
      return liveBorrows.filter(b => b.memberId === currentMemberId && !b.returnDate && (b.status === 'Overdue' || b.status === 'Due Soon'));
    } else {
      return liveBorrows.filter(b => !b.returnDate && (b.status === 'Overdue' || b.status === 'Due Soon'));
    }
  }, [liveBorrows, persona]);

  const filteredBooks = useMemo(() => {
    return books.filter(b => {
      const matchesSearch = b.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
        b.author.toLowerCase().includes(searchQuery.toLowerCase()) ||
        b.isbn.includes(searchQuery) ||
        (b.tags && b.tags.some(t => t.toLowerCase().includes(searchQuery.toLowerCase())));

      const matchesCategory = selectedCategory === 'All' || b.category === selectedCategory;

      const matchesAvailability = availabilityFilter === 'all'
        ? true
        : availabilityFilter === 'available'
          ? b.availableCopies > 0
          : b.availableCopies === 0;

      return matchesSearch && matchesCategory && matchesAvailability;
    }).sort((a, b) => {
      if (sortBy === 'title') return a.title.localeCompare(b.title);
      if (sortBy === 'rating') return b.rating - a.rating;
      if (sortBy === 'copies') return b.availableCopies - a.availableCopies;
      if (sortBy === 'year') return b.publishYear - a.publishYear;
      return 0;
    });
  }, [books, searchQuery, selectedCategory, availabilityFilter, sortBy]);

  const stats = useMemo(() => {
    const totalTitles = books.length;
    const totalCopies = books.reduce((acc, b) => acc + Number(b.totalCopies), 0);
    const availableCopies = books.reduce((acc, b) => acc + Number(b.availableCopies), 0);
    const activeBorrows = liveBorrows.filter(b => !b.returnDate).length;
    const overdueBorrows = liveBorrows.filter(b => b.status === 'Overdue' && !b.returnDate).length;
    const totalFinesCollected = liveBorrows.reduce((acc, b) => acc + (b.finePaid || 0), 0);
    const pendingFines = liveBorrows.filter(b => !b.returnDate).reduce((acc, b) => acc + (b.currentFine || 0), 0);
    const totalMembers = members.length;

    const categoryCounts = {};
    books.forEach(b => {
      categoryCounts[b.category] = (categoryCounts[b.category] || 0) + b.totalCopies;
    });

    const borrowCounts = {};
    borrows.forEach(b => {
      borrowCounts[b.bookId] = (borrowCounts[b.bookId] || 0) + 1;
    });

    return {
      totalTitles,
      totalCopies,
      availableCopies,
      activeBorrows,
      overdueBorrows,
      totalFinesCollected,
      pendingFines,
      totalMembers,
      categoryCounts,
      borrowCounts
    };
  }, [books, liveBorrows, members, borrows]);

  const handleIssueBook = ({ bookId, memberId, loanDays = 14 }) => {
    const targetBook = books.find(b => b.id === bookId);
    const targetMember = members.find(m => m.id === memberId);

    if (!targetBook || targetBook.availableCopies <= 0) {
      showToast('Selected book has no available copies in stock!', 'error');
      return;
    }

    const currentActiveCount = liveBorrows.filter(b => b.memberId === memberId && !b.returnDate).length;
    if (currentActiveCount >= targetMember.maxBooksAllowed) {
      showToast(`Borrow limit reached for ${targetMember.name} (Max ${targetMember.maxBooksAllowed} books).`, 'error');
      return;
    }

    const issueDate = new Date().toISOString().split('T')[0];
    const due = new Date();
    due.setDate(due.getDate() + Number(loanDays));
    const dueDate = due.toISOString().split('T')[0];

    const newBorrow = {
      id: `TXN-${Math.floor(1000 + Math.random() * 9000)}`,
      bookId: targetBook.id,
      bookTitle: targetBook.title,
      memberId: targetMember.id,
      memberName: targetMember.name,
      memberRoll: targetMember.rollNo,
      issueDate,
      dueDate,
      returnDate: null,
      status: 'Active',
      finePaid: 0,
      renewCount: 0
    };

    setBooks(prev => prev.map(b => b.id === bookId ? { ...b, availableCopies: b.availableCopies - 1 } : b));
    setBorrows(prev => [newBorrow, ...prev]);
    setShowIssueModal(false);
    setPreselectedBookForIssue(null);
    showToast(`Book "${targetBook.title}" issued to ${targetMember.name}!`, 'success');
  };

  const handleReturnBook = (borrowId) => {
    const borrow = liveBorrows.find(b => b.id === borrowId);
    if (!borrow || borrow.returnDate) return;

    const returnDate = new Date().toISOString().split('T')[0];
    const fineInfo = calculateFine(borrow.dueDate, returnDate);

    setBorrows(prev => prev.map(b => {
      if (b.id === borrowId) {
        return {
          ...b,
          returnDate,
          status: 'Returned',
          finePaid: fineInfo.fine
        };
      }
      return b;
    }));

    setBooks(prev => prev.map(b => {
      if (b.id === borrow.bookId) {
        return { ...b, availableCopies: Math.min(b.totalCopies, b.availableCopies + 1) };
      }
      return b;
    }));

    const receiptData = {
      ...borrow,
      returnDate,
      finePaid: fineInfo.fine,
      overdueDays: fineInfo.overdueDays
    };

    setReturnReceiptModal(receiptData);
    showToast(`"${borrow.bookTitle}" returned successfully!`, 'success');
  };

  const handleRenewBook = (borrowId) => {
    setBorrows(prev => prev.map(b => {
      if (b.id === borrowId) {
        const currentDue = new Date(b.dueDate);
        currentDue.setDate(currentDue.getDate() + 7);
        return {
          ...b,
          dueDate: currentDue.toISOString().split('T')[0],
          renewCount: (b.renewCount || 0) + 1
        };
      }
      return b;
    }));
    showToast('Due date extended by 7 days!', 'success');
  };

  const handleSaveBook = (bookData) => {
    if (editingBook) {
      setBooks(prev => prev.map(b => b.id === editingBook.id ? { ...bookData, id: b.id } : b));
      showToast(`Updated "${bookData.title}".`, 'success');
    } else {
      const newId = `B${String(books.length + 1).padStart(3, '0')}`;
      const newBook = {
        ...bookData,
        id: newId,
        availableCopies: Number(bookData.totalCopies),
        rating: 4.8,
        coverGradient: 'from-indigo-600 to-purple-800',
        tags: bookData.tags || [bookData.category, 'Textbook']
      };
      setBooks(prev => [newBook, ...prev]);
      showToast(`Added "${newBook.title}" to catalog!`, 'success');
    }
    setShowAddEditBookModal(false);
    setEditingBook(null);
  };

  const handleAddMember = (memberData) => {
    const newId = `M${String(members.length + 1).padStart(3, '0')}`;
    const initials = memberData.name.split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();
    const newMember = {
      ...memberData,
      id: newId,
      joinedDate: new Date().toISOString().split('T')[0],
      status: 'Active',
      avatar: initials || 'MB'
    };
    setMembers(prev => [...prev, newMember]);
    setShowAddMemberModal(false);
    showToast(`Registered member ${newMember.name}!`, 'success');
  };

  const handleExportCSV = () => {
    const headers = ['Book ID', 'Title', 'Author', 'Category', 'ISBN', 'Total Copies', 'Available Copies', 'Shelf Location', 'Publisher', 'Year'];
    const rows = books.map(b => [
      b.id,
      `"${b.title.replace(/"/g, '""')}"`,
      `"${b.author.replace(/"/g, '""')}"`,
      `"${b.category}"`,
      b.isbn,
      b.totalCopies,
      b.availableCopies,
      `"${b.shelfLocation}"`,
      `"${b.publisher}"`,
      b.publishYear
    ]);
    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `ACET_Library_Catalog_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('Catalog exported as CSV!', 'success');
  };

  const handleResetData = () => {
    if (window.confirm('Reset library data to default sample state?')) {
      localStorage.removeItem('acet_lms_books');
      localStorage.removeItem('acet_lms_members');
      localStorage.removeItem('acet_lms_borrows');
      setBooks(INITIAL_BOOKS);
      setMembers(INITIAL_MEMBERS);
      setBorrows(INITIAL_BORROWS);
      showToast('Reset data to initial state.', 'info');
    }
  };

  return (
    <div className="flex flex-col min-h-screen bg-slate-50 text-slate-800">
      {/* Toast notifications */}
      <div className="fixed top-5 right-5 z-50 flex flex-col gap-2 pointer-events-none">
        {toasts.map(t => (
          <div
            key={t.id}
            className={`pointer-events-auto px-4 py-3 rounded-xl shadow-xl text-sm font-medium flex items-center gap-2 ${
              t.type === 'success' ? 'bg-emerald-600 text-white' :
              t.type === 'error' ? 'bg-rose-600 text-white' : 'bg-indigo-600 text-white'
            }`}
          >
            <span>{t.type === 'success' ? '✓' : t.type === 'error' ? '✕' : 'ℹ'}</span>
            <span>{t.message}</span>
          </div>
        ))}
      </div>

      {/* College Mini-Project Notice Header */}
      <div className="bg-slate-900 text-slate-300 text-xs px-4 py-1.5 flex justify-between items-center border-b border-slate-800">
        <div className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
          <span className="font-semibold text-slate-100">Aligarh College of Engineering and Technology</span>
          <span className="hidden sm:inline text-slate-400">• Mini Project: Library Management System</span>
        </div>
        <div className="flex items-center gap-4 text-slate-400">
          <span>By: <strong className="text-indigo-400">Gazal Gupta (2401090100037)</strong></span>
          <span className="hidden md:inline">Guide: <strong className="text-amber-400">Dr. Anand sir (HOD)</strong></span>
        </div>
      </div>

      {/* Main Navbar */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex justify-between items-center h-16">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-brand-600 text-white flex items-center justify-center font-bold text-lg shadow-md shadow-brand-500/20">
              📚
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-900 tracking-tight leading-none">ACET Central Library</h1>
              <p className="text-xs text-slate-500 font-medium mt-0.5">Automated Digital Resource Center</p>
            </div>
          </div>

          <nav className="hidden md:flex items-center gap-1">
            <button
              onClick={() => setActiveTab('catalog')}
              className={`px-3.5 py-2 rounded-lg text-sm font-semibold transition ${
                activeTab === 'catalog' ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              📚 Book Catalog
            </button>
            {persona === 'member' && (
              <button
                onClick={() => setActiveTab('mybooks')}
                className={`px-3.5 py-2 rounded-lg text-sm font-semibold transition flex items-center gap-1.5 ${
                  activeTab === 'mybooks' ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <span>🎒 My Borrowed Books</span>
                {liveBorrows.filter(b => b.memberId === currentMemberId && !b.returnDate).length > 0 && (
                  <span className="px-1.5 py-0.2 bg-brand-600 text-white text-xs rounded-full">
                    {liveBorrows.filter(b => b.memberId === currentMemberId && !b.returnDate).length}
                  </span>
                )}
              </button>
            )}
            {persona === 'admin' && (
              <>
                <button
                  onClick={() => setActiveTab('transactions')}
                  className={`px-3.5 py-2 rounded-lg text-sm font-semibold transition flex items-center gap-1.5 ${
                    activeTab === 'transactions' ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  <span>🔄 Issues & Returns</span>
                  {stats.overdueBorrows > 0 && (
                    <span className="px-1.5 py-0.2 bg-rose-500 text-white text-xs rounded-full">
                      {stats.overdueBorrows}
                    </span>
                  )}
                </button>
                <button
                  onClick={() => setActiveTab('members')}
                  className={`px-3.5 py-2 rounded-lg text-sm font-semibold transition ${
                    activeTab === 'members' ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  👥 Members
                </button>
                <button
                  onClick={() => setActiveTab('analytics')}
                  className={`px-3.5 py-2 rounded-lg text-sm font-semibold transition ${
                    activeTab === 'analytics' ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  📊 Reports & Analytics
                </button>
              </>
            )}
          </nav>

          <div className="flex items-center gap-3">
            {/* Notifications */}
            <div className="relative">
              <button
                onClick={() => setShowNotificationsDropdown(!showNotificationsDropdown)}
                className="p-2 rounded-xl text-slate-500 hover:bg-slate-100 relative"
                title="Notifications"
              >
                <span className="text-lg">🔔</span>
                {pendingNotifications.length > 0 && (
                  <span className="absolute top-1 right-1 w-4 h-4 bg-rose-500 text-white text-[10px] font-bold rounded-full flex items-center justify-center">
                    {pendingNotifications.length}
                  </span>
                )}
              </button>

              {showNotificationsDropdown && (
                <div className="absolute right-0 mt-2 w-80 bg-white rounded-2xl shadow-2xl border border-slate-200 z-50 p-4">
                  <div className="flex justify-between items-center pb-2 border-b border-slate-100 mb-2">
                    <span className="font-bold text-xs text-slate-800">Return Reminders ({pendingNotifications.length})</span>
                    <button onClick={() => setShowNotificationsDropdown(false)} className="text-xs text-slate-400">Close</button>
                  </div>
                  {pendingNotifications.length === 0 ? (
                    <div className="text-xs text-slate-500 text-center py-4">No due alerts! 🎉</div>
                  ) : (
                    <div className="space-y-2 max-h-60 overflow-y-auto">
                      {pendingNotifications.map(item => (
                        <div key={item.id} className="p-2.5 rounded-xl border text-xs bg-rose-50 border-rose-200 text-rose-900">
                          <div className="font-bold">{item.bookTitle}</div>
                          <div className="text-[10px] text-slate-500 flex justify-between mt-1">
                            <span>{item.memberName}</span>
                            <span>Due: {item.dueDate}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Persona Switcher */}
            <div className="bg-slate-100 p-1 rounded-xl flex items-center border border-slate-200">
              <button
                onClick={() => { setPersona('member'); setActiveTab('catalog'); }}
                className={`px-3 py-1 rounded-lg text-xs font-semibold ${
                  persona === 'member' ? 'bg-white text-brand-700 shadow-sm' : 'text-slate-600'
                }`}
              >
                Student
              </button>
              <button
                onClick={() => { setPersona('admin'); setActiveTab('catalog'); }}
                className={`px-3 py-1 rounded-lg text-xs font-semibold ${
                  persona === 'admin' ? 'bg-slate-900 text-white shadow-sm' : 'text-slate-600'
                }`}
              >
                Admin
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Body */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Render catalog or tabs */}
        {activeTab === 'catalog' && (
          <div className="space-y-6">
            {/* Search and Filters */}
            <div className="bg-white p-4 sm:p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-center">
                <div className="md:col-span-6 relative">
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Search by title, author, ISBN..."
                    className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                  />
                  <span className="absolute left-3 top-2.5 text-slate-400">🔍</span>
                </div>
                <div className="md:col-span-3">
                  <select
                    value={selectedCategory}
                    onChange={(e) => setSelectedCategory(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                  >
                    {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div className="md:col-span-3">
                  <select
                    value={availabilityFilter}
                    onChange={(e) => setAvailabilityFilter(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                  >
                    <option value="all">All Inventory</option>
                    <option value="available">In Stock Only</option>
                    <option value="borrowed">Out of Stock</option>
                  </select>
                </div>
              </div>
            </div>

            {/* Book Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
              {filteredBooks.map(book => (
                <div key={book.id} className="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-start text-xs mb-2">
                      <span className="font-mono text-slate-400">{book.id}</span>
                      <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                        book.availableCopies > 0 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                      }`}>
                        {book.availableCopies > 0 ? `${book.availableCopies} available` : 'Out of Stock'}
                      </span>
                    </div>
                    <h3 className="font-bold text-slate-900 text-sm line-clamp-2">{book.title}</h3>
                    <p className="text-xs text-slate-500 mt-1 line-clamp-1">✍️ {book.author}</p>
                    <p className="text-xs text-brand-600 mt-2 font-medium">📍 {book.shelfLocation}</p>
                  </div>

                  <div className="mt-4 pt-3 border-t border-slate-100 flex gap-2">
                    <button
                      onClick={() => setSelectedBookModal(book)}
                      className="flex-1 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold"
                    >
                      Details
                    </button>
                    {persona === 'member' ? (
                      <button
                        onClick={() => handleIssueBook({ bookId: book.id, memberId: currentMemberId })}
                        disabled={book.availableCopies <= 0}
                        className={`py-1.5 px-3 rounded-xl text-xs font-semibold ${
                          book.availableCopies > 0 ? 'bg-brand-600 hover:bg-brand-700 text-white' : 'bg-slate-200 text-slate-400'
                        }`}
                      >
                        Borrow
                      </button>
                    ) : (
                      <button
                        onClick={() => { setPreselectedBookForIssue(book); setShowIssueModal(true); }}
                        disabled={book.availableCopies <= 0}
                        className="py-1.5 px-3 rounded-xl text-xs font-semibold bg-emerald-600 text-white"
                      >
                        Issue
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* My Books tab */}
        {activeTab === 'mybooks' && (
          <div className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm space-y-4">
            <h2 className="text-lg font-bold text-slate-900">Books Borrowed by Gazal Gupta (2401090100037)</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {liveBorrows.filter(b => b.memberId === currentMemberId && !b.returnDate).map(b => (
                <div key={b.id} className="p-4 rounded-2xl border border-slate-200 bg-slate-50 space-y-2 text-xs">
                  <div className="flex justify-between font-bold">
                    <span>{b.bookTitle}</span>
                    <span className={b.status === 'Overdue' ? 'text-rose-600' : 'text-emerald-600'}>{b.status}</span>
                  </div>
                  <div className="text-slate-500">Due Date: {b.dueDate}</div>
                  {b.currentFine > 0 && <div className="text-rose-600 font-bold">Late Fine: ₹{b.currentFine}</div>}
                  <div className="flex gap-2 pt-2">
                    <button onClick={() => handleReturnBook(b.id)} className="flex-1 py-1.5 rounded-lg bg-slate-900 text-white font-semibold">
                      Return Book
                    </button>
                    <button onClick={() => handleRenewBook(b.id)} className="py-1.5 px-3 rounded-lg bg-white border border-slate-300 font-semibold">
                      Renew (+7d)
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Transactions tab (admin) */}
        {activeTab === 'transactions' && persona === 'admin' && (
          <div className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm space-y-4">
            <div className="flex justify-between items-center">
              <h2 className="text-lg font-bold text-slate-900">Circulation Records</h2>
              <button onClick={() => setShowIssueModal(true)} className="px-3 py-1.5 bg-brand-600 text-white text-xs font-semibold rounded-xl">
                + Issue Book
              </button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-400 uppercase font-mono">
                  <tr>
                    <th className="p-3">ID</th>
                    <th className="p-3">Book</th>
                    <th className="p-3">Member</th>
                    <th className="p-3">Due</th>
                    <th className="p-3">Fine</th>
                    <th className="p-3">Status</th>
                    <th className="p-3 text-right">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {liveBorrows.map(item => (
                    <tr key={item.id}>
                      <td className="p-3 font-mono">{item.id}</td>
                      <td className="p-3 font-bold">{item.bookTitle}</td>
                      <td className="p-3">{item.memberName} ({item.memberRoll})</td>
                      <td className="p-3">{item.dueDate}</td>
                      <td className="p-3 font-bold text-rose-600">₹{item.currentFine}</td>
                      <td className="p-3 font-bold">{item.status}</td>
                      <td className="p-3 text-right">
                        {!item.returnDate && (
                          <button onClick={() => handleReturnBook(item.id)} className="px-3 py-1 bg-slate-900 text-white rounded-lg font-semibold">
                            Accept Return
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Members tab */}
        {activeTab === 'members' && persona === 'admin' && (
          <div className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm space-y-4">
            <div className="flex justify-between items-center">
              <h2 className="text-lg font-bold text-slate-900">Registered Members</h2>
              <button onClick={() => setShowAddMemberModal(true)} className="px-3 py-1.5 bg-brand-600 text-white text-xs font-semibold rounded-xl">
                + Add Member
              </button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {members.map(m => (
                <div key={m.id} className="p-4 rounded-2xl border border-slate-200 bg-slate-50 text-xs space-y-1">
                  <div className="font-bold text-sm text-slate-900">{m.name}</div>
                  <div className="font-mono text-slate-500">{m.rollNo}</div>
                  <div className="text-slate-600">{m.department} • {m.year}</div>
                  <div className="text-slate-400">{m.email}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Analytics tab */}
        {activeTab === 'analytics' && persona === 'admin' && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-2xl border border-slate-200 text-center">
                <div className="text-xs text-slate-400">Total Book Copies</div>
                <div className="text-2xl font-bold text-indigo-600">{stats.totalCopies}</div>
              </div>
              <div className="bg-white p-4 rounded-2xl border border-slate-200 text-center">
                <div className="text-xs text-slate-400">Available on Shelf</div>
                <div className="text-2xl font-bold text-emerald-600">{stats.availableCopies}</div>
              </div>
              <div className="bg-white p-4 rounded-2xl border border-slate-200 text-center">
                <div className="text-xs text-slate-400">Currently Issued</div>
                <div className="text-2xl font-bold text-amber-600">{stats.activeBorrows}</div>
              </div>
              <div className="bg-white p-4 rounded-2xl border border-slate-200 text-center">
                <div className="text-xs text-slate-400">Fines Collected</div>
                <div className="text-2xl font-bold text-slate-800">₹{stats.totalFinesCollected}</div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Return Receipt Modal */}
      {returnReceiptModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-3xl max-w-sm w-full p-6 space-y-4 text-xs text-center shadow-xl">
            <h3 className="font-bold text-base text-slate-900">Return Receipt</h3>
            <p className="text-slate-500">{returnReceiptModal.bookTitle}</p>
            <div className="bg-slate-50 p-3 rounded-xl text-left font-mono space-y-1">
              <div>Borrower: {returnReceiptModal.memberName}</div>
              <div>Returned: {returnReceiptModal.returnDate}</div>
              <div>Late Fine: ₹{returnReceiptModal.finePaid}</div>
            </div>
            <button onClick={() => setReturnReceiptModal(null)} className="w-full py-2 bg-slate-900 text-white rounded-xl font-semibold">
              Done
            </button>
          </div>
        </div>
      )}

      {/* Quick Issue Modal */}
      {showIssueModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-3xl max-w-sm w-full p-6 space-y-4 text-xs shadow-xl">
            <h3 className="font-bold text-base text-slate-900">Issue Book</h3>
            <form onSubmit={(e) => {
              e.preventDefault();
              const fd = new FormData(e.target);
              handleIssueBook({ bookId: fd.get('bookId'), memberId: fd.get('memberId'), loanDays: fd.get('loanDays') });
            }} className="space-y-3">
              <div>
                <label className="block font-bold mb-1">Book</label>
                <select name="bookId" defaultValue={preselectedBookForIssue?.id || books[0]?.id} className="w-full p-2 border rounded-xl">
                  {books.map(b => (
                    <option key={b.id} value={b.id} disabled={b.availableCopies <= 0}>{b.title} ({b.availableCopies} left)</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block font-bold mb-1">Member</label>
                <select name="memberId" className="w-full p-2 border rounded-xl">
                  {members.map(m => (
                    <option key={m.id} value={m.id}>{m.name} ({m.rollNo})</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block font-bold mb-1">Duration</label>
                <select name="loanDays" className="w-full p-2 border rounded-xl">
                  <option value="7">7 Days</option>
                  <option value="14">14 Days</option>
                  <option value="21">21 Days</option>
                </select>
              </div>
              <div className="flex gap-2 pt-2">
                <button type="submit" className="flex-1 py-2 bg-emerald-600 text-white font-semibold rounded-xl">Issue</button>
                <button type="button" onClick={() => setShowIssueModal(false)} className="px-4 py-2 bg-slate-100 rounded-xl">Cancel</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
