# 🧾 Leave & Attendance Management System

An **internal HR Leave & Attendance Management System** built with **Flask + SQLAlchemy**, designed for **offline plant-level HR operations**.

This system replaces Excel-based workflows with a **ledger-driven, auditable architecture** where all balances are calculated from transactions instead of stored values.

---

## 📌 Purpose

- Centralize employee leave & attendance records
- Enforce HR policies via system rules (not manual discipline)
- Maintain a complete audit trail
- Operate fully offline on local machines

---

## ✨ Key Features

### 👤 Employee Management
- Add / edit employees
- Department, designation, plant tracking
- Hire date & promotion handling
- Employee exit (left) locking

---

### 📝 Leave Management
- Paid & unpaid leave
- Planned / Unplanned / Sick classification
- Approver tracking
- Recorder (who entered the data) tracking
- Edit leave → **Developer only**
- Delete leave with full audit record

---

### 📊 Ledger-Based Leave Balance
- Monthly automatic accruals
- Leave deductions as transactions
- Promotion-based recalculation
- Manual balance correction via **delta override**
- Balance always computed from ledger

---

### 🕒 Attendance Modules
- Comp-Off
- Early / Late coming
- Outdoor Duty (Full / Half day)

Each module:
- Blocks left employees
- Tracks approvals
- Supports Excel export

---

### 📂 Reports & History
- Filterable leave history
- Monthly HR report (Excel)
- Yearly consolidated report
- Employee profile export (multi-sheet Excel)

---

### 🔐 Role-Based Access Control

| Role | Access |
|----|------|
| admin_1 | Daily HR operations |
| admin_master | Overrides, delete, restore |
| developer | Full system authority |

Rules enforced:
- Only developer can edit historical leave
- Manual balance changes are always audited
- Left employees are locked by default

---

### 💾 Backup & Restore
- SQLite database backup
- Manual restore with pre-restore snapshot
- Fully offline & local

---

## 🧱 Tech Stack

- **Language:** Python 3.10+
- **Backend:** Flask
- **ORM:** SQLAlchemy
- **Database:** SQLite
- **Frontend:** Jinja2, HTML, CSS, Bootstrap
- **Exports:** openpyxl, pandas
- **Auth:** Session-based authentication

---

## 📂 Project Structure

```text
Leave-Management-System/
│
├── app.py                 # Main application
├── models.py              # Database models
├── requirements.txt
├── README.md
├── .gitignore
│
├── templates/             # Jinja templates
├── static/
│   ├── css/
│   ├── js/
│   └── profile/
│
├── backups/               # DB backups (ignored)
├── uploads/               # Runtime uploads (ignored)
├── venv/                  # Virtual environment (ignored)

---

### 🚀 Setup (Local)

##1️⃣ Create virtual environment
python -m venv venv

##2️⃣ Activate environment
Windows
venv\Scripts\activate

##3️⃣ Install dependencies
pip install -r requirements.txt

##4️⃣ Run application
python app.py

##Open browser:
http://127.0.0.1:5000

---
