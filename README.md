# 🧾 Leave Management System (Flask + SQLite)

A full-stack **Leave Management System** built with **Flask and SQLAlchemy** that handles employee leave requests, approvals, balances, and basic analytics.  

The goal of this project is to simulate a **real-world HR leave workflow** with clean backend logic and a usable web interface.

---

## ✨ Features

- 📝 **Leave Requests Workflow**  
  - Create leave requests with dates, reason, and leave type.  
  - Status flow: `Pending → Approved / Rejected`.  
  - Automatic leave balance deduction on approval.

- 📊 **Leave Balances & Types**  
  - Different leave types (e.g., Casual, Sick, Earned).  
  - Track remaining balance per user.  
  - Prevent overbooking or negative balance.

- ✅ **Validation & Rules**  
  - Prevent overlapping leave requests.  
  - Validate date ranges and allowed durations.  
  - Basic rule-based checks for policy-style constraints.

- 📈 **Basic Analytics (Optional / If Implemented)**  
  - View leaves by status, type, or user.  
  - Simple HR-style overview of upcoming leaves.

---

## 🧱 Tech Stack

- **Language:** Python  
- **Framework:** Flask  
- **ORM:** SQLAlchemy  
- **Database:** SQLite  
- **Frontend:** HTML, CSS, Jinja templates  
- **Others:** `virtualenv`, `pip`, `requirements.txt`

---

## 📂 Project Structure

```bash
leave-management-system/
│
├── app.py                # Main Flask application
├── models.py             # SQLAlchemy models (User, Leave, LeaveType, etc.)
├── config.py             # Configuration (DB URI, debug settings, etc.)
├── init_db.py            # Script to initialize / reset the database
│
├── requirements.txt      # Python dependencies
├── README.md             # Project documentation (this file)
├── .gitignore            # Files/folders to ignore in Git
├── .env.example          # Example environment variables
│
├── templates/            # HTML templates (Jinja2)
│   ├── base.html
│   ├── index.html
│   ├── login.html
│   ├── dashboard.html
│   ├── leave_request_form.html
│   ├── leave_list.html
│   ├── leave_detail.html
│   └── ...
│
├── static/
│   ├── css/
│   │   └── styles.css    # Custom styles
│   ├── js/
│   │   └── main.js       
│   └── img/              
│
└── instance/
    └── app.db            
