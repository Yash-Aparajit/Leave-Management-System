# init_db.py — safe merged initializer
"""
Run: python init_db.py

This script:
  - Ensures the SQLite DB file exists (app.db in project root)
  - Calls SQLAlchemy db.create_all() using your models.py
  - Adds a set of historically-added columns via ALTER TABLE if they're missing
  - Seeds default users (only if the users table is empty)

Place this file next to app.py and models.py.
"""

import os
import sys
import sqlite3
from werkzeug.security import generate_password_hash
from flask import Flask
from models import (
    db,
    User,
    Employee,
    LeaveType,
    LeaveEntry,
    Transaction,
    CompOffRecord,
    EarlyLateRecord,
    OutdoorDuty
)



# --- make adjustments if needed ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "app.db")
SECRET_KEY = "change-this-secret-string-to-something-secret"
SYSTEM_VERSION = "1.0.0-init"


# Create a minimal Flask app and initialize the SQLAlchemy instance
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SECRET_KEY"] = SECRET_KEY

# bind SQLAlchemy to this app context for create_all / sessions
db.init_app(app)


def table_columns_sqlite(dbfile, table_name):
    """Return a set of column names for a sqlite table (or empty set if table not exists)."""
    if not os.path.exists(dbfile):
        return set()
    conn = sqlite3.connect(dbfile)
    cur = conn.cursor()
    try:
        cur.execute(f"PRAGMA table_info('{table_name}')")
        rows = cur.fetchall()
        cols = {r[1] for r in rows}
    except sqlite3.Error:
        cols = set()
    finally:
        conn.close()
    return cols


def add_column_if_missing(dbfile, table, column_name, column_type):
    """
    Add a column to a table if missing. Returns True if added.
    e.g. add_column_if_missing(DB_PATH, 'employees', 'plant_location', 'TEXT')
    """
    cols = table_columns_sqlite(dbfile, table)
    if column_name in cols:
        print(f"  - {table}.{column_name} present (skipping)")
        return False

    conn = sqlite3.connect(dbfile)
    cur = conn.cursor()
    try:
        sql = f"ALTER TABLE {table} ADD COLUMN {column_name} {column_type};"
        cur.execute(sql)
        conn.commit()
        print(f"  + Added column {table}.{column_name} ({column_type})")
        return True
    except sqlite3.Error as e:
        print(f"  ! Failed to add column {table}.{column_name}: {e}")
        return False
    finally:
        conn.close()


def ensure_tables_and_columns():
    """Create tables via SQLAlchemy and then ensure a list of optional columns exist."""
    print("Ensuring tables (create_all)...")
    with app.app_context():
        db.create_all()
    print("Tables ensured.")

    # Only attempt ALTER on an existing DB file (create_all will create the DB file)
    if not os.path.exists(DB_PATH):
        print("DB file not present after create_all() — aborting column checks.")
        return

    print("\nChecking & adding commonly-used columns if missing:")

    employees_cols = [
        ("plant_location", "TEXT"),
        ("manual_balance", "REAL"),
        ("initial_accrual_rate", "REAL"),
        ("left_date", "DATE"),
        ("middle_name", "TEXT"),
        ("status", "TEXT"),
        ("department", "TEXT"),
        ("designation", "TEXT"),
        ("contact_number", "TEXT"),
        ("emergency_number", "TEXT"),
    ]
    for name, typ in employees_cols:
        add_column_if_missing(DB_PATH, "employees", name, typ)

    # leave_entries optional columns (situation, approver, and recorder name)
    leave_entries_cols = [
        ("situation", "TEXT"),
        ("approver", "TEXT"),
        ("recorder_name", "TEXT"),   
    ]
    for name, typ in leave_entries_cols:
        add_column_if_missing(DB_PATH, "leave_entries", name, typ)


    # transactions optional helper columns
    transactions_cols = [
        ("period", "TEXT"),
        ("reference_id", "INTEGER"),
        ("note", "TEXT"),
    ]
    for name, typ in transactions_cols:
        add_column_if_missing(DB_PATH, "transactions", name, typ)

    # early_late_records optional columns
    early_late_cols = [
        ("approved_by", "TEXT"),
    ]
    for name, typ in early_late_cols:
        add_column_if_missing(DB_PATH, "early_late_records", name, typ)


def seed_defaults():
    """Seed default users and leave types if not present."""
    with app.app_context():
        # Seed leave types (Paid / Unpaid)
        try:
            lt_paid = LeaveType.query.filter_by(name="Paid").first()
            lt_unpaid = LeaveType.query.filter_by(name="Unpaid").first()
        except Exception as e:
            print("  ! Could not query LeaveType (maybe table missing). Error:", e)
            lt_paid = lt_unpaid = None

        if not lt_paid:
            try:
                db.session.add(LeaveType(name="Paid", is_paid=True))
                print("  + Seeded leave type: Paid")
            except Exception as e:
                print("  ! Failed to seed Paid leave type:", e)
        if not lt_unpaid:
            try:
                db.session.add(LeaveType(name="Unpaid", is_paid=False))
                print("  + Seeded leave type: Unpaid")
            except Exception as e:
                print("  ! Failed to seed Unpaid leave type:", e)
        try:
            db.session.commit()
        except Exception:
            db.session.rollback()

        # Seed users only if users table is empty
        try:
            users_count = db.session.query(User).count()
        except Exception as e:
            print("  ! Failed to query users table:", e)
            users_count = 0

        if users_count == 0:
            print("\nSeeding default users (only because users table is empty):")
            seeds = [
                ("admin_1", "admin@jeena", "admin_1"),
                ("admin_master", "master@jeena", "admin_master"),
                ("developer", "dev@jeena@123", "developer"),
            ]
            for uname, pwd, role in seeds:
                try:
                    u = User(username=uname, password_hash=generate_password_hash(pwd), role=role, force_password_change=True)
                    db.session.add(u)
                    print(f"  + {uname} / {pwd} (role: {role})")
                except Exception as e:
                    print(f"  ! Failed to add user {uname}: {e}")
            try:
                db.session.commit()
            except Exception as e:
                db.session.rollback()
                print("  ! Failed to commit seeded users:", e)
        else:
            print(f"\nUsers table has {users_count} record(s) — skipping user seeding.")


