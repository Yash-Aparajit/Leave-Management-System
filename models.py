# models.py
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.sql import func

db = SQLAlchemy()


def now():
    return datetime.utcnow()


class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(64), nullable=False)
    force_password_change = db.Column(db.Boolean, nullable=False, default=True)
    session_token = db.Column(db.String(128), nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())


    def __repr__(self):
        return f'<User {self.username}>'


class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(64), unique=True, nullable=False, index=True)
    first_name = db.Column(db.String(120), nullable=False)
    middle_name = db.Column(db.String(120), nullable=True)
    last_name = db.Column(db.String(120), nullable=False)

    # Dates
    hire_date = db.Column(db.Date, nullable=True)
    promotion_date = db.Column(db.Date, nullable=True)
    left_date = db.Column(db.Date, nullable=True)

    # Accrual / balances
    accrual_rate = db.Column(db.Float, nullable=True)           # current accrual rate (e.g. 1.5 or 2.5)
    initial_accrual_rate = db.Column(db.Float, nullable=True)   # original accrual rate at import/create (helpful for recalcs)
    manual_balance = db.Column(db.Float, nullable=True)         # optional manual override (audit must be kept in Transaction)

    status = db.Column(db.String(32), nullable=False, default='active')  # active / left

    # Additional profile fields
    plant_location = db.Column(db.String(100), nullable=True)
    department = db.Column(db.String(100), nullable=True)
    designation = db.Column(db.String(100), nullable=True)
    contact_number = db.Column(db.String(50), nullable=True)
    emergency_number = db.Column(db.String(50), nullable=True)

    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())


    # relationships
    leaves = db.relationship('LeaveEntry', backref='employee', lazy='dynamic')
    transactions = db.relationship('Transaction', backref='employee_rel', lazy='dynamic')

    def __repr__(self):
        return f'<Employee {self.employee_id} {self.first_name} {self.last_name}>'


class LeaveType(db.Model):
    __tablename__ = 'leave_types'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False, unique=True)  # "Paid" or "Unpaid"
    is_paid = db.Column(db.Boolean, nullable=False, default=True) # whether this leave deducts balance
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())

    def __repr__(self):
        return f'<LeaveType {self.name}>'


class LeaveEntry(db.Model):
    __tablename__ = 'leave_entries'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date_from = db.Column(db.Date, nullable=True)
    date_to = db.Column(db.Date, nullable=True)
    days = db.Column(db.Float, nullable=False, default=0.0)
    leave_type_id = db.Column(db.Integer, db.ForeignKey('leave_types.id'), nullable=True)
    situation = db.Column(db.String(32), nullable=True)
    reason = db.Column(db.Text, nullable=True)
    approver = db.Column(db.String(200), nullable=True)
    recorder_name = db.Column(db.String(200), nullable=True)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())
    leave_type = db.relationship('LeaveType', foreign_keys=[leave_type_id])



    # relationship to leave type
    leave_type = db.relationship('LeaveType', foreign_keys=[leave_type_id])

    def __repr__(self):
        return f'<LeaveEntry emp={self.employee_id} {self.date_from}..{self.date_to}>'


class Transaction(db.Model):
    __tablename__ = 'transactions'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=True)
    type = db.Column(db.String(80), nullable=False)   # ACCRUAL / LEAVE_TAKEN / MANUAL_OVERRIDE / PROMOTION / ADJUSTMENT / OVERRIDE / PROMOTION_ADJUST etc.
    period = db.Column(db.String(16), nullable=True)  # e.g. '2025-11' for accruals; None for one-off
    amount = db.Column(db.Float, nullable=False, default=0.0)
    reference_id = db.Column(db.Integer, nullable=True)  # link to leave_entries.id etc.
    note = db.Column(db.Text, nullable=True)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())


    def __repr__(self):
        return f'<Transaction {self.type} {self.amount} emp={self.employee_id}>'
    

from datetime import datetime, date

class CompOffRecord(db.Model):
    __tablename__ = 'comp_offs'

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False, index=True)
    emp_code = db.Column(db.String(64), nullable=False)        # denormalized employee.employee_id
    emp_name = db.Column(db.String(255), nullable=False)       # denormalized name snapshot
    department = db.Column(db.String(128), nullable=True)      # denormalized department snapshot

    earned_on = db.Column(db.Date, nullable=False)             # date compoff earned
    taken_on = db.Column(db.Date, nullable=True)               # optional date compoff taken
    approved_by = db.Column(db.String(255), nullable=True)
    note = db.Column(db.Text, nullable=True)

    created_by = db.Column(db.Integer, nullable=True)          # user id who recorded
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, server_default=func.now(), onupdate=func.now())

    # optional relationship back to Employee (read-only convenience)
    employee = db.relationship('Employee', backref=db.backref('comp_offs', lazy='dynamic'))

    def to_export_row(self):
        """Return a dict/tuple suitable for export merging with leaves."""
        return {
            'record_type': 'COMP_OFF',
            'earned_on': self.earned_on.isoformat() if self.earned_on else '',
            'taken_on': self.taken_on.isoformat() if self.taken_on else '',
            'employee_id': self.emp_code,
            'employee_name': self.emp_name,
            'department': self.department or '',
            'approved_by': self.approved_by or '',
            'note': self.note or '',
            'created_by': str(self.created_by or ''),
            'created_at': self.created_at.isoformat() if self.created_at else ''
        }

