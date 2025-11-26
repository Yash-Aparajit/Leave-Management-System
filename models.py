# models.py
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(80), nullable=False, default='viewer_admin')
    force_password_change = db.Column(db.Boolean, default=False)
    session_token = db.Column(db.String(128), nullable=True)

class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(64), unique=True, nullable=False)  
    first_name = db.Column(db.String(120), nullable=False)
    middle_name = db.Column(db.String(120), nullable=True)
    last_name = db.Column(db.String(120), nullable=False)
    hire_date = db.Column(db.Date, nullable=True)
    promotion_date = db.Column(db.Date, nullable=True)
    accrual_rate = db.Column(db.Float, nullable=False, default=1.5)
    manual_balance = db.Column(db.Float, nullable=True)  
    status = db.Column(db.String(40), nullable=False, default='active')

    department = db.Column(db.String(80), nullable=True)        
    designation = db.Column(db.String(120), nullable=True)
    contact_number = db.Column(db.String(40), nullable=True)
    emergency_number = db.Column(db.String(40), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class LeaveType(db.Model):
    __tablename__ = 'leave_types'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    is_paid = db.Column(db.Boolean, default=True)

class LeaveEntry(db.Model):
    __tablename__ = 'leave_entries'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date_from = db.Column(db.Date, nullable=False)
    date_to = db.Column(db.Date, nullable=False)
    days = db.Column(db.Float, nullable=False)
    leave_type_id = db.Column(db.Integer, db.ForeignKey('leave_types.id'))
    reason = db.Column(db.String(500), nullable=True)
    approver = db.Column(db.String(200), nullable=True)  
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Transaction(db.Model):
    __tablename__ = 'transactions'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=True)
    type = db.Column(db.String(80), nullable=False)  
    period = db.Column(db.String(16), nullable=True)  
    amount = db.Column(db.Float, nullable=False, default=0.0)
    reference_id = db.Column(db.Integer, nullable=True)  
    note = db.Column(db.String(1000), nullable=True)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
