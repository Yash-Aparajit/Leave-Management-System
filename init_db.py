# init_db.py
import os
from werkzeug.security import generate_password_hash

from app import app  
from models import db, Employee, LeaveType, LeaveEntry, Transaction, User

def init():
    with app.app_context():
        print("Dropping all existing tables (if any)...")
        db.drop_all()

        print("Creating tables...")
        db.create_all()

        seed_leave_types = [
            ("Paid", True),
            ("Sick", True),
            ("Casual", True),
            ("Unpaid", False),
        ]
        for name, is_paid in seed_leave_types:
            existing = LeaveType.query.filter_by(name=name).first()
            if existing:
                print(f"LeaveType '{name}' already exists. Skipping.")
            else:
                lt = LeaveType(name=name, is_paid=is_paid)
                db.session.add(lt)
                print(f"Seeded LeaveType: {name} (is_paid={is_paid})")

        seed_users = [
            ("admin_1", "admin@jeena", "viewer_admin"),
            ("admin_master", "master@jeena", "admin_override"),
            ("developer", "dev@jeena@123", "developer"),
        ]

        for username, raw_pw, role in seed_users:
            existing = User.query.filter_by(username=username).first()
            pw_hash = generate_password_hash(raw_pw)
            if existing:
                existing.password_hash = pw_hash
                existing.role = role
                existing.force_password_change = True
                existing.session_token = None
                print(f"Updated user: {username} (role={role})")
            else:
                user = User(
                    username=username,
                    password_hash=pw_hash,
                    role=role,
                    force_password_change=True,
                    session_token=None
                )
                db.session.add(user)
                print(f"Created user: {username} (role={role})")

        db.session.commit()

        print("\nInitialization complete!")
        print("Seeded leave types and test users.")
        print("\nTest accounts:")
        print("  viewer_admin / Viewer@1234   (role: viewer_admin)")
        print("  admin_override / AdminOVR@123   (role: admin_override)")
        print("  developer / DevFull@1234   (role: developer)")
        print("\nAll seeded users require password change at first login.\n")

if __name__ == '__main__':
    init()
