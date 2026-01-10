# create_db.py
from werkzeug.security import generate_password_hash
from app import app, db
from models import User, LeaveType

SEED_USERS = [
    ("admin_1", "admin@jeena", "viewer_admin"),
    ("admin_master", "master@jeena", "admin_master"),
    ("developer", "dev@jeena@123", "developer"),
]

def create():
    with app.app_context():
        print("Dropping and creating all tables (fresh DB).")
        db.drop_all()
        db.create_all()

        # seed leave types
        lt1 = LeaveType(name='Paid', is_paid=True)
        lt2 = LeaveType(name='Unpaid', is_paid=False)
        db.session.add_all([lt1, lt2])
        db.session.commit()
        print("Seeded leave types.")

        # seed users
        for username, pwd, role in SEED_USERS:
            u = User(username=username,
                     password_hash=generate_password_hash(pwd),
                     role=role,
                     force_password_change=False)
            db.session.add(u)
        db.session.commit()
        print("Seeded users:", [u[0] for u in SEED_USERS])

if __name__ == '__main__':
    create()
    print("DB created. Now run: python app.py")
