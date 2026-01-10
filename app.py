# app.p
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func, and_, or_

# Excel support
from openpyxl import Workbook

from models import db, Employee, LeaveType, LeaveEntry, Transaction, User

# ---------- App setup ----------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "app.db")   
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
BACKUPS_FOLDER = os.path.join(BASE_DIR, 'backups')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BACKUPS_FOLDER, exist_ok=True)

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"), static_folder=os.path.join(BASE_DIR, "static"))
app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{DB_PATH}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'change-this-secret-string-to-something-secret'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SYSTEM_VERSION'] = '1.0.4'
app.config['BUILD_DATE'] = '2025-11-25'

db.init_app(app)

# ---------- Roles & permissions ----------
ROLE_PERMISSIONS = {
    'viewer_admin': {
        'can_override': False,
        'can_restore_db': False,
        'can_delete_leave': False,
        'can_edit_employee': True,
        'can_set_manual_balance': False,
    },
    'admin_1': {   
        'can_override': False,
        'can_restore_db': False,
        'can_delete_leave': False,
        'can_edit_employee': True,
        'can_set_manual_balance': False,
    },
    'admin_master': {  
        'can_override': True,
        'can_restore_db': True,
        'can_delete_leave': True,
        'can_edit_employee': True,
        'can_set_manual_balance': True,
    },
    'developer': {  
        'can_override': True,
        'can_restore_db': True,
        'can_delete_leave': True,
        'can_edit_employee': True,
        'can_set_manual_balance': True,
    }
}

def normalize_role(role):
    if role == 'admin_override':
        return 'admin_master'
    return role

def has_permission(permission_name):
    uid = session.get('user_id')
    if not uid:
        return False
    user = User.query.get(uid)
    if not user:
        return False
    role = normalize_role(user.role)
    perms = ROLE_PERMISSIONS.get(role, {})
    return perms.get(permission_name, False)

def require_permission(permission_name):
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not has_permission(permission_name):
                flash('Permission denied', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return wrapper
    return decorator

# ---------- session & auth ----------
@app.before_request
def validate_session_token():
    endpoint = (request.endpoint or '')
    public_endpoints = ('login', 'create_tables_and_seed', 'static')
    if endpoint.startswith('static') or endpoint in public_endpoints:
        return
    uid = session.get('user_id')
    token = session.get('session_token')
    if not uid or not token:
        return
    user = User.query.get(uid)
    if not user or not user.session_token or user.session_token != token:
        session.clear()
        flash('Session expired — please log in again.', 'info')
        return redirect(url_for('login'))

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated

@app.template_filter('nice_date')
def nice_date(d):
    return d.strftime('%Y-%m-%d') if d else ''

@app.context_processor
def inject_permissions():
    return dict(has_permission=has_permission, session=session)

# ---------- accrual & balance helpers ----------
def apply_missing_accruals_for_employee(emp):
    
    if not emp.hire_date:
        return
    today = date.today()
    cur = emp.hire_date.replace(day=1)
    last_month = today.replace(day=1)
    while cur <= last_month:
        period = cur.strftime('%Y-%m')
        exists = Transaction.query.filter_by(employee_id=emp.id, type='ACCRUAL', period=period).first()
        if not exists:
            t = Transaction(
                employee_id=emp.id,
                type='ACCRUAL',
                period=period,
                amount=round(float(emp.accrual_rate or 0), 2),
                note=f'Auto-accrual for {period}',
                created_by=session.get('user_id')
            )
            db.session.add(t)
        if cur.month == 12:
            cur = cur.replace(year=cur.year+1, month=1)
        else:
            cur = cur.replace(month=cur.month+1)
    db.session.commit()

def compute_balance(emp):
    """
    If manual_balance (override) is set on employee, return that;
    otherwise ensure accruals are present and compute sum of transactions.
    """
    if getattr(emp, 'manual_balance', None) is not None:
        try:
            return round(float(emp.manual_balance), 2)
        except Exception:
            pass
    try:
        apply_missing_accruals_for_employee(emp)
    except Exception:
        pass
    total = db.session.query(func.coalesce(func.sum(Transaction.amount), 0)).filter(Transaction.employee_id == emp.id).scalar()
    return round(float(total or 0), 2)

# ---------- Authentication routes ----------
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password_hash, password):
            role = normalize_role(user.role)
            if user.role != role:
                user.role = role
            token = uuid.uuid4().hex
            user.session_token = token
            db.session.commit()
            session['session_token'] = token
            session['user_id'] = user.id
            session['username'] = user.username
            session['role'] = role
            if getattr(user, 'force_password_change', False):
                flash('You must change the default password before continuing.', 'info')
                return redirect(url_for('change_password'))
            flash('Logged in', 'success')
            return redirect(url_for('index'))
        flash('Invalid credentials', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    try:
        uid = session.get('user_id')
        if uid:
            user = User.query.get(uid)
            if user:
                user.session_token = None
                db.session.commit()
    except Exception:
        pass
    session.clear()
    flash('Logged out', 'info')
    return redirect(url_for('login'))

@app.route('/change-password', methods=['GET','POST'])
@login_required
def change_password():
    user = User.query.get(session.get('user_id'))
    if not user:
        flash('User not found', 'danger')
        return redirect(url_for('login'))
    if request.method == 'POST':
        current = request.form.get('current_password','')
        new = request.form.get('new_password','')
        confirm = request.form.get('confirm_password','')
        if not check_password_hash(user.password_hash, current):
            flash('Current password is incorrect', 'danger')
            return redirect(url_for('change_password'))
        if len(new) < 8:
            flash('New password must be at least 8 characters', 'danger')
            return redirect(url_for('change_password'))
        if new != confirm:
            flash('New password and confirmation do not match', 'danger')
            return redirect(url_for('change_password'))
        user.password_hash = generate_password_hash(new)
        user.force_password_change = False
        user.session_token = None
        db.session.commit()
        session.clear()
        flash('Password changed. Please log in again.', 'success')
        return redirect(url_for('login'))
    return render_template('change_password.html')

# ---------- Dashboard & general pages ----------
@app.route('/')
@login_required
def index():
    total_employees = Employee.query.filter(Employee.status=='active').count()
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    week_count = LeaveEntry.query.filter(LeaveEntry.date_from >= week_start).count()
    month_count = LeaveEntry.query.filter(func.strftime('%Y-%m', LeaveEntry.date_from) == today.strftime('%Y-%m')).count()
    year_count = LeaveEntry.query.filter(func.strftime('%Y', LeaveEntry.date_from) == str(today.year)).count()

    leave_types = LeaveType.query.all()
    counts_by_type = {lt.name: LeaveEntry.query.filter(LeaveEntry.leave_type_id == lt.id).count() for lt in leave_types}

    employees = Employee.query.filter(Employee.status=='active').order_by(Employee.last_name).limit(50).all()
    balances = { e.id: compute_balance(e) for e in employees }

    return render_template('index.html',
                           total_employees=total_employees,
                           week_count=week_count,
                           month_count=month_count,
                           year_count=year_count,
                           employees=employees,
                           balances=balances,
                           counts_by_type=counts_by_type,
                           system_version=app.config['SYSTEM_VERSION'],
                           build_date=app.config['BUILD_DATE'])

@app.route('/help')
@login_required
def help_page():
    return render_template('help.html', system_version=app.config.get('SYSTEM_VERSION',''), build_date=app.config.get('BUILD_DATE',''))

# ---------- Employee listing & search ----------
@app.route('/employees')
@login_required
def employees():
    return redirect(url_for('employees_list'))

@app.route('/employees_list')
@login_required
def employees_list():
    q = request.args.get('q', '').strip()
    status = request.args.get('status', 'active')
    try:
        page = int(request.args.get('page', '1'))
    except Exception:
        page = 1
    per_page = 40

    base = Employee.query
    if status:
        base = base.filter(Employee.status == status)
    if q:
        base = base.filter(or_(
            Employee.employee_id.ilike(f'%{q}%'),
            Employee.first_name.ilike(f'%{q}%'),
            Employee.middle_name.ilike(f'%{q}%'),
            Employee.last_name.ilike(f'%{q}%')
        ))
    total = base.count()
    employees = base.order_by(Employee.employee_id).offset((page-1)*per_page).limit(per_page).all()

    balances = {}
    try:
        for e in employees:
            try:
                balances[e.id] = compute_balance(e)
            except Exception:
                balances[e.id] = 0
    except Exception:
        balances = { e.id: 0 for e in employees }

    return render_template('employees_list.html',
                           employees=employees,
                           q=q,
                           status=status,
                           page=page,
                           per_page=per_page,
                           total=total,
                           balances=balances)

@app.route('/search')
@login_required
def search():
    q = request.args.get('q','').strip()
    if not q:
        return redirect(url_for('index'))
    emp = Employee.query.filter_by(employee_id=q).first()
    if emp:
        return redirect(url_for('employee_detail', emp_id=emp.id))
    parts = q.split()
    if len(parts) == 1:
        employees = Employee.query.filter(or_(
            Employee.first_name.ilike(f'%{q}%'),
            Employee.middle_name.ilike(f'%{q}%'),
            Employee.last_name.ilike(f'%{q}%')
        )).all()
    else:
        employees = Employee.query.filter(and_(Employee.first_name.ilike(f'%{parts[0]}%'), Employee.last_name.ilike(f'%{parts[-1]}%'))).all()
    return render_template('employees.html', employees=employees)

# ---------- Add / Remove / Promote / Edit employee ----------
@app.route('/employees/add', methods=['GET','POST'])
@login_required
def add_employee():
    departments = ['Purchase', 'ME', 'HR Admin', 'Production', 'Quality', 'Store', 'Other']
    if request.method == 'POST':
        emp_id = request.form.get('employee_id','').strip()
        first = request.form.get('first_name','').strip()
        middle = request.form.get('middle_name','').strip() or None
        last = request.form.get('last_name','').strip()
        try:
            hire_date = dt.strptime(request.form.get('hire_date'), '%Y-%m-%d').date()
        except Exception:
            flash('Invalid hire date', 'danger')
            return redirect(url_for('add_employee'))
        accrual_rate = float(request.form.get('accrual_rate', 1.5))
        department = request.form.get('department','').strip() or None
        designation = request.form.get('designation','').strip() or None
        contact = request.form.get('contact_number','').strip() or None
        emergency = request.form.get('emergency_number','').strip() or None

        if Employee.query.filter_by(employee_id=emp_id).first():
            flash('Employee ID already exists', 'danger')
            return redirect(url_for('add_employee'))
        emp = Employee(employee_id=emp_id, first_name=first, middle_name=middle, last_name=last, hire_date=hire_date, accrual_rate=accrual_rate, status='active',
                       department=department, designation=designation, contact_number=contact, emergency_number=emergency)
        db.session.add(emp)
        db.session.commit()
        flash('Employee added', 'success')
        return redirect(url_for('employees_list'))
    return render_template('add_employee.html', departments=['Purchase','ME','HR Admin','Production','Quality','Store','Other'])

@app.route('/employees/remove', methods=['GET','POST'])
@login_required
def remove_employee():
    if request.method == 'POST':
        emp_id = request.form.get('employee_id','').strip()
        emp = Employee.query.filter_by(employee_id=emp_id).first()
        if not emp:
            flash('Employee not found', 'danger')
            return redirect(url_for('remove_employee'))
        emp.status = 'left'
        db.session.commit()
        flash(f'{emp.employee_id} marked as left', 'success')
        return redirect(url_for('employees_list'))
    return render_template('remove_employee.html')

@app.route('/promote', methods=['GET','POST'])
@login_required
def promote_employee():
    if request.method == 'POST':
        emp_id = request.form.get('employee_id','').strip()
        try:
            new_rate = float(request.form.get('new_rate'))
        except Exception:
            flash('Invalid new rate', 'danger')
            return redirect(url_for('promote_employee'))
        try:
            effective = dt.strptime(request.form.get('effective_date'), '%Y-%m-%d').date()
        except Exception:
            flash('Invalid effective date', 'danger')
            return redirect(url_for('promote_employee'))
        emp = Employee.query.filter_by(employee_id=emp_id).first()
        if not emp:
            flash('Employee not found', 'danger')
            return redirect(url_for('promote_employee'))
        emp.promotion_date = effective
        emp.accrual_rate = new_rate
        db.session.commit()
        flash(f'{emp.employee_id} updated to accrual rate {new_rate}', 'success')
        return redirect(url_for('employees_list'))
    return render_template('promote.html')

@app.route('/employees/edit/<int:emp_id>', methods=['GET', 'POST'])
@login_required
def edit_employee(emp_id):
    """
    Edit employee profile: allows editing of contact, emergency contact,
    department, designation, accrual rate, hire/promotion dates, status,
    and (for admin_master/developer) manual_balance override.
    """
    emp = Employee.query.get_or_404(emp_id)

    if not has_permission('can_edit_employee') and session.get('role') not in ('developer', 'admin_master'):
        flash('No permission to edit employee', 'danger')
        return redirect(url_for('employee_detail', emp_id=emp.id))

    # Departments list for dropdown
    departments = ['Purchase', 'ME', 'HR Admin', 'Production', 'Quality', 'Store', 'Other']

    if request.method == 'POST':
        emp.employee_id = request.form.get('employee_id', emp.employee_id).strip()
        emp.first_name = request.form.get('first_name', emp.first_name).strip()
        emp.middle_name = request.form.get('middle_name','').strip() or None
        emp.last_name = request.form.get('last_name', emp.last_name).strip()

        # Dates
        try:
            hire_date_str = request.form.get('hire_date')
            if hire_date_str:
                emp.hire_date = dt.strptime(hire_date_str, '%Y-%m-%d').date()
        except Exception:
            flash('Invalid hire date', 'danger')
            return redirect(url_for('edit_employee', emp_id=emp.id))

        try:
            promo_str = request.form.get('promotion_date')
            emp.promotion_date = dt.strptime(promo_str, '%Y-%m-%d').date() if promo_str else None
        except Exception:
            flash('Invalid promotion date', 'danger')
            return redirect(url_for('edit_employee', emp_id=emp.id))

        try:
            emp.accrual_rate = float(request.form.get('accrual_rate', emp.accrual_rate))
        except Exception:
            flash('Invalid accrual rate', 'danger')
            return redirect(url_for('edit_employee', emp_id=emp.id))

        emp.status = request.form.get('status', emp.status)

        emp.department = request.form.get('department','').strip() or None
        emp.designation = request.form.get('designation','').strip() or None
        emp.contact_number = request.form.get('contact_number','').strip() or None
        emp.emergency_number = request.form.get('emergency_number','').strip() or None

        # Manual balance override handling:
        new_manual_raw = request.form.get('manual_balance', '').strip()
        new_manual_val = None
        if new_manual_raw != '':
            try:
                new_manual_val = float(new_manual_raw)
            except Exception:
                flash('Invalid manual balance value', 'danger')
                return redirect(url_for('edit_employee', emp_id=emp.id))

        # Only developer and admin_master can set manual balance
        if new_manual_raw != '' and session.get('role') not in ('developer', 'admin_master'):
            flash('Permission denied for manual balance change', 'danger')
            return redirect(url_for('edit_employee', emp_id=emp.id))

        previous_effective = None
        try:
            if getattr(emp, 'manual_balance', None) is not None:
                previous_effective = float(emp.manual_balance)
            else:
                previous_effective = compute_balance(emp)
        except Exception:
            previous_effective = None

        emp.manual_balance = new_manual_val

        db.session.commit()

        try:
            new_effective = float(new_manual_val) if new_manual_val is not None else compute_balance(emp)
            delta = None
            if previous_effective is not None:
                delta = round(float(new_effective) - float(previous_effective), 2)
            note = f'MANUAL_OVERRIDE by {session.get("username")}: previous={previous_effective}, new={new_effective}'
            tr = Transaction(
                employee_id=emp.id,
                type='MANUAL_OVERRIDE',
                period=None,
                amount=delta if delta is not None else 0.0,
                reference_id=None,
                note=note,
                created_by=session.get('user_id')
            )
            db.session.add(tr)
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            flash('Warning: manual override saved but audit record failed: ' + str(e), 'warning')

        flash('Employee updated', 'success')
        return redirect(url_for('employee_detail', emp_id=emp.id))

    return render_template('edit_employee.html', e=emp, departments=departments)


# ---------- Record / Edit / Delete leave ----------


@app.route('/leave/edit/<int:leave_id>', methods=['GET','POST'])
@login_required
def edit_leave(leave_id):
    le = LeaveEntry.query.get_or_404(leave_id)
    emp = Employee.query.get_or_404(le.employee_id)
    leave_types = LeaveType.query.all()
    if not has_permission('can_edit_employee') and session.get('role') not in ('developer', 'admin_master'):
        flash('No permission to edit leave', 'danger')
        return redirect(url_for('employee_detail', emp_id=emp.id))
    if request.method == 'POST':
        try:
            date_from = dt.strptime(request.form.get('date_from'), '%Y-%m-%d').date()
            date_to = dt.strptime(request.form.get('date_to'), '%Y-%m-%d').date()
            days = float(request.form.get('days'))
            lt_id = int(request.form.get('leave_type_id'))
            reason = request.form.get('reason','').strip()
            approver = request.form.get('approver','').strip()
        except Exception:
            flash('Invalid data', 'danger')
            return redirect(url_for('edit_leave', leave_id=leave_id))
        le.date_from = date_from
        le.date_to = date_to
        le.days = days
        le.leave_type_id = lt_id
        le.reason = reason
        try:
            setattr(le, 'approver', approver or None)
        except Exception:
            pass
        db.session.commit()
        tr = Transaction.query.filter_by(reference_id=le.id, type='LEAVE_TAKEN').first()
        if tr:
            tr.amount = round(-abs(days),2)
            tr.note = f'Edited leave {date_from} to {date_to}' + (f' — Approver: {approver}' if approver else '')
            db.session.commit()
        flash('Leave updated', 'success')
        return redirect(url_for('employee_detail', emp_id=emp.id))
    return render_template('edit_leave.html', le=le, leave_types=leave_types, emp=emp)

@app.route('/leave/delete/<int:leave_id>', methods=['POST'])
@login_required
def delete_leave(leave_id):
    if not has_permission('can_delete_leave') and session.get('role') not in ('developer', 'admin_master'):
        flash('Delete permission required', 'danger')
        return redirect(url_for('index'))
    le = LeaveEntry.query.get_or_404(leave_id)
    emp = Employee.query.get_or_404(le.employee_id)
    tr_list = Transaction.query.filter_by(reference_id=le.id).all()
    for tr in tr_list:
        db.session.delete(tr)
    db.session.delete(le)
    adj = Transaction(employee_id=emp.id, type='ADJUSTMENT', period=None, amount=0.0, note=f'Deleted leave id {leave_id} by {session.get("username")}', created_by=session.get('user_id'))
    db.session.add(adj)
    db.session.commit()
    flash('Leave deleted', 'success')
    return redirect(url_for('employee_detail', emp_id=emp.id))

# ---------- Employee detail, print & export ----------
@app.route('/employee/<int:emp_id>')
@login_required
def employee_detail(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    bal = compute_balance(emp)
    leaves = LeaveEntry.query.filter_by(employee_id=emp.id).order_by(LeaveEntry.date_from.desc()).all()
    transactions = Transaction.query.filter_by(employee_id=emp.id).order_by(Transaction.created_at.desc()).limit(200).all()
    return render_template('employee_detail.html', e=emp, bal=bal, leaves=leaves, transactions=transactions)

@app.route('/employee/print/<int:emp_id>')
@login_required
def employee_print(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    bal = compute_balance(emp)
    leaves = LeaveEntry.query.filter_by(employee_id=emp.id).order_by(LeaveEntry.date_from).all()
    transactions = Transaction.query.filter_by(employee_id=emp.id).order_by(Transaction.created_at).all()
    return render_template('employee_print.html', e=emp, bal=bal, leaves=leaves, transactions=transactions)

@app.route('/check_balance', methods=['GET','POST'])
@login_required
def check_balance():
    """
    Show a small form to lookup an employee by Employee ID and display current balance + recent leaves.
    """
    if request.method == 'POST':
        emp_code = request.form.get('employee_id','').strip()
        emp = Employee.query.filter_by(employee_id=emp_code).first()
        if not emp:
            flash('Employee not found', 'danger')
            return redirect(url_for('check_balance'))
        bal = compute_balance(emp)
        leaves = LeaveEntry.query.filter_by(employee_id=emp.id).order_by(LeaveEntry.date_from.desc()).all()
        return render_template('balance.html', emp=emp, bal=bal, leaves=leaves)

    return render_template('check_balance.html')

@app.route('/export_employee/<emp_code>')
@login_required
def export_employee(emp_code):
    """
    Exports an employee report CSV that includes approver (if present) in the leaves section.
    """
    emp = Employee.query.filter_by(employee_id=emp_code).first_or_404()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Employee ID', emp.employee_id])
    writer.writerow(['Name', f"{emp.first_name} {emp.middle_name or ''} {emp.last_name}"])
    writer.writerow(['Department', emp.department or ''])
    writer.writerow(['Designation', emp.designation or ''])
    writer.writerow(['Contact', emp.contact_number or ''])
    writer.writerow(['Emergency', emp.emergency_number or ''])
    writer.writerow(['Hire Date', emp.hire_date.isoformat() if emp.hire_date else ''])
    writer.writerow([])
    writer.writerow(['Leaves: date_from','date_to','days','type','approver','reason'])
    leaves = LeaveEntry.query.filter_by(employee_id=emp.id).order_by(LeaveEntry.date_from).all()
    for l in leaves:
        lt = LeaveType.query.get(l.leave_type_id)
        approver = getattr(l, 'approver', '') or ''
        writer.writerow([l.date_from.isoformat() if l.date_from else '', l.date_to.isoformat() if l.date_to else '', l.days, lt.name if lt else l.leave_type_id, approver, l.reason])
    writer.writerow([])
    writer.writerow(['Transactions: date','type','amount','note'])
    trans = Transaction.query.filter_by(employee_id=emp.id).order_by(Transaction.created_at).all()
    for t in trans:
        writer.writerow([t.created_at.isoformat() if t.created_at else '', t.type, t.amount, t.note])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode('utf-8')), mimetype='text/csv', as_attachment=True, download_name=f'{emp.employee_id}_report.csv')

@app.route('/export_employees')
@login_required
def export_employees():
    """
    Export a CSV or Excel (.xlsx) for all employees (or filtered by status/q via query params).
    format=xlsx -> Excel, otherwise CSV.
    """
    status = request.args.get('status', '')  
    q = request.args.get('q', '').strip()
    fmt = request.args.get('format', '').lower()

    base = Employee.query
    if status:
        base = base.filter(Employee.status == status)
    if q:
        base = base.filter(or_(Employee.employee_id.ilike(f'%{q}%'), Employee.first_name.ilike(f'%{q}%'), Employee.last_name.ilike(f'%{q}%')))

    employees = base.order_by(Employee.employee_id).all()

    # XLSX export
    if fmt == 'xlsx':
        wb = Workbook()
        ws = wb.active
        ws.title = "Employees"

        headers = ['Employee ID','First name','Middle name','Last name','Department','Designation','Contact','Emergency','Hire Date','Status','Accrual rate','Manual balance']
        ws.append(headers)
        for e in employees:
            row = [
                e.employee_id,
                e.first_name,
                e.middle_name or '',
                e.last_name,
                e.department or '',
                e.designation or '',
                e.contact_number or '',
                e.emergency_number or '',
                e.hire_date.isoformat() if e.hire_date else '',
                e.status or '',
                e.accrual_rate if e.accrual_rate is not None else '',
                e.manual_balance if getattr(e, 'manual_balance', None) is not None else ''
            ]
            ws.append(row)

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    val = str(cell.value or '')
                except Exception:
                    val = ''
                if len(val) > max_len:
                    max_len = len(val)
            adjusted_width = (max_len + 2)
            ws.column_dimensions[col_letter].width = adjusted_width

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        filename = f'employees_{dt.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

    si = io.StringIO()
    writer = csv.writer(si)
    writer.writerow(['Employee ID','First name','Middle name','Last name','Department','Designation','Contact','Emergency','Hire Date','Status','Accrual rate','Manual balance'])
    for e in employees:
        writer.writerow([e.employee_id, e.first_name, e.middle_name or '', e.last_name, e.department or '', e.designation or '', e.contact_number or '', e.emergency_number or '', e.hire_date.isoformat() if e.hire_date else '', e.status, e.accrual_rate, e.manual_balance if getattr(e, 'manual_balance', None) is not None else ''])
    si.seek(0)
    return send_file(io.BytesIO(si.getvalue().encode('utf-8')), mimetype='text/csv', as_attachment=True, download_name='employees_list.csv')

# ---------- History & exports ----------
@app.route('/history')
@login_required
def history():
    from_date = request.args.get('from_date')
    to_date = request.args.get('to_date')
    emp_code = request.args.get('employee_id')
    ttype = request.args.get('type')

    q = Transaction.query
    if from_date:
        try:
            fd = dt.strptime(from_date, '%Y-%m-%d')
            q = q.filter(Transaction.created_at >= fd)
        except Exception:
            pass
    if to_date:
        try:
            td = dt.strptime(to_date, '%Y-%m-%d') + timedelta(days=1)
            q = q.filter(Transaction.created_at < td)
        except Exception:
            pass
    if emp_code:
        emp = Employee.query.filter_by(employee_id=emp_code).first()
        if emp:
            q = q.filter(Transaction.employee_id == emp.id)
        else:
            q = q.filter(Transaction.employee_id == -1)
    if ttype:
        q = q.filter(Transaction.type == ttype)

    q = q.order_by(Transaction.created_at.desc()).limit(500)
    txs = q.all()

    unique = {}
    for t in txs:
        unique[t.id] = t
    txs_unique = list(unique.values())

    emp_ids = {t.employee_id for t in txs_unique if t.employee_id is not None}
    employees = Employee.query.filter(Employee.id.in_(list(emp_ids))).all() if emp_ids else []
    employees_map = { e.id: e for e in employees }

    user_ids = {t.created_by for t in txs_unique if t.created_by}
    users = User.query.filter(User.id.in_(list(user_ids))).all() if user_ids else []
    users_map = { u.id: u.username for u in users }

    leave_types = [lt.name for lt in LeaveType.query.all()]

    return render_template('history.html', transactions=txs_unique,
                           employees_map=employees_map, users_map=users_map,
                           filters={'from_date': from_date, 'to_date': to_date, 'employee_id': emp_code, 'type': ttype},
                           leave_types=leave_types)

@app.route('/export_history')
@login_required
def export_history():
    from_date = request.args.get('from_date')
    to_date = request.args.get('to_date')
    emp_code = request.args.get('employee_id')
    ttype = request.args.get('type')

    q = Transaction.query
    if from_date:
        try:
            fd = dt.strptime(from_date, '%Y-%m-%d')
            q = q.filter(Transaction.created_at >= fd)
        except Exception:
            pass
    if to_date:
        try:
            td = dt.strptime(to_date, '%Y-%m-%d') + timedelta(days=1)
            q = q.filter(Transaction.created_at < td)
        except Exception:
            pass
    if emp_code:
        emp = Employee.query.filter_by(employee_id=emp_code).first()
        if emp:
            q = q.filter(Transaction.employee_id == emp.id)
        else:
            q = q.filter(Transaction.employee_id == -1)
    if ttype:
        q = q.filter(Transaction.type == ttype)
    q = q.order_by(Transaction.created_at.desc()).all()

    si = io.StringIO()
    writer = csv.writer(si)
    writer.writerow(['Date', 'Employee ID', 'Employee Name', 'Type', 'Amount', 'Performed by', 'Note'])
    emp_ids = {t.employee_id for t in q if t.employee_id is not None}
    employees = Employee.query.filter(Employee.id.in_(list(emp_ids))).all() if emp_ids else []
    emp_map = { e.id: e for e in employees }
    user_ids = {t.created_by for t in q if t.created_by}
    users = User.query.filter(User.id.in_(list(user_ids))).all() if user_ids else []
    user_map = { u.id: u.username for u in users }
    for t in q:
        emp = emp_map.get(t.employee_id)
        emp_code = emp.employee_id if emp else (t.employee_id or '')
        emp_name = f"{emp.first_name} {emp.last_name}" if emp else ''
        performed = user_map.get(t.created_by, 'System')
        writer.writerow([t.created_at.isoformat() if t.created_at else '', emp_code, emp_name, t.type, t.amount, performed, (t.note or '')])
    si.seek(0)
    return send_file(io.BytesIO(si.getvalue().encode('utf-8')), mimetype='text/csv', as_attachment=True, download_name='history_export.csv')

@app.route('/leaves_history', methods=['GET'])
@login_required
def leaves_history():
    """
    Leaves history (only leave entries) — filterable by date range, employee id, leave type, department.
    """
    from_date = request.args.get('from_date', '')
    to_date = request.args.get('to_date', '')
    emp_code = request.args.get('employee_id', '').strip()
    leave_type_id = request.args.get('leave_type_id', '')
    dept = request.args.get('department', '')

    q = LeaveEntry.query

    if from_date:
        try:
            fd = dt.strptime(from_date, '%Y-%m-%d').date()
            q = q.filter(LeaveEntry.date_from >= fd)
        except Exception:
            pass
    if to_date:
        try:
            td = dt.strptime(to_date, '%Y-%m-%d').date()
            q = q.filter(LeaveEntry.date_to <= td)
        except Exception:
            pass

    if emp_code:
        emp = Employee.query.filter_by(employee_id=emp_code).first()
        if emp:
            q = q.filter(LeaveEntry.employee_id == emp.id)
        else:
            q = q.filter(LeaveEntry.employee_id == -1)

    if leave_type_id:
        try:
            q = q.filter(LeaveEntry.leave_type_id == int(leave_type_id))
        except Exception:
            pass

    if dept:
        emp_ids = [e.id for e in Employee.query.filter(Employee.department == dept).all()]
        if emp_ids:
            q = q.filter(LeaveEntry.employee_id.in_(emp_ids))
        else:
            q = q.filter(LeaveEntry.employee_id == -1)

    leaves = q.order_by(LeaveEntry.date_from.desc()).limit(2000).all()

    emp_ids = {l.employee_id for l in leaves if l.employee_id is not None}
    employees = Employee.query.filter(Employee.id.in_(list(emp_ids))).all() if emp_ids else []
    emp_map = {e.id: e for e in employees}

    leave_types = LeaveType.query.all()
    departments = ['Purchase','ME','HR Admin','Production','Quality','Store','Other']

    return render_template('leaves_history.html',
                           leaves=leaves,
                           emp_map=emp_map,
                           leave_types=leave_types,
                           departments=departments,
                           filters={'from_date': from_date, 'to_date': to_date, 'employee_id': emp_code, 'leave_type_id': leave_type_id, 'department': dept})

@app.route('/month_report')
@login_required
def month_report():
    """
    Small UI to pick year/month before exporting the month report.
    """
    return render_template('month_report.html')

@app.route('/export_month_report', methods=['GET'])
@login_required
def export_month_report():
    """
    Export a month-end report as .xlsx.
    Query params:
      year=YYYY, month=MM  (or month_iso=YYYY-MM)
    If missing, defaults to current month.
    """
    year = request.args.get('year')
    month = request.args.get('month')
    month_iso = request.args.get('month_iso')  # YYYY-MM

    if month_iso:
        try:
            year, month = month_iso.split('-', 1)
        except Exception:
            pass

    try:
        if not year or not month:
            today = date.today()
            year = str(today.year)
            month = str(today.month).zfill(2)
        year_i = int(year)
        month_i = int(month)
    except Exception:
        today = date.today()
        year_i = today.year
        month_i = today.month

    first_day = date(year_i, month_i, 1)
    if month_i == 12:
        last_day = date(year_i+1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(year_i, month_i+1, 1) - timedelta(days=1)

    leaves = LeaveEntry.query.filter(
        LeaveEntry.date_to >= first_day,
        LeaveEntry.date_from <= last_day
    ).order_by(LeaveEntry.date_from).all()

    emp_ids = {l.employee_id for l in leaves}
    employees = Employee.query.filter(Employee.id.in_(list(emp_ids))).all() if emp_ids else []
    emp_map = {e.id: e for e in employees}

    lt_ids = {l.leave_type_id for l in leaves if l.leave_type_id is not None}
    leave_types = LeaveType.query.filter(LeaveType.id.in_(list(lt_ids))).all() if lt_ids else []
    lt_map = {lt.id: lt.name for lt in leave_types}

    wb = Workbook()
    ws = wb.active
    ws.title = f'Leaves_{year_i}_{str(month_i).zfill(2)}'

    headers = ['Employee ID', 'Name', 'Department', 'Leave Type', 'Date From', 'Date To', 'Days', 'Approver', 'Recorded By', 'Reason']
    ws.append(headers)

    user_ids = {l.created_by for l in leaves if l.created_by}
    users = User.query.filter(User.id.in_(list(user_ids))).all() if user_ids else []
    user_map = {u.id: u.username for u in users}

    for l in leaves:
        emp = emp_map.get(l.employee_id)
        if emp:
            emp_code = emp.employee_id
            emp_name = f"{emp.first_name} {emp.middle_name or ''} {emp.last_name}".strip()
            department = emp.department or ''
        else:
            emp_code = ''
            emp_name = ''
            department = ''
        leave_type_name = lt_map.get(l.leave_type_id, (LeaveType.query.get(l.leave_type_id).name if l.leave_type_id else ''))
        approver = getattr(l, 'approver', '') or ''
        recorded_by = user_map.get(l.created_by, '')
        row = [
            emp_code,
            emp_name,
            department,
            leave_type_name,
            l.date_from.isoformat() if l.date_from else '',
            l.date_to.isoformat() if l.date_to else '',
            l.days,
            approver,
            recorded_by,
            l.reason or ''
        ]
        ws.append(row)

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                val = str(cell.value or '')
            except Exception:
                val = ''
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max_len + 4

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    filename = f'month_report_{year_i}_{str(month_i).zfill(2)}.xlsx'
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=filename)

# ---------- Backup & Restore ----------
@app.route('/backup')
@login_required
def backup():
    """Create timestamped backup and send to user (any logged-in user)."""
    db_path = DB_PATH
    if not os.path.exists(db_path) or os.path.getsize(db_path) == 0:
        flash('DB not found or empty', 'danger')
        return redirect(url_for('index'))

    timestamp = dt.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f'app_backup_{timestamp}.db'
    backup_path = os.path.join(BACKUPS_FOLDER, backup_name)
    try:
        shutil.copy2(db_path, backup_path)
    except Exception as e:
        flash(f'Failed to create backup: {e}', 'danger')
        return redirect(url_for('index'))

    try:
        return send_file(backup_path, as_attachment=True, download_name=backup_name)
    except Exception as e:
        flash(f'Backup created but failed to send file: {e}', 'warning')
        return redirect(url_for('index'))

@app.route('/restore', methods=['GET','POST'])
@login_required
def restore():
    """
    Restore the DB from an uploaded file.
    Only developer or admin_master allowed.
    """
    if session.get('role') not in ('developer', 'admin_master'):
        flash('Restore requires developer or admin_master role', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        if 'db_file' not in request.files:
            flash('No file uploaded', 'danger')
            return redirect(url_for('restore'))
        f = request.files['db_file']
        if f.filename == '':
            flash('No selected file', 'danger')
            return redirect(url_for('restore'))
        filename = secure_filename(f.filename)
        dest = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(dest)

        try:
            if os.path.exists(DB_PATH):
                pre_bak = os.path.join(BACKUPS_FOLDER, 'pre_restore_' + dt.now().strftime('%Y%m%d_%H%M%S') + '.db')
                shutil.copy2(DB_PATH, pre_bak)
            shutil.copy2(dest, DB_PATH)
        except Exception as e:
            flash(f'Failed to restore DB: {e}', 'danger')
            return redirect(url_for('restore'))

        flash('Database restored. Please restart the application to ensure a clean state.', 'success')
        return redirect(url_for('index'))

    return render_template('restore.html')

@app.route('/employees/import', methods=['GET', 'POST'])
@login_required
def import_employees():
    if not has_permission('can_edit_employee') and session.get('role') not in ('developer','admin_master'):
        flash('Import employees requires permission to edit employees', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        f = request.files.get('file')
        if not f or f.filename == '':
            flash('No file uploaded', 'danger')
            return redirect(url_for('import_employees'))
        # only accept xlsx
        if not f.filename.lower().endswith('.xlsx'):
            flash('Only .xlsx files accepted', 'danger')
            return redirect(url_for('import_employees'))

        fname = secure_filename(f.filename)
        uid = uuid.uuid4().hex
        dest = os.path.join(app.config['UPLOAD_FOLDER'], f'import_{uid}_{fname}')
        f.save(dest)

        rows, file_errors = parse_xlsx_file(dest)
        if file_errors:
            flash('File errors: ' + '; '.join(file_errors), 'danger')
            return redirect(url_for('import_employees'))

        # determine row-level status: exists in DB? mark as error unless user picks update mode (we default to reject existing)
        to_create = []
        to_reject = []
        for r in rows:
            empid = r['data']['employee_id']
            if Employee.query.filter_by(employee_id=empid).first():
                r['errors'].append('Employee exists in database (duplicate). Use Edit or enable update mode.')
                to_reject.append(r)
            elif r['errors']:
                to_reject.append(r)
            else:
                to_create.append(r)

        # Save parsed result to temp JSON for confirm (safe)
        preview_token = uuid.uuid4().hex
        preview_path = os.path.join(app.config['UPLOAD_FOLDER'], f'preview_{preview_token}.json')
        with open(preview_path, 'w', encoding='utf-8') as fh:
            json.dump({
                'source_filename': fname,
                'uploaded_path': dest,
                'rows': rows
            }, fh, default=str)

        # render preview template
        return render_template('import_preview.html',
                               preview_token=preview_token,
                               to_create=to_create,
                               to_reject=to_reject,
                               total=len(rows),
                               created_count=len(to_create),
                               rejected_count=len(to_reject))

    # GET -> show upload form
    return render_template('import_employees.html')

# ---------- Run server ----------
if __name__ == '__main__':
    with app.app_context():
        try:
            User.query.update({User.session_token: None})
            db.session.commit()
        except Exception:
            pass
    app.run(host='127.0.0.1', port=5000, debug=True)
