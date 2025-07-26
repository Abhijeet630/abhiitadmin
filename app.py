from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps # Import for decorators

# Initialize the Flask application
app = Flask(__name__)

# Database Configuration
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'your_super_secret_key_here' # IMPORTANT: Change this to a strong, random, and secret key in production!

# Initialize SQLAlchemy object
db = SQLAlchemy(app)

# Initialize Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login' # This tells Flask-Login where your login route is
login_manager.login_message_category = 'warning' # Category for the default "Please log in..." message

# Make the datetime object globally available in Jinja2 templates
@app.context_processor
def inject_datetime():
    return {'datetime': datetime}

# --- Database Models ---

# User Model for authentication and roles
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=True) # Added email for future password reset
    password_hash = db.Column(db.String(128))
    role = db.Column(db.String(80), default='user') # 'user' or 'admin'

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.username}>'

# ComputerSystem Model (Existing)
class ComputerSystem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    floor = db.Column(db.String(100))
    department = db.Column(db.String(100))
    host_name = db.Column(db.String(100), nullable=False)
    employee_name = db.Column(db.String(100))
    ip_address = db.Column(db.String(100))
    operating_system = db.Column(db.String(100))
    product_model = db.Column(db.String(100))
    pc_type = db.Column(db.String(100))
    processor = db.Column(db.String(100))
    price = db.Column(db.String(100))
    ram_size = db.Column(db.String(100))
    hard_disk_type = db.Column(db.String(100))
    hard_disk_size = db.Column(db.String(100))
    hard_disk_sn = db.Column(db.String(100))
    ssd_disk_type = db.Column(db.String(100))
    ssd_disk_size = db.Column(db.String(100))
    ssd_hard_disk_sn = db.Column(db.String(100))
    adapter_mac_address = db.Column(db.String(100))
    external_lancard = db.Column(db.String(100))
    display_make_model = db.Column(db.String(100))
    display_serial_number = db.Column(db.String(100))

# Router Model (Existing)
class Router(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    department = db.Column(db.String(100))
    router_name = db.Column(db.String(100))
    router_model_name = db.Column(db.String(100))
    serial_no = db.Column(db.String(100))
    router_connected = db.Column(db.String(100))
    price_list = db.Column(db.String(100))


# --- Flask-Login User Loader ---
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# --- Login, Logout Routes ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        flash('You are already logged in!', 'info')
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()

        if user and user.check_password(password):
            login_user(user)
            flash('Logged in successfully!', 'success')
            next_page = request.args.get('next') # Redirect to the page user tried to access
            return redirect(next_page or url_for('index'))
        else:
            flash('Invalid username or password.', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required # Only logged-in users can log out
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))


# --- Admin-only Decorator ---
def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            flash('You do not have administrative access to view this page.', 'danger')
            return redirect(url_for('index')) # Redirect to home if not admin
        return f(*args, **kwargs)
    return decorated_function


# --- Admin User Management Routes ---

@app.route('/admin/users')
@login_required
@admin_required
def admin_users():
    users = User.query.all()
    return render_template('admin_users.html', users=users)

@app.route('/admin/users/add', methods=['GET', 'POST'])
@login_required
@admin_required
def add_user():
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        role = request.form.get('role', 'user') # Default to 'user' if not specified

        if not username or not password:
            flash('Username and Password are required.', 'danger')
            return render_template('admin_user_form.html', user=None, title='Add New User')

        if User.query.filter_by(username=username).first():
            flash('Username already exists.', 'danger')
            return render_template('admin_user_form.html', user=None, title='Add New User')
        if email and User.query.filter_by(email=email).first():
            flash('Email already exists.', 'danger')
            return render_template('admin_user_form.html', user=None, title='Add New User')

        new_user = User(username=username, email=email, role=role)
        new_user.set_password(password)
        try:
            db.session.add(new_user)
            db.session.commit()
            flash(f'User "{username}" added successfully!', 'success')
            return redirect(url_for('admin_users'))
        except Exception as e:
            flash(f'Error adding user: {e}', 'danger')
            db.session.rollback()
    return render_template('admin_user_form.html', user=None, title='Add New User')

@app.route('/admin/users/edit/<int:id>', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_user(id):
    user = User.query.get_or_404(id)
    if request.method == 'POST':
        user.username = request.form.get('username')
        user.email = request.form.get('email')
        new_password = request.form.get('password') # Optional password change
        user.role = request.form.get('role')

        # Prevent changing username/email to an already existing one
        if User.query.filter(User.username == user.username, User.id != id).first():
            flash('Username already exists.', 'danger')
            return render_template('admin_user_form.html', user=user, title='Edit User')
        if user.email and User.query.filter(User.email == user.email, User.id != id).first():
            flash('Email already exists.', 'danger')
            return render_template('admin_user_form.html', user=user, title='Edit User')

        if new_password:
            user.set_password(new_password)

        try:
            db.session.commit()
            flash(f'User "{user.username}" updated successfully!', 'success')
            return redirect(url_for('admin_users'))
        except Exception as e:
            flash(f'Error updating user: {e}', 'danger')
            db.session.rollback()
    return render_template('admin_user_form.html', user=user, title='Edit User')

@app.route('/admin/users/delete/<int:id>', methods=['POST'])
@login_required
@admin_required
def delete_user(id):
    user = User.query.get_or_404(id)
    if user.id == current_user.id:
        flash('You cannot delete your own account.', 'danger')
        return redirect(url_for('admin_users'))
    if user.username == 'admin' and user.id == 1: # Prevent deleting the initial admin if it's user id 1
        flash('Cannot delete the primary admin account.', 'danger')
        return redirect(url_for('admin_users'))

    try:
        db.session.delete(user)
        db.session.commit()
        flash(f'User "{user.username}" deleted successfully!', 'success')
    except Exception as e:
        flash(f'Error deleting user: {e}', 'danger')
        db.session.rollback()
    return redirect(url_for('admin_users'))


# --- Existing Routes, now protected ---

# Home route
@app.route('/')
@login_required # Protect this route
def index():
    return render_template('index.html')

# System Info List
@app.route('/system_info')
@login_required # Protect this route
def system_info_list():
    systems = ComputerSystem.query.order_by(ComputerSystem.id.desc()).all()
    return render_template('system_info.html', systems=systems)

# Add System Info Route
@app.route('/system_info/add', methods=['GET', 'POST'])
@login_required # Protect this route
def add_system_info():
    if request.method == 'POST':
        new_system = ComputerSystem(
            floor=request.form.get('floor'),
            department=request.form.get('department'),
            host_name=request.form.get('host_name'),
            employee_name=request.form.get('employee_name'),
            ip_address=request.form.get('ip_address'),
            operating_system=request.form.get('operating_system'),
            product_model=request.form.get('product_model'),
            pc_type=request.form.get('pc_type'),
            processor=request.form.get('processor'),
            price=request.form.get('price'),
            ram_size=request.form.get('ram_size'),
            hard_disk_type=request.form.get('hard_disk_type'),
            hard_disk_size=request.form.get('hard_disk_size'),
            hard_disk_sn=request.form.get('hard_disk_sn'),
            ssd_disk_type=request.form.get('ssd_disk_type'),
            ssd_disk_size=request.form.get('ssd_disk_size'),
            ssd_hard_disk_sn=request.form.get('ssd_hard_disk_sn'),
            adapter_mac_address=request.form.get('adapter_mac_address'),
            external_lancard=request.form.get('external_lancard'),
            display_make_model=request.form.get('display_make_model'),
            display_serial_number=request.form.get('display_serial_number')
        )
        try:
            db.session.add(new_system)
            db.session.commit()
            flash('Computer System Information added successfully!', 'success')
            return redirect(url_for('system_info_list'))
        except Exception as e:
            flash(f'Error adding system information: {e}', 'danger')
            db.session.rollback()
    return render_template('system_info_form.html', title='Add New System Information')

# Edit System Info Route
@app.route('/system_info/edit/<int:id>', methods=['GET', 'POST'])
@login_required # Protect this route
def edit_system_info(id):
    system = ComputerSystem.query.get_or_404(id)
    if request.method == 'POST':
        system.floor = request.form.get('floor')
        system.department = request.form.get('department')
        system.host_name = request.form.get('host_name')
        system.employee_name = request.form.get('employee_name')
        system.ip_address = request.form.get('ip_address')
        system.operating_system = request.form.get('operating_system')
        system.product_model = request.form.get('product_model')
        system.pc_type = request.form.get('pc_type')
        system.processor = request.form.get('processor')
        system.price = request.form.get('price')
        system.ram_size = request.form.get('ram_size')
        system.hard_disk_type = request.form.get('hard_disk_type')
        system.hard_disk_size = request.form.get('hard_disk_size')
        system.hard_disk_sn = request.form.get('hard_disk_sn')
        system.ssd_disk_type = request.form.get('ssd_disk_type')
        system.ssd_disk_size = request.form.get('ssd_disk_size')
        system.ssd_hard_disk_sn = request.form.get('ssd_hard_disk_sn')
        system.adapter_mac_address = request.form.get('adapter_mac_address')
        system.external_lancard = request.form.get('external_lancard')
        system.display_make_model = request.form.get('display_make_model')
        system.display_serial_number = request.form.get('display_serial_number')

        try:
            db.session.commit()
            flash('Computer System Information updated successfully!', 'success')
            return redirect(url_for('system_info_list'))
        except Exception as e:
            flash(f'Error updating system information: {e}', 'danger')
            db.session.rollback()
    return render_template('system_info_form.html', system=system, title='Edit System Information')

# Delete System Info Route
@app.route('/system_info/delete/<int:id>', methods=['POST'])
@login_required # Protect this route
def delete_system_info(id):
    system = ComputerSystem.query.get_or_404(id)
    try:
        db.session.delete(system)
        db.session.commit()
        flash('Computer System Information deleted successfully!', 'success')
    except Exception as e:
        flash(f'Error deleting system information: {e}', 'danger')
        db.session.rollback()
    return redirect(url_for('system_info_list'))


# Router Info List
@app.route('/router_info')
@login_required # Protect this route
def router_info_list():
    routers = Router.query.order_by(Router.id.desc()).all()
    return render_template('router_info.html', routers=routers)

# Add Router Information Route
@app.route('/router_info/add', methods=['GET', 'POST'])
@login_required # Protect this route
def add_router_info():
    if request.method == 'POST':
        new_router = Router(
            department=request.form.get('department'),
            router_name=request.form.get('router_name'),
            router_model_name=request.form.get('router_model_name'),
            serial_no=request.form.get('serial_no'),
            router_connected=request.form.get('router_connected'),
            price_list=request.form.get('price_list')
        )
        try:
            db.session.add(new_router)
            db.session.commit()
            flash('Router Information added successfully!', 'success')
            return redirect(url_for('router_info_list'))
        except Exception as e:
            flash(f'Error adding router information: {e}', 'danger')
            db.session.rollback()
    return render_template('router_info_form.html', title='Add New Router Information')

# Edit Router Information Route
@app.route('/router_info/edit/<int:id>', methods=['GET', 'POST'])
@login_required # Protect this route
def edit_router_info(id):
    router = Router.query.get_or_404(id)
    if request.method == 'POST':
        router.department = request.form.get('department')
        router.router_name = request.form.get('router_name')
        router.router_model_name = request.form.get('router_model_name')
        router.serial_no = request.form.get('serial_no')
        router.router_connected = request.form.get('router_connected')
        router.price_list = request.form.get('price_list')

        try:
            db.session.commit()
            flash('Router Information updated successfully!', 'success')
            return redirect(url_for('router_info_list'))
        except Exception as e:
            flash(f'Error updating router information: {e}', 'danger')
            db.session.rollback()
    return render_template('router_info_form.html', router=router, title='Edit Router Information')

# Delete Router Information Route
@app.route('/router_info/delete/<int:id>', methods=['POST'])
@login_required # Protect this route
def delete_router_info(id):
    router = Router.query.get_or_404(id)
    try:
        db.session.delete(router)
        db.session.commit()
        flash('Router Information deleted successfully!', 'success')
    except Exception as e:
        flash(f'Error deleting router information: {e}', 'danger')
        db.session.rollback()
    return redirect(url_for('router_info_list'))

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Create 'Abhijeet' admin user if one doesn't exist
        if not User.query.filter_by(username='Abhijeet').first():
            abhijeet_user = User(username='Abhijeet', email='abhijeet@example.com', role='admin') # Add an email
            abhijeet_user.set_password('Abhi@218')
            db.session.add(abhijeet_user)
            db.session.commit()
            print("Admin user 'Abhijeet' created with password 'Abhi@218'.")

        # Create a regular 'user' if one doesn't exist
        if not User.query.filter_by(username='testuser').first():
            test_user = User(username='testuser', email='test@example.com', role='user')
            test_user.set_password('password123')
            db.session.add(test_user)
            db.session.commit()
            print("Regular user 'testuser' created with password 'password123'.")

    app.run(debug=True, port=5000)