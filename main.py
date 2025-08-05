from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, g, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
import io
import csv
import openpyxl
import subprocess
import sys
import json
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
import threading
import queue

# Initialize the Flask application
app = Flask(__name__)

# Database Configuration
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///combined_db.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'your_super_secret_key_here' # IMPORTANT: Change this to a strong, random, and secret key in production!

# Initialize SQLAlchemy and Flask-Login
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login' # This is the function name for the login route
login_manager.login_message_category = 'warning'

# A queue to hold backup requests
backup_request_queue = queue.Queue()

# Make the datetime object globally available in Jinja2 templates
@app.context_processor
def inject_datetime():
    return {'datetime': datetime}

# ====================================================================================
# USER AUTHENTICATION MODEL
# ====================================================================================
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(50), default='user')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# ====================================================================================
# DATABASE MODELS FOR ALL APPLICATIONS
# ====================================================================================

# --- Models from System Inventory (System_inventory.py) ---
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
    
    def __repr__(self):
        return f'<ComputerSystem {self.host_name}>'

class Router(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    department = db.Column(db.String(100))
    router_name = db.Column(db.String(100))
    router_model_name = db.Column(db.String(100))
    serial_no = db.Column(db.String(100))
    router_connected = db.Column(db.String(100))
    price_list = db.Column(db.String(100))

    def __repr__(self):
        return f'<Router {self.router_name}>'

# --- Models from Web Patch (Web_Patch.py) ---
class Rack(db.Model):
    __tablename__ = 'racks'
    rack_id = db.Column(db.Integer, primary_key=True)
    rack_name = db.Column(db.String(100), unique=True, nullable=False)
    panels = db.relationship('Panel', backref='rack', lazy='dynamic', cascade="all, delete-orphan")

class Panel(db.Model):
    __tablename__ = 'panels'
    panel_id = db.Column(db.Integer, primary_key=True)
    rack_id = db.Column(db.Integer, db.ForeignKey('racks.rack_id', ondelete='CASCADE'), nullable=False)
    panel_label = db.Column(db.String(100), nullable=False)
    num_ports = db.Column(db.Integer, default=24, nullable=False)
    __table_args__ = (db.UniqueConstraint('rack_id', 'panel_label', name='_rack_panel_uc'),)

class PortDescription(db.Model):
    __tablename__ = 'port_descriptions'
    port_label = db.Column(db.String(100), primary_key=True)
    description = db.Column(db.String(255))

class PatchedConnection(db.Model):
    __tablename__ = 'patched_connections'
    id = db.Column(db.Integer, primary_key=True)
    port_a_label = db.Column(db.String(100), nullable=False)
    port_b_label = db.Column(db.String(100), nullable=False)
    __table_args__ = (db.UniqueConstraint('port_a_label', 'port_b_label', name='_patch_uc_ab'),
                      db.UniqueConstraint('port_b_label', 'port_a_label', name='_patch_uc_ba'))

class PortStatus(db.Model):
    __tablename__ = 'port_status'
    port_label = db.Column(db.String(100), primary_key=True)
    status = db.Column(db.String(10), nullable=False, default='down')

# --- New Models from Outlook.py ---
class SystemInfo(db.Model):
    __tablename__ = 'systems'
    hostname = db.Column(db.String(255), primary_key=True)
    ip_address = db.Column(db.String(100))
    os = db.Column(db.String(100))
    os_release = db.Column(db.String(100))
    username = db.Column(db.String(100))
    last_updated = db.Column(db.String(100)) # Store as string for simplicity, can be DateTime

class OutlookAccount(db.Model):
    __tablename__ = 'outlook_accounts'
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), db.ForeignKey('systems.hostname'))
    account_name = db.Column(db.String(255))
    email_address = db.Column(db.String(255))
    account_type = db.Column(db.String(100))
    default_store_path = db.Column(db.String(255))
    last_scanned = db.Column(db.String(100)) # Store as string for simplicity

class PstFile(db.Model):
    __tablename__ = 'pst_files'
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), db.ForeignKey('systems.hostname'))
    pst_name = db.Column(db.String(255))
    pst_path = db.Column(db.String(255), unique=True)
    pst_size_mb = db.Column(db.Float)
    pst_type = db.Column(db.String(100))
    source = db.Column(db.String(100))
    last_scanned = db.Column(db.String(100))
    linked_email_address = db.Column(db.String(255))
    linked_account_type = db.Column(db.String(100))

class BackupHistory(db.Model):
    __tablename__ = 'backup_history'
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), db.ForeignKey('systems.hostname'))
    username = db.Column(db.String(100))
    file_name = db.Column(db.String(255))
    original_path = db.Column(db.String(255))
    backup_path = db.Column(db.String(255))
    status = db.Column(db.String(50))
    message = db.Column(db.String(500))
    time_taken_seconds = db.Column(db.Float)
    backup_timestamp = db.Column(db.String(100))
    robocopy_output = db.Column(db.Text)


# Create database tables within the application context
with app.app_context():
    db.create_all()
    # Create a default user if none exists
    # Username: Abhijeet, Password: Abhi@218
    if not User.query.filter_by(username='Abhijeet').first():
        admin = User(username='Abhijeet', role='admin')
        admin.set_password('Abhi@218')
        db.session.add(admin)
        db.session.commit()
        print("Default admin user 'Abhijeet' created.")


# ====================================================================================
# GLOBAL CONFIGURATION FOR OUTLOOK BACKUP
# ====================================================================================
# IMPORTANT: This path refers to where the client_script.py is located
# on the client machine.
CLIENT_SCRIPT_PATH = r"C:\Users\Abhi\Desktop\Outlook PST Website\client_script.py"
PYTHON_EXE_PATH = sys.executable # Path to the Python executable running Flask

# ====================================================================================
# ROUTES FOR ALL APPLICATIONS
# ====================================================================================

# Login Route
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user, remember=request.form.get('remember'))
            flash('Login successful!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid username or password.', 'danger')
    return render_template('login.html')

# Logout Route
@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

# Homepage route (for the dashboard) - Now protected by login
@app.route('/')
@login_required
def index():
    return render_template('index.html')

# Admin Panel Route
@app.route('/admin_panel', methods=['GET', 'POST'])
@login_required
def admin_panel():
    if current_user.role != 'admin':
        flash('Access denied: Admin privileges required.', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        role = request.form.get('role') or 'user'

        if not username or not password:
            flash('username and password are required.', 'danger')
            return redirect(url_for('admin_panel'))

        existing_user = User.query.filter_by(username=username).first()
        if existing_user:
            flash('User already exists.', 'danger')
        else:
            new_user = User(username=username, role=role)
            new_user.set_password(password)
            try:
                db.session.add(new_user)
                db.session.commit()
                flash(f'User "{username}" with role "{role}" created successfully.', 'success')
            except Exception as e:
                db.session.rollback()
                flash(f'User creation failed: {e}', 'danger')
        return redirect(url_for('admin_panel'))

    users = User.query.all()
    return render_template('admin_panel.html', users=users)

@app.route('/admin_panel/edit_user/<int:user_id>', methods=['GET', 'POST'])
@login_required
def edit_user(user_id):
    if current_user.role != 'admin':
        flash('You do not have access to the admin panel.', 'danger')
        return redirect(url_for('index'))
    
    user_to_edit = User.query.get_or_404(user_id)
    if request.method == 'POST':
        new_username = request.form.get('username')
        new_password = request.form.get('password')
        new_role = request.form.get('role', 'user')

        if not new_username:
            flash('Username is required!', 'danger')
            return redirect(url_for('edit_user', user_id=user_id))
        
        existing_user = User.query.filter(User.username == new_username, User.id != user_id).first()
        if existing_user:
            flash(f'The username "{new_username}" is already in use.', 'danger')
            return redirect(url_for('edit_user', user_id=user_id))

        user_to_edit.username = new_username
        user_to_edit.role = new_role
        if new_password:
            user_to_edit.set_password(new_password)
        
        try:
            db.session.commit()
            flash(f'User "{user_to_edit.username}" updated successfully!', 'success')
            return redirect(url_for('admin_panel'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error updating user: {e}', 'danger')

    return render_template('admin_user_form.html', user=user_to_edit, title='Edit User')


@app.route('/admin_panel/delete_user/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if current_user.username != 'Abhijeet':
        flash('not access only Abhijeet user access.', 'danger')
        return redirect(url_for('index'))

    user_to_delete = User.query.get_or_404(user_id)
    if user_to_delete.username == 'Abhijeet':
        flash('only abhijeet user can delete.', 'danger')
        return redirect(url_for('admin_panel'))

    try:
        db.session.delete(user_to_delete)
        db.session.commit()
        flash(f'User "{user_to_delete.username}" deleted successfully.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'User delete failed: {e}', 'danger')
    return redirect(url_for('admin_panel'))

@app.route('/change_password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        old_password = request.form.get('old_password')
        new_password = request.form.get('new_password')
        confirm_password = request.form.get('confirm_password')

        if not current_user.check_password(old_password):
            flash('incorrect old password.', 'danger')
        elif new_password != confirm_password:
            flash('new password and confirm password do not match.', 'danger')
        else:
            current_user.set_password(new_password)
            try:
                db.session.commit()
                flash('Password changed successfully.', 'success')
                return redirect(url_for('index'))
            except Exception as e:
                db.session.rollback()
                flash(f'Password change failed: {e}', 'danger')
    return render_template('change_password.html')

with app.app_context():
    db.create_all()
    if not User.query.filter_by(username='Abhijeet').first():
        admin = User(username='Abhijeet', role='admin')
        admin.set_password('Abhi@218')
        db.session.add(admin)
        db.session.commit()
        print("Default admin user 'Abhijeet' created.")

# --- Routes for System Inventory (from System_inventory.py) ---
@app.route('/inventory')
def inventory():
    return render_template('system.html')
@app.route('/inventory/systems')
@login_required
def system_info_list():
    systems = ComputerSystem.query.all()
    return render_template('system_info.html', systems=systems)

@app.route('/inventory/systems/add', methods=['GET', 'POST'])
@login_required
def add_system_info():
    if request.method == 'POST':
        floor = request.form.get('floor')
        department = request.form.get('department')
        host_name = request.form['host_name']
        employee_name = request.form.get('employee_name')
        ip_address = request.form.get('ip_address')
        operating_system = request.form.get('operating_system')
        product_model = request.form.get('product_model')
        pc_type = request.form.get('pc_type')
        processor = request.form.get('processor')
        price = request.form.get('price')
        ram_size = request.form.get('ram_size')
        hard_disk_type = request.form.get('hard_disk_type')
        hard_disk_size = request.form.get('hard_disk_size')
        hard_disk_sn = request.form.get('hard_disk_sn')
        ssd_disk_type = request.form.get('ssd_disk_type')
        ssd_disk_size = request.form.get('ssd_disk_size')
        ssd_hard_disk_sn = request.form.get('ssd_hard_disk_sn')
        adapter_mac_address = request.form.get('adapter_mac_address')
        external_lancard = request.form.get('external_lancard')
        display_make_model = request.form.get('display_make_model')
        display_serial_number = request.form.get('display_serial_number')

        if not host_name:
            flash('Host Name is required!', 'danger')
            return redirect(url_for('add_system_info'))

        new_system = ComputerSystem(
            floor=floor, department=department, host_name=host_name, employee_name=employee_name,
            ip_address=ip_address, operating_system=operating_system, product_model=product_model,
            pc_type=pc_type, processor=processor, price=price, ram_size=ram_size,
            hard_disk_type=hard_disk_type, hard_disk_size=hard_disk_size, hard_disk_sn=hard_disk_sn,
            ssd_disk_type=ssd_disk_type, ssd_disk_size=ssd_disk_size, ssd_hard_disk_sn=ssd_hard_disk_sn,
            adapter_mac_address=adapter_mac_address, external_lancard=external_lancard,
            display_make_model=display_make_model, display_serial_number=display_serial_number
        )
        try:
            db.session.add(new_system)
            db.session.commit()
            flash('System Information added successfully!', 'success')
            return redirect(url_for('system_info_list'))
        except Exception as e:
            flash(f'Error adding system information: {e}', 'danger')
            db.session.rollback()
    return render_template('system_info_form.html', title='Add System Information')

@app.route('/inventory/systems/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_system_info(id):
    system = ComputerSystem.query.get_or_404(id)
    if request.method == 'POST':
        system.floor = request.form.get('floor')
        system.department = request.form.get('department')
        system.host_name = request.form['host_name']
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

        if not system.host_name:
            flash('Host Name is required!', 'danger')
            return redirect(url_for('edit_system_info', id=id))

        try:
            db.session.commit()
            flash('System Information updated successfully!', 'success')
            return redirect(url_for('system_info_list'))
        except Exception as e:
            flash(f'Error updating system information: {e}', 'danger')
            db.session.rollback()
    return render_template('system_info_form.html', system=system, title='Edit System Information')

@app.route('/inventory/systems/delete/<int:id>', methods=['POST'])
@login_required
def delete_system_info(id):
    system = ComputerSystem.query.get_or_404(id)
    try:
        db.session.delete(system)
        db.session.commit()
        flash('System Information deleted successfully!', 'success')
    except Exception as e:
        flash(f'Error deleting system information: {e}', 'danger')
        db.session.rollback()
    return redirect(url_for('system_info_list'))

# --- Routes for Router Information (from System_inventory.py) ---
@app.route('/inventory/routers')
@login_required
def router_info_list():
    routers = Router.query.all()
    return render_template('router_info.html', routers=routers)

@app.route('/inventory/routers/add', methods=['GET', 'POST'])
@login_required
def add_router_info():
    if request.method == 'POST':
        department = request.form.get('department')
        router_name = request.form.get('router_name')
        router_model_name = request.form.get('router_model_name')
        serial_no = request.form.get('serial_no')
        router_connected = request.form.get('router_connected')
        price_list = request.form.get('price_list')

        new_router = Router(
            department=department, router_name=router_name,
            router_model_name=router_model_name, serial_no=serial_no,
            router_connected=router_connected, price_list=price_list
        )
        try:
            db.session.add(new_router)
            db.session.commit()
            flash('Router Information added successfully!', 'success')
            return redirect(url_for('router_info_list'))
        except Exception as e:
            flash(f'Error adding router information: {e}', 'danger')
            db.session.rollback()
    return render_template('router_info_form.html', title='Add Router Information')

@app.route('/inventory/routers/edit/<int:id>', methods=['GET', 'POST'])
@login_required
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

@app.route('/inventory/routers/delete/<int:id>', methods=['POST'])
@login_required
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

# --- Routes for Web Patch (from Web_Patch.py) ---
@app.route('/webpatch')
@login_required
def webpatch_index():
    return render_template('webpatch.html')
def get_webpatch_state():
    """Fetches the current application state (racks, panels, connections, descriptions)."""
    racks_data = Rack.query.order_by(Rack.rack_id).all()
    racks = []
    for rack_row in racks_data:
        panels_data = Panel.query.filter_by(rack_id=rack_row.rack_id).order_by(Panel.panel_id).all()
        panels = [{'panel_id': p.panel_id, 'panel_label': p.panel_label, 'num_ports': p.num_ports} for p in panels_data]
        racks.append({
            'rack_id': rack_row.rack_id,
            'rack_name': rack_row.rack_name,
            'panels': panels
        })
    port_descriptions = {pd.port_label: pd.description for pd in PortDescription.query.all()}
    patched_connections = [{'port_a': pc.port_a_label, 'port_b': pc.port_b_label} for pc in PatchedConnection.query.all()]
    up_ports = [ps.port_label for ps in PortStatus.query.filter_by(status='up').all()]
    return {
        'racks': racks,
        'port_descriptions': port_descriptions,
        'patched_connections': patched_connections,
        'up_ports': up_ports
    }
@app.route('/webpatch/api/state', methods=['GET'])
@login_required
def get_webpatch_state_api():
    state = get_webpatch_state()
    return jsonify(state)
@app.route('/webpatch/api/add_rack', methods=['POST'])
@login_required
def add_rack():
    data = request.json
    new_rack_name = data.get('rack_name')
    if not new_rack_name:
        return jsonify({'success': False, 'message': 'Rack name is required.'}), 400
    try:
        new_rack = Rack(rack_name=new_rack_name)
        db.session.add(new_rack)
        db.session.commit()
        return jsonify({'success': True, 'message': f'Rack {new_rack_name} added successfully.', 'state': get_webpatch_state()}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'A rack with the name "{new_rack_name}" already exists.'}), 409
@app.route('/webpatch/api/rename_rack', methods=['POST'])
@login_required
def rename_rack():
    data = request.json
    rack_id = data.get('rack_id')
    new_rack_name = data.get('new_name')
    if not all([rack_id, new_rack_name]):
        return jsonify({'success': False, 'message': 'Rack ID and new name are required.'}), 400
    rack = Rack.query.get(rack_id)
    if not rack:
        return jsonify({'success': False, 'message': 'Rack not found.'}), 404
    old_rack_name = rack.rack_name
    try:
        rack.rack_name = new_rack_name # Update related data ports_to_update = PortDescription.query.filter(PortDescription.port_label.like(f"{old_rack_name}-%")).all() for port in ports_to_update: port.port_label = port.port_label.replace(old_rack_name, new_rack_name, 1) connections_to_update = PatchedConnection.query.filter( (PatchedConnection.port_a_label.like(f"{old_rack_name}-%")) | (PatchedConnection.port_b_label.like(f"{old_rack_name}-%")) ).all() for conn in connections_to_update: if conn.port_a_label.startswith(f"{old_rack_name}-"): conn.port_a_label = conn.port_a_label.replace(old_rack_name, new_rack_name, 1) if conn.port_b_label.startswith(f"{old_rack_name}-"): conn.port_b_label = conn.port_b_label.replace(old_rack_name, new_rack_name, 1) statuses_to_update = PortStatus.query.filter(PortStatus.port_label.like(f"{old_rack_name}-%")).all() for status in statuses_to_update: status.port_label = status.port_label.replace(old_rack_name, new_rack_name, 1) db.session.commit() return jsonify({'success': True, 'message': 'Rack renamed successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to rename rack: {e}'}), 500
@app.route('/webpatch/api/delete_rack', methods=['POST'])
@login_required
def delete_rack():
    data = request.json
    rack_id = data.get('rack_id')
    if not rack_id:
        return jsonify({'success': False, 'message': 'Rack ID is required.'}), 400
    rack = Rack.query.get(rack_id)
    if not rack:
        return jsonify({'success': False, 'message': 'Rack not found.'}), 404
    try:
        db.session.delete(rack)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Rack deleted successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to delete rack: {e}'}), 500
@app.route('/webpatch/api/add_panel', methods=['POST'])
@login_required
def add_panel():
    data = request.json
    rack_id = data.get('rack_id')
    panel_label = data.get('panel_label')
    if not all([rack_id, panel_label]):
        return jsonify({'success': False, 'message': 'Rack ID and panel label are required.'}), 400
    try:
        new_panel = Panel(rack_id=rack_id, panel_label=panel_label)
        db.session.add(new_panel)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Panel added successfully.', 'state': get_webpatch_state()}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to add panel: {e}'}), 500
@app.route('/webpatch/api/update_panel_ports', methods=['POST'])
@login_required
def update_panel_ports():
    data = request.json
    panel_id = data.get('panel_id')
    num_ports = data.get('num_ports')
    if not all([panel_id, num_ports is not None]):
        return jsonify({'success': False, 'message': 'Panel ID and number of ports are required.'}), 400
    panel = Panel.query.get(panel_id)
    if not panel:
        return jsonify({'success': False, 'message': 'Panel not found.'}), 404
    try:
        panel.num_ports = num_ports
        db.session.commit()
        return jsonify({'success': True, 'message': 'Number of ports updated successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to update ports: {e}'}), 500
@app.route('/webpatch/api/delete_panel', methods=['POST'])
@login_required
def delete_panel():
    data = request.json
    panel_id = data.get('panel_id')
    if not panel_id:
        return jsonify({'success': False, 'message': 'Panel ID is required.'}), 400
    panel = Panel.query.get(panel_id)
    if not panel:
        return jsonify({'success': False, 'message': 'Panel not found.'}), 404
    try:
        db.session.delete(panel)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Panel deleted successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to delete panel: {e}'}), 500
@app.route('/webpatch/api/add_description', methods=['POST'])
@login_required
def add_description():
    data = request.json
    port_label = data.get('port_label')
    description = data.get('description')
    if not port_label:
        return jsonify({'success': False, 'message': 'Port label is required.'}), 400
    try:
        port_desc = PortDescription.query.get(port_label)
        if port_desc:
            port_desc.description = description
        else:
            new_desc = PortDescription(port_label=port_label, description=description)
            db.session.add(new_desc)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Description updated successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to update description: {e}'}), 500
@app.route('/webpatch/api/delete_description', methods=['POST'])
@login_required
def delete_description():
    data = request.json
    port_label = data.get('port_label')
    if not port_label:
        return jsonify({'success': False, 'message': 'Port label is required.'}), 400
    port_desc = PortDescription.query.get(port_label)
    if not port_desc:
        return jsonify({'success': False, 'message': 'Description not found.'}), 404
    try:
        db.session.delete(port_desc)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Description deleted successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to delete description: {e}'}), 500
@app.route('/webpatch/api/patch', methods=['POST'])
@login_required
def patch_ports():
    data = request.json
    port_a_label = data.get('port_a_label')
    port_b_label = data.get('port_b_label')
    if not all([port_a_label, port_b_label]):
        return jsonify({'success': False, 'message': 'Both port labels are required.'}), 400
    existing_connection = PatchedConnection.query.filter(
        (PatchedConnection.port_a_label == port_a_label and PatchedConnection.port_b_label == port_b_label) |
        (PatchedConnection.port_a_label == port_b_label and PatchedConnection.port_b_label == port_a_label)
    ).first()
    if existing_connection:
        return jsonify({'success': False, 'message': 'These two ports are already patched.'}), 409
    try:
        new_patch = PatchedConnection(port_a_label=port_a_label, port_b_label=port_b_label)
        db.session.add(new_patch)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Ports patched successfully.', 'state': get_webpatch_state()}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to patch ports: {e}'}), 500
@app.route('/webpatch/api/unpatch', methods=['POST'])
@login_required
def unpatch_ports():
    data = request.json
    port_a_label = data.get('port_a_label')
    port_b_label = data.get('port_b_label')
    if not all([port_a_label, port_b_label]):
        return jsonify({'success': False, 'message': 'Both port labels are required.'}), 400
    connection = PatchedConnection.query.filter(
        (PatchedConnection.port_a_label == port_a_label and PatchedConnection.port_b_label == port_b_label) |
        (PatchedConnection.port_a_label == port_b_label and PatchedConnection.port_b_label == port_a_label)
    ).first()
    if not connection:
        return jsonify({'success': False, 'message': 'Patch connection not found.'}), 404
    try:
        db.session.delete(connection)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Ports unpatched successfully.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to unpatch ports: {e}'}), 500
@app.route('/webpatch/api/set_port_status', methods=['POST'])
@login_required
def set_port_status():
    data = request.json
    port_label = data.get('port_label')
    status = data.get('status')
    if not all([port_label, status]):
        return jsonify({'success': False, 'message': 'Port label and status are required.'}), 400
    if status not in ['up', 'down']:
        return jsonify({'success': False, 'message': 'Status must be "up" or "down".'}), 400
    try:
        port_status = PortStatus.query.get(port_label)
        if port_status:
            port_status.status = status
        else:
            new_status = PortStatus(port_label=port_label, status=status)
            db.session.add(new_status)
        db.session.commit()
        return jsonify({'success': True, 'message': f'Port {port_label} status set to {status}.', 'state': get_webpatch_state()}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Failed to set port status: {e}'}), 500

# --- Routes for Outlook Management ---
@app.route('/outlook')
@login_required
def outlook_dashboard():
    systems = SystemInfo.query.order_by(SystemInfo.hostname).all()
    outlook_accounts = OutlookAccount.query.order_by(OutlookAccount.hostname).all()
    pst_files = PstFile.query.order_by(PstFile.hostname, PstFile.pst_name).all()
    backup_history = BackupHistory.query.order_by(BackupHistory.backup_timestamp.desc()).limit(20).all()
    return render_template('outlook.html', systems=systems, outlook_accounts=outlook_accounts, pst_files=pst_files, backup_history=backup_history)

@app.route('/outlook/api/system_info', methods=['POST'])
def receive_system_info():
    """Receives system and PST information from the client script and stores it in the database."""
    data = request.json
    system_data = data.get('system_info', {})
    accounts_data = data.get('outlook_accounts', [])
    pst_files_data = data.get('pst_files', [])
    
    hostname = system_data.get('Hostname')
    if not hostname:
        return jsonify({"status": "error", "message": "Hostname is required"}), 400
    
    try:
        # Update or create SystemInfo entry
        system = SystemInfo.query.get(hostname)
        if system:
            system.ip_address = system_data.get('IPAddress')
            system.os = system_data.get('OS')
            system.os_release = system_data.get('OSRelease')
            system.username = system_data.get('Username')
            system.last_updated = datetime.now().isoformat()
        else:
            system = SystemInfo(
                hostname=hostname,
                ip_address=system_data.get('IPAddress'),
                os=system_data.get('OS'),
                os_release=system_data.get('OSRelease'),
                username=system_data.get('Username'),
                last_updated=datetime.now().isoformat()
            )
            db.session.add(system)
        
        # Clear old accounts and PST files for this system
        OutlookAccount.query.filter_by(hostname=hostname).delete()
        PstFile.query.filter_by(hostname=hostname).delete()
        
        # Add new accounts
        for account_data in accounts_data:
            new_account = OutlookAccount(
                hostname=hostname,
                account_name=account_data.get('AccountName'),
                email_address=account_data.get('EmailAddress'),
                account_type=account_data.get('AccountType'),
                default_store_path=account_data.get('DefaultStore'),
                last_scanned=datetime.now().isoformat()
            )
            db.session.add(new_account)
        
        # Add new PST/OST files
        for pst_data in pst_files_data:
            # Check if this file is linked to an account
            linked_email = None
            linked_type = None
            for account in accounts_data:
                if account.get('DefaultStore') and account['DefaultStore'].lower() == pst_data['Path'].lower():
                    linked_email = account['EmailAddress']
                    linked_type = account['AccountType']
                    break

            new_pst = PstFile(
                hostname=hostname,
                pst_name=pst_data.get('Name'),
                pst_path=pst_data.get('Path'),
                pst_size_mb=pst_data.get('SizeMB'),
                pst_type=pst_data.get('Type'),
                source=pst_data.get('Source'),
                last_scanned=datetime.now().isoformat(),
                linked_email_address=linked_email,
                linked_account_type=linked_type
            )
            db.session.add(new_pst)
        
        db.session.commit()
        return jsonify({"status": "success", "message": "System info and PST files updated"}), 200
        
    except Exception as e:
        db.session.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/outlook/api/request_backup', methods=['POST'])
@login_required
def request_backup():
    """Endpoint for the web UI to request a backup."""
    data = request.get_json()
    hostname = data.get('hostname')
    file_path = data.get('file_path')

    if not all([hostname, file_path]):
        return jsonify({'status': 'error', 'message': 'Hostname and file_path are required.'}), 400
    
    # Add the request to the queue
    backup_request_queue.put({'hostname': hostname, 'file_path': file_path})
    return jsonify({'status': 'success', 'message': f'Backup request for {file_path} on {hostname} queued.'}), 200

@app.route('/outlook/api/get_backup_request', methods=['GET'])
def get_backup_request():
    """Endpoint for the client script to poll for backup requests."""
    try:
        # Get a request from the queue without blocking
        request_data = backup_request_queue.get_nowait()
        return jsonify({'status': 'success', 'request': request_data}), 200
    except queue.Empty:
        return jsonify({'status': 'no_requests', 'message': 'No backup requests available.'}), 200

@app.route('/outlook/api/backup_status', methods=['POST'])
def receive_backup_status():
    """Receives and stores the backup status from the client script."""
    status_data = request.json
    try:
        new_history = BackupHistory(
            hostname=status_data.get('hostname'),
            username=status_data.get('username'),
            file_name=status_data.get('file_name'),
            original_path=status_data.get('original_path'),
            backup_path=status_data.get('backup_path'),
            status=status_data.get('status'),
            message=status_data.get('message'),
            time_taken_seconds=status_data.get('time_taken_seconds'),
            backup_timestamp=status_data.get('backup_timestamp'),
            robocopy_output=status_data.get('robocopy_output')
        )
        db.session.add(new_history)
        db.session.commit()
        return jsonify({"status": "success", "message": "Backup status logged"}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    # You might want to run this in a production-ready server like Gunicorn
    # For development, you can use the built-in server
    app.run(host='0.0.0.0', port=5000, debug=True)