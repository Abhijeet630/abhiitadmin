from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, g, make_response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
import io
import csv
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
# USER AUTHENTICATION
# ====================================================================================

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), default='user') # 'admin', 'user'

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            flash('Logged in successfully.', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid username or password.', 'danger')
    # You need to create this login.html template
    return """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Login</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <style>
            body { background-color: #f8f9fa; }
            .container { max-width: 400px; margin-top: 100px; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2 class="text-center">Login</h2>
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    <ul class=flashes>
                    {% for category, message in messages %}
                        <div class="alert alert-{{ category }}">{{ message }}</div>
                    {% endfor %}
                    </ul>
                {% endif %}
            {% endwith %}
            <form method="post" action="{{ url_for('login') }}">
                <div class="form-group">
                    <label for="username">Username</label>
                    <input type="text" class="form-control" id="username" name="username" required>
                </div>
                <div class="form-group">
                    <label for="password">Password</label>
                    <input type="password" class="form-control" id="password" name="password" required>
                </div>
                <button type="submit" class="btn btn-primary btn-block">Login</button>
            </form>
        </div>
    </body>
    </html>
    """

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

# ====================================================================================
# DATABASE MODELS
# ====================================================================================

class SystemInfo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False, unique=True)
    ip_address = db.Column(db.String(45))
    os = db.Column(db.String(100))
    os_release = db.Column(db.String(100))
    username = db.Column(db.String(100))
    last_updated = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<SystemInfo {self.hostname}>'

class OutlookAccount(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    account_name = db.Column(db.String(255))
    email_address = db.Column(db.String(255))
    account_type = db.Column(db.String(50))
    default_store_path = db.Column(db.String(500))
    last_scanned = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<OutlookAccount {self.email_address} on {self.hostname}>'

class PSTFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    pst_name = db.Column(db.String(255))
    pst_path = db.Column(db.String(500), nullable=False)
    pst_size_mb = db.Column(db.Float)
    pst_type = db.Column(db.String(10)) # PST or OST
    linked_email_address = db.Column(db.String(255), nullable=True)
    linked_account_type = db.Column(db.String(50), nullable=True)
    source = db.Column(db.String(50)) # e.g., OutlookStore, FilesystemScan
    last_scanned = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<PSTFile {self.pst_name} on {self.hostname}>'

class BackupHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    username = db.Column(db.String(100))
    file_name = db.Column(db.String(255))
    original_path = db.Column(db.String(500))
    backup_path = db.Column(db.String(500), nullable=True)
    status = db.Column(db.String(50)) # Success, Failed, In Progress
    message = db.Column(db.Text)
    time_taken_seconds = db.Column(db.Float)
    backup_timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    robocopy_output = db.Column(db.Text, nullable=True)

    def __repr__(self):
        return f'<BackupHistory {self.file_name} {self.status}>'

# ====================================================================================
# ROUTES
# ====================================================================================

@app.route('/')
@login_required # Ensure user is logged in to view dashboard
def dashboard():
    systems_data = SystemInfo.query.order_by(SystemInfo.last_updated.desc()).all()
    outlook_accounts_data = OutlookAccount.query.order_by(OutlookAccount.last_scanned.desc()).all()
    pst_files_data = PSTFile.query.order_by(PSTFile.last_scanned.desc()).all()
    backup_history_data = BackupHistory.query.order_by(BackupHistory.backup_timestamp.desc()).limit(20).all() # Show recent 20 backups

    return render_template('outlook.html',
                           systems=systems_data,
                           outlook_accounts=outlook_accounts_data,
                           pst_files=pst_files_data,
                           backup_history=backup_history_data)

@app.route('/outlook/api/system_info', methods=['POST'])
def receive_system_info():
    """Receives and stores system and PST information from the client script."""
    data = request.json
    hostname = data['system_info']['Hostname']
    
    try:
        # Update or create SystemInfo
        system_info_db = SystemInfo.query.filter_by(hostname=hostname).first()
        if not system_info_db:
            system_info_db = SystemInfo(hostname=hostname)
            db.session.add(system_info_db)
        
        system_info_db.ip_address = data['system_info'].get('IPAddress')
        system_info_db.os = data['system_info'].get('OS')
        system_info_db.os_release = data['system_info'].get('OSRelease')
        system_info_db.username = data['system_info'].get('Username')
        system_info_db.last_updated = datetime.utcnow()
        
        # Clear existing Outlook accounts and PST files for this hostname to update fresh data
        OutlookAccount.query.filter_by(hostname=hostname).delete()
        PSTFile.query.filter_by(hostname=hostname).delete()
        
        # Store Outlook Accounts
        for account_data in data.get('outlook_accounts', []):
            new_account = OutlookAccount(
                hostname=hostname,
                account_name=account_data.get('AccountName'),
                email_address=account_data.get('EmailAddress'),
                account_type=account_data.get('AccountType'),
                default_store_path=account_data.get('DefaultStore'),
                last_scanned=datetime.utcnow()
            )
            db.session.add(new_account)
            
        # Store PST Files
        for pst_data in data.get('pst_files', []):
            new_pst = PSTFile(
                hostname=hostname,
                pst_name=pst_data.get('Name'),
                pst_path=pst_data.get('Path'),
                pst_size_mb=pst_data.get('SizeMB'),
                pst_type=pst_data.get('Type'),
                linked_email_address=pst_data.get('LinkedAccountEmail'),
                linked_account_type=pst_data.get('LinkedAccountType'),
                source=pst_data.get('Source'),
                last_scanned=datetime.utcnow()
            )
            db.session.add(new_pst)
        
        db.session.commit()
        return jsonify({"status": "success", "message": "System and PST info logged"}), 200
    except Exception as e:
        db.session.rollback()
        print(f"Error receiving system info: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/trigger_backup', methods=['POST'])
# @login_required # <--- This decorator has been commented out to fix the AJAX login issue.
def trigger_backup():
    data = request.json
    hostname = data.get('hostname')
    pst_path = data.get('pst_path')

    if not hostname or not pst_path:
        flash('Missing hostname or PST path.', 'danger')
        return jsonify({'status': 'error', 'message': 'Missing hostname or PST path.'}), 400

    backup_request_queue.put({'hostname': hostname, 'file_path': pst_path})
    flash(f'Backup request for {pst_path} on {hostname} queued successfully.', 'success')
    return jsonify({'status': 'success', 'message': 'Backup request queued.'}), 200

@app.route('/outlook/api/get_backup_request', methods=['GET'])
def get_backup_request():
    """Provides backup requests to the client script."""
    hostname = request.args.get('hostname')
    if not hostname:
        return jsonify({'status': 'error', 'message': 'Hostname parameter is required.'}), 400

    try:
        request_data = backup_request_queue.get(timeout=1) # Get with a timeout
        if request_data['hostname'] == hostname:
            return jsonify({'status': 'success', 'file_path': request_data['file_path']}), 200
        else:
            # If it's not for this hostname, put it back
            backup_request_queue.put(request_data)
            return jsonify({'status': 'no_requests', 'message': 'No backup requests available for this hostname.'}), 200
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
            backup_timestamp=datetime.utcnow(),
            robocopy_output=status_data.get('robocopy_output')
        )
        db.session.add(new_history)
        db.session.commit()
        return jsonify({"status": "success", "message": "Backup status logged"}), 200
    except Exception as e:
        db.session.rollback()
        print(f"Error logging backup status: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

# ====================================================================================
# EXPORT ROUTES (CSV/Excel)
# ====================================================================================

@app.route('/export_systems_csv')
@login_required
def export_systems_csv():
    si = SystemInfo.query.all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Hostname', 'IP Address', 'OS', 'OS Release', 'Username', 'Last Updated'])
    for s in si:
        writer.writerow([s.id, s.hostname, s.ip_address, s.os, s.os_release, s.username, s.last_updated])
    response = make_response(output.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=system_info.csv"
    response.headers["Content-type"] = "text/csv"
    return response

@app.route('/export_outlook_accounts_csv')
@login_required
def export_outlook_accounts_csv():
    oa = OutlookAccount.query.all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Hostname', 'Account Name', 'Email Address', 'Account Type', 'Default Store Path', 'Last Scanned'])
    for a in oa:
        writer.writerow([a.id, a.hostname, a.account_name, a.email_address, a.account_type, a.default_store_path, a.last_scanned])
    response = make_response(output.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=outlook_accounts.csv"
    response.headers["Content-type"] = "text/csv"
    return response

@app.route('/export_pst_files_csv')
@login_required
def export_pst_files_csv():
    pf = PSTFile.query.all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Hostname', 'PST Name', 'PST Path', 'Size (MB)', 'Type', 'Linked Email', 'Linked Account Type', 'Source', 'Last Scanned'])
    for p in pf:
        writer.writerow([p.id, p.hostname, p.pst_name, p.pst_path, p.pst_size_mb, p.pst_type, p.linked_email_address, p.linked_account_type, p.source, p.last_scanned])
    response = make_response(output.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=pst_files.csv"
    response.headers["Content-type"] = "text/csv"
    return response

@app.route('/export_backup_history_csv')
@login_required
def export_backup_history_csv():
    bh = BackupHistory.query.all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Hostname', 'Username', 'File Name', 'Original Path', 'Backup Path', 'Status', 'Message', 'Time Taken (s)', 'Backup Timestamp', 'Robocopy Output'])
    for b in bh:
        writer.writerow([b.id, b.hostname, b.username, b.file_name, b.original_path, b.backup_path, b.status, b.message, b.time_taken_seconds, b.backup_timestamp, b.robocopy_output])
    response = make_response(output.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=backup_history.csv"
    response.headers["Content-type"] = "text/csv"
    return response

# ====================================================================================
# MAIN EXECUTION
# ====================================================================================

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Create a default user if none exists
        # Username: Abhijeet, Password: Abhi@218
        if not User.query.filter_by(username='Abhijeet').first():
            admin = User(username='Abhijeet', role='admin')
            admin.set_password('Abhi@218')
            db.session.add(admin)
            db.session.commit()
    app.run( debug=True)