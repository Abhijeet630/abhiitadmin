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

app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///combined_db.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'your_super_secret_key_here'

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message_category = 'warning'

backup_request_queue = queue.Queue()

@app.context_processor
def inject_datetime():
    return {'datetime': datetime}

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), default='user')

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
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

class SystemInfo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False, unique=True)
    ip_address = db.Column(db.String(45))
    os = db.Column(db.String(100))
    os_release = db.Column(db.String(100))
    username = db.Column(db.String(100))
    last_updated = db.Column(db.DateTime, default=datetime.utcnow)

class OutlookAccount(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    account_name = db.Column(db.String(255))
    email_address = db.Column(db.String(255))
    account_type = db.Column(db.String(50))
    default_store_path = db.Column(db.String(500))
    last_scanned = db.Column(db.DateTime, default=datetime.utcnow)

class PSTFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    pst_name = db.Column(db.String(255))
    pst_path = db.Column(db.String(500), nullable=False)
    pst_size_mb = db.Column(db.Float)
    pst_type = db.Column(db.String(10))
    linked_email_address = db.Column(db.String(255), nullable=True)
    linked_account_type = db.Column(db.String(50), nullable=True)
    source = db.Column(db.String(50))
    last_scanned = db.Column(db.DateTime, default=datetime.utcnow)

class BackupHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    hostname = db.Column(db.String(255), nullable=False)
    username = db.Column(db.String(100))
    file_name = db.Column(db.String(255))
    original_path = db.Column(db.String(500))
    backup_path = db.Column(db.String(500), nullable=True)
    status = db.Column(db.String(50))
    message = db.Column(db.Text)
    time_taken_seconds = db.Column(db.Float)
    backup_timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    robocopy_output = db.Column(db.Text, nullable=True)

@app.route('/')
@login_required
def dashboard():
    systems_data = SystemInfo.query.order_by(SystemInfo.last_updated.desc()).all()
    outlook_accounts_data = OutlookAccount.query.order_by(OutlookAccount.last_scanned.desc()).all()
    pst_files_data = PSTFile.query.order_by(PSTFile.last_scanned.desc()).all()
    backup_history_data = BackupHistory.query.order_by(BackupHistory.backup_timestamp.desc()).limit(20).all()

    return render_template('outlook.html',
                           systems=systems_data,
                           outlook_accounts=outlook_accounts_data,
                           pst_files=pst_files_data,
                           backup_history=backup_history_data)

@app.route('/outlook/api/system_info', methods=['POST'])
def receive_system_info():
    data = request.json
    hostname = data['system_info']['Hostname']
    try:
        system_info_db = SystemInfo.query.filter_by(hostname=hostname).first()
        if not system_info_db:
            system_info_db = SystemInfo(hostname=hostname)
            db.session.add(system_info_db)

        system_info_db.ip_address = data['system_info'].get('IPAddress')
        system_info_db.os = data['system_info'].get('OS')
        system_info_db.os_release = data['system_info'].get('OSRelease')
        system_info_db.username = data['system_info'].get('Username')
        system_info_db.last_updated = datetime.utcnow()

        OutlookAccount.query.filter_by(hostname=hostname).delete()
        PSTFile.query.filter_by(hostname=hostname).delete()

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
@login_required
def trigger_backup():
    data = request.get_json()
    hostname = data.get('hostname')
    pst_path = data.get('pst_path')

    if not hostname or not pst_path:
        return jsonify({'status': 'error', 'message': 'Missing hostname or PST path.'}), 400

    backup_request_queue.put({'hostname': hostname, 'file_path': pst_path})
    return jsonify({'status': 'success', 'message': 'Backup request queued.'}), 200

@app.route('/outlook/api/get_backup_request', methods=['GET'])
def get_backup_request():
    hostname = request.args.get('hostname')
    if not hostname:
        return jsonify({'status': 'error', 'message': 'Hostname parameter is required.'}), 400

    try:
        request_data = backup_request_queue.get(timeout=1)
        if request_data['hostname'] == hostname:
            return jsonify({'status': 'success', 'file_path': request_data['file_path']}), 200
        else:
            backup_request_queue.put(request_data)
            return jsonify({'status': 'no_requests', 'message': 'No backup requests available for this hostname.'}), 200
    except queue.Empty:
        return jsonify({'status': 'no_requests', 'message': 'No backup requests available.'}), 200

@app.route('/outlook/api/backup_status', methods=['POST'])
def receive_backup_status():
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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
