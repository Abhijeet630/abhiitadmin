import socket
import platform
import json
import requests
import os
import time
import sys
import psutil  # For getting current username
try:
    import win32com.client  # For Outlook automation
    print("Info: 'pywin32' imported successfully.", file=sys.stderr)
except ImportError:
    print("Error: 'pywin32' not found. Outlook COM features will be disabled.", file=sys.stderr)
    print("Please install it: pip install pywin32", file=sys.stderr)
    win32com = None
import subprocess  # For running robocopy
from datetime import datetime # Import datetime for timestamping

# --- Configuration ---
# *** Adjust this to your Flask server's IP and port ***
WEB_API_URL_INFO = "http://127.0.0.1:5000/outlook/api/system_info"  # Endpoint to send system/PST info
WEB_API_URL_BACKUP_STATUS = "http://127.0.0.1:5000/outlook/api/backup_status"  # Endpoint to send backup status
WEB_API_URL_BACKUP_REQUEST = "http://127.0.0.1:5000/outlook/api/get_backup_request" # Endpoint to poll for backup requests

# *** Adjust this to your NAS share UNC path ***
NAS_BACKUP_PATH = r"\\172.16.17.162\data\IT ADMIN"
# Example: r"\\192.168.1.100\SharedFolder\PST_Backups"
# Make sure the user running this script has write access to the NAS path.

# --- Helper Functions ---
def get_system_info():
    """Collects hostname, IP address, OS details, and current username."""
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    current_username = None
    try:
        current_username = psutil.Process(os.getpid()).username()
    except Exception:
        current_username = os.getlogin()  # Fallback

    return {
        "Hostname": hostname,
        "IPAddress": ip_address,
        "OS": platform.system(),
        "OSRelease": platform.release(),
        "Username": current_username
    }

def get_outlook_info():
    """
    Uses Outlook COM object to collect Outlook account details (type, email)
    and data file paths (PST/OST). Returns lists of accounts and data files.
    """
    outlook_accounts = []
    pst_ost_files_from_outlook = []

    if win32com is None:
        print("Info: Outlook COM access skipped because 'pywin32' is unavailable.", file=sys.stderr)
        return outlook_accounts, pst_ost_files_from_outlook

    try:
        print("Info: Trying to access Outlook Application COM object...", file=sys.stderr)
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        print("Info: Successfully accessed Outlook Namespace.", file=sys.stderr)

        # Get Outlook account details
        if namespace.Accounts.Count > 0:
            print(f"Info: Found {namespace.Accounts.Count} Outlook accounts.", file=sys.stderr)
            for account in namespace.Accounts:
                account_type = "Unknown"
                if account.AccountType == 0: account_type = "Exchange"
                elif account.AccountType == 1: account_type = "IMAP"
                elif account.AccountType == 2: account_type = "POP3"
                elif account.AccountType == 3: account_type = "HTTP/Outlook.com"
                elif account.AccountType == 4: account_type = "Other"

                default_store_path = None
                try:
                    if hasattr(account, 'DefaultStore') and account.DefaultStore and hasattr(account.DefaultStore, 'FilePath'):
                        default_store_path = account.DefaultStore.FilePath
                    else:
                        print(f"Warning: DefaultStore or FilePath not found for account {account.DisplayName}", file=sys.stderr)
                except Exception as ex:
                    print(f"Warning: Failed to get DefaultStore for account {account.DisplayName}: {ex}", file=sys.stderr)

                outlook_accounts.append({
                    "AccountName": account.DisplayName,
                    "EmailAddress": account.SmtpAddress,
                    "AccountType": account_type,
                    "DefaultStore": default_store_path
                })
                print(f"Info: Account added: {account.DisplayName} (Mapped Type: {account_type}, Store: {default_store_path})", file=sys.stderr)
        else:
            print("Info: No Outlook accounts found via COM object.", file=sys.stderr)

        # Get PST/OST files from Outlook stores
        if namespace.Stores.Count > 0:
            print(f"Info: Found {namespace.Stores.Count} Outlook data stores.", file=sys.stderr)
            for store in namespace.Stores:
                if store.FilePath:
                    file_type = "Unknown"
                    if store.FilePath.lower().endswith('.pst'): file_type = "PST"
                    elif store.FilePath.lower().endswith('.ost'): file_type = "OST"

                    try:
                        size_bytes = os.path.getsize(store.FilePath)
                        pst_ost_files_from_outlook.append({
                            "Name": os.path.basename(store.FilePath),
                            "Path": store.FilePath,
                            "SizeMB": round(size_bytes / (1024 * 1024), 2),
                            "Type": file_type,
                            "Source": "OutlookStore"
                        })
                    except Exception as e:
                        print(f"Warning: Failed to get file info for Outlook store file {store.FilePath}: {e}", file=sys.stderr)
                        pst_ost_files_from_outlook.append({
                            "Name": os.path.basename(store.FilePath),
                            "Path": store.FilePath,
                            "SizeMB": 0,
                            "Type": file_type,
                            "Source": "OutlookStore (Access Error)"
                        })
        else:
            print("Info: No Outlook data stores found via COM object.", file=sys.stderr)

    except Exception as e:
        print(f"Critical Error: Outlook COM access failed. This usually means Outlook is not running or pywin32 is misconfigured. Error: {e}", file=sys.stderr)
        return outlook_accounts, pst_ost_files_from_outlook
    
    return outlook_accounts, pst_ost_files_from_outlook

def find_pst_files_filesystem():
    """
    Scans common Outlook data paths on the filesystem for PST/OST files.
    """
    print("Info: Starting file system scan for PST/OST files...", file=sys.stderr)
    found_files = []
    common_outlook_data_paths = [
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Outlook'),
        os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Outlook'),
        os.path.join(os.environ.get('USERPROFILE', ''), 'Documents', 'Outlook Files')
    ]
    # Add common drive letters C-E
    for drive in ['C:', 'D:', 'E:']:
        common_outlook_data_paths.append(os.path.join(drive, 'Users', get_system_info().get("Username"), 'Documents', 'Outlook Files'))
        
    for path in common_outlook_data_paths:
        if os.path.exists(path):
            print(f"Info: Searching in {path}", file=sys.stderr)
            for root, _, files in os.walk(path):
                for file in files:
                    if file.lower().endswith(('.pst', '.ost')):
                        full_path = os.path.join(root, file)
                        try:
                            size_bytes = os.path.getsize(full_path)
                            file_type = "PST" if file.lower().endswith('.pst') else "OST"
                            found_files.append({
                                "Name": file,
                                "Path": full_path,
                                "SizeMB": round(size_bytes / (1024 * 1024), 2),
                                "Type": file_type,
                                "Source": "FilesystemScan"
                            })
                        except Exception as e:
                            print(f"Warning: Failed to get file info for file system file {full_path}: {e}", file=sys.stderr)
                            found_files.append({
                                "Name": file,
                                "Path": full_path,
                                "SizeMB": 0,
                                "Type": "Unknown",
                                "Source": "FilesystemScan (Access Error)"
                            })
    print(f"Info: Found {len(found_files)} files during filesystem scan.", file=sys.stderr)
    return found_files

def send_info_to_server(system_info, outlook_accounts, pst_files):
    """Sends collected system and PST information to the Flask server."""
    payload = {
        "system_info": system_info,
        "outlook_accounts": outlook_accounts,
        "pst_files": pst_files
    }
    try:
        response = requests.post(WEB_API_URL_INFO, json=payload, timeout=30)
        print(f"Info: Sent data to server. Status Code: {response.status_code}", file=sys.stderr)
        return response
    except requests.exceptions.RequestException as e:
        print(f"Error: Failed to send data to server: {e}", file=sys.stderr)
        return None

def send_backup_status_to_server(status_data):
    """Sends backup status information to the Flask server."""
    try:
        response = requests.post(WEB_API_URL_BACKUP_STATUS, json=status_data, timeout=60)
        print(f"Info: Sent backup status to server. Status Code: {response.status_code}", file=sys.stderr)
        return response
    except requests.exceptions.RequestException as e:
        print(f"Error: Failed to send backup status to server: {e}", file=sys.stderr)
        return None

def run_backup(file_path):
    """
    Uses robocopy to copy the specified file to the NAS path.
    Logs output and sends status to the server.
    """
    print(f"Info: Starting backup for file: {file_path}", file=sys.stderr)
    start_time = time.time()
    
    file_name = os.path.basename(file_path)
    hostname = get_system_info().get("Hostname")
    
    # Create a destination folder on the NAS for this hostname and user
    backup_dest_folder = os.path.join(NAS_BACKUP_PATH, hostname, get_system_info().get("Username"))
    
    if not os.path.exists(backup_dest_folder):
        try:
            os.makedirs(backup_dest_folder)
            print(f"Info: Created destination folder on NAS: {backup_dest_folder}", file=sys.stderr)
        except Exception as e:
            message = f"Error: Failed to create NAS directory: {e}"
            print(message, file=sys.stderr)
            end_time = time.time()
            send_backup_status_to_server({
                "hostname": hostname,
                "username": get_system_info().get("Username"),
                "file_name": file_name,
                "original_path": file_path,
                "backup_path": None,
                "status": "Failed",
                "message": message,
                "time_taken_seconds": end_time - start_time,
                "backup_timestamp": datetime.now().isoformat(),
                "robocopy_output": ""
            })
            return

    backup_dest_path = os.path.join(backup_dest_folder, file_name)
    
    log_file_path = os.path.join(os.environ.get('TEMP'), f"robocopy_log_{hostname}_{file_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.log")

    # Robocopy command
    # /E - copy subdirectories, including empty ones (useful for folders)
    # /ZB - restartable mode, and use backup mode (for permissions issues)
    # /R:5 - retry 5 times on failed copies
    # /W:5 - wait 5 seconds between retries
    # /NP - no progress
    # /LOG:file - output to log file
    # /MT:8 - multithreaded copy (8 threads)
    robocopy_command = [
        "robocopy", os.path.dirname(file_path), backup_dest_folder, file_name,
        "/V", "/MT:8", "/R:5", "/W:5", "/ZB", "/NP", "/LOG+:" + log_file_path
    ]

    # Handle source path with spaces
    if ' ' in os.path.dirname(file_path):
        robocopy_command[1] = f'"{os.path.dirname(file_path)}"'
    if ' ' in backup_dest_folder:
        robocopy_command[2] = f'"{backup_dest_folder}"'

    try:
        # Run the command with a timeout
        result = subprocess.run(robocopy_command, capture_output=True, text=True, check=False, timeout=3600)
        
        end_time = time.time()
        time_taken = end_time - start_time
        
        # Read the log file
        with open(log_file_path, 'r') as log_file:
            robocopy_output = log_file.read()

        # Robocopy exit codes: 0=Success, 1=Copied with differences, 2=Extra files, 3-7=Other issues.
        # We consider codes 0 and 1 as successful backups.
        if result.returncode <= 1:
            status = "Success"
            message = "Backup completed successfully."
        else:
            status = "Failed"
            message = f"Robocopy failed with exit code: {result.returncode}"
            
        # Send status to server
        send_backup_status_to_server({
            "hostname": hostname,
            "username": get_system_info().get("Username"),
            "file_name": file_name,
            "original_path": file_path,
            "backup_path": backup_dest_path,
            "status": status,
            "message": message,
            "time_taken_seconds": time_taken,
            "backup_timestamp": datetime.now().isoformat(),
            "robocopy_output": robocopy_output
        })
        
        print(f"Info: Backup for {file_name} finished. Status: {status}", file=sys.stderr)
        
    except FileNotFoundError:
        message = "Error: robocopy command not found. Is it in your PATH?"
        print(message, file=sys.stderr)
        send_backup_status_to_server({
            "hostname": hostname,
            "username": get_system_info().get("Username"),
            "file_name": file_name,
            "original_path": file_path,
            "backup_path": None,
            "status": "Failed",
            "message": message,
            "time_taken_seconds": 0,
            "backup_timestamp": datetime.now().isoformat(),
            "robocopy_output": ""
        })
    except subprocess.TimeoutExpired:
        message = "Error: robocopy process timed out after 1 hour."
        print(message, file=sys.stderr)
        send_backup_status_to_server({
            "hostname": hostname,
            "username": get_system_info().get("Username"),
            "file_name": file_name,
            "original_path": file_path,
            "backup_path": None,
            "status": "Failed",
            "message": message,
            "time_taken_seconds": 3600,
            "backup_timestamp": datetime.now().isoformat(),
            "robocopy_output": "Process timed out."
        })
    except Exception as e:
        message = f"Error: An unexpected error occurred during backup: {e}"
        print(message, file=sys.stderr)
        send_backup_status_to_server({
            "hostname": hostname,
            "username": get_system_info().get("Username"),
            "file_name": file_name,
            "original_path": file_path,
            "backup_path": None,
            "status": "Failed",
            "message": message,
            "time_taken_seconds": time.time() - start_time,
            "backup_timestamp": datetime.now().isoformat(),
            "robocopy_output": ""
        })

def check_for_backup_requests():
    """Polls the Flask server for new backup requests."""
    try:
        response = requests.get(WEB_API_URL_BACKUP_REQUEST)
        if response.status_code == 200:
            requests_data = response.json()
            if requests_data.get('status') == 'success':
                file_path = requests_data.get('file_path')
                if file_path:
                    print(f"Info: Received backup request for {file_path}", file=sys.stderr)
                    # Run the backup in a non-blocking way to keep the loop running
                    # This is a simple implementation; a real-world scenario might use threading or multiprocessing
                    run_backup(file_path)
    except requests.exceptions.RequestException as e:
        print(f"Warning: Failed to check for backup requests: {e}", file=sys.stderr)

# --- Main execution loop ---
def main_loop():
    """Main loop for the client to collect info and check for backup requests."""
    print("Starting client main loop...", file=sys.stderr)
    while True:
        system_info = get_system_info()
        outlook_accounts, pst_files_from_outlook = get_outlook_info()
        pst_files_from_filesystem = find_pst_files_filesystem()
        all_pst_files = pst_files_from_outlook + pst_files_from_filesystem
        
        # Deduplicate files based on path
        seen_paths = set()
        unique_pst_files = []
        for file in all_pst_files:
            if file['Path'] not in seen_paths:
                unique_pst_files.append(file)
                seen_paths.add(file['Path'])

        # Send all collected data to the server
        send_info_to_server(system_info, outlook_accounts, unique_pst_files)
        
        # Check for backup requests from the server
        check_for_backup_requests()
        
        print("Info: Sleeping for 60 seconds before next scan...", file=sys.stderr)
        time.sleep(60)

if __name__ == "__main__":
    main_loop()