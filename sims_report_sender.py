import os
import csv
import re
import base64
import requests
import json
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import msal

# ==========================================
# --- CONFIGURATION ---
# ==========================================

SCOPES = ['https://graph.microsoft.com/Mail.Send.Shared']

# Locate the user's AppData directory to prevent Permission Denied errors
appdata_dir = os.getenv('APPDATA')
# Fallback to the user's home directory if APPDATA is not found
if appdata_dir is None:
    appdata_dir = os.path.expanduser('~')

APP_FOLDER = os.path.join(appdata_dir, 'StudentReportSender')

# Ensure the application directory exists
if not os.path.exists(APP_FOLDER):
    os.makedirs(APP_FOLDER)

CONFIG_FILE = os.path.join(APP_FOLDER, 'app_config.json')

# ==========================================
# --- CORE LOGIC FUNCTIONS ---
# ==========================================

def load_app_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as file:
                return json.load(file)
        except json.JSONDecodeError:
            print("  -> ERROR: Corrupted config file. Loading defaults.")
    return {
        "tenant_id": "",
        "client_id": "",
        "sender_email": "",
        "subject": "Student Report: {name}",
        "body": "Dear Parent/Guardian,\n\nPlease find attached the latest report.\n\nKind regards,\nSchool Administration"
    }

def save_app_config(config_data):
    with open(CONFIG_FILE, 'w') as file:
        json.dump(config_data, file, indent=4)

def load_sent_log(log_file):
    if os.path.exists(log_file):
        try:
            with open(log_file, 'r') as file:
                return set(json.load(file))
        except json.JSONDecodeError:
            print("  -> ERROR: Corrupted JSON log. Starting fresh.")
            return set()
    return set()

def save_sent_log(log_set, log_file):
    with open(log_file, 'w') as file:
        json.dump(list(log_set), file, indent=4)

def clear_sent_log(log_file):
    with open(log_file, 'w') as file:
        json.dump([], file)

def get_access_token(tenant_id, client_id, status_callback=None):
    if status_callback: status_callback("Waiting for browser authentication...")
    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    print("\nWaiting for user to authenticate in the browser...")
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" in result:
        return result["access_token"]
    else:
        error_message = result.get("error_description", result.get("error", "Unknown Error"))
        raise Exception(f"Failed to retrieve access token: {error_message}")

def load_sims_emails(csv_path):
    email_dict = {}
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            next(reader, None) 
            for row in reader:
                row += [''] * (5 - len(row))
                forename, surname = row[0].strip(), row[1].strip()
                if forename and surname:
                    current_student_key = f"{surname}-{forename}".lower()
                    if current_student_key not in email_dict:
                        email_dict[current_student_key] = set()
                    raw_emails = [row[2].strip(), row[3].strip(), row[4].strip()]
                    for email in raw_emails:
                        if email: 
                            email_dict[current_student_key].add(email)
        return email_dict
    except Exception as e:
        print(f"CRITICAL ERROR loading CSV: {e}")
        return None

def extract_name_from_filename(filename):
    match = re.search(r'^(.*?)-\d', filename)
    if match: return match.group(1).lower()
    return None

def determine_file_properties(filename):
    name, ext = os.path.splitext(filename)
    ext = ext.lower()
    if ext == '.pdf': return filename, 'application/pdf'
    elif ext == '.doc': return filename, 'application/msword'
    elif ext == '.docx': return filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif ext == '.xml': return f"{name}.doc", 'application/msword'
    else: return None, None

def send_graph_email(access_token, sender_email, parent_email, subject, body, filepath, attachment_name, mime_type):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    try:
        with open(filepath, 'rb') as f:
            encoded_file = base64.b64encode(f.read()).decode('utf-8')
    except Exception as e:
        print(f"  -> ERROR reading file {filepath}: {e}")
        return False

    email_payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": parent_email}}],
            "from": {"emailAddress": {"address": sender_email}},
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment_name,
                "contentType": mime_type,
                "contentBytes": encoded_file
            }]
        },
        "saveToSentItems": "true"
    }

    try:
        response = requests.post(url, headers=headers, json=email_payload)
        if response.status_code == 202:
            print(f"  -> SUCCESS: Sent {attachment_name} to {parent_email}")
            return True
        else:
            print(f"  -> API ERROR sending to {parent_email}: {response.text}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"  -> NETWORK ERROR sending to {parent_email}: {e}")
        return False

# ==========================================
# --- GUI & THREADING LOGIC ---
# ==========================================

class ThreadSafeStdoutRedirector:
    def __init__(self, text_widget, root):
        self.text_widget = text_widget
        self.root = root

    def write(self, string):
        self.root.after(0, self._write, string)

    def _write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class ReportSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Report Sender")
        self.root.geometry("650x880")
        self.root.resizable(False, False)
        
        self.bg_color = "#202020"
        self.fg_color = "#e0e0e0"
        self.entry_bg = "#333333"
        self.btn_bg = "#444444"
        self.btn_active = "#555555"
        self.btn_primary = "#005a9e"
        
        self.root.configure(bg=self.bg_color)
        self.apply_dark_theme()

        self.csv_path = tk.StringVar()
        self.reports_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        self.is_dry_run = tk.BooleanVar(value=True)
        
        self.config_data = load_app_config()
        self.tenant_id = tk.StringVar(value=self.config_data.get("tenant_id", ""))
        self.client_id = tk.StringVar(value=self.config_data.get("client_id", ""))
        self.sender_email = tk.StringVar(value=self.config_data.get("sender_email", ""))
        self.subject_text = tk.StringVar(value=self.config_data.get("subject", "Student Report: {name}"))
        
        self.setup_ui()
        
        sys.stdout = ThreadSafeStdoutRedirector(self.console_text, self.root)
        sys.stderr = sys.stdout

        print("--- APPLICATION STARTED ---")
        print("Please configure your Microsoft Entra credentials, data sources, and message settings.")
        print("Note: DRY RUN is currently enabled.\n")

    def apply_dark_theme(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(".", background=self.bg_color, foreground=self.fg_color, fieldbackground=self.entry_bg, insertcolor=self.fg_color)
        style.configure("TLabelframe", background=self.bg_color, foreground=self.fg_color, bordercolor="#555555")
        style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.fg_color, font=("Segoe UI", 9, "bold"))
        style.configure("TButton", background=self.btn_bg, foreground=self.fg_color, borderwidth=1, bordercolor="#555555", focuscolor="none")
        style.map("TButton", background=[("active", self.btn_active)])
        style.configure("Primary.TButton", background=self.btn_primary, foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#0078d4")])
        style.configure("TCheckbutton", background=self.bg_color, foreground=self.fg_color, focuscolor="none")
        style.map("TCheckbutton", background=[("active", self.bg_color)])

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill="both", expand=True)

        # --- Microsoft Entra ID Frame ---
        frame_auth = ttk.LabelFrame(main_frame, text=" Microsoft Entra ID Configuration ")
        frame_auth.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_auth, text="Tenant ID:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        ttk.Entry(frame_auth, textvariable=self.tenant_id, width=63).grid(row=0, column=1, padx=(0, 10), pady=(10, 5), sticky="w")

        ttk.Label(frame_auth, text="Client ID:").grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ttk.Entry(frame_auth, textvariable=self.client_id, width=63).grid(row=1, column=1, padx=(0, 10), pady=(0, 10), sticky="w")

        # --- File Selection Frame ---
        frame_files = ttk.LabelFrame(main_frame, text=" Data Sources ")
        frame_files.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_files, text="Student Emails (CSV):").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        ttk.Entry(frame_files, textvariable=self.csv_path, width=50, state='readonly').grid(row=0, column=1, padx=(0, 10), pady=(10, 5))
        ttk.Button(frame_files, text="SELECT", width=10, command=self.browse_csv).grid(row=0, column=2, padx=(0, 10), pady=(10, 5))

        ttk.Label(frame_files, text="XML Reports Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ttk.Entry(frame_files, textvariable=self.reports_path, width=50, state='readonly').grid(row=1, column=1, padx=(0, 10), pady=(0, 10))
        ttk.Button(frame_files, text="SELECT", width=10, command=self.browse_folder).grid(row=1, column=2, padx=(0, 10), pady=(0, 10))

        # --- Message Configuration Frame ---
        frame_msg = ttk.LabelFrame(main_frame, text=" Message Configuration ")
        frame_msg.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_msg, text="Send From Email:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        ttk.Entry(frame_msg, textvariable=self.sender_email, width=63).grid(row=0, column=1, padx=(0, 10), pady=(10, 5), sticky="w")

        ttk.Label(frame_msg, text="Subject Line:\n(Use {name} for student)").grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))
        ttk.Entry(frame_msg, textvariable=self.subject_text, width=63).grid(row=1, column=1, padx=(0, 10), pady=(0, 5), sticky="w")

        ttk.Label(frame_msg, text="Email Body:").grid(row=2, column=0, sticky="nw", padx=10, pady=(0, 10))
        
        self.body_text = tk.Text(frame_msg, height=4, width=47, bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color, font=("Segoe UI", 9))
        self.body_text.grid(row=2, column=1, padx=(0, 10), pady=(0, 10), sticky="w")
        self.body_text.insert("1.0", self.config_data.get("body", ""))

        # --- Console Output Frame ---
        frame_console = ttk.LabelFrame(main_frame, text=" Execution Log ")
        frame_console.pack(fill="both", expand=True, pady=(0, 10))

        self.console_text = scrolledtext.ScrolledText(
            frame_console, wrap=tk.WORD, state='disabled', 
            bg="#121212", fg="#00ff00", insertbackground="#ffffff", 
            font=("Consolas", 9), height=10, borderwidth=1, relief="flat"
        )
        self.console_text.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Action Buttons Frame ---
        frame_actions = ttk.Frame(main_frame)
        frame_actions.pack(fill="x", pady=(5, 5))

        self.btn_refresh = ttk.Button(frame_actions, text="Clear Sent Log", command=self.ui_clear_log)
        self.btn_refresh.pack(side="left")

        self.chk_dry_run = ttk.Checkbutton(frame_actions, text="Enable DRY RUN", variable=self.is_dry_run)
        self.chk_dry_run.pack(side="left", padx=15)

        self.btn_confirm = ttk.Button(frame_actions, text="START EXECUTION", style="Primary.TButton", command=self.start_processing_thread)
        self.btn_confirm.pack(side="right")

        # --- Status Bar ---
        status_bar = tk.Label(self.root, textvariable=self.status_text, bg="#0078d4", fg="#ffffff", anchor="w", padx=10, font=("Segoe UI", 8))
        status_bar.pack(side="bottom", fill="x")

    def browse_csv(self):
        path = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV Files", "*.csv")])
        if path: self.csv_path.set(path)

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select Reports Folder")
        if path: self.reports_path.set(path)

    def get_sent_log_path(self):
        if not self.reports_path.get(): return None
        return os.path.join(self.reports_path.get(), 'sent_log.json')

    def ui_clear_log(self):
        log_path = self.get_sent_log_path()
        if not log_path:
            messagebox.showwarning("Missing Folder", "Select the Reports Folder first to clear its specific log.")
            return
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear the sent log for this folder? Reports will be re-sent on the next run."):
            clear_sent_log(log_path)
            print("\n  -> LOG CLEARED: The tracking file has been reset for a new reporting cycle.\n")
            self.status_text.set("Log cleared.")

    def toggle_buttons(self, state):
        if state == tk.DISABLED:
            self.btn_confirm.state(['disabled'])
            self.btn_refresh.state(['disabled'])
            self.chk_dry_run.state(['disabled'])
        else:
            self.btn_confirm.state(['!disabled'])
            self.btn_refresh.state(['!disabled'])
            self.chk_dry_run.state(['!disabled'])

    def start_processing_thread(self):
        if not self.tenant_id.get().strip() or not self.client_id.get().strip():
            messagebox.showerror("Validation Error", "Please provide your Microsoft Entra Tenant ID and Client ID.")
            return
            
        if not self.csv_path.get() or not self.reports_path.get() or not self.sender_email.get().strip():
            messagebox.showerror("Validation Error", "Please ensure CSV, Reports Folder, and Sender Email are all provided.")
            return

        self.config_data = {
            "tenant_id": self.tenant_id.get().strip(),
            "client_id": self.client_id.get().strip(),
            "sender_email": self.sender_email.get().strip(),
            "subject": self.subject_text.get(),
            "body": self.body_text.get("1.0", tk.END).strip()
        }
        save_app_config(self.config_data)

        self.toggle_buttons(tk.DISABLED)
        threading.Thread(target=self.execute_sending_logic, daemon=True).start()

    def update_status(self, text):
        self.status_text.set(text)

    def execute_sending_logic(self):
        csv_file = self.csv_path.get()
        rep_folder = self.reports_path.get()
        log_file = self.get_sent_log_path()
        current_dry_run = self.is_dry_run.get()
        
        t_id = self.config_data["tenant_id"]
        c_id = self.config_data["client_id"]
        sender = self.config_data["sender_email"]
        subject_template = self.config_data["subject"]
        body_content = self.config_data["body"]

        self.update_status("Starting process...")
        print(f"\n--- STARTING PROCESS (DRY RUN = {current_dry_run}) ---")
        
        email_map = load_sims_emails(csv_file)
        if email_map is None:
            print("Stopping: Data mapping failed.")
            self.update_status("Failed: Data mapping error.")
            self.toggle_buttons(tk.NORMAL)
            return

        sent_log = load_sent_log(log_file)
        access_token = None

        if not current_dry_run:
            try:
                print("Authenticating with Microsoft Entra ID...")
                access_token = get_access_token(t_id, c_id, status_callback=self.update_status)
                print("Authentication Successful.")
                self.update_status("Authentication successful. Processing files...")
            except Exception as e:
                print(f"Authentication Failed: {e}")
                self.update_status("Authentication failed.")
                self.toggle_buttons(tk.NORMAL)
                return
        else:
            print("Running in DRY RUN mode. Authentication bypassed.")
            self.update_status("DRY RUN active. Simulating processing...")

        print("-" * 40)

        files = [f for f in os.listdir(rep_folder) if not os.path.isdir(os.path.join(rep_folder, f))]
        total_files = len(files)
        processed_count = 0

        for filename in files:
            processed_count += 1
            self.update_status(f"Processing file {processed_count} of {total_files}...")
            
            filepath = os.path.join(rep_folder, filename)
            print(f"\nProcessing File: {filename}")
            
            if filename in sent_log:
                print(f"  -> LOGGED: This report has already been sent. Skipping.")
                continue
            
            attachment_name, mime_type = determine_file_properties(filename)
            if not attachment_name:
                print(f"  -> ERROR: Unsupported filetype detected. Skipping.")
                continue

            sims_name_key = extract_name_from_filename(filename)
            if not sims_name_key:
                print("  -> ERROR: Could not parse a valid SIMS name structure. Skipping.")
                continue

            if sims_name_key not in email_map:
                print(f"  -> ERROR: Name key '{sims_name_key}' not found in CSV. Skipping.")
                continue

            target_emails = email_map[sims_name_key]
            if not target_emails:
                print(f"  -> ERROR: Match found, but no valid email addresses exist. Skipping.")
                continue

            name_parts = sims_name_key.split('-')
            display_name = f"{name_parts[-1].title()} {'-'.join(name_parts[:-1]).title()}" if len(name_parts) >= 2 else sims_name_key.title()
            
            formatted_subject = subject_template.replace("{name}", display_name)
            
            print(f"  -> Match Confirmed: '{display_name}'")
            transmission_success = False

            for email_address in target_emails:
                if current_dry_run:
                    print(f"  -> [DRY RUN] Would send '{attachment_name}' to: {email_address} (From: {sender})")
                else:
                    print(f"  -> [SENDING] Transmitting '{attachment_name}' to: {email_address}...")
                    success = send_graph_email(access_token, sender, email_address, formatted_subject, body_content, filepath, attachment_name, mime_type)
                    if success: transmission_success = True
            
            if not current_dry_run and transmission_success:
                sent_log.add(filename)
                save_sent_log(sent_log, log_file)
            
        print("\n--- PROCESS COMPLETE ---")
        self.update_status("Process complete.")
        self.toggle_buttons(tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportSenderApp(root)
    root.mainloop()