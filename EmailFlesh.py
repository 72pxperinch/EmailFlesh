import imaplib
import email
from email.header import decode_header
import os
import json
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os.path
import logging

# Global flags and state variables
stop_requested = False
progress_data = {
    'emails': {}  # Will store progress for each email address
}

# File to store progress
PROGRESS_FILE = os.path.join(os.path.expanduser("~/Library/Application Support/EmailFlesh"), "progress.json")

# Ensure the application support directory exists
os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)

# Set up logging
log_file = os.path.join(os.path.expanduser("~/Library/Logs/EmailFlesh"), "app.log")
os.makedirs(os.path.dirname(log_file), exist_ok=True)

logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

try:
    logging.debug("Application starting...")
    logging.debug(f"Python version: {sys.version}")
    logging.debug(f"Current working directory: {os.getcwd()}")
    logging.debug(f"System platform: {sys.platform}")
except Exception as e:
    logging.error(f"Error during startup logging: {e}")

def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions by showing them in a message box"""
    import traceback
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    messagebox.showerror('Error', f'An unexpected error occurred:\n\n{error_msg}')

# Install the error handler
sys.excepthook = handle_exception

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def save_progress(email_address, last_processed):
    # Store progress for specific email address
    if 'emails' not in progress_data:
        progress_data['emails'] = {}
        
    progress_data['emails'][email_address] = {
        'last_processed': last_processed,
        'last_updated': time.time()  # Add timestamp
    }
    try:
        # Ensure directory exists
        os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
        
        with open(PROGRESS_FILE, "w") as f:
            json.dump(progress_data, f, indent=4)  # Added indent for readability
    except Exception as e:
        logging.error(f"Error saving progress: {e}")

def load_progress():
    global progress_data
    try:
        # Ensure directory exists
        os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
        
        if os.path.exists(PROGRESS_FILE):
            with open(PROGRESS_FILE, "r") as f:
                content = f.read().strip()
                if content:
                    loaded_data = json.loads(content)
                    # Validate structure
                    if isinstance(loaded_data, dict) and 'emails' in loaded_data:
                        progress_data = loaded_data
                    else:
                        progress_data = {'emails': {}}
                else:
                    progress_data = {'emails': {}}
        else:
            # Create new file with empty structure
            progress_data = {'emails': {}}
            with open(PROGRESS_FILE, "w") as f:
                json.dump(progress_data, f, indent=4)
    except Exception as e:
        logging.error(f"Error loading progress: {e}")
        progress_data = {'emails': {}}

def get_email_progress(email_address):
    # Get progress for specific email address
    if email_address in progress_data['emails']:
        return progress_data['emails'][email_address]['last_processed']
    return 0

def download_attachments(email_address, password, folder_name, progress):
    global stop_requested

    def log(message):
        progress.insert(tk.END, f"{message}\n")
        progress.see(tk.END)
        root.update()

    try:
        log("Connecting to the email server...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(email_address, password)
        log("Logged in successfully.")

        log("Selecting the inbox...")
        mail.select("inbox")

        log("Searching for emails...")
        status, messages = mail.search(None, "ALL")
        if status != "OK":
            raise Exception("Failed to retrieve email list.")
        messages = messages[0].split()
        log(f"Found {len(messages)} emails.")

        # Load the last processed email for this specific email address
        last_processed_index = get_email_progress(email_address)
        if last_processed_index > 0:
            log(f"Resuming from email {last_processed_index + 1} for {email_address}.")

        # Create folder for attachments if it doesn't exist
        log(f"Using folder: {folder_name}")
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        # Process emails
        for i, msg in enumerate(messages[last_processed_index:], last_processed_index + 1):
            if stop_requested:
                log("Download stopped by user.")
                break

            log(f"Processing email {i}/{len(messages)}...")
            status, msg_data = mail.fetch(msg, "(RFC822)")
            if status != "OK":
                log(f"Failed to fetch email {i}. Skipping...")
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    for part in msg.walk():
                        if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
                            continue

                        filename = part.get_filename()
                        if filename:
                            log(f"Downloading attachment: {filename}")
                            filepath = os.path.join(folder_name, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            log(f"Saved attachment to: {filepath}")
            
            # Save progress after processing each email
            save_progress(email_address, i)

        if not stop_requested:
            log("All emails processed.")
            messagebox.showinfo("Success", "Attachments downloaded successfully!")
    except Exception as e:
        log(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        if 'mail' in locals():
            log("Closing the connection.")
            mail.logout()
        reset_buttons()

def start_download():
    global stop_requested
    stop_requested = False

    email_address = email_entry.get()
    password = password_entry.get()
    folder_name = folder_entry.get()

    if not email_address or not password or not folder_name:
        messagebox.showerror("Error", "Please fill in all fields!")
        return

    # Create a unique folder for each email address
    user_folder = os.path.join(folder_name, email_address.split('@')[0])
    if not os.path.exists(user_folder):
        os.makedirs(user_folder)

    # Update button states
    start_button["state"] = "disabled"
    stop_button["state"] = "normal"

    # Start the download in a separate thread
    download_thread = threading.Thread(
        target=download_attachments, 
        args=(email_address, password, user_folder, progress)
    )
    download_thread.daemon = True
    download_thread.start()

def stop_download():
    global stop_requested
    stop_requested = True
    reset_buttons()

def reset_buttons():
    # Reset button states to their initial values
    start_button["state"] = "normal"
    stop_button["state"] = "disabled"

def choose_folder():
    selected_dir = filedialog.askdirectory()
    if selected_dir:
        output_folder = os.path.join(selected_dir, "EmailFlesh Downloads")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, output_folder)

def ensure_default_folder():
    # Set default folder to a subfolder in the user's Downloads directory
    default_dir = os.path.expanduser("~/Downloads")
    default_folder = os.path.join(default_dir, "EmailFlesh Downloads")
    
    # Create the "EmailFlesh Downloads" folder if it doesn't exist
    if not os.path.exists(default_folder):
        os.makedirs(default_folder)
    
    # Set this default folder in the folder_entry widget
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, default_folder)

def show_info():
    messagebox.showinfo("Information", "To get an app-specific password for Gmail:\n"
                                       "1. Go to your Google Account settings.\n"
                                       "2. Navigate to Security > App Passwords.\n"
                                       "3. Generate a new app password for 'Mail'.\n"
                                       "4. Use the generated password in this application.")

def reset_progress():
    global progress_data
    try:
        email_address = email_entry.get()
        
        if not email_address:
            messagebox.showerror("Error", "Please enter an email address first!")
            return
            
        # Reset progress only for the current email
        if email_address in progress_data['emails']:
            del progress_data['emails'][email_address]
            
            # Save updated progress to file
            with open(PROGRESS_FILE, "w") as f:
                json.dump(progress_data, f, indent=4)
                
            # Clear progress text area
            progress.delete(1.0, tk.END)
            progress.insert(tk.END, f"Progress has been reset for {email_address}.\n")
            
            messagebox.showinfo("Success", f"Progress has been reset for {email_address}!")
        else:
            messagebox.showinfo("Info", f"No progress data found for {email_address}.")
            
    except Exception as e:
        logging.error(f"Error resetting progress: {e}")
        messagebox.showerror("Error", f"Failed to reset progress: {e}")

# Load previous progress
load_progress()

# Set up the Tkinter UI
root = tk.Tk()
root.title("Email Attachment Downloader")

frame = ttk.Frame(root, padding=20)
frame.grid(sticky="NSEW")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Email input
ttk.Label(frame, text="Email:").grid(column=0, row=0, sticky="W")
email_entry = ttk.Entry(frame, width=30)
email_entry.grid(column=1, row=0, pady=5)

# Password input
ttk.Label(frame, text="App Password:").grid(column=0, row=1, sticky="W")
password_entry = ttk.Entry(frame, width=30, show="*")
password_entry.grid(column=1, row=1, pady=5)

# Info button
info_button = ttk.Button(frame, text="i", command=show_info)
info_button.grid(column=2, row=1, pady=5)

# Folder input
ttk.Label(frame, text="Save Folder:").grid(column=0, row=2, sticky="W")
folder_entry = ttk.Entry(frame, width=30)
folder_entry.grid(column=1, row=2, pady=5)
choose_button = ttk.Button(frame, text="Choose Folder", command=choose_folder)
choose_button.grid(column=2, row=2, pady=5)

# Ensure default folder is set
ensure_default_folder()

# Progress area
progress_label = ttk.Label(frame, text="Progress:")
progress_label.grid(column=0, row=3, sticky="W")
progress = tk.Text(frame, height=10, width=50, state="normal")
progress.grid(column=0, row=4, columnspan=3, pady=10, sticky="NSEW")

# Control buttons
button_frame = ttk.Frame(frame)
button_frame.grid(column=0, row=5, columnspan=3, pady=5)

start_button = ttk.Button(button_frame, text="Start", command=start_download)
start_button.pack(side=tk.LEFT, padx=5)

stop_button = ttk.Button(button_frame, text="Stop", command=stop_download, state="disabled")
stop_button.pack(side=tk.LEFT, padx=5)

reset_progress_button = ttk.Button(button_frame, text="Reset Progress", command=reset_progress)
reset_progress_button.pack(side=tk.LEFT, padx=5)

# Add after creating the root window
if sys.platform == 'darwin':
    try:
        # Make the app more native-looking on macOS
        root.createcommand('tk::mac::ReopenApplication', lambda: root.lift())
        
        # Add About menu item to the application menu
        def show_about():
            messagebox.showinfo("About Email Attachment Downloader", 
                              "Email Attachment Downloader v1.0\n\n"
                              "Download email attachments easily.\n\n"
                              "© 2024 Your Name")
        
        root.createcommand('tk::mac::ShowPreferences', lambda: None)  # Disable preferences
        root.createcommand('tk::mac::ShowHelp', lambda: show_info())  # Show help
        root.createcommand('tk::mac::AboutDialog', lambda: show_about())  # Show about dialog
    except Exception as e:
        logging.error(f"Error setting up macOS integration: {e}")

root.mainloop()
