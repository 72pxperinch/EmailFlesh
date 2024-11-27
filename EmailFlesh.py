import imaplib
import email
from email.header import decode_header
import os
import json
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Global flags and state variables
paused = False
stop_requested = False
progress_data = {}

# File to store progress
PROGRESS_FILE = "progress.json"

def save_progress(last_processed):
    # Store only the email ID or index, not the entire message object
    progress_data["last_processed"] = last_processed
    try:
        with open(PROGRESS_FILE, "w") as f:
            json.dump(progress_data, f)
    except Exception as e:
        print(f"Error saving progress: {e}")

def load_progress():
    global progress_data
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r") as f:
                # Attempt to load progress from the file
                content = f.read().strip()
                if content:  # Check if file is not empty
                    progress_data = json.loads(content)
                else:
                    print("Progress file is empty, starting fresh.")
                    progress_data = {"last_processed": None}
        except (json.JSONDecodeError, Exception) as e:
            print(f"Error loading progress: {e}. Starting fresh.")
            progress_data = {"last_processed": None}
    else:
        progress_data = {"last_processed": None}


def download_attachments(email_address, password, folder_name, progress):
    global paused, stop_requested

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

        # Load the last processed email
        last_processed_index = 0
        if progress_data["last_processed"]:
            last_processed_index = int(progress_data["last_processed"])
            log(f"Resuming from email {last_processed_index + 1}.")

        # Create folder for attachments if it doesn't exist
        log(f"Using folder: {folder_name}")
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        # Process emails
        for i, msg in enumerate(messages[last_processed_index:], last_processed_index + 1):
            if stop_requested:
                log("Download stopped by user.")
                break

            while paused:
                log("Paused. Waiting to resume...")
                time.sleep(1)

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
            
            # Save progress after processing each email by index (or Message-ID)
            save_progress(i)  # Save the index of the last processed email

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
        # Reset buttons on completion
        reset_buttons()

def start_download():
    global paused, stop_requested
    paused = False
    stop_requested = False

    email_address = email_entry.get()
    password = password_entry.get()
    folder_name = folder_entry.get()

    if not email_address or not password or not folder_name:
        messagebox.showerror("Error", "Please fill in all fields!")
        return

    # Disable and enable appropriate buttons
    start_button["state"] = "disabled"
    pause_button["state"] = "normal"
    resume_button["state"] = "disabled"
    stop_button["state"] = "normal"

    # Start the download in a separate thread
    download_thread = threading.Thread(target=download_attachments, args=(email_address, password, folder_name, progress))
    download_thread.daemon = True  # Make sure the thread will close when the program exits
    download_thread.start()

def pause_download():
    global paused
    paused = True
    pause_button["state"] = "disabled"
    resume_button["state"] = "normal"

def resume_download():
    global paused
    paused = False
    pause_button["state"] = "normal"
    resume_button["state"] = "disabled"

def stop_download():
    global stop_requested
    stop_requested = True
    reset_buttons()

def reset_buttons():
    # Reset button states to their initial values
    start_button["state"] = "normal"
    pause_button["state"] = "disabled"
    resume_button["state"] = "disabled"
    stop_button["state"] = "disabled"

def choose_folder():
    selected_dir = filedialog.askdirectory()
    if selected_dir:
        output_folder = os.path.join(selected_dir, "Email Attachments")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, output_folder)

# Ensure the Output folder is created on initialization
def ensure_default_folder():
    # Set default folder to the "Output" directory in the current working directory
    default_dir = os.getcwd()
    default_folder = os.path.join(default_dir, "Email Attachments")
    
    # Create the "Output" folder if it doesn't exist
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
start_button = ttk.Button(frame, text="Start", command=start_download)
start_button.grid(column=0, row=5, pady=5)

pause_button = ttk.Button(frame, text="Pause", command=pause_download, state="disabled")
pause_button.grid(column=1, row=5, pady=5)

resume_button = ttk.Button(frame, text="Resume", command=resume_download, state="disabled")
resume_button.grid(column=2, row=5, pady=5)

stop_button = ttk.Button(frame, text="Stop", command=stop_download, state="disabled")
stop_button.grid(column=3, row=5, pady=5)

root.mainloop()
