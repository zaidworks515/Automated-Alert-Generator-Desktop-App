import os
import pandas as pd
import smtplib
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import threading
import time
from db_queries import get_all_designations, get_all_recipients
from email_management import add_email, delete_email, list_all_emails
import aspose.tasks as tasks

def read_ms_project(file_path):
    project = tasks.Project(file_path)
    print('PROJECT FILE READ SUCCESSFULLY')
    delayed_tasks = []

    root_task = project.root_task
    all_tasks = root_task.select_all_child_tasks()

    for task in all_tasks:
        task_name = task.name
        wbs = task.wbs
        start_date = task.start
        finish_date = task.finish
        cost_variance = task.cost_variance
        duration_variance = task.duration_variance
        percent_complete = task.percent_complete

        if task_name.lower().startswith('milestone'):
            continue

        delayed_tasks.append({
            "wbs": wbs,
            "name": task_name,
            "start_date": start_date,
            "finish_date": finish_date,
            "cost_variance": cost_variance,
            "duration_variance": duration_variance,
            "percent_complete": percent_complete
        })

    return delayed_tasks

def send_email(subject, body, end_text, recipient_emails):
    sender_email = "zaidworks515@gmail.com"  # Replace with your email

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = subject

    full_body = f"{body}<br><br>{end_text}"

    msg.attach(MIMEText(full_body, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, 'ubeg uddv dkqi cpmu')  # Replace with your email password
        text = msg.as_string()
        server.sendmail(sender_email, recipient_emails, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def check_delayed_tasks(file_path, recipient_emails):
    delayed_tasks = read_ms_project(file_path)

    project_name = os.path.splitext(os.path.basename(file_path))[0]
    df = pd.DataFrame(delayed_tasks)
    df.to_csv('extracted_data.csv', index=False)

    data = pd.read_csv('extracted_data.csv')
    data.dropna(inplace=True)
    data['duration_variance'] = pd.to_numeric(data['duration_variance'].str.extract('(\d+)')[0], errors='coerce')
    data['start_date'] = pd.to_datetime(data['start_date']).dt.date
    data['finish_date'] = pd.to_datetime(data['finish_date']).dt.date

    today = datetime.today().date()
    potential_delay_makers = data[
        (data['finish_date'] > today) & 
        ((data['cost_variance'] > 0) | ((data['duration_variance'] > 0) & (data['percent_complete'] < 100)))
    ]

    delayed_task = data[
        (data['finish_date'] < today) & 
        ((data['cost_variance'] > 0) | ((data['duration_variance'] > 0) & (data['percent_complete'] < 100)))
    ]

    if not potential_delay_makers.empty or not delayed_task.empty:
        body = f"""
        <html>
        <body>
            <p>Dear Team,</p>
            <p>Please be informed about the following project delays and potential issues:</p>
            <h3>Overall Project Details:</h3>
            {data.head(1).to_html(index=False)}<br>
        """

        if not potential_delay_makers.empty:
            body += "<h3>Tasks with potential delay:</h3>"
            body += potential_delay_makers.iloc[1:].to_html(index=False) + "<br>"

        if not delayed_task.empty:
            body += "<h3>Delayed Task:</h3>"
            body += delayed_task.to_html(index=False) + "<br>"

        end_text = "============== This is a computer-generated email, please refer to the administrator in case of any error =============="
        send_email("Project Delay Alert", body, end_text, recipient_emails)
        
    file_path = 'extracted_data.csv'
    try:
        if os.path.isfile(file_path):
            os.remove(file_path)
            print(f"Successfully deleted: {file_path}")
        else:
            print(f"The file {file_path} does not exist.")
    except Exception as e:
        print(f"Error deleting file: {e}")
        
    print("EMAIL SENDED SUCCESSFULLY AND DELETED THE CREATED CSV RECORD FILE")
        

def get_recipient_emails(selected_designations):
    recipients = get_all_recipients()
    return [recipient[2] for recipient in recipients if recipient[3] in selected_designations]

def run_task(file_path, selected_designations):
    while True:
        now = datetime.now()
        recipient_emails = get_recipient_emails(selected_designations)
        if recipient_emails:
            check_delayed_tasks(file_path, recipient_emails)
        time.sleep(86400)  # Check every 24 hours


class AlertSenderApp:
    def __init__(self, master):
        self.frame = ttk.Frame(master, padding=10)
        self.frame.pack(expand=True, fill=tk.BOTH)

        self.label = ttk.Label(self.frame, text="Select Project File:")
        self.label.pack(pady=(0, 10))

        self.file_path_label = ttk.Label(self.frame, text="No file selected", relief="sunken", width=40)
        self.file_path_label.pack(pady=(0, 10))

        self.browse_button = ttk.Button(self.frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(pady=(0, 10))

        self.checkbox_frame = ttk.LabelFrame(self.frame, text="Select Designations")
        self.checkbox_frame.pack(pady=(10, 10), fill=tk.BOTH, expand=True)

        self.check_vars = {}
        self.check_buttons = []

        self.populate_designations()

        self.start_button = ttk.Button(self.frame, text="Send Alert", command=self.start_alerts)
        self.start_button.pack(pady=(20, 0))

        self.file_path = None  # Initialize file_path to store the selected project file path

        # Event binding to refresh the tab when the notebook tab changes
        master.bind("<<NotebookTabChanged>>", self.refresh)

    def populate_designations(self):
        # Clear existing checkboxes
        for check_button in self.check_buttons:
            check_button.destroy()
        self.check_buttons.clear()
        self.check_vars.clear()

        designations = get_all_designations()  # Assuming this function retrieves designations
        for designation in designations:
            var = tk.BooleanVar()
            check_button = ttk.Checkbutton(self.checkbox_frame, text=designation, variable=var)
            check_button.pack(anchor='w')
            self.check_vars[designation] = var
            self.check_buttons.append(check_button)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Project File",
            filetypes=[("Project Files", "*.mpp"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_path_label.config(text=os.path.basename(file_path))

    def start_alerts(self):
        if self.file_path:
            selected_designations = [designation for designation, var in self.check_vars.items() if var.get()]
            if selected_designations:
                threading.Thread(target=run_task, args=(self.file_path, selected_designations), daemon=True).start()
                messagebox.showinfo("Success", "Alert sending started.")
            else:
                messagebox.showerror("Error", "Please select at least one designation.")
        else:
            messagebox.showerror("Error", "Please select a project file before starting alerts.")

    def refresh(self, event=None):
        self.file_path_label.config(text="No file selected")
        self.file_path = None
        
        self.populate_designations()

class EmailManagementApp:
    def __init__(self, master):
        self.master = master

        # Create main frame
        self.frame = ttk.Frame(master, padding=10)
        self.frame.pack(expand=True, fill=tk.BOTH)

        # Title for Email Management
        title_label = ttk.Label(self.frame, text="Email Management", font=("Arial", 16))
        title_label.pack(pady=(0, 10))

        self.email_listbox = tk.Listbox(self.frame)
        self.email_listbox.pack(fill=tk.BOTH, expand=True)

        self.refresh_email_list()

        self.delete_button = ttk.Button(self.frame, text="Delete Selected Email", command=self.delete_selected_email)
        self.delete_button.pack(pady=(5, 0))

        self.add_button = ttk.Button(self.frame, text="Add Email", command=self.add_email)
        self.add_button.pack(pady=(5, 0))

        # Labels for input fields
        ttk.Label(self.frame, text="Name:").pack(pady=(5, 0))
        self.name_entry = ttk.Entry(self.frame, width=50)
        self.name_entry.pack(pady=(5, 0))
        
        ttk.Label(self.frame, text="Email:").pack(pady=(5, 0))
        self.email_entry = ttk.Entry(self.frame, width=50)
        self.email_entry.pack(pady=(5, 0))

        ttk.Label(self.frame, text="Designation:").pack(pady=(5, 0))
        self.designation_entry = ttk.Entry(self.frame, width=50)
        self.designation_entry.pack(pady=(5, 0))

    def refresh_email_list(self):
        self.email_listbox.delete(0, tk.END)
        emails = list_all_emails()  # Assume this returns a list of tuples (id, name, email, designation)
        for email in emails:
            self.email_listbox.insert(tk.END, f"{email[1]} ({email[2]}) - {email[3]}")  # Format as "Name (Email) - Designation"

    def delete_selected_email(self):
        selected_email_index = self.email_listbox.curselection()
        if selected_email_index:
            # Get the selected email text
            email_text = self.email_listbox.get(selected_email_index)

            # Retrieve the actual email address from the list of emails
            emails = list_all_emails()  # Assume this retrieves all emails
            email_to_delete = None

            for email in emails:
                # Assuming email is a tuple like (id, name, email_address, ...)
                if email[1] in email_text:  # Check the name, which should be the second element
                    email_to_delete = email[2]  # Get the actual email address
                    print(email_to_delete)
                    break

            if email_to_delete is not None:
                delete_email(email_to_delete)  # Call delete function with the email
                self.refresh_email_list()  # Refresh the list to show changes
                messagebox.showinfo("Success", "Email deleted successfully.")
            else:
                messagebox.showerror("Error", "Could not find the selected email address.")
        else:
            messagebox.showerror("Error", "Please select an email to delete.")


    def add_email(self):
        name = self.name_entry.get()
        email = self.email_entry.get()
        designation = self.designation_entry.get()
        if email and name and designation:
            add_email(name, email, designation)
            self.refresh_email_list()
            self.email_entry.delete(0, tk.END)
            self.name_entry.delete(0, tk.END)
            self.designation_entry.delete(0, tk.END)
            messagebox.showinfo("Success", "Email added successfully.")
        else:
            messagebox.showerror("Error", "Please fill in all fields.")



def main():
    root = tk.Tk()
    root.title("Email Alert System")

    # Set up the tab control
    tab_control = ttk.Notebook(root)

    # Create tabs
    alert_sender_tab = AlertSenderApp(tab_control)
    email_management_tab = EmailManagementApp(tab_control)

    # Add tabs to the control
    tab_control.add(alert_sender_tab.frame, text="Alert Sender")
    tab_control.add(email_management_tab.frame, text="Email Management")
    
    # Set the style for tabs
    style = ttk.Style()
    style.configure("TNotebook.Tab", padding=(10, 5), background="#f0f0f0")  # Adjust tab style as needed
    style.map("TNotebook.Tab", background=[("selected", "#d0d0d0")])  # Selected tab color

    tab_control.pack(expand=1, fill="both")
    
    root.mainloop()

if __name__ == "__main__":
    main()
