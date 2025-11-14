import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import csv
import pandas as pd
import os
from threading import Thread

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender System")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variables
        self.recipients_list = []
        self.attachment_path = None
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # SMTP Settings Section
        settings_frame = ttk.LabelFrame(main_frame, text="SMTP Settings", padding="10")
        settings_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        settings_frame.columnconfigure(1, weight=1)
        
        ttk.Label(settings_frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.smtp_server = ttk.Entry(settings_frame, width=40)
        self.smtp_server.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.smtp_server.insert(0, "smtp.gmail.com")
        
        ttk.Label(settings_frame, text="SMTP Port:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.smtp_port = ttk.Entry(settings_frame, width=40)
        self.smtp_port.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.smtp_port.insert(0, "587")
        
        ttk.Label(settings_frame, text="Your Email:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.sender_email = ttk.Entry(settings_frame, width=40)
        self.sender_email.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        ttk.Label(settings_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.sender_password = ttk.Entry(settings_frame, width=40, show="*")
        self.sender_password.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Recipients Section
        recipients_frame = ttk.LabelFrame(main_frame, text="Recipients", padding="10")
        recipients_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        recipients_frame.columnconfigure(1, weight=1)
        
        ttk.Button(recipients_frame, text="Load from CSV", 
                  command=self.load_csv).grid(row=0, column=0, padx=5, pady=2)
        ttk.Button(recipients_frame, text="Load from Excel", 
                  command=self.load_excel).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(recipients_frame, text="Add Manual", 
                  command=self.add_manual_recipient).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(recipients_frame, text="Clear All", 
                  command=self.clear_recipients).grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Label(recipients_frame, text="Recipients List:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.recipients_count = ttk.Label(recipients_frame, text="Count: 0")
        self.recipients_count.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        self.recipients_listbox = tk.Listbox(recipients_frame, height=6)
        self.recipients_listbox.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=2)
        
        scrollbar = ttk.Scrollbar(recipients_frame, orient=tk.VERTICAL, command=self.recipients_listbox.yview)
        scrollbar.grid(row=2, column=4, sticky=(tk.N, tk.S))
        self.recipients_listbox.config(yscrollcommand=scrollbar.set)
        
        # Email Content Section
        content_frame = ttk.LabelFrame(main_frame, text="Email Content", padding="10")
        content_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        content_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        ttk.Label(content_frame, text="Subject:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.subject = ttk.Entry(content_frame, width=60)
        self.subject.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        ttk.Label(content_frame, text="Message:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=2)
        self.message = scrolledtext.ScrolledText(content_frame, width=60, height=10)
        self.message.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=2)
        content_frame.rowconfigure(1, weight=1)
        
        # Attachment Section
        attachment_frame = ttk.Frame(content_frame)
        attachment_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(attachment_frame, text="Add Attachment", 
                  command=self.add_attachment).pack(side=tk.LEFT, padx=5)
        self.attachment_label = ttk.Label(attachment_frame, text="No attachment")
        self.attachment_label.pack(side=tk.LEFT, padx=5)
        
        # Send Button
        send_frame = ttk.Frame(main_frame)
        send_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.send_button = ttk.Button(send_frame, text="Send Emails", 
                                      command=self.send_emails, style="Accent.TButton")
        self.send_button.pack(side=tk.LEFT, padx=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(send_frame, length=300, 
                                           variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, padx=5)
        
        # Status Section
        status_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10")
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, width=80, height=8)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
    def log_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def load_csv(self):
        filename = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as file:
                    csv_reader = csv.reader(file)
                    headers = next(csv_reader, None)
                    
                    for row in csv_reader:
                        if row and row[0]:  # Assuming email is in first column
                            email = row[0].strip()
                            if email and '@' in email:
                                self.recipients_list.append(email)
                
                self.update_recipients_display()
                self.log_status(f"Loaded {len(self.recipients_list)} recipients from CSV")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")
    
    def load_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                # Assuming email is in first column
                for email in df.iloc[:, 0]:
                    email = str(email).strip()
                    if email and '@' in email and email != 'nan':
                        self.recipients_list.append(email)
                
                self.update_recipients_display()
                self.log_status(f"Loaded {len(self.recipients_list)} recipients from Excel")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel: {str(e)}")
    
    def add_manual_recipient(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Recipient")
        dialog.geometry("300x100")
        
        ttk.Label(dialog, text="Email Address:").pack(pady=10)
        email_entry = ttk.Entry(dialog, width=30)
        email_entry.pack(pady=5)
        
        def add_email():
            email = email_entry.get().strip()
            if email and '@' in email:
                self.recipients_list.append(email)
                self.update_recipients_display()
                dialog.destroy()
            else:
                messagebox.showwarning("Invalid Email", "Please enter a valid email address")
        
        ttk.Button(dialog, text="Add", command=add_email).pack(pady=10)
    
    def clear_recipients(self):
        self.recipients_list.clear()
        self.update_recipients_display()
        self.log_status("Cleared all recipients")
    
    def update_recipients_display(self):
        self.recipients_listbox.delete(0, tk.END)
        for email in self.recipients_list:
            self.recipients_listbox.insert(tk.END, email)
        self.recipients_count.config(text=f"Count: {len(self.recipients_list)}")
    
    def add_attachment(self):
        filename = filedialog.askopenfilename(title="Select file to attach")
        if filename:
            self.attachment_path = filename
            self.attachment_label.config(text=os.path.basename(filename))
    
    def send_emails(self):
        # Validation
        if not self.smtp_server.get() or not self.smtp_port.get():
            messagebox.showerror("Error", "Please enter SMTP server and port")
            return
        
        if not self.sender_email.get() or not self.sender_password.get():
            messagebox.showerror("Error", "Please enter your email and password")
            return
        
        if not self.recipients_list:
            messagebox.showerror("Error", "Please add at least one recipient")
            return
        
        if not self.subject.get():
            messagebox.showerror("Error", "Please enter a subject")
            return
        
        # Start sending in a separate thread
        Thread(target=self._send_emails_thread, daemon=True).start()
    
    def _send_emails_thread(self):
        self.send_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        try:
            # Connect to SMTP server
            self.log_status("Connecting to SMTP server...")
            server = smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get()))
            server.starttls()
            server.login(self.sender_email.get(), self.sender_password.get())
            self.log_status("Successfully connected and authenticated")
            
            total = len(self.recipients_list)
            success_count = 0
            
            for i, recipient in enumerate(self.recipients_list):
                try:
                    # Create message
                    msg = MIMEMultipart()
                    msg['From'] = self.sender_email.get()
                    msg['To'] = recipient
                    msg['Subject'] = self.subject.get()
                    
                    # Add message body
                    msg.attach(MIMEText(self.message.get("1.0", tk.END), 'plain'))
                    
                    # Add attachment if exists
                    if self.attachment_path:
                        with open(self.attachment_path, 'rb') as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', 
                                          f'attachment; filename={os.path.basename(self.attachment_path)}')
                            msg.attach(part)
                    
                    # Send email
                    server.send_message(msg)
                    success_count += 1
                    self.log_status(f"✓ Sent to {recipient}")
                    
                except Exception as e:
                    self.log_status(f"✗ Failed to send to {recipient}: {str(e)}")
                
                # Update progress
                progress = ((i + 1) / total) * 100
                self.progress_var.set(progress)
                self.root.update_idletasks()
            
            server.quit()
            self.log_status(f"\nCompleted: {success_count}/{total} emails sent successfully")
            messagebox.showinfo("Success", f"Sent {success_count}/{total} emails successfully")
            
        except Exception as e:
            self.log_status(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to send emails: {str(e)}")
        
        finally:
            self.send_button.config(state=tk.NORMAL)
            self.progress_var.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()