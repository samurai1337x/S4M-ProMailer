import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import smtplib
import csv
import os
import re
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

class BulkEmailApp:
    def __init__(self, master):
        self.master = master
        master.title("Bulk Email Marketing Tool")
        master.geometry("700x800")
        master.configure(bg='#2C3E50')

        # Dark Theme Colors
        self.bg_color = '#2C3E50'
        self.fg_color = '#ECF0F1'
        self.entry_bg = '#34495E'
        self.button_bg = '#3498DB'
        self.button_fg = 'white'

        # Custom Styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, font=('Arial', 10))
        self.style.configure('TButton', background=self.button_bg, foreground=self.button_fg)
        self.style.configure('TCombobox', background=self.entry_bg, foreground=self.fg_color)

        # SMTP Server Configuration
        self.smtp_servers = {
            "Gmail": ("smtp.gmail.com", 587),
            "Outlook": ("smtp.office365.com", 587),
            "Yahoo": ("smtp.mail.yahoo.com", 587),
            "Custom": ("", 0)
        }

        self.create_widgets()

    def create_widgets(self):
        # Main Container
        main_frame = tk.Frame(self.master, bg=self.bg_color)
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        # SMTP Server Selection
        smtp_frame = tk.Frame(main_frame, bg=self.bg_color)
        smtp_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(smtp_frame, text="SMTP Server:").pack(side=tk.LEFT, padx=(0,10))
        self.smtp_var = tk.StringVar(value="Select SMTP Server")
        self.smtp_dropdown = ttk.Combobox(smtp_frame, textvariable=self.smtp_var, 
                                          values=list(self.smtp_servers.keys()), 
                                          width=30, state="readonly")
        self.smtp_dropdown.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.smtp_dropdown.bind("<<ComboboxSelected>>", self.toggle_custom_smtp)

        # Custom SMTP Server Entries
        self.custom_smtp_frame = tk.Frame(main_frame, bg=self.bg_color)
        self.custom_smtp_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(self.custom_smtp_frame, text="Custom Server:").pack(side=tk.LEFT, padx=(0,10))
        self.custom_server_entry = tk.Entry(self.custom_smtp_frame, width=30, 
                                            bg=self.entry_bg, fg=self.fg_color, 
                                            insertbackground=self.fg_color)
        self.custom_server_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,10))
        
        ttk.Label(self.custom_smtp_frame, text="Port:").pack(side=tk.LEFT)
        self.custom_port_entry = tk.Entry(self.custom_smtp_frame, width=10, 
                                          bg=self.entry_bg, fg=self.fg_color, 
                                          insertbackground=self.fg_color)
        self.custom_port_entry.pack(side=tk.LEFT, padx=(0,10))
        
        # Hide custom SMTP frame initially
        self.custom_smtp_frame.pack_forget()

        # Sender Details
        sender_frame = tk.Frame(main_frame, bg=self.bg_color)
        sender_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(sender_frame, text="Sender Email:").pack(side=tk.LEFT, padx=(0,10))
        self.sender_email = tk.Entry(sender_frame, width=40, 
                                     bg=self.entry_bg, fg=self.fg_color, 
                                     insertbackground=self.fg_color)
        self.sender_email.pack(side=tk.LEFT, expand=True, fill=tk.X)

        sender_pass_frame = tk.Frame(main_frame, bg=self.bg_color)
        sender_pass_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(sender_pass_frame, text="Email Password:").pack(side=tk.LEFT, padx=(0,10))
        self.sender_password = tk.Entry(sender_pass_frame, show="*", width=40, 
                                        bg=self.entry_bg, fg=self.fg_color, 
                                        insertbackground=self.fg_color)
        self.sender_password.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Email Subject
        subject_frame = tk.Frame(main_frame, bg=self.bg_color)
        subject_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(subject_frame, text="Email Subject:").pack(side=tk.LEFT, padx=(0,10))
        self.email_subject = tk.Entry(subject_frame, width=40, 
                                      bg=self.entry_bg, fg=self.fg_color, 
                                      insertbackground=self.fg_color)
        self.email_subject.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Email Body File
        body_file_frame = tk.Frame(main_frame, bg=self.bg_color)
        body_file_frame.pack(fill=tk.X, pady=10)
        
        self.body_file_path = tk.StringVar()
        body_button = ttk.Button(body_file_frame, text="Select Email Body File", command=self.select_body_file)
        body_button.pack(side=tk.LEFT, padx=(0,10))
        body_label = ttk.Label(body_file_frame, textvariable=self.body_file_path)
        body_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Attachments
        attach_frame = tk.Frame(main_frame, bg=self.bg_color)
        attach_frame.pack(fill=tk.X, pady=10)
        
        self.attachments = []
        attach_button = ttk.Button(attach_frame, text="Add Attachments", command=self.add_attachments)
        attach_button.pack(side=tk.LEFT, padx=(0,10))
        
        self.attachments_listbox = tk.Listbox(main_frame, width=50, height=3, 
                                               bg=self.entry_bg, fg=self.fg_color)
        self.attachments_listbox.pack(pady=5)

        # Recipients
        recipients_frame = tk.Frame(main_frame, bg=self.bg_color)
        recipients_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(recipients_frame, text="Recipients:").pack(side=tk.LEFT, padx=(0,10))
        csv_button = ttk.Button(recipients_frame, text="Upload CSV", command=self.upload_csv)
        csv_button.pack(side=tk.LEFT, padx=(0,10))

        self.recipients_text = tk.Text(main_frame, height=5, width=50, 
                                       bg=self.entry_bg, fg=self.fg_color, 
                                       insertbackground=self.fg_color)
        self.recipients_text.pack(pady=5)

        # Action Buttons
        action_frame = tk.Frame(main_frame, bg=self.bg_color)
        action_frame.pack(pady=10, fill=tk.X)
        
        preview_button = ttk.Button(action_frame, text="Preview Email", command=self.preview_email)
        preview_button.pack(side=tk.LEFT, expand=True, padx=5)
        
        send_button = ttk.Button(action_frame, text="Send Emails", command=self.send_bulk_emails)
        send_button.pack(side=tk.LEFT, expand=True, padx=5)

        # Status Log
        ttk.Label(main_frame, text="Status Log:").pack(pady=(10, 0))
        self.status_log = tk.Text(main_frame, height=6, width=70, 
                                  bg=self.entry_bg, fg=self.fg_color, 
                                  state=tk.DISABLED, 
                                  insertbackground=self.fg_color)
        self.status_log.pack(pady=5)

    def toggle_custom_smtp(self, event):
        if self.smtp_var.get() == "Custom":
            self.custom_smtp_frame.pack(fill=tk.X, pady=10)
        else:
            self.custom_smtp_frame.pack_forget()

    def validate_email(self, email):
        """Validate email address format"""
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(email_regex, email) is not None

    def select_body_file(self):
        """Select email body file with error handling"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("HTML Files", "*.html")],
            title="Select Email Body File"
        )
        if file_path:
            self.body_file_path.set(file_path)

    def add_attachments(self):
        """Add attachments with duplicate prevention"""
        files = filedialog.askopenfilenames(title="Select Attachments")
        for file in files:
            if file not in self.attachments:
                self.attachments.append(file)
                self.attachments_listbox.insert(tk.END, os.path.basename(file))

    def upload_csv(self):
        """Upload CSV with advanced error handling"""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv")], 
            title="Upload Recipient CSV"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r') as csvfile:
                reader = csv.reader(csvfile)
                # Validate email addresses
                emails = [
                    row[0].strip() for row in reader 
                    if row and self.validate_email(row[0].strip())
                ]
                
                if not emails:
                    messagebox.showwarning("CSV Upload", "No valid email addresses found.")
                    return

                self.recipients_text.delete('1.0', tk.END)
                self.recipients_text.insert(tk.END, ", ".join(emails))
        except Exception as e:
            messagebox.showerror("CSV Upload Error", f"Error reading CSV: {str(e)}")

    def log_status(self, message):
        """Enhanced status logging"""
        self.status_log.configure(state=tk.NORMAL)
        self.status_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.status_log.configure(state=tk.DISABLED)
        self.status_log.see(tk.END)
        self.master.update_idletasks()

    def get_smtp_details(self):
        """Get SMTP details with comprehensive validation"""
        if self.smtp_var.get() == "Custom":
            server = self.custom_server_entry.get().strip()
            port = self.custom_port_entry.get().strip()
            
            if not server or not port:
                raise ValueError("Custom SMTP server and port must be provided")
            
            try:
                port = int(port)
                if port <= 0 or port > 65535:
                    raise ValueError("Invalid port number")
            except ValueError:
                raise ValueError("Port must be a valid integer between 1 and 65535")
            
            return server, port
        else:
            return self.smtp_servers[self.smtp_var.get()]

    def preview_email(self):
        """Enhanced email preview with error handling"""
        try:
            # Validate required fields
            if not self.validate_preview_inputs():
                return

            preview_window = tk.Toplevel(self.master)
            preview_window.title("Email Preview")
            preview_window.geometry("500x500")
            preview_window.configure(bg=self.bg_color)

            # Load email body
            body_content = self.load_email_body()

            preview_text = tk.Text(preview_window, wrap=tk.WORD, 
                                   bg=self.entry_bg, fg=self.fg_color, 
                                   insertbackground=self.fg_color)
            preview_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

            preview_text.insert(tk.END, f"Subject: {self.email_subject.get()}\n\n")
            preview_text.insert(tk.END, "Body:\n")
            preview_text.insert(tk.END, body_content)
            preview_text.insert(tk.END, "\n\nAttachments:\n")
            
            for attachment in self.attachments:
                preview_text.insert(tk.END, f"- {os.path.basename(attachment)}\n")

            preview_text.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("Preview Error", str(e))

    def validate_preview_inputs(self):
        """Validate inputs before preview or sending"""
        if not self.sender_email.get() or not self.validate_email(self.sender_email.get()):
            messagebox.showerror("Validation Error", "Invalid sender email address")
            return False
        
        if not self.email_subject.get():
            messagebox.showerror("Validation Error", "Email subject is required")
            return False
        
        if not self.body_file_path.get():
            messagebox.showerror("Validation Error", "Please select an email body file")
            return False
        
        return True

    def load_email_body(self):
        """Load email body with error handling"""
        try:
            with open(self.body_file_path.get(), 'r', encoding='utf-8') as file:
                return file.read()
        except Exception as e:
            raise ValueError(f"Could not read email body file: {e}")

    def send_bulk_emails(self):
        """Comprehensive email sending with enhanced error handling"""
        try:
            # Validate sender email
            sender_email = self.sender_email.get().strip()
            if not self.validate_email(sender_email):
                messagebox.showerror("Validation Error", "Invalid sender email address")
                return

            # Validate password
            sender_password = self.sender_password.get()
            if not sender_password:
                messagebox.showerror("Validation Error", "Password is required")
                return

            # Validate subject
            email_subject = self.email_subject.get().strip()
            if not email_subject:
                messagebox.showerror("Validation Error", "Email subject is required")
                return

            # Validate body file
            body_file_path = self.body_file_path.get()
            if not body_file_path or not os.path.exists(body_file_path):
                messagebox.showerror("Validation Error", "Please select a valid email body file")
                return

            # Get recipients with validation
            recipients_text = self.recipients_text.get('1.0', tk.END).strip()
            if not recipients_text:
                messagebox.showerror("Recipients Error", "No recipient emails found")
                return

            # Split and validate recipients
            recipients = [
                email.strip() for email in re.split(r'[,;\s]+', recipients_text) 
                if email.strip() and self.validate_email(email.strip())
            ]
            
            if not recipients:
                messagebox.showerror("Recipients Error", "No valid recipient emails found")
                return

            # Get SMTP details with explicit error handling
            try:
                smtp_server, smtp_port = self.get_smtp_details()
            except ValueError as smtp_error:
                messagebox.showerror("SMTP Configuration Error", str(smtp_error))
                return

            # Load email body
            try:
                with open(body_file_path, 'r', encoding='utf-8') as file:
                    body_content = file.read()
            except Exception as body_error:
                messagebox.showerror("File Error", f"Could not read email body: {body_error}")
                return

            # Confirm sending
            confirm = messagebox.askyesno(
                "Confirm Sending", 
                f"Are you sure you want to send emails to {len(recipients)} recipients?"
            )
            if not confirm:
                return

            # Locate send button (safely)
            send_button = None
            for widget in self.master.winfo_children():
                if isinstance(widget, ttk.Button) and widget.cget('text') == 'Send Emails':
                    send_button = widget
                    break

            # Disable send button during sending
            if send_button:
                send_button.config(state=tk.DISABLED)

            # Reset log
            self.status_log.config(state=tk.NORMAL)
            self.status_log.delete('1.0', tk.END)
            self.status_log.config(state=tk.DISABLED)

            # Prepare to send emails
            try:
                # Create SMTP connection
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                
                # Login to email
                try:
                    server.login(sender_email, sender_password)
                except smtplib.SMTPAuthenticationError:
                    messagebox.showerror("Authentication Failed", 
                        "Unable to login. Check your email and password.")
                    server.quit()
                    return
                except Exception as auth_error:
                    messagebox.showerror("Login Error", str(auth_error))
                    server.quit()
                    return

                # Send emails
                successful_sends = 0
                failed_sends = 0
                total_recipients = len(recipients)

                for index, recipient in enumerate(recipients, 1):
                    try:
                        # Construct email
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = recipient
                        msg['Subject'] = email_subject

                        # Attach body
                        msg.attach(MIMEText(body_content, 
                            'html' if body_file_path.lower().endswith('.html') else 'plain'))

                        # Attach files
                        for attachment_path in self.attachments:
                            try:
                                with open(attachment_path, 'rb') as file:
                                    part = MIMEApplication(file.read(), 
                                        Name=os.path.basename(attachment_path))
                                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                                    msg.attach(part)
                            except Exception as attach_error:
                                self.log_status(f"Attachment error for {attachment_path}: {attach_error}")

                        # Send email
                        server.send_message(msg)
                        
                        # Log successful send
                        self.log_status(f"Email sent to {recipient} ({index}/{total_recipients})")
                        successful_sends += 1

                        # Progress update
                        self.master.update_idletasks()

                        # Delay to prevent spam
                        time.sleep(5)

                    except Exception as send_error:
                        self.log_status(f"Failed to send to {recipient}: {send_error}")
                        failed_sends += 1

                # Close SMTP connection
                server.quit()

                # Final summary
                summary_message = (
                    f"Email Sending Complete\n"
                    f"Total Recipients: {total_recipients}\n"
                    f"Successful Sends: {successful_sends}\n"
                    f"Failed Sends: {failed_sends}"
                )
                messagebox.showinfo("Sending Complete", summary_message)

            except Exception as smtp_error:
                messagebox.showerror("SMTP Connection Error", str(smtp_error))

        except Exception as general_error:
            messagebox.showerror("Unexpected Error", str(general_error))

        finally:
            # Re-enable send button
            if send_button:
                send_button.config(state=tk.NORMAL)

    def get_smtp_details(self):
        """Get SMTP details with comprehensive validation"""
        selected_server = self.smtp_var.get()
        
        if selected_server == "Custom":
            server = self.custom_server_entry.get().strip()
            port = self.custom_port_entry.get().strip()
            
            if not server:
                raise ValueError("Custom SMTP server must be provided")
            
            try:
                port = int(port)
                if port <= 0 or port > 65535:
                    raise ValueError("Invalid port number")
            except ValueError:
                raise ValueError("Port must be a valid integer between 1 and 65535")
            
            return server, port
        elif selected_server in self.smtp_servers:
            return self.smtp_servers[selected_server]
        else:
            raise ValueError("Invalid SMTP server selection")

    def validate_email(self, email):
        """Validate email address format"""
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(email_regex, email) is not None

    def log_status(self, message):
        """Enhanced status logging"""
        self.status_log.configure(state=tk.NORMAL)
        self.status_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.status_log.configure(state=tk.DISABLED)
        self.status_log.see(tk.END)
        self.master.update_idletasks()

def main():
    root = tk.Tk()
    app = BulkEmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()