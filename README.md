# 🚀 Email Sender System

A lightweight, user-friendly desktop application to send bulk emails via SMTP (Gmail supported). Built with Python and Tkinter, this tool supports CSV/Excel recipient import, single attachment support, multithreaded sending, and a responsive GUI that remains interactive during sends.

Author: SauravDnj • GitHub: https://github.com/SauravDnj

---

## 📋 Table of contents
- Features
- Project structure
- Requirements
- Installation
- Quick start
- CSV / Excel format
- Security (Gmail)
- Troubleshooting
- Development & contribution
- License & contact

---

## ✨ Features
- 🔐 SMTP configuration (server, port, username, password) with TLS support  
- 📧 Gmail defaults: smtp.gmail.com : 587  
- 📥 Import recipients from CSV (.csv) and Excel (.xlsx/.xls) — first column = email  
- ➕ Add recipients manually and Clear All option  
- 🧾 Real-time recipient count and send status log (success / failure)  
- 📨 Subject and message editor (plain text)  
- 📎 Single file attachment support (any file type)  
- ⚙️ Multithreaded sending to keep the UI responsive  
- ⏱ Progress bar indicating send progress

---

## 🛠 Requirements
- Python 3.7+
- Built-in: tkinter, smtplib, email, threading, csv, os
- External:
  - pandas (for reading Excel files)

Install pandas:
```bash
pip install pandas
```

Note: On Linux, you may need to install tkinter separately (e.g., `sudo apt-get install python3-tk`).

---

## ⚡ Installation
Clone the repository and run:
```bash
git clone https://github.com/SauravDnj/Email-Sender-System.git
cd Email-Sender-System
pip install pandas
python email_sender.py
```
The GUI will launch.

---

## 🧭 Quick start — how to use
1. Configure SMTP
   - Enter your SMTP username (email) and password (for Gmail, use an App Password).
   - Server: `smtp.gmail.com`  Port: `587` (default for Gmail + STARTTLS).

2. Add recipients (choose one)
   - Load from CSV — first column must contain email addresses.
   - Load from Excel (.xlsx/.xls) — first column must contain email addresses.
   - Add manually — enter a single email and click Add.
   - Use Clear All to remove recipients.

3. Compose your email
   - Enter Subject and the Message body (plain text).
   - (Optional) Add one attachment.

4. Send
   - Click "Send Emails".
   - The application will authenticate and send messages in a background thread.
   - Watch progress via the progress bar and the status log for each recipient.

---

## 📄 CSV / Excel format
- Place recipient email addresses in the first column.
- Additional columns (e.g., name) are permitted but not used unless you implement personalization.
Example CSV:
```csv
email,name
saurav@example.com,Saurav
sauravdnj@example.com,SauravDnj
```

---

## 🔒 Gmail users — App Password (required)
To use Gmail SMTP securely:
1. Enable 2-Step Verification for your Google account.
2. Go to Security → App passwords.
3. Create an app password (Mail or Other) and use it in the app instead of your main password.

Never commit credentials to source control.

---

## 🐞 Troubleshooting
- Authentication fails:
  - Ensure email and app password are correct; 2FA must be enabled for App Passwords.
- Connection errors:
  - Verify SMTP host/port. Ensure network/firewall permits outbound SMTP on the chosen port.
- Excel import fails:
  - Confirm `pandas` is installed and file is a valid `.xlsx`/`.xls`.
- Tkinter not available:
  - Install tkinter for your OS (e.g., `python3-tk` on Debian/Ubuntu).

For further debugging, check the status/log output in the GUI for per-recipient errors.

---

## 🛠 Development notes & future improvements
Potential enhancements:
- HTML email templates & inline images
- Multiple attachments
- Personalization using CSV/Excel columns (e.g., names)
- Scheduling, batching, and rate-limiting for large lists
- Save/load SMTP profiles and improved credential storage (encrypted)
- Dark mode and accessibility improvements

---

## 🤝 Contributing
Contributions are welcome. Suggested workflow:
1. Fork the repo
2. Create a feature branch: git checkout -b feature/your-feature
3. Commit and push your changes
4. Open a Pull Request and describe the change

Please include tests and update this README for user-facing changes.

---

## ✉️ Contact
Author: SauravDnj  
GitHub: https://github.com/SauravDnj  
Email: sauravdanej24@gmail.com
