# Thunderbird Email Automation

A Python-based automation toolkit for Mozilla Thunderbird that enables sending, reading, and downloading email attachments programmatically — without relying on Thunderbird's UI.

## Features

- **Send Email** — Compose and send emails with or without attachments via SMTP
- **Read Email** — Fetch and parse emails from your inbox using MBOX File
- **Download Attachments** — Automatically save email attachments to a local directory

## Requirements

- Python 3.8+
- Mozilla Thunderbird installed and configured with at least one email account
- IMAP and SMTP access enabled on your email account

### Python Dependencies

```bash
pip install -r requirements.txt
requirements.txt

smtplib
email
mailbox
BytesParser
Note: packages are part of the Python standard library. No additional packages are required for basic usage.

Configuration
Create a config.py or .env file with your email credentials:

EMAIL_ADDRESS = "your-email@example.com"
EMAIL_PASSWORD = "your-app-password"
SMTP_SERVER   = "smtp.example.com"   # e.g., smtp.gmail.com
SMTP_PORT     = 587
Security tip: Use an app-specific password rather than your main account password. Never commit credentials to version control.

# Usage
Send an Email
from Send_Email.py import send_email

send_email_full(
    to=recipient_addresses,
    cc=cc_list,
    subject=email_subject,
    body=body_email,
    attachment_list=file_path,
    password_email=password_email
)

# Read Emails
# Download Attachments
from ReadEmail.py import check_emails

check_emails(mbox_path,inbox_dir,excel_tracking_file,day_yesterday)

# Project Structure
thunderbird-email-automation/
├── ReadEmail.py              # Attachment download logic
├── Send_Email.py             # SMTP email sending logic
└── README.md

# How It Works
Module	Protocol	Description
ReadEmail.py	MBOX Local File	Scans emails for attachments and saves them locally
Send_Email.py	SMTP	Connects to SMTP server, composes MIME messages, sends email
Common Issues
Authentication error

Make sure IMAP/SMTP access is enabled in your email provider settings
For Gmail, generate an App Password under your Google Account security settings
Connection timeout

Verify the IMAP/SMTP server addresses and ports in config.py
Check that your firewall allows outbound connections on ports 587 (SMTP) and 993 (IMAP)
Thunderbird profile

This tool communicates directly with your mail server via standard protocols — Thunderbird does not need to be running
License
MIT License. See LICENSE for details.