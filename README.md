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
IMAP_PORT     = 993
Security tip: Use an app-specific password rather than your main account password. Never commit credentials to version control.

# Usage
Send an Email
from sender import send_email

send_email(
    to="recipient@example.com",
    subject="Hello from Thunderbird Automation",
    body="This email was sent automatically.",
    attachment_path="report.pdf"  # optional
)

# Read Emails
from reader import fetch_emails

emails = fetch_emails(folder="INBOX", limit=10)
for mail in emails:
    print(mail["from"], mail["subject"], mail["date"])
    print(mail["body"])

# Download Attachments
from downloader import download_attachments

download_attachments(
    folder="INBOX",
    save_dir="./downloads",
    limit=20
)


# Project Structure
thunderbird-email-automation/
├── config.py              # Email credentials and server settings
├── sender.py              # SMTP email sending logic
├── reader.py              # IMAP email reading logic
├── downloader.py          # Attachment download logic
├── requirements.txt       # Python dependencies
└── README.md

# How It Works
Module	Protocol	Description
sender.py	SMTP	Connects to SMTP server, composes MIME messages, sends email
reader.py	IMAP	Connects to IMAP server, fetches and parses email headers/body
downloader.py	IMAP	Scans emails for attachments and saves them locally
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

Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you'd like to change.
