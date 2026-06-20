import mimetypes
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import html
import os
from datetime import datetime
import pandas as pd

def detail_email(recipient_name,df_html):
    email_subject = "TEST SENDING EMAIL FROM MOZILLA THUNDERBIRD"
    html_body = html_table(df_html)
    message = f"""\
            <html>
                <head>
                </head>
                <body>
                    <p><b>Dear {recipient_name},</b></p>
                    <p>This is an automation testing Email. Please ignore the message.</p>
                    Thank you in advance.</p>
                    <br>
                    {html_body}
                    <br>
                    <p>Best Regards,</p>
                    <p>Sender | Division Team | Ext: 021xxxxxx</p>
                </body>
            </html>
    
            """
    return email_subject, message

def html_table(df_html):
    html_table = df_html.to_html(index=False, header=True, border=1)
    temp_html_file = os.path.join(os.getcwd(), "table.html")
    with open(temp_html_file, "w", encoding="utf-8") as f:
        f.write(html_table)
    return html_table

def send_email_full(to_list, cc_list, subject_email, body_email, attachment_list=None, password_email=None):
    # SMTP server details (for Thunderbird SMTP)
    smtp_server = "server@domain.co.id"
    smtp_port = 123 #input port Server
    smtp_user = "username@email.co.id"
    smtp_password = password_email

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['Subject'] = subject_email
    msg['To'] = ",".join(to_list)
    msg['Cc'] = ",".join(cc_list)

    msg.attach(MIMEText(body_email, 'html')) #write body email in HTML or plain text

    # Path to the attachment file
    attachment_path = attachment_list

    for file_attachment in attachment_path:
        if os.path.exists(file_attachment):
            filename = os.path.basename(file_attachment)

            ctype, encoding = mimetypes.guess_type(file_attachment)
            if ctype is None:
                ctype = "application/octet-stream"

            maintype, subtype = ctype.split("/", 1)

            with open(file_attachment, "rb") as f:
                part = MIMEBase(maintype, subtype)
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=filename)
            msg.attach(part)
        else:
            print(f"Attachment not found at {attachment_path}")

    # Connect to the SMTP server
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        # Log in to the SMTP server
        server.login(smtp_user, smtp_password)
        text = msg.as_string()
        server.sendmail(smtp_user, to_list, text)
        print(f"Email {subject_email} has been sent successfully to {to_list}")
    except Exception as e:
        print(f"Error: {e}")


def send_email():
    """Initialize Variables"""
    file_path = r"C:\Users\Downloads\ComputerSales.csv"
    password_email = "test.123"

    df_html = pd.DataFrame({
    "Date": ["2026-06-01", "2026-06-02", "2026-06-03", "2026-06-04", "2026-06-05"],
    "Product": ["Laptop", "Mouse", "Keyboard", "Monitor", "Headset"],
    "Units_Sold": [12, 45, 30, 18, 25],
    "Unit_Price": [1200, 25, 50, 300, 80]
    })

    recipient_addresses = ["email1@domain.co.id", "email2@domain.co.id", "email3@domain.co.id", "email4@domain.co.id"]
    cc_list = ["email5@domain.co.id"]
    recipient_name = "NAME"
    email_subject, body_email = detail_email(recipient_name,df_html)

    # compose_email_in_thunderbird(email_list, email_subject, body_new, file_upload)
    send_email_full(recipient_addresses, cc_list, email_subject, body_email, attachment_list=file_path, password_email=password_email)

    now = datetime.now().strftime('%Y_%m_%d %H:%M:%S')
    print(f"Sending E-mail {email_subject} to {recipient_name} has been run successfully at {now}...")

if __name__ == '__main__':
    send_email()
