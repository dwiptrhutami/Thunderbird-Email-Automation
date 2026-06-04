#import packages
import os
import mailbox
import email
from email import policy
from email.parser import BytesParser
import openpyxl
import pandas as pd
from openpyxl import load_workbook
import datetime
import time
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
from email import header
from openpyxl.styles import PatternFill
from openpyxl.workbook import Workbook
import json
import schedule
import re
import base64

def save_config(config):
    with open("config.json","w") as f:
        json.dump(config,f,indent=4, cls=CustomEncoder)
    print("Configuration saved!")

def load_config():
    with open("config.json","r") as f:
        config = json.load(f)
    return config

def generate_unique_filename(output_dir,filename):
    """Generate an unique filename by appending a counter or timestamp"""
    base,extension = os.path.splitext(filename)
    counter=1
    unique_filename = filename

    """Check if the file already exists and generate a new name"""
    while os.path.exists(os.path.join(output_dir,unique_filename)):
        unique_filename = f"{base}_{counter}{extension}"
        counter+=1
    return unique_filename

def safe_windows_path(filepath):
    """Rename file path with limitation characters"""
    max_path = 250 #for safely running
    base_dir = os.path.abspath(output_dir)
    ext = os.path.splitext(filepath)[1]
    filename = os.path.split(filepath)[1]

    """Count max character for filename"""
    max_filename_len = max_path - len(base_dir) - len(os.sep) - len(ext)

    # Minimal len filename
    max_filename_len = max(10, max_filename_len)
    name_only = os.path.splitext(filename)[0]
    safe_name = name_only[:max_filename_len]
    return os.path.join(base_dir, safe_name + ext)

def save_attachment(part, filename, email_from,subject,date):
    """Save the attachment if it has not been saved before"""
    # create a file path to save the attachment
    unique_filename = generate_unique_filename(output_dir,filename)
    #save the attachment
    filepath = os.path.join(output_dir,unique_filename)
    filepath = safe_windows_path(filepath)
    with open(filepath, 'wb') as f:
        f.write(part)
        print(f"Attachment saved: {filepath} from Subject: {subject}")

    #update the excel tracking file
    update_tracking_file(unique_filename,filepath,email_from,subject,date)
    return True

def update_tracking_file(filename, filepath, email_from, subject,date):
    """Update the Excel tracking file with the latest attachment info"""
    if os.path.exists(excel_tracking_file):
        workbook = openpyxl.load_workbook(excel_tracking_file)
        sheet = workbook["Folder A"]
    else:
        workbook = Workbook()
        workbook.active.title = "Folder A"
        #add headers for the first time
        sheet = workbook.active
        sheet.append(["Subject", "Filename", "Filepath", "Email From", "Timestamp","Email Date", "Status", "Time Update"])

    # append the latest data
    tracking_data = (subject, filename, filepath, email_from, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), date)
    sheet.append(tracking_data)
    workbook.save(excel_tracking_file)
    print(f"Tracking updated in {excel_tracking_file}")

#update tracking file
def update_subject_tracking(subject,email_from,date,status):
    """Update the Excel tracking file with the latest attachment info"""
    if os.path.exists(excel_tracking_file):
        workbook = openpyxl.load_workbook(excel_tracking_file)
        sheet2 = workbook["Subject Email List"]
    else:
        workbook = Workbook()
        sheet2 = workbook.create_sheet(title="Subject Email List",index=0)
        #add headers for the first time
        sheet2.append(["Subject", "Email From", "Timestamp", "Email Date", "Status"])

    tracking_data2 = (subject, email_from, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), date, status)
    sheet2.append(tracking_data2)
    workbook.save(excel_tracking_file)
    print(f"Subject updated in {excel_tracking_file}")

    color_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    if status == "FORWARD MESSAGE" or status == "CHECK MANUAL EMAIL":
        for col in sheet2.iter_cols(min_row=1,max_row=1):
            header = col[0].column
            sheet2.cell(row=sheet2.max_row,column=header).fill = color_fill
            #save the workbook
            workbook.save(excel_tracking_file)

def load_processed_subjects_from_excel():
    """Generate an Excel tracking file to track and record reading email process"""
    processed=set()
    try:
        workbook = openpyxl.load_workbook(excel_tracking_file)
        sheet = workbook["Subject Email List"]
        for row in sheet.iter_rows(min_row=2,values_only=True):
            subject = row[1]
            received_date_str = row[3]
            if subject and received_date_str:
                processed.add((subject,received_date_str))
    except FileNotFoundError:
        workbook = Workbook()
        # add headers for the first time
        sheet = workbook.create_sheet(title="Subject Email List",index=0)
        sheet.append(["Run Time", "Subject", "Email From", "Email Received Date", "Remarks"])
        workbook.save(excel_tracking_file)
        print(f"Tracking updated in {excel_tracking_file}")
    return processed

def encode_subject(subject):
    """Cleaning Encoded Subject in utf-8 format to convert Chinese Characters"""
    decode_header = header.decode_header(subject)
    subject_clean = ""
    for part, encoding in decode_header:
        if isinstance(part,bytes):
            subject_clean += part.decode(encoding or 'utf-8',errors='replace')
        else:
            subject_clean += part
    return subject_clean

def character_emails(message):
    email_date_str = re.findall(r"^Date:\s*(.+)$", message.as_string(), flags=re.IGNORECASE | re.MULTILINE)
    email_date = email.utils.parsedate_to_datetime(email_date_str[0]).date()
    received_date = email.utils.parsedate_to_datetime(email_date_str[0]).strftime('%Y-%m-%d %H:%M:%S')
    email_from = encode_subject(
        (re.findall(r"^From:\s*(.+)$", message.as_string(), flags=re.IGNORECASE | re.MULTILINE)[0]))
    subject = encode_subject(
        (re.findall(r"^Subject:\s*(.+)$", message.as_string(), flags=re.IGNORECASE | re.MULTILINE)[0]))
    return (email_from,subject,email_date,received_date)


def download_attachments(message,subject,email_from,received_date):
    # loop through all the emails in the mbox file
    msg = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
    filename_list = {}
    if msg.is_multipart():
        # loop through the email parts to find attachments
        for part in msg.walk():
            raw = part.get_payload()
            filename = part.get_filename()
            if filename and not filename.endswith('.eml'):
                b64_data = re.sub(r'\s+', '', raw)
                try:
                    decoded = base64.b64decode(b64_data)
                    filename_list[filename] = decoded
                except Exception as e:
                    print(f"Failed decoding {filename}: {e}")
            elif filename and filename.endswith('.eml'):
                try:
                    filename_list[filename] = part
                except Exception as e:
                    print(f"Failed decoding {filename}: {e}")
    else:
        raw_bytes = message.as_bytes()

        text = raw_bytes.decode("utf-8", errors="ignore")
        # Find MIME boundary (if it's any)
        boundary_match = re.search(r'boundary="([^"]+)"', text, re.IGNORECASE)
        if not boundary_match:
            # raise ValueError("No MIME boundary found")
            filename = "CHECK MANUAL"
            filename_list[filename] = None
            print(f"No MIME boundary found: {subject} from {email_from}: {received_date} ")
        else:
            boundary = boundary_match.group(1)
            delimiter = "--" + boundary
            parts = text.split(delimiter)
            for part in parts:
                # Skip empty segments
                if "Content-Disposition" not in part:
                    continue

                # Extract normal filename=
                normal = re.search(r'filename="([^"]+)"', part)
                filename = normal.group(1) if normal else None
                # Extract Split filename= (filename*0, filename*1)
                split_parts = {}
                for m in re.finditer(r'filename\*(\d+)="([^"]+)"', part):
                    idx = int(m.group(1))
                    split_parts[idx] = m.group(2)
                if split_parts:
                    filename = "".join(split_parts[i] for i in sorted(split_parts.keys()))

                if not filename:
                    continue  # skip if no filename on this MIME part

                # Extract Base64 content in this part only
                b64_match = re.search(r'\n\n([\s\S]+?)$', part.strip())
                if not b64_match:
                    print("No Base64 found in part for:", filename)
                b64_data = re.sub(r'\s+', '', b64_match.group(1))

                # Fix padding
                pad = len(b64_data) % 4
                if pad:
                    b64_data += "=" * (4 - pad)

                try:
                    decoded = base64.b64decode(b64_data)
                    filename_list[filename] = decoded
                    print("Decoded:", filename)

                except Exception as e:
                    print(f"Failed decoding {filename}: {e}")

    check_eml = False
    for file_list, part_file in filename_list.items():
        if '.eml' in file_list:
            check_eml = True

    return (check_eml,filename_list)


#function to check latest email today (reversed inbox date)
def read_emails(mbox_path,output_dir,excel_tracking_file,wanted_extensions,day_yesterday):
    print(f"Checking emails at {datetime.now().strftime('%H:%M')}")

    """Initialize Folder"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    """Initialize Time Range"""
    today = datetime.now().date()
    yesterday = today - timedelta(days=day_yesterday)

    """Load Previously Processed Subject Emails from Tracking Excel File"""
    processed_subjects = load_processed_subjects_from_excel()

    """Initialize MBOX Path File"""
    mbox = mailbox.mbox(mbox_path)

    for message in reversed(mbox): #using reversed to run process reading email from Newest to Oldest
        email_from,subject,email_date,received_date = character_emails(message)

        if email_date > yesterday:
            # print(subject, received_date)
            if (subject, received_date) not in processed_subjects:
                check_eml, filename_list = download_attachments(message, subject, email_from, received_date)
                if check_eml:
                    update_tracking_file("<FORWARD MESSAGE>", "MUST CHECK MANUAL", email_from, subject, received_date)
                    update_subject_tracking(subject, email_from, received_date, "FORWARD MESSAGE")
                elif filename_list == {}:
                    update_subject_tracking(subject, email_from, received_date, "NO ATTACHMENT")
                elif "CHECK MANUAL" in filename_list.keys():
                    update_subject_tracking(subject, email_from, received_date, "CHECK MANUAL EMAIL")
                else:
                    for file_list, part_file in filename_list.items():
                        if any(file_list.lower().endswith(ext) for ext in wanted_extensions):
                            print(f"Attachment saved: {file_list} from Subject: {subject}")
                            # update_tracking_file(file_list, output_dir + file_list, email_from, subject, received_date)
                            save_attachment(part_file, file_list, email_from, subject, received_date)
                            update_subject_tracking(subject, email_from, received_date, "OK")
                        else:
                            print(f"Skipping {file_list} from Subject: {subject}")
                            update_subject_tracking(subject, email_from, received_date, "CHECK MANUAL")
            else:
                print(f"Skipping email with subject: {subject} {received_date} already processed")

    # import RunAll
    # RunAll.backup()
    # import RenameFile
    # RenameFile.run_program_rename()

def run_program_all(scheduled_time):
    print(f"Saving emails at {datetime.now().strftime('%H:%M')} (scheduled for: {scheduled_time})")
    check_emails(mbox_path, output_dir, excel_tracking_file, wanted_extensions, day_yesterday)

def schedule_tasks():
    schedulers = ["08:03", "10:25", "12:55", "14:55"]
    for hour in schedulers:
        schedule.every().day.at(hour).do(run_program_all, hour)

def run_pending_task():
    while True:
        schedule.run_pending()
        time.sleep(1)

def run_all():
    config = {}
    mbox_path = config.get("mbox_path")
    # directory to save attachments
    output_dir = config.get("output_dir")
    # path to Excel Tracking File
    excel_tracking_file = config.get("excel_tracking_file")
    # define a list of unwanted file extensions
    wanted_extensions = config.get("wanted_extensions")
    day_yesterday = config.get("day_yesterday")
    save_config(config)


#start the email check process
if __name__ == '__main__':
    config = load_config()
    # path to your thunderbird mbox file
    mbox_path = config.get("mbox_path")
    # directory to save attachments
    output_dir = config.get("output_dir")
    # path to Excel Tracking File
    excel_tracking_file = config.get("excel_tracking_file")
    # define a list of unwanted file extensions
    wanted_extensions = config.get("wanted_extensions")
    day_yesterday = config.get("day_yesterday")
    print("Scheduler started. Waiting for tasks..")
    schedule_tasks()
    run_pending_task()

