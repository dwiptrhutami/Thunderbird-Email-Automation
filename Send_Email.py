import mimetypes
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import encode_rfc2231
import html
import time
import win32com.client as win32
import os
from datetime import datetime
import pandas as pd
from RenameFile6 import def_max_row
from RenameFile6 import check_directory
import numpy as np
from io import StringIO
from GetEmail12 import load_config

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
    header_table = ['MODEL NAME', 'ARTICLE', 'STAGE', 'TANGGAL UPDATE', 'CWA LAMA', 'CWA', 'PIC', 'PIC LAMA', 'SEASON',
                    'FACTORY']
    factory_col = ['UPLOAD-PCI-CFM', 'UPLOAD-PGD-CFM', 'UPLOAD-PGS-CFM', 'UPLOAD-PCA-CFM']
    filename_list = pd.DataFrame([s.split(",") for s in file_upload], columns=['File Name'])

    data = []
        data.append(
            {"MODEL NAME": model_table, "ARTICLE": art_table, "STAGE": stage_table, "TANGGAL UPDATE": time_table,
             "CWA LAMA": cwa_value[1], "CWA": cwa_value[0], "PIC": pic_table, "PIC LAMA": pic_lama,
             "SEASON": season_table, "FACTORY": factory})
    df_table = pd.DataFrame(data, columns=header_table)
    html_table = df_table.to_html(index=False, header=True, border=1)
    temp_html_file = os.path.join(os.getcwd(), "table.html")
    with open(temp_html_file, "w", encoding="utf-8") as f:
        f.write(html_table)
    return html_table

def send_email_full(to_list, cc_list, subject_email, body_new, attachment_list=None, password_email=None):
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

    msg.attach(MIMEText(body_new, 'html')) #write body email in HTML or plain text

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
    config = load_config()
    file_path = config.get("file_path")
    file_path2 = config.get("file_path2")
    excel_tracking_file = config.get("excel_tracking_file")
    # season_sheet = config.get("season_sheet")
    df_html = pd.read_excel(file_path2, sheet_name="Tracking_New_Season", header=0)
    df_html2 = pd.read_excel(file_path2, sheet_name="ArtUploadedSAP", header=0)
    try:
        df_html2 = df_html2.applymap(lambda x: x.replace('\xa0', '').strip() if isinstance(x, str) else x)
    except Exception as e:
        print(f"{e} for df_html2")
    cc_list = []

    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    excel_app.DisplayAlerts = False

    workbook = excel_app.Workbooks.Open(file_path)

    sheet2 = workbook.Sheets("Email PIC Record")
    usedRange = sheet2.UsedRange.Value
    df = pd.DataFrame(usedRange[1:], columns=usedRange[0])
    start_send = def_max_row(df, 'Sent Email Time')
    max_send = def_max_row(df, 'Updated E-memo File Name')

    sheet3 = workbook.Sheets("PIC")
    usedRange3 = sheet3.UsedRange.Value
    df2 = pd.DataFrame(usedRange3[1:], columns=usedRange3[0])

    sheet4 = workbook.Sheets("ALL UPDATE")
    usedRange4 = sheet4.UsedRange.Value
    df3 = pd.DataFrame(usedRange4[1:], columns=usedRange4[0])

    teams_dict = {}
    list_div = df2['TEAM DIVISION'].dropna().tolist()
    name = df2.iloc[:, 9].dropna().tolist()
    email = df2['EMAIL'].dropna().tolist()
    for members, team in zip(name, list_div):
        if team not in teams_dict:
            teams_dict[team] = []
        teams_dict[team].append(members)
    my_dict = dict(zip(name, email))
    results = {team: {'members': [], 'attachments': [], 'datetime': [], 'index': []} for team in teams_dict.keys()}

    for k in range(start_send, max_send):
        folder_file = df['Updated E-memo File Name'][k]
        file_exits = check_directory(folder_file)
        if file_exits:
            file_name, file_extension = os.path.splitext(os.path.basename(folder_file))
            datetime_sort = file_name[-17:]
            pic_name = file_name.split('_')[0]
            for name, members in teams_dict.items():
                for person in members:
                    if person == pic_name:
                        results[name]['members'].append(person)
                        results[name]['attachments'].append(folder_file)
                        results[name]['datetime'].append(datetime_sort)
                        results[name]['index'].append(k)

    for team, attachment in results.items():
        file_upload = attachment['attachments']
        index_row = attachment['index']
        if not file_upload or team == 'ALL':
            continue
        else:
            to_list = teams_dict[team]
            email_list = set([my_dict[person] for person in to_list if person in my_dict])
            datetime_str = attachment['datetime']
            oldest_date = min(datetime_str)[0:10]
            newest_date = max(datetime_str)[0:10]
            oldest_time = min(attachment['datetime'])
            newest_time = max(attachment['datetime'])
            if oldest_date == newest_date:
                if oldest_time[11:15] > '0800' and newest_time[11:15] <= '1000':
                    times_in_range = f"08:01 - 10:00"
                    dates_in_range = oldest_date
                elif oldest_time[11:15] > '1000' and newest_time[11:15] <= '1300':
                    times_in_range = f"10:01 - 13:00"
                    dates_in_range = oldest_date
                elif oldest_time[11:15] > '1300' and newest_time[11:15] <= '1500':
                    times_in_range = f"13:01 - 15:00"
                    dates_in_range = oldest_date
                elif oldest_time[11:15] < '0800' and newest_time[11:15] <= '0800':
                    times_in_range = f"15:01 - 08:00"
                    dates_in_range = f"{oldest_date[0:8]}{int(oldest_date[8:10]) - 1:02} to {newest_date}"
                elif oldest_time[11:15] > '1500' and newest_time[11:15] > '1500':
                    times_in_range = f"15:01 - 08:00"
                    dates_in_range = f"{oldest_date} to {newest_date[0:8]}{int(newest_date[8:10]) + 1:02}"
                else:
                    dates_in_range = oldest_date
                    times_in_range = f"{oldest_time[11:13]}:01 - {int(newest_time[11:13]) + 1:02}:00"
                email_subject = f"E-MEMO Period: {dates_in_range} [{times_in_range}]"

            else:
                dates_in_range = f"{oldest_date} to {newest_date}"
                if newest_time[11:15] <= '0800':
                    if oldest_time[11:15] > '1500':
                        times_in_range = f"15:01 - 08:00"
                    else:
                        times_in_range = f"{oldest_time[11:13]}:01 - 08:00"
                elif '0800' < newest_time[11:15] <= '1000':
                    if oldest_time[11:15] > '1500':
                        times_in_range = f"15:01 - 10:00"
                    else:
                        times_in_range = f"{oldest_time[11:13]}:01 - 10:00"
                elif '1000' < newest_time[11:15] <= '1300':
                    if oldest_time[11:15] > '1500':
                        times_in_range = f"15:01 - 13:00"
                    else:
                        times_in_range = f"{oldest_time[11:13]}:01 - 13:00"
                elif '1300' < newest_time[11:15] <= '1500':
                    if oldest_time[11:15] > '1500':
                        times_in_range = f"15:01 - 5:00"
                    else:
                        times_in_range = f"{oldest_time[11:13]}:01 - 15:00"
                else:
                    times_in_range = f"{oldest_time[11:13]}:01 - {int(newest_time[11:13]) + 1:02}:00"
                email_subject = f"E-MEMO Period: {dates_in_range} [{times_in_range}]"
            html_body = html_table_ememo(file_upload, df_html, df_html2, df3)
            message = f"""\
                    <html>
                        <head>
                        </head>
                        <body>
                            <p><b>Dear {team} and TEAM,</b></p>
                            <p>Berikut terlampir update E-MEMO per tanggal {dates_in_range} pukul {times_in_range}.</p>
                            Terima kasih.</p>
                            <br>
                            {html_body}
                            <br>
                            <p>Best Regards,</p>
                            <p>Sulis | SAP DC Team | Ext: 6642/6699 | Teams: Sulis Suhartini</p>
                        </body>
                    </html>

                    """
            # compose_email_in_thunderbird(email_list, email_subject, body_new, file_upload)
            send_email_full(email_list, cc_list, email_subject, message, attachment_list=file_upload,
                            password_email="adidas.79")

            now = datetime.now().strftime('%Y_%m_%d %H:%M:%S')
            print(f"Sending E-mail {email_subject} to {team} has been run successfully at {now}...")
            for m in range(len(index_row)):
                row_value = index_row[m]
                df.loc[row_value, 'Sent Email Time'] = now
                df.loc[row_value, 'Updated E-memo File Name'] = get_filename(file_upload[m])
                df.loc[row_value, 'Subject'] = email_subject
                df.loc[row_value, 'Team'] = f"Team {team}"

            row_count = sheet4.Cells(sheet4.Rows.Count, "A").End(3).Row
            table_update = pd.read_html(StringIO(html_body))[0]
            for row_num, row in enumerate(table_update.itertuples(index=False, name=None), start=row_count + 1):
                for col_num, value in enumerate(row, start=1):
                    sheet4.Cells(row_num, col_num).Value = value if pd.notna(value) else ""
                    for i in range(7, 13):
                        border = sheet4.Cells(row_num, col_num).Borders(i)
                        border.LineStyle = 1
                        border.Weight = 2

    for row in range(1, df.shape[0] + 1):
        sheet2.Cells(row + 1, 1).Value = df['Updated E-memo File Name'].loc[row - 1]
        sheet2.Cells(row + 1, 3).Value = df['Sent Email Time'].loc[row - 1]
        sheet2.Cells(row + 1, 4).Value = df['Subject'].loc[row - 1]
        sheet2.Cells(row + 1, 5).Value = df['Team'].loc[row - 1]

    time.sleep(2)
    workbook.Save()
    time.sleep(2)
    workbook.Close(SaveChanges=True)

    excel_app.Quit()

    """REVISED BY DWI 060226"""
    """UPDATE: EFICIENCY CODING"""
    """SEND NA PIC E-MEMO"""
    df5 = pd.read_excel(excel_tracking_file, sheet_name='Folder F_NA PIC', header=0).fillna("")
    start_row = df5["SENT TIME"].replace("", pd.NA).dropna(how="all").shape[0]
    max_row = len(df5)
    file_upload2 = []
    art_upload = []
    received_time = []
    for k in range(start_row, max_row):
        file_upload2.append(df5.loc[k, 'Folder Type'])
        art_upload.append(df5.loc[k, 'Filename'].split("_")[2])
        received_time.append(df5.loc[k, 'Filename'][-22:-5])
    df_table2 = pd.DataFrame(zip(art_upload, received_time), columns=["ARTICLE", "RECEIVED TIME"])
    html_table2 = df_table2.to_html(index=False, header=True, border=1)
    email_subject2 = "E-MEMO BELUM BAGI PIC APPLY 4 PROSES"
    message2 = f"""\
            <html>
                <head>
                </head>
                <body>
                    <p><b>Dear IDA,</b></p>
                    <p>Berikut terlampir ARTICLE belum bagi PIC untuk apply 4 proses.</p>
                    <p>Update E-MEMO per tanggal RECEIVED TIME (tanggal terima Email dari Developer).</p>
                    <p>Terima kasih.</p>
                    <br>
                    {html_table2}
                    <br>
                    <p>Best Regards,</p>
                    <p>Sulis | SAP DC Team | Ext: 6642/6699 | Teams: Sulis Suhartini</p>
                </body>
            </html>
            """
    email_list2 = ['dwi.sap@pci.co.id', 'indah.sap@pci.co.id']
    cc_list2 = []
    send_email_full(email_list2, cc_list2, email_subject2, message2, attachment_list=file_upload2,
                    password_email="adidas.79")
    now = datetime.now().strftime("%Y_%m_%d-%H%M%S")
    df5.loc[start_row:max_row, 'SENT TIME'] = now

    with pd.ExcelWriter(excel_tracking_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df5.to_excel(writer, sheet_name='Folder F_NA PIC', index=False)

    print("Waiting for the next schedule run time...")


if __name__ == '__main__':
    send_email()


