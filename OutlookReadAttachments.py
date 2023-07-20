import imaplib
import email
from email.header import decode_header
import webbrowser
import os
from datetime import datetime, timedelta
import pytz
import pandas as pd


username = ""
password = ""
imap_server = "outlook.office365.com"


def clean(s):
    # clean text for creating a folder
    s_list = s.split()
    folder = "Amex Downloads "
    if "from" in s_list:
        s_ind = s_list.index("from")
        for i in range(s_ind +1,len(s_list)):
            folder += s_list[i]
            folder += " "
    return folder.replace("/","-")

created_file = input("What would you like the control file to be named? ")
if not created_file.endswith('.xlsx') or created_file.endswith('.xls'):
        created_file += ".xlsx"
    

def combine(folder_name, amex_file_name = 0, sheet_name = 0):
    if amex_file_name != 0 and sheet_name != 0:
        pass
    merged_data = []

    onlyfiles = [f for f in os.listdir(folder_name) if os.path.isfile(os.path.join(folder_name, f)) and not f.startswith('.DS_Store')]
    for file in onlyfiles:
        if file.endswith('.xlsx') or file.endswith('.xls'):
            file_path = os.path.join(folder_name, file)
            data = pd.read_excel(file_path)
            merged_data.append(data)

    merged_data = pd.concat(merged_data,ignore_index=True)
    return merged_data
    


imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
imap.login(username, password)

status, messages = imap.select("INBOX")

#specify date and time range to read emails
start_date = datetime(2023,6,26)
end_date = datetime(2023,6,29)

#set days equal to desired delta, if hours, replace days= with hours=
start_date_str = (start_date - timedelta(days=1)).strftime("%d-%b-%Y")
end_date_str = (end_date + timedelta(days=1)).strftime("%d-%b-%Y")


search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
status, messages = imap.search(None, search_criteria)

message_ids = messages[0].split()

for message_id in message_ids:
    
    res, msg = imap.fetch(message_id, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                if encoding is not None:
                    # if it's bytes and the encoding is not None, decode to str
                    subject = subject.decode(encoding)
                else:
                    # if the encoding is None, assume it's already a Unicode string
                    subject = subject.decode()
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)
            print("Subject:", subject)
            print("From:", From)

            date_str = msg["Date"]
            date_str = date_str.split(' (')[0]
            date = datetime.strptime(date_str, "%a, %d %b %Y %H:%M:%S %z")
            # Convert the timestamp to your desired timezone
            timezone = pytz.timezone("US/Central")
            date = date.astimezone(timezone)
            print("Sent Date and Time:", date.strftime("%Y-%m-%d %H:%M:%S %Z"))

            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body)
                        pass
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        filename = os.path.basename(filename)
                        if filename:
                            desktop_path = os.path.expanduser("~/Desktop")
                            folder_name = clean(subject)
                            folder_path = os.path.join(desktop_path, folder_name)
                            inner_path = os.path.join(folder_path,"Updated Downloaded Files")
                            if not os.path.isdir(folder_path):
                                # make a folder for this email (named after the subject)
                                os.makedirs(folder_path)
                                os.makedirs(inner_path)
                            filepath = os.path.join(inner_path, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                webbrowser.open(filepath)
        print("="*100)
if os.path.exists(folder_path):
    control_file_path = os.path.join(folder_path,created_file)
    control_data = combine(inner_path)
    control_data.to_excel(control_file_path)


# close the connection and logout
imap.close()
imap.logout()

print("All files downloaded")





