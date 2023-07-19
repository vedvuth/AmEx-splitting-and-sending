from fileinput import filename
import pandas as pd
import os
from email.mime.text import MIMEText
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


#server connection information
port = 587  # For starttls
smtp_server = "smtp-mail.outlook.com" #For outlook, gmail: smtp.gmail.com
sender_email = "" 
password = "" #For gmail, requires an app-password via two factor authentication


#Create user class for Card Users
class User():
    #initiate email address, card name, link, and email name variables
    def __init__(self, email_address, card_name, link = None, email_name = 0):
        if email_name == 0:
            self.email_name = card_name
        else:
            self.email_name = email_name
        self.card_name = card_name
        self.email_address = email_address
        self.link = link
    #simple print user function
    def __str__(self):
        s = ("Email Name: " + self.email_name + ", " + "Card Name: " + self.card_name + ", "\
            + "Email Adress: " + self.email_address + "Link: " + self.link)
        return s
    
    #Send email function
    def sendEmail(self, p_week, c_week, file_name):
        #Create email content
        subject = 'Action Needed: Amex Expenses needed from ' + p_week + ' to ' + c_week
        #test_subject = "Automated Test Email, Please Ignore"
        #text = "Hello " + self.email_name + "," + '''\n\nPlease reply, enter the empty items, and upload your receipt to the spreadsheet attached below. I have added the link to the folder below for your convenience.\n\n''' + self.link + '''\n\nThank you,\nAccounting Team'''
        text = "Hello {},\n\nPlease open the spreadsheet below, fill in the missing columns, and attach the updated spreadsheet in a reply to this email by Friday. Please also upload your purchase reciepts to the link below.\n\n{}\n\nThank you,\nAccounting Team".format(self.email_name,self.link)
        #test_text = "Hello " + self.email_name + "," + '''\n\nThis is an automated test email. Please ignore.\n\nThank you,\nAccounting Team'''
        # Create a multipart message object
        message = MIMEMultipart()
        message["Subject"] = subject
        message["From"] = sender_email
        message["To"] = self.email_address
        # Add the text to the message
        message.attach(MIMEText(text, "plain"))

        # Add the attachment to the message
        attachpath = os.path.expanduser("~/Desktop")
        file_path = os.path.join(attachpath, file_name)
        attachment_path = file_path  #Specify the actual attachment file path
        attachment_filename = file_name  #Replace with the desired attachment filename
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {attachment_filename}",
            )
            #attach file
            message.attach(part)

        # Create a secure SSL context
        context = ssl.create_default_context()

        # Send the email with attachment
        with smtplib.SMTP(smtp_server, port) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(sender_email, password)
            server.sendmail(sender_email, self.email_address, message.as_string())
        return

#Split excel file by card holder, and download files into specified folder
def download_splits(amex_file_name, sheet_name,folder_name):
    #read in intial amex file excel sheet
    #process the file according to the amex spreadsheet formatting
    data = pd.read_excel(amex_file_name,sheet_name,skiprows = 6)
    keep_col = ["Date", "Description", "Card Member", "Amount", "Entity", "PROPERTY INFO"]
    columns_to_delete = [col for col in data.columns if col not in keep_col]
    data = data.drop(columns=columns_to_delete)
    data["Purchase Description"] = ""
    data["Notes"] = ""
    data.rename(columns = {"PROPERTY INFO":"Property Information"}, inplace=True)
    #group data by Card Member column
    df = data.groupby("Card Member")

    #create and download files
    for name, group in df:
        #create file name
        docname = name + ".xlsx"
        file_name = folder_name + "/" + docname
        #download excel file to computer with specified pathing
        group.to_excel(file_name, index=False)
    return 

#Function to combine excel files in specified folder into one excel file - excel files must be in the same format
def combine(amex_file_name,sheet_name,folder_name):
    merged_data = [] #list for storing files

    #get paths for each indivdual file via iteration through folder
    onlyfiles = [f for f in os.listdir(folder_name) if os.path.isfile(os.path.join(folder_name, f)) and not f.startswith('.DS_Store')]

    #reformat filenames, read-in files as DataFrames, and add to merged_data list
    for file in onlyfiles:
        if file.endswith('.xlsx') or file.endswith('.xls'):
            file_path = os.path.join(folder_name, file)
            data = pd.read_excel(file_path)
            merged_data.append(data)
    
    #concatenate all DataFrames in merged_data list
    merged_data = pd.concat(merged_data,ignore_index=True)
    
    #return single combined DataFrame
    return merged_data

#create folder function
def create_folder(folder_name):
    desktop_path = os.path.expanduser("~/Desktop")
    folder_path = os.path.join(desktop_path, folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

#Create list with user info by reading in provided user info file
def userL(file_name):
    user_info_list = []
    data = pd.read_excel(file_name)
    for i in range(len(data)):
        data_sort = data.loc[i]
        user_info_list.append([data_sort[0],data_sort[1],str(data_sort[2]),str(data_sort[3])])
    return user_info_list

def main():
    userclassList = []

    amex_file = input("What is the AMEX spreadsheet file name? ")
    amex_file = amex_file.strip("'").replace("\\ ", " ")
    # amex_file = os.path.basename(amex_file)
    if amex_file[-1]==" ":
        amex_file = amex_file[:-1]

    sheet = input("What is the name of the sheet in the AMEX spreadsheet? ")

    userinfo_file = input("What is the user info file name? ")
    userinfo_file = userinfo_file.strip("'").replace("\\ ", " ")
    #userinfo_file = os.path.basename(userinfo_file)
    if userinfo_file[-1]==" ":
        userinfo_file = userinfo_file[:-1]

    folder1 = input("What would you like the created folder to be named? ")
    folder2 = "Amex Split Files"
    created_file = input("What would you like the created file to be named? ")
    if not created_file.endswith('.xlsx') or created_file.endswith('.xls'):
        created_file += ".xlsx"

    fol1 = create_folder(folder1)
    x = (os.path.join(folder1,folder2))
    create_folder(x)

    # Get the absolute path to the user's home directory
    home_dir = os.path.expanduser("~")

    # Construct the absolute file paths

    amex_file_path = os.path.abspath(amex_file)
    userinfo_file_path = os.path.abspath(userinfo_file)
    folder_path1 = os.path.join(home_dir, "Desktop", folder1)
    folder_path2 = os.path.join(folder_path1,folder2)
    created_file_path = os.path.join(folder_path1,created_file)


    download_splits(amex_file_path, sheet,folder_path2)
    user_list = userL(userinfo_file_path)
    for user in user_list:
        if user[2] != 'none':
            u = User(user[2],user[1].upper(),user[3],user[0])
            userclassList.append(u)
    
    
    for ppl in userclassList:
        send_file = folder_path2 + "/" + ppl.card_name + ".xlsx"
        if os.path.exists(send_file):
            pass
            #ppl.sendEmail(" "," ",send_file)
            print("="*50)
            print()
            print("Sent Email to:",ppl.email_name)
            print()
            u = User(""," ",ppl.link,ppl.email_name)
            u.sendEmail("5/5/23","12/5/23", send_file)

    if os.path.exists(folder_path1):
        combined_sheet = combine(amex_file_path,sheet,folder_path2)
        combined_sheet.to_excel(created_file_path)



    print("All files downloaded and all emails sent!")

main()