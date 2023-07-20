import pandas as pd
import os

#Get file names
file1 = input("File 1: ")
file2 = input("File 2: ")

#process files names for mac terminal
file1 = file1.strip("'").replace("\\ ", " ")
if file1[-1]==" ":
    file1 = file1[:-1]

file2 = file2.strip("'").replace("\\ ", " ")
if file2[-1]==" ":
    file2 = file2[:-1]

#read in files
df_a = pd.read_excel(file1)
df_b = pd.read_excel(file2)

#Sort excel file length, df_a should be the longer file
if len(df_b) > len(df_a):
    df_a, df_b = df_b, df_a

#Create dictionary to convert into DataFrame
missing_entries = {}
for column in df_a.columns.drop("Unnamed: 0"):
    missing = df_a[~df_a[column].isin(df_b[column])]
    if not missing.empty:
        missing_entries[column] = missing[column].values.tolist()

#Created DataFrame from dictionary
missing_data = pd.DataFrame.from_records(missing_entries)

#Get and process file name
created_file = input("What would you like the comparison file to be named? ")
if not created_file.endswith('.xlsx') or created_file.endswith('.xls'):
        created_file += ".xlsx"

#Get and process folder path
folder_path = input("Drag and drop the folder you would like the file to be located. If desktop, type Desktop. ")
folder_path = folder_path.strip("'").replace("\\ ", " ")
if folder_path[-1]==" ":
    folder_path = folder_path[:-1]

file_path = os.path.join(folder_path,created_file)
if folder_path.lower() == "desktop":
    folder_path = os.path.join(os.path.expanduser("~/Desktop"),created_file)
#Download to excel
missing_data.to_excel(file_path, index=False)




