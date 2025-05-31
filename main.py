import win32com.client as win32
import pandas as pd
import os
import datetime


input_folder = r" " #enter the path of the folder containing the initial spreadsheets
output_folder = r" " #enter the path of the folder where you want to place the final spreadsheets


os.makedirs(output_folder, exist_ok=True)


files = [
    os.path.join(input_folder, file)
    for file in os.listdir(input_folder)
    if file.endswith('.xlsx')
]

if not files:
    print("No .xlsx files found in the folder.")
    exit()

print(f"Found files: {files}")


final_table = pd.DataFrame()
for file in files:
    print(f"Reading: {file}")
    table = pd.read_excel(file)
    final_table = pd.concat([final_table, table], ignore_index=True)


sheet_number = input('Enter the number for the final spreadsheet: ')
output_filename = f"final_spreadsheet_{sheet_number}.xlsx"
output_path = os.path.join(output_folder, output_filename)


final_table.to_excel(output_path, index=False)
print(f'Final spreadsheet saved as {output_path}')


try:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.To = input('Enter the recipient email address: ')
    email.Subject = 'Generated Final Spreadsheet'
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%H:%M:%S")
    sender_name = input('Enter your name: ')
    sender_name.capitalize()

    email.Body = f"""
        Hello,

        Please find attached the final spreadsheet generated at {formatted_time}.

        Best regards,
        {sender_name}
    """
    email.Attachments.Add(output_path)
    print("Opening email for review...")
    email.Send()
except Exception as e:
    print(f"An error occurred while creating or sending the email: {e}")

