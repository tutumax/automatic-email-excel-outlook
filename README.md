# 📧 Excel Merge and Email Automation with Outlook
This Python script automates the process of merging multiple Excel files from a directory, generating a single final spreadsheet, and sending it via Outlook email.

## 🚀 Features
Reads and merges all .xlsx files from a specified folder

Saves the merged data into a new Excel file

Sends the resulting file as an attachment through Microsoft Outlook

User input for customization (file name, recipient email, and sender name)

## 🛠 Requirements
Python 3.8+

Microsoft Outlook installed on your system

Required Python libraries:

```bash
    pip install pandas pywin32 openpyxl
```

## 📌 How to Use

Put the folder path:
```bash
    input_folder = r" " #enter the path of the folder containing the initial spreadsheets
    output_folder = r" " #enter the path of the folder where you want to place the final spreadsheets
```

Run the script:

```bash
    python main.py
```

## Provide the requested inputs:

Number for the final spreadsheet

Email address of the recipient

Your name (for email signature)

The script will open an Outlook email draft with the final file attached and send it automatically.

## 🧠 Author
Created by a young developer passionate about Python and automation. Always learning and building cool tools!

Name: Arthur Máximo Gonçalves
