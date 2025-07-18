## Extract Email Tables
Creating a python Script to extract emails data and convert it to CSV scheduling this script to run on a specific time using PowerShell.

## Pyhton Script

```python
import imaplib
import email
import pandas as pd
import os

# ===== CONFIGURATION =====
IMAP_SERVER = 'imap.yourmail.com'       # e.g., 'imap.gmail.com'
EMAIL_ACCOUNT = 'you@example.com'
PASSWORD = 'your_password'
MAILBOX = 'INBOX'
OUTPUT_DIR = 'email_tables'

# Create output directory
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ===== CONNECT TO MAILBOX =====
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL_ACCOUNT, PASSWORD)
mail.select(MAILBOX)

# Search for all emails (you can change 'ALL' to 'UNSEEN' or 'FROM "xyz@example.com"')
status, data = mail.search(None, 'ALL')
email_ids = data[0].split()

for num in email_ids:
    _, msg_data = mail.fetch(num, '(RFC822)')
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Get HTML part of email
    html_content = None
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                html_content = part.get_payload(decode=True).decode()
                break
    else:
        if msg.get_content_type() == "text/html":
            html_content = msg.get_payload(decode=True).decode()

    if not html_content:
        continue

    # ===== PARSE HTML TABLES TO CSV =====
    try:
        tables = pd.read_html(html_content)  # Extracts all <table> tags
        subject = msg.get('Subject', 'email').replace('/', '_').strip()

        for idx, table in enumerate(tables, start=1):
            file_name = f"{subject}_table_{idx}.csv"
            file_path = os.path.join(OUTPUT_DIR, file_name)
            table.to_csv(file_path, index=False)
            print(f"✅ Saved: {file_path}")

    except ValueError:
        print("❌ No tables found in this email.")

mail.logout()
```

## How It Works

```text
1. Connects to IMAP mailbox (works for Gmail, Outlook, etc.).
2. Downloads email HTML content.
3. Extracts <table> elements using pandas.read_html().
4.  Saves each table as CSV in email_tables/ folder.
```

## Requirements
Install necessary libraries:

```bash
pip install pandas lxml

```
## Steps

## 1. Save Your Python Script
Save your Python script as ExtractEmailTables.py in a folder (e.g., C:\Scripts\ExtractEmailTables.py).

## 2. Save Your Python Script
Save your Power shell script as Powershell.ps1

## Powershell Script

```powershell
# Define task details
$TaskName = "RunPythonEmailExtractor"
$TaskDescription = "Extract email tables and convert to CSV"
$PythonExe = "C:\Python311\python.exe"
$ScriptPath = "C:\Scripts\ExtractEmailTables.py"
$TriggerTime = (Get-Date).AddMinutes(1)  # Schedule 1 minute from now

# Create action to run Python script
$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument $ScriptPath

# Create daily trigger at a specific time (e.g., 10:00 AM)
$Trigger = New-ScheduledTaskTrigger -Daily -At 10:00AM

# Register the scheduled task
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Description $TaskDescription -User "$env:USERNAME" -RunLevel Highest

```

## Key Points:
$TaskName → Name of the task.

$PythonExe → Path to Python executable.

$ScriptPath → Path to your Python script.

Trigger → You can use -Daily, -AtStartup, or -Once.

-RunLevel Highest → Runs with admin privileges if needed.

## To Check the Scheduled Task:
```powershell
Get-ScheduledTask -TaskName "RunPythonEmailExtractor"
```
## To Remove the Scheduled Task:
```powershell
Unregister-ScheduledTask -TaskName "RunPythonEmailExtractor" -Confirm:$false
```
## Alternative: Immediate Run (One-Time Schedule)
For running after 5 minutes from now:

```powershell
$Trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(5))
```
