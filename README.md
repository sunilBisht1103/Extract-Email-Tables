# Extract-Email-Tables
Creating a python Script to extract emails data and convert it to CSV scheduling this script to run on a specific time using PowerShell.


 Python script that connects to your email inbox, extracts HTML table data from emails, and converts it into CSV format:

✅ How It Works
✔ Connects to IMAP mailbox (works for Gmail, Outlook, etc.).
✔ Downloads email HTML content.
✔ Extracts <table> elements using pandas.read_html().
✔ Saves each table as CSV in email_tables/ folder.

 Requirements
Install necessary libraries:

bash
Copy
Edit
pip install pandas lxml

✅ Optional Features:
Filter emails by date, subject, or sender.

Mark emails as read after processing.

Combine all tables into one CSV file.

Handle attachments if needed.



✅ 1. Save Your Python Script
Save your Python script as ExtractEmailTables.py in a folder (e.g., C:\Scripts\ExtractEmailTables.py).



✅ Key Points:
$TaskName → Name of the task.

$PythonExe → Path to Python executable.

$ScriptPath → Path to your Python script.

Trigger → You can use -Daily, -AtStartup, or -Once.

-RunLevel Highest → Runs with admin privileges if needed.

✅ To Check the Scheduled Task:
powershell
Copy
Edit
Get-ScheduledTask -TaskName "RunPythonEmailExtractor"
✅ To Remove the Scheduled Task:
powershell
Copy
Edit
Unregister-ScheduledTask -TaskName "RunPythonEmailExtractor" -Confirm:$false
✅ Alternative: Immediate Run (One-Time Schedule)
For running after 5 minutes from now:

powershell
Copy
Edit
$Trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(5))
