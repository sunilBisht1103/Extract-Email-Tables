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
