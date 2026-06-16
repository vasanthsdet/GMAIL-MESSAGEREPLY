"""Send a sample recruiter email to revathibathina11@gmail.com for testing."""
import base64, pickle, sys, os
import truststore; truststore.inject_into_ssl()
from email.mime.text import MIMEText
from pathlib import Path
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = Path(__file__).parent / "token.pickle"

with open(TOKEN_FILE, "rb") as fh:
    creds = pickle.load(fh)

if creds.expired and creds.refresh_token:
    creds.refresh(Request())

service = build("gmail", "v1", credentials=creds)

body = """\
Hi Revathi,

Hope you're doing well! I came across your profile and wanted to reach out regarding an exciting opportunity.

We have an urgent opening for a QA Automation Engineer with a strong focus on ETA Data Migration Validation at a leading financial services client based in Dallas, TX (Remote OK).

Role: QA Automation Engineer – Data Migration Validation
Duration: 6 months (contract to hire)
Start: ASAP

Required Skills:
- ETL pipeline validation and data migration testing
- SQL (complex queries, data reconciliation)
- Python scripting for automated data validation
- AWS Glue / AWS S3 experience
- Great Expectations or similar data quality frameworks
- Experience with ETA (Extract, Transform, Audit) workflows

Could you please share your updated resume highlighting your experience with data migration validation, ETL testing, and SQL? We'd love to move quickly on this.

Looking forward to hearing from you.

Best,
Michael Torres
Senior Technical Recruiter
DataTalent Solutions
michael.torres@datatalentsolutions.com
"""

msg = MIMEText(body, "plain")
msg["To"]      = "revathibathina11@gmail.com"
msg["From"]    = "revathibathina11@gmail.com"
msg["Subject"] = "Urgent: QA Engineer – ETA Data Migration Validation | Remote | Dallas, TX"

raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
sent = service.users().messages().send(userId="me", body={"raw": raw}).execute()
print(f"Sent! Message ID: {sent['id']}")
print("Check revathibathina11@gmail.com inbox.")
