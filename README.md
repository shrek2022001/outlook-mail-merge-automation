# Outlook Mail Merge Automation

A Python script that automates personalized bulk outreach emails via Outlook, 
using a contact list from Excel.

## What it does
- Reads a contact list (Name, Email) from an Excel file
- Generates a personalized HTML email per contact from a template
- Sends via Outlook (win32com) or opens each as a draft for manual review
- Validates email format and skips malformed rows
- Preserves the sender's existing Outlook signature
- CCs a fixed list of addresses on every send

## Setup
1. `pip install pandas pywin32`
2. Requires Outlook installed and configured on Windows
3. Set up `contacts.xlsx` with `Name` and `Email` columns
4. Edit `BODY_TEMPLATE`, `SUBJECT`, and `CC_EMAILS` in `config.py`

## Safety
`SEND_EMAILS` defaults to `False` — every email opens as a draft for review 
instead of sending automatically. Set to `True` only once templates and 
recipient list are verified.

## Example use case
Built for sending personalized partnership/outreach emails at scale while 
keeping each one reviewable before sending.
