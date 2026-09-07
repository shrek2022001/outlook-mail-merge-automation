"""
Configuration for the Outlook Mail Merge tool.

Just edit the values below to customize the outreach campaign without
changing the main script.
"""

# Path to the Excel file containing contacts. Must have "Name" and "Email" columns.
EXCEL_PATH = "contacts.xlsx"
SHEET_NAME = 0  

# Email subject line
SUBJECT = "Strategic Collaboration Opportunity"

# Addresses to CC on every email sent (semicolon-separated, Outlook format)
CC_EMAILS = "cc1@example.com; cc2@example.com"

# Safety switch: keep this False until templates and recipient list are verified.
# False -> every email opens as a draft for manual review instead of sending.
# True  -> emails are sent automatically.
SEND_EMAILS = False

BODY_TEMPLATE = """
<div style="font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">

  <p style="margin:0 0 12pt 0;">Dear {name},</p>

  <p style="margin:0 0 12pt 0;">Greetings from [Your Company Name] !!</p>

  <p style="margin:0 0 12pt 0;">
    [Your Company Name] is a [brief company description]. Our products/services
    are trusted by [credibility statement, e.g. number of clients or installations].
  </p>

  <p style="margin:0 0 12pt 0;">
    Here's what we offer:
  </p>

  <ol style="margin:0 0 12pt 24pt; padding:0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
    <li style="margin:0 0 6pt 0;">[Capability or product line 1]</li>
    <li style="margin:0 0 6pt 0;">[Capability or product line 2]</li>
    <li style="margin:0 0 6pt 0;">[Capability or product line 3]</li>
  </ol>

  <p style="margin:0 0 12pt 0;">
    I would love to explore how we can collaborate to create value for both
    our businesses. Let's schedule a follow-up discussion at your convenience.
    Looking forward to your thoughts.
  </p>

  <p style="margin:0;">Best Regards,</p>

</div>
"""
