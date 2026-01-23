import pandas as pd
import win32com.client as win32

EXCEL_PATH = "contacts.xlsx"
SHEET_NAME = 0  # or "Sheet1"

SUBJECT = "Strategic Collaboration Opportunity with ImageProVision"

BODY_TEMPLATE = BODY_TEMPLATE = """
<div style="font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">

  <p style="margin:0 0 12pt 0;">Dear {name},</p>

  <p style="margin:0 0 12pt 0;">Greetings from ImageProVision !!</p>

  <p style="margin:0 0 12pt 0;">
    ImageProVision is a global company with offices in Princeton- USA. Our systems are trusted by over
    500 installations across 17 countries, including the USA, Europe, Asia, GCC, and Africa. We serve
    Life Sciences, Pharmaceutical, F&amp;B and other industries, providing US FDA and European
    regulatory-compliant systems that seamlessly integrates with LIMS Or Datalink. Additionally, we
    specialize in developing custom solutions tailored to specific applications.
  </p>

  <p style="margin:0 0 12pt 0;">
    ImageProVision's products are fully compliant with 21 CFR Part 11 standards and offer the following capabilities:
  </p>

  <ol style="margin:0 0 12pt 24pt; padding:0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Particle size and shape analysis with morphology &amp; Reverse Engineering powered by AI/ML
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Globule Size distribution (GSD) on Cream/Ointment.
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Nasal/ Inhalers analysis as per USP Monograph
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Particulate matter count for sub-visible analysis as per EP 2.9.19, USP 1788.2, and USP &lt;788&gt;, &lt;789&gt; (Method- 2) &amp; ISO 8871-3
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Nano-particle analysis using SEM/TEM
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Glass delamination studies as per USP &lt;790&gt; for Iron Sucrose &amp; Propofol Injections.
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Hot stage microscopy for polymorphic studies and multi-component classification based on melting point.
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Automated colony counter- MICROBE AI-400 powered by AI/ML
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      PROOF-READING system for the packaging division- Artwork Management, Proof reading for carton, leaflet, foil, label &amp; Braille, Pharmacode/Barcode inspection.
    </li>
    <li style="margin:0 0 6pt 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Seam Analyser for Soft Gelatine Capsules.
    </li>
    <li style="margin:0 0 0 0; font-family:Aptos, 'Segoe UI', Arial, sans-serif; font-size:12pt; color:#000000;">
      Cell Analysis for Biologics Products.
    </li>
  </ol>

  <p style="margin:0 0 12pt 0;">
    I would love to explore how we can collaborate to create value for both our businesses. Let’s schedule a follow-up discussion at your convenience. Looking forward to your thoughts.
  </p>

  <p style="margin:0;">Best Regards,</p>

</div>
"""


# Always CC these emails
CC_EMAILS = "anil@imageprovision.com; mudit@imageprovision.com"

# SAFETY SWITCH
SEND_EMAILS = False  # Change to True when ready

def is_valid_email(s: str) -> bool:
    return isinstance(s, str) and "@" in s and "." in s

def main():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = [c.strip() for c in df.columns]

    if "Name" not in df.columns or "Email" not in df.columns:
        raise ValueError("Excel must contain columns: Name, Email")

    outlook = win32.Dispatch("Outlook.Application")

    sent = 0
    skipped = 0

    for i, row in df.iterrows():
        name = str(row["Name"]).strip() if pd.notna(row["Name"]) else ""
        email = str(row["Email"]).strip() if pd.notna(row["Email"]) else ""

        if not name or not is_valid_email(email):
            print(f"Skipping row {i}: invalid data")
            skipped += 1
            continue

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.CC = CC_EMAILS
        mail.Subject = SUBJECT

        # Load signature quietly
        mail.Display(False)
        signature = mail.HTMLBody

        # Put your HTML above signature
        mail.HTMLBody = BODY_TEMPLATE.format(name=name) + "<br><br>" + signature

        if SEND_EMAILS:
          mail.Send()
          sent += 1
        else:
          mail.Display()  # open draft for review



        if SEND_EMAILS:
            mail.Send()
            sent += 1
        else:
            mail.Display()  # Opens draft for review

    print(f"Done. Sent={sent}, Skipped={skipped}")

if __name__ == "__main__":
    main()
