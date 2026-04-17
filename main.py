"""
QR Ticket Email Sender — Resend version (Gmail-compliant)
----------------------------------------------------------
Reads your Excel file and sends each student a beautiful
HTML ticket email with their personal QR code embedded.

SETUP:
  1. uv add resend qrcode pillow openpyxl python-dotenv
     OR: pip install resend qrcode[pil] openpyxl pillow python-dotenv
  2. Fill in your .env file
  3. python send_qr_emails.py
"""

import io
import base64
import os
import time
import openpyxl
import qrcode
import resend
from dotenv import load_dotenv

load_dotenv()

# ================================================================
# CONFIG
# ================================================================

EXCEL_FILE         = "students_database.xlsx"
TEMPLATE_FILE      = "ticket_template.html"

COL_NAME           = 1   # A
COL_CLASS          = 2   # B
COL_ID             = 3   # C
COL_EMAIL          = 4   # D
HEADER_ROW         = 1

RESEND_API_KEY     = os.getenv("RESEND_API_KEY")
SENDER_EMAIL       = os.getenv("SENDER_EMAIL")
SENDER_NAME        = os.getenv("SENDER_NAME")
EVENT_NAME         = os.getenv("EVENT_NAME")
EVENT_DATE         = os.getenv("EVENT_DATE")
EVENT_LOCATION     = os.getenv("EVENT_LOCATION")
EVENT_TIME         = os.getenv("EVENT_TIME")

SEND_DELAY_SECONDS = 1.5

# "preview" = save QR images + print to console, no emails sent
# "real"    = actually send emails
MODE = "real"

# ================================================================


def load_template():
    """Load the HTML template from file."""
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"Template file '{TEMPLATE_FILE}' not found. "
            "Make sure it's in the same folder as this script."
        )
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        return f.read()


def make_qr_bytes(name, class_, id_):
    """Generate a QR code PNG containing name|class|id."""
    text = f"{name} | {class_} | {id_}"
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=3,
    )
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1a1a2e", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_plain_text(name, class_, id_):
    """Plain text fallback — important for spam score, always include this."""
    return f"""Ahoj {name},

Tvoj vstupný lístok na {EVENT_NAME} je pripravený.

Podujatie: {EVENT_NAME}
Dátum:     {EVENT_DATE}
Čas:       {EVENT_TIME}
Miesto:    {EVENT_LOCATION}

Tvoje údaje:
  Meno:   {name}
  Trieda: {class_}
  ID:     {id_}

QR kód je priložený k tomuto emailu ako obrázok. Ukáž ho pri vstupe.

---
Tento email bol odoslaný automaticky systémom lístkov pre {EVENT_NAME}.
Odosielateľ: {SENDER_EMAIL}
"""


def load_students(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    students = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        name   = row[COL_NAME  - 1]
        class_ = row[COL_CLASS - 1]
        id_    = row[COL_ID    - 1]
        email  = row[COL_EMAIL - 1]
        if not name and not email:
            continue
        if not email:
            print(f"  ⚠  Skipping {name} — no email address")
            continue
        students.append({
            "name":   str(name).strip(),
            "class_": str(class_).strip(),
            "id":     str(id_).strip(),
            "email":  str(email).strip(),
        })
    return students


def preview_mode(students, template):
    os.makedirs("qr_preview", exist_ok=True)
    print("\n── PREVIEW MODE — no emails sent ──\n")
    for s in students:
        qr_bytes = make_qr_bytes(s["name"], s["class_"], s["id"])
        fname = f"qr_preview/qr_{s['name'].replace(' ', '_')}.png"
        with open(fname, "wb") as f:
            f.write(qr_bytes)

        # Also save a rendered HTML preview so you can open it in a browser
        qr_cid = f"qr_{s['id']}@tickets.rozlucka.me"
        html_preview = template.format(
            name=s["name"],
            class_=s["class_"],
            id_=s["id"],
            qr_cid=qr_cid,
            EVENT_NAME=EVENT_NAME,
            EVENT_DATE=EVENT_DATE,
            EVENT_TIME=EVENT_TIME,
            EVENT_LOCATION=EVENT_LOCATION,
            SENDER_EMAIL=SENDER_EMAIL,
        )
        html_fname = f"qr_preview/preview_{s['name'].replace(' ', '_')}.html"
        with open(html_fname, "w", encoding="utf-8") as f:
            f.write(html_preview)

        print(f"TO:      {s['email']}")
        print(f"SUBJECT: Tvoj lístok na {EVENT_NAME} – {s['name']}")
        print(f"QR:      {fname}")
        print(f"HTML:    {html_fname}  ← open in browser to preview")
        print("─" * 50)
    print(f"\n✓ {len(students)} tickets previewed. Files saved to qr_preview/")
    print('  Set MODE = "real" to send for real.\n')


def send_mode(students, template):
    resend.api_key = RESEND_API_KEY

    sent = 0
    failed = 0

    for i, s in enumerate(students, 1):
        qr_bytes = make_qr_bytes(s["name"], s["class_"], s["id"])
        qr_b64 = base64.b64encode(qr_bytes).decode()
        qr_cid = f"qr_{s['id']}@tickets.rozlucka.me"

        html = template.format(
            name=s["name"],
            class_=s["class_"],
            id_=s["id"],
            qr_cid=qr_cid,
            EVENT_NAME=EVENT_NAME,
            EVENT_DATE=EVENT_DATE,
            EVENT_TIME=EVENT_TIME,
            EVENT_LOCATION=EVENT_LOCATION,
            SENDER_EMAIL=SENDER_EMAIL,
        )
        plain = make_plain_text(s["name"], s["class_"], s["id"])
        subject = f"Tvoj lístok na {EVENT_NAME} – {s['name']}"

        try:
            resend.Emails.send({
                "from": f"{SENDER_NAME} <{SENDER_EMAIL}>",
                "to": [s["email"]],
                "subject": subject,
                "html": html,
                "text": plain,
                "attachments": [
                    {
                        "filename": f"listok_{s['name'].replace(' ', '_')}.png",
                        "content": qr_b64,
                        "content_type": "image/png",
                        "inline": True,
                        "content_id": qr_cid,
                    }
                ],
            })
            print(f"  [{i}/{len(students)}] ✓ Sent to {s['name']} <{s['email']}>")
            sent += 1

            if i < len(students):
                time.sleep(SEND_DELAY_SECONDS)

        except Exception as e:
            print(f"  [{i}/{len(students)}] ✗ Failed for {s['name']}: {e}")
            failed += 1

    print(f"\n── Done: {sent} sent, {failed} failed ──\n")


def check_env():
    """Make sure all required .env variables are set."""
    missing = [k for k in [
        "RESEND_API_KEY", "SENDER_EMAIL", "SENDER_NAME",
        "EVENT_NAME", "EVENT_DATE", "EVENT_LOCATION", "EVENT_TIME"
    ] if not os.getenv(k)]
    if missing:
        print(f"  ✗ Missing .env variables: {', '.join(missing)}")
        print("    Fill them in your .env file and try again.")
        exit(1)


def main():
    check_env()

    print(f"Loading template from {TEMPLATE_FILE} …")
    template = load_template()

    print(f"Loading {EXCEL_FILE} …")
    students = load_students(EXCEL_FILE)
    print(f"  Found {len(students)} students.\n")

    if not students:
        print("No students found. Check your column settings in CONFIG.")
        return

    if MODE == "preview":
        preview_mode(students, template)
    elif MODE == "real":
        confirm = input(f"Send {len(students)} real emails? Type YES to confirm: ")
        if confirm.strip().upper() == "YES":
            send_mode(students, template)
        else:
            print("Cancelled.")
    else:
        print(f"Unknown MODE '{MODE}'. Use 'preview' or 'real'.")


if __name__ == "__main__":
    main()