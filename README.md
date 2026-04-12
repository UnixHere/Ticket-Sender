# 🎟 Ticket Sender

Sends personalised HTML event tickets by email — each student gets a ticket with their own QR code embedded, generated from an Excel spreadsheet.

Built with [Resend](https://resend.com) for reliable email delivery.

---

## How it works

1. You prepare an Excel file (`students_database.xlsx`) with columns: Name, Class, ID, Email
2. The script generates a unique QR code for each student (encodes their name, class, and ID)
3. Each student receives a HTML email with their QR code embedded
4. At the event, scan the QR code at the entrance

---

## Project structure

```
Ticket-Sender/
├── main.py                  # Main script — generates QR codes and sends emails
├── generate_ids.py          # Helper — fills empty ID cells in your Excel file
├── ticket_template.html     # HTML email template (edit to customise design)
├── students_database.xlsx   # Your student data (not included, you create this)
├── .env                     # Your secrets and event config (not committed)
└── pyproject.toml           # Dependencies
```

---

## Setup

### 1. Install dependencies

With `uv` (recommended):
```bash
uv sync
```

Or with pip:
```bash
pip install resend qrcode[pil] openpyxl pillow python-dotenv
```

### 2. Create your `.env` file

```env
RESEND_API_KEY=re_xxxxxxxxxxxxxxxx
SENDER_EMAIL=tickets@yourdomain.com
SENDER_NAME=Your Event Name
EVENT_NAME=Rozlúčka 2025
EVENT_DATE=15. júna 2025
EVENT_TIME=18:00
EVENT_LOCATION=Aula, School Name
```

### 3. Prepare your Excel file

Create `students_database.xlsx` with these columns:

| A — Name | B — Class | C — ID | D — Email |
|----------|-----------|--------|-----------|
| Ján Novák | 4.A | 142 | jan@example.com |

If you don't have IDs yet, run:
```bash
python generate_ids.py
```
This fills all empty ID cells with unique random 3-digit numbers (100–999).

---

## Usage

### Preview first (no emails sent)

Open `main.py` and set:
```python
MODE = "preview"
```

Then run:
```bash
python main.py
```

This saves rendered HTML previews and QR images to `qr_preview/` so you can check everything looks right before sending.

### Send for real

Set:
```python
MODE = "real"
```

Run:
```bash
python main.py
```

You'll be asked to type `YES` to confirm before anything is sent.

---

## Customising the ticket design

Edit `ticket_template.html` — it's a standard HTML file with these placeholders:

| Placeholder | Value |
|-------------|-------|
| `{name}` | Student's full name |
| `{class_}` | Student's class |
| `{id_}` | Student's ID number |
| `{qr_cid}` | QR code image reference (keep as-is) |
| `{EVENT_NAME}` | Event name from `.env` |
| `{EVENT_DATE}` | Event date from `.env` |
| `{EVENT_TIME}` | Event time from `.env` |
| `{EVENT_LOCATION}` | Event location from `.env` |
| `{SENDER_EMAIL}` | Sender email from `.env` |

---

## Notes

- Resend requires a verified sending domain — free tier allows 3,000 emails/month
- Emails are sent with a 1.5 second delay between each to avoid rate limits
- Students with a missing email address are skipped with a warning
- The QR code encodes: `Name | Class | ID`
