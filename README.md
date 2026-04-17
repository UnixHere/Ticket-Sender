# 🎟 Ticket Sender & Verifier

A two-part system for sending personalised QR ticket emails and verifying them at the entrance.

Built with [Resend](https://resend.com) for reliable email delivery.

---

## How it works

1. Prepare an Excel file (`students_database.xlsx`) with columns: Name, Class, ID, Email
2. Run **Ticket Sender** — generates a unique QR code for each student and emails them their ticket
3. At the event, run **Ticket Verifier** — scan QR codes at the entrance to check students in

---

## Project structure

```
Ticket-Sender/
├── main.py                  # Ticket Sender — generates QR codes and sends emails
├── generate_ids.py          # Helper — fills empty ID cells in your Excel file
├── ticket_template.html     # HTML email template (edit to customise design)
├── webapp.py                # Ticket Verifier — Flask backend for the scanner web app
├── index.html               # Ticket Verifier — mobile-friendly scanner UI
├── students_database.xlsx   # Your student data (not included, you create this)
├── .env                     # Your secrets and event config (not committed)
└── pyproject.toml           # Dependencies
```

---

# 📤 Ticket Sender

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

You'll be asked to type `YES` to confirm before anything is sent. Students who have already been sent a ticket (column E = 1) are automatically skipped.

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
- Students with column E already set to `1` (Sent) are skipped to prevent duplicate emails
- The QR code encodes: `Name | Class | ID`

---

# 📷 Ticket Verifier

A locally-hosted mobile-friendly web app that scans QR codes at the entrance and checks them against `students_database.xlsx`.

## Features

- 📷 **QR Scanner** — uses the phone camera to scan ticket QR codes
- ✅ **Instant verification** — checks the QR against your Excel database
- ⚠️ **Duplicate detection** — warns if a ticket has already been scanned
- ✋ **Check-in button** — marks a student as arrived directly in the Excel file (column F)
- 🔍 **Manual lookup** — search by name, ID, class or email
- 📋 **Attendees list** — full list with live arrived/pending stats
- ↩ **Undo** — un-check someone if needed
- 📱 **QR code for the app URL** — prints a scannable QR in the terminal on startup so you don't have to type the address manually

## Setup

### 1. Install dependencies

```bash
pip install flask flask-cors openpyxl "qrcode[pil]"
```

Or with uv:
```bash
uv add flask flask-cors openpyxl "qrcode[pil]"
```

### 2. Copy your database

Put your `students_database.xlsx` in the same folder as `webapp.py`. Column F (`Arrived`) is created automatically on first run.

### 3. Run

```bash
python webapp.py
```

The terminal will print your local IP, a scannable QR code, and a link to download the QR as an image:

```
  → https://192.168.1.42:5000
  (Or visit https://192.168.1.42:5000/qr to download the QR as an image)
```

### 4. Open on phone

Connect your phone to the **same WiFi** and either:
- Scan the QR code printed in the terminal, or
- Open `https://<your-ip>:5000` manually in the browser

> On iPhone: Safari works best for camera access.  
> On Android: Chrome works best.

---

## Excel columns

| Col | Header  | Used by                                      |
|-----|---------|----------------------------------------------|
| A   | Name    | Both                                         |
| B   | Class   | Both                                         |
| C   | ID      | Both                                         |
| D   | Email   | Ticket Sender                                |
| E   | Sent    | Ticket Sender (skips already-sent students)  |
| F   | Arrived | Ticket Verifier (added automatically)        |
