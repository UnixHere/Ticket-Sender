# 🎟 Ticket Sender & Verifier

A two-part system for sending personalised QR ticket emails and checking students in at the entrance.

- **Ticket Sender** (`main.py`) — reads your Excel spreadsheet, generates a unique QR code for each student, and sends them a personalised HTML ticket by email
- **Ticket Verifier** (`webapp.py` + `index.html`) — a locally-hosted web app that scans QR codes at the entrance and marks students as arrived in the same spreadsheet

---

## Project structure

```
Ticket-Sender/
├── main.py                  # Ticket Sender — generates QR codes and sends emails
├── generate_ids.py          # Helper — fills empty ID cells in your Excel file
├── ticket_template.html     # HTML email template (edit to customise design)
├── webapp.py                # Ticket Verifier — Flask backend
├── index.html               # Ticket Verifier — mobile scanner UI
├── students_database.xlsx   # Your student data (not included, create this yourself)
├── .env                     # Secrets and event config (never commit this)
└── pyproject.toml           # Python dependencies
```

---

## Excel column layout

Both tools share the same spreadsheet. Columns must be in this exact order:

| Col | Header   | Description                                              |
|-----|----------|----------------------------------------------------------|
| A   | Name     | Student's full name                                      |
| B   | Class    | Student's class (e.g. `4.A`)                             |
| C   | ID       | Unique 3-digit number — use `generate_ids.py` if missing |
| D   | Email    | Student's email address                                  |
| E   | Sent     | Set to `1` by Ticket Sender after emailing               |
| F   | Arrived  | Set to `1` by Ticket Verifier when scanned at entrance   |

Columns E and F are created automatically on first run if missing.

---

## Setup

### 1. Install dependencies

With `uv` (recommended):
```bash
uv sync
```

Or with pip:
```bash
pip install resend "qrcode[pil]" openpyxl pillow python-dotenv flask flask-cors
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

> Resend requires a verified sending domain. The free tier allows 3,000 emails/month.

---

# 📤 Ticket Sender

## 1. Prepare your spreadsheet

Create `students_database.xlsx` with the column layout above. If students don't have IDs yet, run:

```bash
python generate_ids.py
```

This fills every empty cell in column C with a unique random 3-digit number (100–999). Existing IDs are never overwritten or duplicated. Supports up to 900 students.

You can also pass a custom filename:
```bash
python generate_ids.py other_file.xlsx
```

---

## 2. Preview tickets before sending

Open `main.py` and set:
```python
MODE = "preview"
```

Then run:
```bash
python main.py
```

Preview mode does **not** send any emails. Instead it saves each student's rendered ticket as an HTML file and their QR code as a PNG into a `qr_preview/` folder. Open the HTML files in a browser to check the design before sending.

---

## 3. Send tickets

Set:
```python
MODE = "real"
```

Run:
```bash
python main.py
```

- Only students with column E **not** set to `1` are included — already-sent students are automatically skipped
- You'll be asked to type `YES` to confirm before anything is sent
- After each successful send, column E is updated to `1` in the spreadsheet
- A 1.5 second delay is added between emails to avoid rate limits
- Students with no email address or no name are skipped with a warning

---

## 4. Resend to one student

To resend a ticket to a specific student regardless of their sent status:

```bash
python main.py resend 123
python main.py resend "Jana Nováková"
python main.py resend jana@example.com
```

- Searches by ID, full name, or email (case-insensitive)
- Shows the matched student's details before sending
- Asks for `YES` confirmation before sending
- Does **not** update column E — resend is intentional and manual
- If `MODE` is set to `preview`, it shows the match but does not send

---

## Customising the ticket design

Edit `ticket_template.html`. It's a standard HTML file with these placeholders:

| Placeholder        | Value                      |
|--------------------|----------------------------|
| `{name}`           | Student's full name        |
| `{class_}`         | Student's class            |
| `{id_}`            | Student's ID number        |
| `{qr_cid}`         | QR code image (keep as-is) |
| `{EVENT_NAME}`     | From `.env`                |
| `{EVENT_DATE}`     | From `.env`                |
| `{EVENT_TIME}`     | From `.env`                |
| `{EVENT_LOCATION}` | From `.env`                |
| `{SENDER_EMAIL}`   | From `.env`                |

---

# 📷 Ticket Verifier

A mobile-friendly web app you run locally on a laptop at the entrance. Any phone on the same WiFi can open it and use it as a QR scanner.

## Features

- 📷 **QR Scanner** — uses the phone camera to scan ticket QR codes
- ✅ **Instant verification** — checks the scanned QR against the Excel database
- ⚠️ **Duplicate detection** — warns if a ticket has already been scanned
- ✋ **Check-in button** — marks student as arrived in column F of the spreadsheet
- ↩ **Undo** — un-check a student if they were checked in by mistake
- 🔍 **Manual lookup** — search by name, ID, class, or email
- 📋 **Attendees list** — full list with arrived/pending filters and live stats
- 📱 **QR code for the app URL** — on startup, prints a scannable QR in the terminal so phones don't have to type the address manually

## Setup

### 1. Install dependencies

```bash
pip install flask flask-cors openpyxl "qrcode[pil]"
```

Or with uv:
```bash
uv add flask flask-cors openpyxl "qrcode[pil]"
```

### 2. Place your spreadsheet

Put `students_database.xlsx` in the same folder as `webapp.py`.

### 3. Run

```bash
python webapp.py
```

On startup the terminal will show your local IP, print a scannable QR code, and tell you where to download it as an image:

```
=======================================================
  🎟  Ticket Verifier
=======================================================
  ✓  Database: students_database.xlsx

  Open in browser (same WiFi):
  → https://192.168.1.42:5000

  Scan this QR code to open on any device:

  █▀▀▀▀▀▀▀████ ...

  (Or visit https://192.168.1.42:5000/qr to download the QR as an image)
=======================================================
```

### 4. Open on a phone

Connect the phone to the **same WiFi** and either scan the QR printed in the terminal or open the address manually in the browser.

> iPhone: use Safari for camera access
> Android: use Chrome for camera access

---

## API endpoints

| Method | Endpoint            | Description                                                        |
|--------|---------------------|--------------------------------------------------------------------|
| `POST` | `/api/verify`       | Verify a QR string. Body: `{ "qr": "Name \| Class \| ID" }`       |
| `POST` | `/api/mark_arrived` | Check in or undo. Body: `{ "row": 5, "arrived": true }`           |
| `GET`  | `/api/stats`        | Returns `{ total, arrived, pending }`                             |
| `GET`  | `/api/attendees`    | Returns the full student list                                      |
| `GET`  | `/qr`               | Returns a PNG QR code image pointing to the app URL               |

---

## QR code format

Both tools use the same format:

```
Name | Class | ID
```

For example: `Jana Nováková | 4.A | 142`

The verifier matches on **ID** (exact) and does a soft name check — a mismatch shows a warning but does not reject the ticket.
