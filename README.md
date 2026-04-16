# Email Outreach

A lightweight Windows desktop tool for sending personalized cold outreach emails via Outlook.

## How it works

- Sends emails through the Outlook desktop app using win32com (no SMTP credentials needed)
- Personalizes each email with the recipient's name and company
- Tracks sends in a CSV log
- Auto-increments a URL extension on the portfolio link for tracking visits per outreach

## Setup

**Install dependencies:**
```bash
pip install customtkinter pywin32
```

**Requirements:**
- Windows
- Outlook desktop app installed and logged in with your account

## Usage

Double-click `run.vbs` to launch the app (no terminal window).

Fill in:
- **Name** — recipient's first name
- **Email** — recipient's email address
- **Company** — company name (used in subject line and email body)

Choose **Send Now** or **Schedule**. If scheduling, pick a date (`MM/DD/YYYY`) and time (`HH:MM AM/PM`). Scheduled emails are queued in Outlook's Outbox and sent automatically — the app can be closed after scheduling.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Main application |
| `template.html` | Email body — edit this to change the message |
| `data.json` | Stores the current URL extension counter |
| `log.csv` | Send history (auto-created on first send) |
| `run.vbs` | Launcher — opens app with no terminal window |

## Customizing the template

Edit `template.html` directly. Available placeholders:

| Placeholder | Value |
|-------------|-------|
| `{name}` | Recipient's name |
| `{company}` | Company name |
| `{url_extension}` | Auto-incremented portfolio URL suffix |
