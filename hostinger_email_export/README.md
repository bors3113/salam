# Hostinger Email Export to Excel

Exports the **received** (Inbox) and **sent** emails of a Hostinger mailbox into an organised Excel workbook.

> SMTP only *sends* mail and can't read a mailbox, so the script logs in to the same account over
> **IMAP** (`imap.hostinger.com`, port 993, SSL) with your email address and password.
> It opens folders read-only, so no message gets marked as read.

## Setup

```bash
pip install openpyxl
```

## Usage

```bash
# prompts for the password
python export_emails.py --email you@yourdomain.com

# or pass the password through an environment variable
HOSTINGER_EMAIL_PASSWORD='secret' python export_emails.py --email you@yourdomain.com

# only messages since a date, custom output file
python export_emails.py --email you@yourdomain.com --since 2025-01-01 --output mail.xlsx
```

| Option | Description |
|---|---|
| `--since YYYY-MM-DD` | Only export messages on or after this date |
| `--limit N` | Only the newest N messages per folder |
| `--sent-folder NAME` | Sent folder name, if auto-detection fails (Hostinger usually uses `INBOX.Sent`) |
| `--headers-only` | Faster: skips bodies, so there is no preview and no attachment names |
| `--host`, `--port` | Override the IMAP server (defaults: `imap.hostinger.com:993`) |
| `--output FILE` | Output path (default `emails_<account>_<date>.xlsx`) |

## Workbook layout

| Sheet | Contents |
|---|---|
| **Summary** | Message count and size per folder, date range, top 10 senders and recipients, monthly received/sent counts |
| **Received** | Every Inbox message, newest first |
| **Sent** | Every Sent message, newest first |
| **All Emails** | Both folders combined, newest first |

Columns: Folder, Date, From Name, From Email, To, Cc, Subject, Attachments (count), Attachment Names,
Size (KB), Read, Preview (first 300 characters of the body) and Message-ID.
The header row is frozen and has filters, and the rows are banded.
