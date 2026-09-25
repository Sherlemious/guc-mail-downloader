# guc-mail-downloader

Back up an entire GUC (Exchange / Outlook Web App) mailbox to your own disk
before the university deletes the account.

Each message is saved as the server's original `.eml` file, byte for byte, so
headers, HTML bodies, inline images and attachments all survive. Every folder
is backed up: Inbox, Sent Items, Deleted Items, Drafts, Junk, your own folders
such as *Memoirs*, and the Online Archive if you have one.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 1. check that login works and see how many items each folder has
python guc_mail_backup.py probe --email you@student.guc.edu.eg

# 2. download everything (you can stop and re-run it; it picks up where it left off)
python guc_mail_backup.py backup --email you@student.guc.edu.eg

# 3. optional: also pack each folder into an .mbox file (for Thunderbird etc.)
python guc_mail_backup.py mbox
```

The script asks for your password. To skip the prompt, set the `GUC_PASSWORD`
environment variable.

## What you get

```
guc-mail-backup/
  mail/Inbox/2026-09-21_152500_Registration Procedures for Grad..._1a2b3c4d.eml
  mail/Sent Items/...
  mail/Memoirs/...
  index.csv          # one row per message: date, folder, from, subject, file
  manifest.jsonl     # what has been downloaded (used to resume)
  failures.jsonl     # anything that failed (re-run to retry it)
  mbox/...           # only after running the `mbox` command
```

`.eml` files open directly in Outlook, Thunderbird, Apple Mail, and most other
mail clients. To import everything into Gmail or another account, you can add
the `mbox/` files to Thunderbird with the ImportExportTools NG add-on, or drag
the `.eml` files into a folder of an IMAP account.

## If the login fails

| Symptom | Try |
|---|---|
| Wrong server / DNS error | Use the host from your OWA address bar: `--server <host>` (default `mail.guc.edu.eg`) |
| 401 Unauthorized | `--user "GUC\first.last"` (domain\username), or `--auth ntlm` / `--auth basic` |
| Certificate error | `--insecure` |
| EWS blocked for students | `backup --method imap` (only works if the university has IMAP enabled; `probe` checks it) |
| Mailbox huge / throttled | The script waits and retries on its own; you can also back up one folder at a time with `--folder Inbox` |

## How it works

It connects to Exchange Web Services (`https://<server>/EWS/Exchange.asmx`),
the same API Outlook for Mac uses, through
[exchangelib](https://github.com/ecederstrand/exchangelib). It walks every
folder, lists the item IDs, and then fetches each item's raw MIME content in
small batches. If EWS is unavailable, IMAP is the fallback.
