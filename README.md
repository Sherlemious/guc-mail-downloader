# guc-mail-downloader

Save a copy of **every email** in your university (GUC / Outlook Web App)
mailbox to your own computer before the account is deleted. It covers every
folder (Inbox, Sent, Deleted, Drafts, Junk, your own folders, and the Online
Archive), keeps all attachments, and can be stopped and resumed at any time.
Afterwards you get a **searchable offline copy** that opens in any web browser.

<img width="1917" height="911" alt="Screenshot of guc-mail-downloader" src="https://github.com/user-attachments/assets/2ecbf584-d7a0-42d0-84cc-fcaef2d5fa3b" />

## How to use it (no technical knowledge needed)

1. **Download this tool**: on the GitHub page click **Code → Download ZIP**,
   then unzip it.
2. **Install Python** if you don't have it: <https://www.python.org/downloads/>.
   On Windows, tick **"Add python.exe to PATH"** in the installer.
3. **Double-click the launcher**:
   - Windows: `Backup my email (Windows).bat`
   - Mac: `Backup my email (Mac-Linux).command`. If macOS blocks it,
     right-click it, choose **Open**, then **Open** again.
   - Linux: `bash "Backup my email (Mac-Linux).command"`
4. Type your university email and password. The password doesn't show while
   you type; that's normal.
5. It shows how many emails each folder has. Press Enter and wait.
6. At the end, say **yes** to the browsable copy. It opens a page where you
   can search and read all your emails, even offline.

The first run takes a minute to set itself up. If the window closes, your
internet drops, or you stop it, just run it again: it skips everything
already downloaded and carries on.

Several people can use the same copy of the tool; each email address gets its
own folder under `guc-mail-backup/`.

### What you get

```
guc-mail-backup/you@student.guc.edu.eg/
  mail/<folder>/<date>_<subject>_<id>.eml   every email, original and complete
  viewer/index.html                         searchable archive for any browser
  index.csv                                 spreadsheet of all emails
  attachments/...                           only if you run the "extract" command
  mbox/...                                  only if you ask for it (Thunderbird/Gmail import)
```

`.eml` files open directly in Outlook, Thunderbird, Apple Mail, and most other
mail clients.

### Your password

Your password is only sent to your university's mail server, the same one
your browser talks to. It is never stored on disk.

### If something goes wrong

- **"The server rejected that email/password"**: check the password. After 3
  tries it asks whether your login uses a different username, such as
  `GUC\first.last`.
- **"Could not reach the mail server"**: open your webmail in a browser, copy
  the address from the address bar, and paste it when asked. This also lets
  you use the tool with other universities that run Outlook Web App.
- **Some emails failed**: run it again. Failed emails are retried each time.

## Command line

```bash
pip install -r requirements.txt
python guc_mail_backup.py                 # guided mode (same as the launchers)
python guc_mail_backup.py COMMAND --help  # every option of a command
```

| Command | What it does | Needs login |
|---|---|---|
| `probe` | Check the login works and count items per folder | yes |
| `backup` | Download everything (resumable) | yes |
| `html` | Build the searchable offline viewer (`viewer/index.html`) | no |
| `extract` | Save all attachments as normal files under `attachments/` | no |
| `mbox` | Pack each folder into an `.mbox` file | no |
| `stats` | Counts and sizes per folder and year, plus top senders | no |
| `verify` | Check every saved file is still there and intact (`--fix` re-queues broken ones) | no |

### Useful `backup` options

```bash
# everything, then build the viewer
python guc_mail_backup.py backup --email you@student.guc.edu.eg --html

# preview only: how much would be downloaded
python guc_mail_backup.py backup --email you@student.guc.edu.eg --dry-run

# only email folders, skip junk, only 2024 onwards, 8 parallel downloads
python guc_mail_backup.py backup --email you@student.guc.edu.eg \
    --mail-only --exclude "Junk E-Mail" --since 2024-01-01 --workers 8

# one folder (and its subfolders)
python guc_mail_backup.py backup --email you@student.guc.edu.eg --folder Inbox
```

| Option | Meaning |
|---|---|
| `--folder PATH` / `--exclude PATH` | Include or skip folders, including their subfolders. Repeatable; `*` wildcards work. |
| `--mail-only` | Skip Calendar, Contacts, Tasks and Notes |
| `--since` / `--until YYYY-MM-DD` | Date range, inclusive |
| `--dry-run` | Count only, download nothing |
| `--workers N` / `--batch N` | Parallel requests and messages per request (defaults 4 and 20) |
| `--html` / `--mbox` | Build the viewer or the mbox files when done |
| `--out DIR` | Where to save (default `guc-mail-backup/<email>`) |
| `--server HOST` | Your webmail host (default `mail.guc.edu.eg`, or `$GUC_SERVER`) |
| `--user NAME` | Login name if it isn't your email, e.g. `GUC\first.last` |
| `--method ews/imap`, `--auth ntlm/basic`, `--autodiscover`, `--insecure` | Connection tweaks, rarely needed |
| `--password-stdin` | Read the password from stdin (otherwise `$GUC_PASSWORD`, otherwise a prompt) |
| `--json`, `-q` | Machine-readable result on stdout, and no progress output |

Exit codes: `0` ok, `1` error, `2` some messages failed, `3` login rejected,
`4` server unreachable, `130` interrupted.

## How it works

It logs in to Exchange Web Services (`https://<server>/EWS/Exchange.asmx`),
the API Outlook for Mac uses, through
[exchangelib](https://github.com/ecederstrand/exchangelib). It lists every
item in every folder directly, not page by page like the webmail, then
downloads each item's original MIME content in parallel batches. If that
fails, it tries NTLM and basic login, then autodiscover, then IMAP. A
manifest records what has been saved, which is what makes resuming work.

## Development

```bash
python -m unittest discover -s tests -v   # offline tests using fake mail servers
```

See [AGENTS.md](AGENTS.md) for the code map and conventions.
