# guc-mail-downloader

Save a copy of **every email** in your university (GUC / Outlook Web App)
mailbox to your own computer before the account is deleted. It covers every
folder (Inbox, Sent, Deleted, Drafts, Junk, your own folders, and the Online
Archive), keeps all attachments, and can be stopped and resumed at any time.

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

The first run takes a minute to set itself up. If the window closes, your
internet drops, or you stop it, just run it again: it skips everything
already downloaded and carries on.

Your emails end up in `guc-mail-backup/<your email>/` next to the launcher:

| What | Where |
|---|---|
| Every email as an `.eml` file (double-click to open it in Outlook, Mail, or Thunderbird) | `mail/<folder>/` |
| A spreadsheet listing every email (date, folder, sender, subject) | `index.csv` |
| Optional `.mbox` files for importing into Thunderbird or Gmail | `mbox/` |

Several people can use the same copy of the tool; each email address gets its
own folder.

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

## Advanced: command line

```bash
pip install -r requirements.txt
python guc_mail_backup.py                      # guided mode (same as the launchers)
python guc_mail_backup.py probe  --email you@student.guc.edu.eg
python guc_mail_backup.py backup --email you@student.guc.edu.eg [--folder Inbox] [--method imap]
python guc_mail_backup.py mbox   --out guc-mail-backup/you@student.guc.edu.eg
```

Other options: `--server`, `--user`, `--auth ntlm|basic`, `--autodiscover`,
`--insecure`, and `--no-archive`. Set `GUC_PASSWORD` and `GUC_SERVER` as
environment variables to skip the prompts.

## How it works

It logs in to Exchange Web Services (`https://<server>/EWS/Exchange.asmx`),
the API Outlook for Mac uses, through
[exchangelib](https://github.com/ecederstrand/exchangelib). It lists every
item in every folder directly, not page by page like the webmail, then
downloads each item's original MIME content in small batches. If EWS is
unavailable, it tries NTLM and basic login, then autodiscover, then IMAP.
