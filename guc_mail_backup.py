#!/usr/bin/env python3
"""Back up an entire Exchange / Outlook Web App mailbox to local .eml files.

Built for the GUC mailbox (on-prem Exchange 2016, the one behind the
"Outlook Web App" page), but works with any Exchange server that exposes
EWS or IMAP.

Every message is saved byte-for-byte as the server's original MIME, so
attachments, inline images, headers and HTML bodies are all preserved.
Each folder becomes a directory; an index.csv summarises everything, and
the run can be interrupted and resumed at any time.

Usage:
    python guc_mail_backup.py                 # guided mode, just answer the questions
    python guc_mail_backup.py probe  --email you@student.guc.edu.eg
    python guc_mail_backup.py backup --email you@student.guc.edu.eg
    python guc_mail_backup.py mbox   --out guc-mail-backup/you@student.guc.edu.eg
"""

import argparse
import csv
import getpass
import hashlib
import imaplib
import json
import mailbox
import os
import re
import sys
from email import message_from_bytes, policy
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from pathlib import Path

DEFAULT_SERVER = "mail.guc.edu.eg"
DEFAULT_OUT = "guc-mail-backup"
MANIFEST = "manifest.jsonl"
FAILURES = "failures.jsonl"


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def safe_name(text, limit=80):
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", text or "").strip(" .")
    text = re.sub(r"\s+", " ", text)
    return (text[:limit].rstrip(" .") or "no-subject")


def decode_mime_words(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def headers_from_mime(raw):
    """Pull subject/from/date out of raw MIME for naming and the index."""
    msg = message_from_bytes(raw, policy=policy.compat32)
    subject = decode_mime_words(msg.get("Subject"))
    sender = decode_mime_words(msg.get("From"))
    date = None
    if msg.get("Date"):
        try:
            date = parsedate_to_datetime(msg["Date"])
        except Exception:
            date = None
    return subject, sender, date


class Store:
    """Writes .eml files and keeps a resumable manifest of what is done."""

    def __init__(self, out_dir):
        self.out = Path(out_dir)
        self.out.mkdir(parents=True, exist_ok=True)
        self.manifest_path = self.out / MANIFEST
        self.failures_path = self.out / FAILURES
        self.done = set()
        if self.manifest_path.exists():
            with open(self.manifest_path, encoding="utf-8") as fh:
                for line in fh:
                    try:
                        self.done.add(json.loads(line)["key"])
                    except Exception:
                        pass
        self._manifest = open(self.manifest_path, "a", encoding="utf-8")
        self._failures = open(self.failures_path, "a", encoding="utf-8")

    def has(self, key):
        return key in self.done

    def save(self, key, folder, raw, fallback_date=None):
        subject, sender, date = headers_from_mime(raw)
        date = date or fallback_date
        stamp = date.strftime("%Y-%m-%d_%H%M%S") if date else "0000-00-00_000000"
        digest = hashlib.sha1(key.encode()).hexdigest()[:8]
        folder_dir = self.out / "mail" / Path(*[safe_name(p, 60) for p in folder.split("/") if p])
        folder_dir.mkdir(parents=True, exist_ok=True)
        path = folder_dir / f"{stamp}_{safe_name(subject)}_{digest}.eml"
        tmp = path.with_suffix(".part")
        tmp.write_bytes(raw)
        os.replace(tmp, path)
        rec = {
            "key": key,
            "folder": folder,
            "path": str(path.relative_to(self.out)),
            "subject": subject,
            "from": sender,
            "date": date.isoformat() if date else "",
            "size": len(raw),
        }
        self._manifest.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self._manifest.flush()
        self.done.add(key)
        return path

    def fail(self, key, folder, error):
        rec = {"key": key, "folder": folder, "error": str(error)[:500]}
        self._failures.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self._failures.flush()

    def write_index(self):
        rows = {}
        with open(self.manifest_path, encoding="utf-8") as fh:
            for line in fh:
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                rows[rec["key"]] = rec
        index = self.out / "index.csv"
        with open(index, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["date", "folder", "from", "subject", "size_bytes", "file"])
            for r in sorted(rows.values(), key=lambda r: (r["folder"], r["date"])):
                w.writerow([r["date"], r["folder"], r["from"], r["subject"], r["size"], r["path"]])
        return index, len(rows)

    def close(self):
        self._manifest.close()
        self._failures.close()


def progress(folder, i, total, saved, skipped, failed):
    sys.stdout.write(
        f"\r  {folder[:50]:<50} {i}/{total}  saved={saved} skipped={skipped} failed={failed}   "
    )
    sys.stdout.flush()


# --------------------------------------------------------------------------
# connecting (tries several methods so people don't have to)
# --------------------------------------------------------------------------

class AuthFailed(Exception):
    """The server was reached but rejected the email/password."""


class CannotConnect(Exception):
    """No connection method worked, and it wasn't clearly a password problem."""


def _exchangelib():
    try:
        import exchangelib
        return exchangelib
    except ImportError:
        sys.exit("exchangelib is not installed. Run:  pip install -r requirements.txt")


def ews_try(opts, password, auth=None, autodiscover=False):
    """Log in over EWS and make one real request, failing fast."""
    x = _exchangelib()
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    if opts.insecure:
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    creds = x.Credentials(username=opts.user or opts.email, password=password)
    if autodiscover:
        account = x.Account(primary_smtp_address=opts.email, credentials=creds,
                            autodiscover=True, access_type=x.DELEGATE,
                            config=x.Configuration(credentials=creds, retry_policy=x.FailFast()))
    else:
        kw = {"auth_type": {"ntlm": x.NTLM, "basic": x.BASIC}[auth]} if auth else {}
        config = x.Configuration(server=opts.server, credentials=creds,
                                 retry_policy=x.FailFast(), **kw)
        account = x.Account(primary_smtp_address=opts.email, config=config,
                            autodiscover=False, access_type=x.DELEGATE)
    account.inbox.total_count  # forces an authenticated round trip
    # connected: from now on, wait out throttling instead of failing
    account.protocol.config.retry_policy = x.FaultTolerance(max_wait=900)
    return account


def imap_connect(opts, password):
    conn = imaplib.IMAP4_SSL(opts.server, opts.imap_port, timeout=30)
    conn.login(opts.user or opts.email, password)
    return conn


def _looks_like_auth_error(e):
    from exchangelib.errors import UnauthorizedError
    text = str(e).lower()
    return (isinstance(e, (UnauthorizedError, imaplib.IMAP4.error))
            or "401" in text or "unauthorized" in text or "authentication failed" in text
            or "wrong username or password" in text)


def connect(opts, password, log=print):
    """Try EWS (auto/NTLM/basic auth, then autodiscover) and then IMAP.

    Returns ("ews", Account) or ("imap", IMAP4_SSL).
    """
    attempts = []
    if opts.method in (None, "ews"):
        if opts.autodiscover:
            attempts.append(("EWS via autodiscover", None, True))
        else:
            auths = [None, "ntlm", "basic"] if opts.auth in (None, "auto") else [opts.auth]
            for a in auths:
                attempts.append((f"EWS on {opts.server} ({a or 'auto'} login)", a, False))
            attempts.append(("EWS via autodiscover", None, True))

    auth_errors, last = 0, None
    for label, auth, auto in attempts:
        log(f"  trying {label} ...")
        try:
            return "ews", ews_try(opts, password, auth=auth, autodiscover=auto)
        except Exception as e:
            last = e
            if _looks_like_auth_error(e):
                auth_errors += 1
                log("    password rejected")
            else:
                log(f"    no luck ({type(e).__name__})")

    if opts.method in (None, "imap"):
        log(f"  trying IMAP on {opts.server}:{opts.imap_port} ...")
        try:
            return "imap", imap_connect(opts, password)
        except Exception as e:
            last = e
            if _looks_like_auth_error(e):
                auth_errors += 1
                log("    password rejected")
            else:
                log(f"    no luck ({type(e).__name__})")

    if auth_errors:
        raise AuthFailed(str(last))
    raise CannotConnect(f"{type(last).__name__}: {last}")


# --------------------------------------------------------------------------
# EWS (Exchange Web Services) - preferred
# --------------------------------------------------------------------------

def ews_folders(account, include_archive=True):
    """Yield (relative_path, folder) for every folder in the mailbox."""
    roots = [("", account.msg_folder_root)]
    if include_archive:
        try:
            roots.append(("Online Archive", account.archive_msg_folder_root))
        except Exception:
            pass
    for prefix, root in roots:
        try:
            base = root.absolute
            for f in root.walk():
                rel = f.absolute[len(base):].strip("/")
                yield (f"{prefix}/{rel}" if prefix else rel), f
        except Exception:
            if not prefix:
                raise  # the main mailbox must work; a missing archive is fine


def ews_summary(account, opts):
    rows = []
    for path, f in ews_folders(account, include_archive=not opts.no_archive):
        try:
            n = f.total_count
        except Exception:
            n = None
        if n:
            rows.append((path, n))
    return rows


def ews_backup(account, opts):
    from exchangelib.items import Item

    store = Store(opts.out)
    folders = list(ews_folders(account, include_archive=not opts.no_archive))
    grand = {"saved": 0, "skipped": 0, "failed": 0}

    for path, folder in folders:
        if opts.folder and not any(path.lower().startswith(x.lower()) for x in opts.folder):
            continue
        try:
            count = folder.total_count
        except Exception:
            count = None
        if not count:
            continue
        print(f"\n[{path}] {count} items")

        # 1) list ids only (cheap), 2) fetch MIME for the ones not yet saved
        try:
            ids = list(folder.all().only("id", "changekey", "datetime_received"))
        except Exception as e:
            print(f"  ! cannot list folder: {e}")
            store.fail(f"folder:{path}", path, e)
            continue

        todo = [i for i in ids if isinstance(i, Item) and not store.has(f"ews:{i.id}")]
        skipped = len(ids) - len(todo)
        saved = failed = 0
        by_id = {i.id: i for i in todo}

        for start in range(0, len(todo), opts.batch):
            chunk = todo[start:start + opts.batch]
            try:
                fetched = list(account.fetch(ids=chunk, only_fields=["mime_content"]))
            except Exception as e:
                fetched = [e] * len(chunk)
            for src, item in zip(chunk, fetched):
                key = f"ews:{src.id}"
                if isinstance(item, Exception) or not getattr(item, "mime_content", None):
                    # retry this one on its own before giving up
                    try:
                        item = list(account.fetch(ids=[src], only_fields=["mime_content"]))[0]
                        if isinstance(item, Exception):
                            raise item
                        if not item.mime_content:
                            raise ValueError("server returned no MIME content")
                    except Exception as e:
                        store.fail(key, path, e)
                        failed += 1
                        continue
                raw = item.mime_content
                if isinstance(raw, str):
                    raw = raw.encode("utf-8")
                store.save(key, path, raw, fallback_date=by_id[src.id].datetime_received)
                saved += 1
            progress(path, skipped + saved + failed, len(ids), saved, skipped, failed)

        progress(path, len(ids), len(ids), saved, skipped, failed)
        grand["saved"] += saved
        grand["skipped"] += skipped
        grand["failed"] += failed

    return finish(store, grand)


# --------------------------------------------------------------------------
# IMAP - fallback if EWS is blocked
# --------------------------------------------------------------------------

LIST_RE = re.compile(r'\((?P<flags>[^)]*)\) (?P<delim>"[^"]*"|NIL) (?P<name>.+)')


def imap_decode_folder(name):
    """Decode IMAP modified UTF-7 folder names."""
    import base64

    def repl(m):
        s = m.group(1)
        if s == "":
            return "&"
        s = s.replace(",", "/")
        s += "=" * (-len(s) % 4)
        return base64.b64decode(s).decode("utf-16-be")
    return re.sub(r"&([^-]*)-", repl, name)


def imap_folders(conn):
    typ, data = conn.list()
    folders = []
    for line in data:
        m = LIST_RE.match(line.decode("utf-8", "replace"))
        if not m or "\\Noselect" in m.group("flags"):
            continue
        name = m.group("name").strip()
        if name.startswith('"'):
            name = name[1:-1].replace('\\"', '"')
        delim = m.group("delim").strip('"')
        path = imap_decode_folder(name)
        if delim and delim != "/":
            path = path.replace(delim, "/")
        folders.append((name, path))
    return folders


def imap_backup(conn, opts):
    store = Store(opts.out)
    grand = {"saved": 0, "skipped": 0, "failed": 0}

    for raw_name, path in imap_folders(conn):
        if opts.folder and not any(path.lower().startswith(x.lower()) for x in opts.folder):
            continue
        typ, _ = conn.select(f'"{raw_name}"', readonly=True)
        if typ != "OK":
            print(f"  ! cannot open {path}")
            continue
        uidvalidity = conn.response("UIDVALIDITY")[1][0].decode()
        typ, data = conn.uid("search", None, "ALL")
        uids = data[0].split() if data and data[0] else []
        if not uids:
            continue
        print(f"\n[{path}] {len(uids)} messages")
        saved = skipped = failed = 0
        for n, uid in enumerate(uids, 1):
            key = f"imap:{raw_name}:{uidvalidity}:{uid.decode()}"
            if store.has(key):
                skipped += 1
                continue
            try:
                typ, msg = conn.uid("fetch", uid, "(BODY.PEEK[])")
                raw = next(p[1] for p in msg if isinstance(p, tuple))
                store.save(key, path, raw)
                saved += 1
            except Exception as e:
                store.fail(key, path, e)
                failed += 1
            if n % 10 == 0 or n == len(uids):
                progress(path, n, len(uids), saved, skipped, failed)
        grand["saved"] += saved
        grand["skipped"] += skipped
        grand["failed"] += failed
    conn.logout()
    return finish(store, grand)


# --------------------------------------------------------------------------
# shared steps
# --------------------------------------------------------------------------

def finish(store, grand):
    index, total = store.write_index()
    store.close()
    print(f"\n\nDone. new={grand['saved']} already-had={grand['skipped']} failed={grand['failed']}")
    print(f"Total messages in backup: {total}")
    if grand["failed"]:
        print(f"{grand['failed']} message(s) failed (see {store.failures_path.name}). "
              "Run the backup again to retry them.")
    return grand


def run_backup(method, handle, opts):
    if method == "imap":
        return imap_backup(handle, opts)
    return ews_backup(handle, opts)


def make_mbox(out):
    """Bundle the saved .eml files into one .mbox per folder (for Thunderbird etc.)."""
    root = Path(out) / "mail"
    dest = Path(out) / "mbox"
    if not root.exists():
        sys.exit(f"No backup found in {root}")
    for folder in sorted({p.parent for p in root.rglob("*.eml")}):
        rel = folder.relative_to(root)
        target = dest / rel.with_name(rel.name + ".mbox")
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            target.unlink()
        box = mailbox.mbox(str(target))
        box.lock()
        n = 0
        for eml in sorted(folder.glob("*.eml")):
            box.add(mailbox.mboxMessage(eml.read_bytes()))
            n += 1
        box.flush()
        box.unlock()
        box.close()
        print(f"  {n:>6}  {target}")
    return dest


def default_out(email):
    return str(Path(DEFAULT_OUT) / safe_name(email.lower(), 120))


def server_from_input(text):
    """Accept 'mail.x.edu', 'https://mail.x.edu/owa/...' etc. and return the host."""
    text = text.strip()
    m = re.match(r"^(?:[a-z]+://)?([^/\s:]+)", text, re.I)
    return m.group(1) if m else text


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
    except Exception:
        pass


# --------------------------------------------------------------------------
# guided mode (what the double-click launchers run)
# --------------------------------------------------------------------------

def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or default


def yes(prompt, default=True):
    value = input(f"{prompt} [{'Y/n' if default else 'y/N'}]: ").strip().lower()
    return default if not value else value.startswith("y")


def wizard():
    print("=" * 64)
    print("  University mailbox backup")
    print("  Saves every email (with attachments) from every folder to this computer.")
    print("  Your password is only sent to your university's mail server.")
    print("=" * 64 + "\n")

    email = ""
    while "@" not in email:
        email = ask("Your university email address")
    opts = argparse.Namespace(
        email=email, user=None, server=os.environ.get("GUC_SERVER", DEFAULT_SERVER),
        auth="auto", autodiscover=False, insecure=False, imap_port=993, method=None,
        no_archive=False, out=default_out(email), folder=None, batch=20,
    )

    method = handle = None
    password = os.environ.get("GUC_PASSWORD")
    tries = 0
    while handle is None:
        password = password or getpass.getpass("Password (it will not show while you type): ")
        print("\nConnecting...")
        try:
            method, handle = connect(opts, password)
        except AuthFailed:
            password = None
            tries += 1
            print("\nThe server rejected that email/password.")
            if tries >= 3:
                if yes("Does your login use a different username than your email "
                       "(e.g. GUC\\first.last)?", default=False):
                    opts.user = ask("Username")
                elif not yes("Try again?"):
                    return 1
        except CannotConnect as e:
            print(f"\nCould not reach the mail server {opts.server}.\n  ({e})")
            print("Open your webmail in a browser and copy the address from the address bar.")
            new = ask("Webmail address (leave empty to quit)")
            if not new:
                return 1
            opts.server = server_from_input(new)

    print(f"\nConnected ({method.upper()}). Backup folder: {Path(opts.out).resolve()}\n")
    if method == "ews":
        rows = ews_summary(handle, opts)
        for path, n in rows:
            print(f"  {n:>7}  {path}")
        print(f"  {sum(n for _, n in rows):>7}  items in total\n")
    input("Press Enter to start downloading (you can close the window any time and "
          "run it again later to continue)...")

    result = run_backup(method, handle, opts)

    if yes("\nAlso create .mbox files (handy for importing into Thunderbird/Gmail)?", default=False):
        make_mbox(opts.out)

    print(f"\nYour emails are in: {Path(opts.out).resolve()}")
    print("Open index.csv for a list of everything, or open any .eml file to read it.")
    open_folder(str(Path(opts.out).resolve()))
    return 0 if not result["failed"] else 2


# --------------------------------------------------------------------------
# command line (for scripting / advanced use)
# --------------------------------------------------------------------------

def get_password(opts):
    return os.environ.get("GUC_PASSWORD") or getpass.getpass(f"Password for {opts.user or opts.email}: ")


def cli_connect(opts):
    try:
        return connect(opts, get_password(opts))
    except AuthFailed:
        sys.exit("Login rejected: wrong password, or try --user 'DOMAIN\\username'.")
    except CannotConnect as e:
        sys.exit(f"Could not connect: {e}\nCheck --server (the host in your webmail address bar), "
                 "or try --insecure / --autodiscover.")


def cmd_probe(opts):
    method, handle = cli_connect(opts)
    print(f"Connected via {method.upper()}.")
    if method == "ews":
        rows = ews_summary(handle, opts)
        for path, n in rows:
            print(f"  {n:>7}  {path}")
        print(f"  {sum(n for _, n in rows):>7}  items in total")
    else:
        for _, path in imap_folders(handle):
            print(f"  {path}")
        handle.logout()


def cmd_backup(opts):
    method, handle = cli_connect(opts)
    print(f"Connected via {method.upper()}. Saving to {Path(opts.out).resolve()}")
    result = run_backup(method, handle, opts)
    sys.exit(0 if not result["failed"] else 2)


def cmd_mbox(opts):
    make_mbox(opts.out)


def main():
    if len(sys.argv) == 1:
        # no arguments (e.g. double-clicked): guided mode
        try:
            code = wizard()
        except (KeyboardInterrupt, EOFError):
            print("\nStopped. Progress is saved; run it again to continue.")
            code = 130
        if sys.stdin.isatty():
            try:
                input("\nPress Enter to close.")
            except (KeyboardInterrupt, EOFError):
                pass
        sys.exit(code)

    p = argparse.ArgumentParser(
        description="Back up an Exchange/OWA mailbox to local .eml files. "
                    "Run with no arguments for guided mode.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
        sp.add_argument("--email", required=True, help="your mailbox address")
        sp.add_argument("--user", help="login name if different from --email, e.g. 'GUC\\\\first.last'")
        sp.add_argument("--server", default=os.environ.get("GUC_SERVER", DEFAULT_SERVER),
                        help=f"mail server host from your webmail URL (default {DEFAULT_SERVER})")
        sp.add_argument("--method", choices=["ews", "imap"],
                        help="force a method (default: try EWS, then IMAP)")
        sp.add_argument("--auth", choices=["auto", "ntlm", "basic"], default="auto")
        sp.add_argument("--autodiscover", action="store_true", help="find the EWS server automatically")
        sp.add_argument("--insecure", action="store_true", help="skip TLS certificate verification")
        sp.add_argument("--imap-port", type=int, default=993)
        sp.add_argument("--no-archive", action="store_true", help="skip the Online Archive mailbox")
        sp.add_argument("--out", help=f"output directory (default {DEFAULT_OUT}/<email>)")

    sp = sub.add_parser("probe", help="check that login works and count items per folder")
    common(sp)
    sp.set_defaults(func=cmd_probe)

    sp = sub.add_parser("backup", help="download every message (resumable)")
    common(sp)
    sp.add_argument("--folder", action="append",
                    help="only back up folders starting with this path (repeatable), e.g. --folder Inbox")
    sp.add_argument("--batch", type=int, default=20, help="messages fetched per EWS request")
    sp.set_defaults(func=cmd_backup)

    sp = sub.add_parser("mbox", help="also pack the .eml files into per-folder .mbox files")
    sp.add_argument("--out", required=True, help=f"the backup directory, e.g. {DEFAULT_OUT}/<email>")
    sp.set_defaults(func=cmd_mbox)

    opts = p.parse_args()
    if getattr(opts, "email", None) and not opts.out:
        opts.out = default_out(opts.email)
    opts.folder = getattr(opts, "folder", None)
    try:
        opts.func(opts)
    except KeyboardInterrupt:
        print("\nInterrupted - progress is saved, re-run the same command to continue.")
        sys.exit(130)


if __name__ == "__main__":
    main()
