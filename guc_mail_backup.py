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
    python guc_mail_backup.py probe  --email you@student.guc.edu.eg
    python guc_mail_backup.py backup --email you@student.guc.edu.eg
    python guc_mail_backup.py mbox   --out guc-mail-backup
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
import time
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
# EWS (Exchange Web Services) - preferred
# --------------------------------------------------------------------------

def ews_account(args, password):
    try:
        from exchangelib import (DELEGATE, NTLM, BASIC, Account, Configuration,
                                 Credentials, FaultTolerance)
        from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    except ImportError:
        sys.exit("exchangelib is not installed. Run:  pip install -r requirements.txt")

    if args.insecure:
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    creds = Credentials(username=args.user or args.email, password=password)
    retry = FaultTolerance(max_wait=900)  # ride out Exchange throttling
    if args.autodiscover:
        return Account(primary_smtp_address=args.email, credentials=creds,
                       autodiscover=True, access_type=DELEGATE,
                       config=Configuration(credentials=creds, retry_policy=retry))
    auth = {"ntlm": NTLM, "basic": BASIC}.get(args.auth)
    kw = {"auth_type": auth} if auth else {}
    config = Configuration(server=args.server, credentials=creds, retry_policy=retry, **kw)
    return Account(primary_smtp_address=args.email, config=config,
                   autodiscover=False, access_type=DELEGATE)


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
        except Exception as e:
            if prefix:
                print(f"  (no online archive: {type(e).__name__})")
            else:
                raise


def ews_backup(args, password):
    from exchangelib.items import Item

    account = ews_account(args, password)
    print(f"Connected via EWS as {account.primary_smtp_address}")
    store = Store(args.out)

    folders = list(ews_folders(account, include_archive=not args.no_archive))
    print(f"Found {len(folders)} folders")
    grand = {"saved": 0, "skipped": 0, "failed": 0}

    for path, folder in folders:
        if args.folder and not any(path.lower().startswith(x.lower()) for x in args.folder):
            continue
        try:
            count = folder.total_count
        except Exception:
            count = None
        if not count:
            continue
        print(f"\n[{path}] {count} items ({folder.folder_class or 'folder'})")

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

        for start in range(0, len(todo), args.batch):
            chunk = todo[start:start + args.batch]
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

    finish(store, grand)


# --------------------------------------------------------------------------
# IMAP - fallback if EWS is blocked
# --------------------------------------------------------------------------

LIST_RE = re.compile(r'\((?P<flags>[^)]*)\) (?P<delim>"[^"]*"|NIL) (?P<name>.+)')


def imap_decode_folder(name):
    """Decode IMAP modified UTF-7 folder names."""
    def repl(m):
        s = m.group(1)
        if s == "":
            return "&"
        s = s.replace(",", "/")
        s += "=" * (-len(s) % 4)
        import base64
        return base64.b64decode(s).decode("utf-16-be")
    return re.sub(r"&([^-]*)-", repl, name)


def imap_connect(args, password):
    conn = imaplib.IMAP4_SSL(args.server, args.imap_port)
    conn.login(args.user or args.email, password)
    return conn


def imap_backup(args, password):
    conn = imap_connect(args, password)
    print(f"Connected via IMAP to {args.server}")
    store = Store(args.out)
    grand = {"saved": 0, "skipped": 0, "failed": 0}

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
        folders.append((name, delim))
    print(f"Found {len(folders)} folders")

    for raw_name, delim in folders:
        path = imap_decode_folder(raw_name)
        if delim and delim != "/":
            path = path.replace(delim, "/")
        if args.folder and not any(path.lower().startswith(x.lower()) for x in args.folder):
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
    finish(store, grand)


# --------------------------------------------------------------------------
# commands
# --------------------------------------------------------------------------

def finish(store, grand):
    index, total = store.write_index()
    store.close()
    print(f"\n\nDone. saved={grand['saved']} already-had={grand['skipped']} failed={grand['failed']}")
    print(f"Total messages in backup: {total}")
    print(f"Index: {index}")
    if grand["failed"]:
        print(f"Failures logged to {store.failures_path} - just re-run the same command to retry them.")


def get_password(args):
    return os.environ.get("GUC_PASSWORD") or getpass.getpass(f"Password for {args.user or args.email}: ")


def cmd_probe(args):
    password = get_password(args)
    print(f"Trying EWS on https://{args.server}/EWS/Exchange.asmx ...")
    try:
        account = ews_account(args, password)
        total = 0
        for path, f in ews_folders(account, include_archive=not args.no_archive):
            if f.total_count:
                print(f"  {f.total_count:>7}  {path}")
                total += f.total_count
        print(f"EWS works. {total} items in total. Run the 'backup' command.")
        return
    except Exception as e:
        print(f"  EWS failed: {type(e).__name__}: {e}")
    print(f"Trying IMAP on {args.server}:{args.imap_port} ...")
    try:
        conn = imap_connect(args, password)
        typ, data = conn.list()
        print(f"IMAP works ({len(data)} folders). Run:  backup --method imap")
        conn.logout()
    except Exception as e:
        print(f"  IMAP failed: {type(e).__name__}: {e}")
        print("\nNeither EWS nor IMAP worked. Check the server name (the host in your OWA "
              "address bar), try --user 'GUC\\username', --auth ntlm/basic, or --insecure.")


def cmd_backup(args):
    password = get_password(args)
    if args.method == "imap":
        imap_backup(args, password)
    else:
        ews_backup(args, password)


def cmd_mbox(args):
    """Bundle the saved .eml files into one .mbox per folder (for Thunderbird etc.)."""
    root = Path(args.out) / "mail"
    dest = Path(args.out) / "mbox"
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


def main():
    p = argparse.ArgumentParser(description="Back up an Exchange/OWA mailbox to local .eml files.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
        sp.add_argument("--email", required=True, help="your mailbox address")
        sp.add_argument("--user", help="login name if different from --email, e.g. 'GUC\\\\first.last'")
        sp.add_argument("--server", default=DEFAULT_SERVER,
                        help=f"mail server host from your OWA URL (default {DEFAULT_SERVER})")
        sp.add_argument("--auth", choices=["auto", "ntlm", "basic"], default="auto")
        sp.add_argument("--autodiscover", action="store_true", help="find the EWS server automatically")
        sp.add_argument("--insecure", action="store_true", help="skip TLS certificate verification")
        sp.add_argument("--imap-port", type=int, default=993)
        sp.add_argument("--no-archive", action="store_true", help="skip the Online Archive mailbox")
        sp.add_argument("--out", default=DEFAULT_OUT, help=f"output directory (default {DEFAULT_OUT})")

    sp = sub.add_parser("probe", help="check that login works and count items per folder")
    common(sp)
    sp.set_defaults(func=cmd_probe)

    sp = sub.add_parser("backup", help="download every message (resumable)")
    common(sp)
    sp.add_argument("--method", choices=["ews", "imap"], default="ews")
    sp.add_argument("--folder", action="append",
                    help="only back up folders starting with this path (repeatable), e.g. --folder Inbox")
    sp.add_argument("--batch", type=int, default=20, help="messages fetched per EWS request")
    sp.set_defaults(func=cmd_backup)

    sp = sub.add_parser("mbox", help="also pack the .eml files into per-folder .mbox files")
    sp.add_argument("--out", default=DEFAULT_OUT)
    sp.set_defaults(func=cmd_mbox)

    args = p.parse_args()
    try:
        args.func(args)
    except KeyboardInterrupt:
        print("\nInterrupted - progress is saved, re-run the same command to continue.")
        sys.exit(130)


if __name__ == "__main__":
    main()
