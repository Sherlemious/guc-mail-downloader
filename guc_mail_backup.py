#!/usr/bin/env python3
"""Back up an entire Exchange / Outlook Web App mailbox to local .eml files.

Built for the GUC mailbox (on-prem Exchange 2016 behind Outlook Web App), but
works with any Exchange server that exposes EWS or IMAP.

Every message is saved byte-for-byte as the server's original MIME, so
attachments, inline images, headers and HTML bodies are all preserved. Each
folder becomes a directory, an index.csv summarises everything, and runs can
be interrupted and resumed at any time.

Usage:
    python guc_mail_backup.py                     # guided mode, just answer the questions
    python guc_mail_backup.py backup --email you@student.guc.edu.eg
    python guc_mail_backup.py --help              # all commands and options

See README.md for people and AGENTS.md for AI agents / automation.
"""

import argparse
import base64
import concurrent.futures as cf
import csv
import fnmatch
import getpass
import hashlib
import html
import imaplib
import json
import mailbox
import os
import re
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from email import message_from_bytes, policy
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from urllib.parse import quote

__version__ = "2.0.0"

DEFAULT_SERVER = "mail.guc.edu.eg"
DEFAULT_ROOT = "guc-mail-backup"
MANIFEST = "manifest.jsonl"
FAILURES = "failures.jsonl"

# exit codes (documented in AGENTS.md)
EXIT_OK = 0
EXIT_ERROR = 1
EXIT_PARTIAL = 2
EXIT_AUTH = 3
EXIT_CONNECT = 4
EXIT_INTERRUPTED = 130

QUIET = False


class CliError(Exception):
    def __init__(self, message, code=EXIT_ERROR):
        super().__init__(message)
        self.code = code


def log(*args):
    """Human-readable progress. Always stderr, so stdout stays clean for --json."""
    if not QUIET:
        print(*args, file=sys.stderr, flush=True)


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


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def parse_date(text, end=False):
    """'2024-01-31' or a full ISO timestamp -> aware datetime (local time zone).

    With end=True a bare date means the end of that day, so --until is inclusive.
    """
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        raise CliError(f"Bad date {text!r}; use YYYY-MM-DD")
    if end and len(text) <= 10:
        dt += timedelta(days=1)
    if dt.tzinfo is None:
        dt = dt.astimezone()
    return dt


def in_range(dt, opts):
    if dt is None:
        return True
    if opts.since and dt < opts.since:
        return False
    if opts.until and dt >= opts.until:
        return False
    return True


def folder_selected(path, opts, folder_class=None):
    def match(p, pattern):
        p, pattern = p.lower(), pattern.lower().rstrip("/")
        return p == pattern or p.startswith(pattern + "/") or fnmatch.fnmatch(p, pattern)

    if opts.folder and not any(match(path, x) for x in opts.folder):
        return False
    if opts.exclude and any(match(path, x) for x in opts.exclude):
        return False
    if opts.mail_only and folder_class and not folder_class.startswith("IPF.Note"):
        return False
    return True


class Store:
    """Writes .eml files and keeps a resumable manifest of what is done."""

    def __init__(self, out_dir):
        self.out = Path(out_dir)
        self.out.mkdir(parents=True, exist_ok=True)
        self.manifest_path = self.out / MANIFEST
        self.failures_path = self.out / FAILURES
        self.done = {rec["key"] for rec in read_manifest(self.out)}
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
            "path": path.relative_to(self.out).as_posix(),
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
        rows = read_manifest(self.out)
        index = self.out / "index.csv"
        with open(index, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["date", "folder", "from", "subject", "size_bytes", "file"])
            for r in sorted(rows, key=lambda r: (r["folder"], r["date"])):
                w.writerow([r["date"], r["folder"], r["from"], r["subject"], r["size"], r["path"]])
        return index, len(rows)

    def close(self):
        self._manifest.close()
        self._failures.close()


def read_manifest(out):
    """All saved messages, de-duplicated by key (last entry wins)."""
    path = Path(out) / MANIFEST
    recs = {}
    if path.exists():
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                try:
                    rec = json.loads(line)
                    recs[rec["key"]] = rec
                except Exception:
                    pass
    return list(recs.values())


def write_manifest(out, records):
    path = Path(out) / MANIFEST
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as fh:
        for rec in records:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    os.replace(tmp, path)


class Progress:
    """Live one-line progress on a terminal; one summary line per folder otherwise."""

    def __init__(self, label, total):
        self.label, self.total = label, total
        self.live = sys.stderr.isatty() and not QUIET

    def update(self, done, saved, skipped, failed):
        if self.live:
            sys.stderr.write(f"\r  {self.label[:40]:<40} {done}/{self.total}  "
                             f"new={saved} had={skipped} failed={failed}   ")
            sys.stderr.flush()

    def close(self, saved, skipped, failed):
        if self.live:
            self.update(self.total, saved, skipped, failed)
            sys.stderr.write("\n")
        else:
            log(f"  {self.label}: new={saved} had={skipped} failed={failed}")


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
        raise CliError("exchangelib is not installed. Run:  pip install -r requirements.txt")


def ews_try(opts, password, auth=None, autodiscover=False):
    """Log in over EWS and make one real request, failing fast."""
    x = _exchangelib()
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    if opts.insecure:
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    creds = x.Credentials(username=opts.user or opts.email, password=password)
    conns = max(2, getattr(opts, "workers", None) or 4)
    if autodiscover:
        account = x.Account(primary_smtp_address=opts.email, credentials=creds,
                            autodiscover=True, access_type=x.DELEGATE,
                            config=x.Configuration(credentials=creds, retry_policy=x.FailFast(),
                                                   max_connections=conns))
    else:
        kw = {"auth_type": {"ntlm": x.NTLM, "basic": x.BASIC}[auth]} if auth else {}
        config = x.Configuration(server=opts.server, credentials=creds,
                                 retry_policy=x.FailFast(), max_connections=conns, **kw)
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
    try:
        from exchangelib.errors import UnauthorizedError
    except ImportError:
        UnauthorizedError = ()
    text = str(e).lower()
    return (isinstance(e, (UnauthorizedError, imaplib.IMAP4.error))
            or "401" in text or "unauthorized" in text or "authentication failed" in text
            or "wrong username or password" in text)


def connect(opts, password, say=log):
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
        say(f"  trying {label} ...")
        try:
            return "ews", ews_try(opts, password, auth=auth, autodiscover=auto)
        except Exception as e:
            last = e
            if _looks_like_auth_error(e):
                auth_errors += 1
                say("    password rejected")
            else:
                say(f"    no luck ({type(e).__name__})")

    if opts.method in (None, "imap"):
        say(f"  trying IMAP on {opts.server}:{opts.imap_port} ...")
        try:
            return "imap", imap_connect(opts, password)
        except Exception as e:
            last = e
            if _looks_like_auth_error(e):
                auth_errors += 1
                say("    password rejected")
            else:
                say(f"    no luck ({type(e).__name__})")

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
            rows.append({"folder": path, "items": n, "class": getattr(f, "folder_class", None) or ""})
    return rows


def ews_fetch_chunk(account, chunk):
    """Fetch MIME for a batch; returns [(item, bytes | Exception)]. Runs in worker threads."""
    try:
        fetched = list(account.fetch(ids=chunk, only_fields=["mime_content"]))
    except Exception as e:
        fetched = [e] * len(chunk)
    out = []
    for src, item in zip(chunk, fetched):
        if isinstance(item, Exception) or not getattr(item, "mime_content", None):
            # retry this one on its own before giving up
            try:
                item = list(account.fetch(ids=[src], only_fields=["mime_content"]))[0]
                if isinstance(item, Exception):
                    raise item
                if not item.mime_content:
                    raise ValueError("server returned no MIME content")
            except Exception as e:
                out.append((src, e))
                continue
        raw = item.mime_content
        out.append((src, raw.encode("utf-8") if isinstance(raw, str) else raw))
    return out


def ews_backup(account, opts):
    store = Store(opts.out)
    summary = new_summary("ews", opts)

    for path, folder in ews_folders(account, include_archive=not opts.no_archive):
        if not folder_selected(path, opts, getattr(folder, "folder_class", None)):
            continue
        try:
            count = folder.total_count
        except Exception:
            count = None
        if not count:
            continue
        log(f"\n[{path}] {count} items")

        # 1) list ids only (cheap), 2) fetch MIME for the ones not yet saved
        try:
            ids = list(folder.all().only("id", "changekey", "datetime_received"))
        except Exception as e:
            log(f"  ! cannot list folder: {e}")
            store.fail(f"folder:{path}", path, e)
            summary["failed"] += 1
            continue

        ids = [i for i in ids if not isinstance(i, Exception) and getattr(i, "id", None)]
        ids = [i for i in ids if in_range(getattr(i, "datetime_received", None), opts)]
        todo = [i for i in ids if not store.has(f"ews:{i.id}")]
        stats = {"folder": path, "new": 0, "had": len(ids) - len(todo), "failed": 0}
        if opts.dry_run:
            stats["would_download"] = len(todo)
            log(f"  would download {len(todo)}, already have {stats['had']}")
            add_folder(summary, stats)
            continue

        by_id = {i.id: i for i in todo}
        progress = Progress(path, len(ids))
        chunks = [todo[s:s + opts.batch] for s in range(0, len(todo), opts.batch)]
        pool = cf.ThreadPoolExecutor(max_workers=max(1, opts.workers))
        try:
            for results in pool.map(lambda c: ews_fetch_chunk(account, c), chunks):
                for src, raw in results:
                    key = f"ews:{src.id}"
                    if isinstance(raw, Exception):
                        store.fail(key, path, raw)
                        stats["failed"] += 1
                    else:
                        store.save(key, path, raw, fallback_date=by_id[src.id].datetime_received)
                        stats["new"] += 1
                progress.update(stats["had"] + stats["new"] + stats["failed"],
                                stats["new"], stats["had"], stats["failed"])
        finally:
            pool.shutdown(wait=False, cancel_futures=True)
        progress.close(stats["new"], stats["had"], stats["failed"])
        add_folder(summary, stats)

    return finish(store, summary)


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


def imap_date(dt):
    return dt.strftime("%d-%b-%Y")


def imap_backup(conn, opts):
    store = Store(opts.out)
    summary = new_summary("imap", opts)

    criteria = []
    if opts.since:
        criteria += ["SINCE", imap_date(opts.since)]
    if opts.until:
        criteria += ["BEFORE", imap_date(opts.until)]

    for raw_name, path in imap_folders(conn):
        if not folder_selected(path, opts):
            continue
        typ, _ = conn.select(f'"{raw_name}"', readonly=True)
        if typ != "OK":
            log(f"  ! cannot open {path}")
            continue
        uidvalidity = conn.response("UIDVALIDITY")[1][0].decode()
        typ, data = conn.uid("search", None, *(criteria or ["ALL"]))
        uids = data[0].split() if data and data[0] else []
        if not uids:
            continue
        log(f"\n[{path}] {len(uids)} messages")
        keys = [(uid, f"imap:{raw_name}:{uidvalidity}:{uid.decode()}") for uid in uids]
        todo = [(u, k) for u, k in keys if not store.has(k)]
        stats = {"folder": path, "new": 0, "had": len(keys) - len(todo), "failed": 0}
        if opts.dry_run:
            stats["would_download"] = len(todo)
            log(f"  would download {len(todo)}, already have {stats['had']}")
            add_folder(summary, stats)
            continue
        progress = Progress(path, len(keys))
        for n, (uid, key) in enumerate(todo, 1):
            try:
                typ, msg = conn.uid("fetch", uid, "(BODY.PEEK[])")
                raw = next(p[1] for p in msg if isinstance(p, tuple))
                store.save(key, path, raw)
                stats["new"] += 1
            except Exception as e:
                store.fail(key, path, e)
                stats["failed"] += 1
            if n % 10 == 0:
                progress.update(stats["had"] + n, stats["new"], stats["had"], stats["failed"])
        progress.close(stats["new"], stats["had"], stats["failed"])
        add_folder(summary, stats)
    conn.logout()
    return finish(store, summary)


# --------------------------------------------------------------------------
# shared backup steps
# --------------------------------------------------------------------------

def new_summary(method, opts):
    return {"ok": True, "command": "backup", "method": method, "dry_run": bool(opts.dry_run),
            "out": str(Path(opts.out).resolve()), "new": 0, "had": 0, "failed": 0, "folders": []}


def add_folder(summary, stats):
    summary["folders"].append(stats)
    summary["new"] += stats["new"]
    summary["had"] += stats["had"]
    summary["failed"] += stats["failed"]
    if "would_download" in stats:
        summary["would_download"] = summary.get("would_download", 0) + stats["would_download"]


def finish(store, summary):
    index, total = store.write_index()
    store.close()
    summary["total_in_backup"] = total
    summary["index"] = str(index.resolve())
    if summary["dry_run"]:
        log(f"\nDry run: would download {summary.get('would_download', 0)}, "
            f"already have {summary['had']}.")
        return summary
    log(f"\nDone. new={summary['new']} already-had={summary['had']} failed={summary['failed']}")
    log(f"Total messages in backup: {total}")
    if summary["failed"]:
        summary["ok"] = False
        log(f"{summary['failed']} message(s) failed (see {store.failures_path.name}). "
            "Run the backup again to retry them.")
    return summary


def run_backup(method, handle, opts):
    if method == "imap":
        return imap_backup(handle, opts)
    return ews_backup(handle, opts)


# --------------------------------------------------------------------------
# working with a finished backup (no network needed)
# --------------------------------------------------------------------------

def default_out(email):
    return str(Path(DEFAULT_ROOT) / safe_name(email.lower(), 120))


def resolve_out(opts):
    """--out, else the folder for --email, else the only backup under guc-mail-backup/."""
    if getattr(opts, "out", None):
        return opts.out
    if getattr(opts, "email", None):
        return default_out(opts.email)
    root = Path(DEFAULT_ROOT)
    found = sorted(p.parent for p in root.glob(f"*/{MANIFEST}")) if root.exists() else []
    if len(found) == 1:
        return str(found[0])
    if not found:
        raise CliError(f"No backup found under {root}/. Pass --out <backup folder>.")
    raise CliError("Several backups found; pick one with --out:\n  " + "\n  ".join(map(str, found)))


def load_backup(out):
    recs = read_manifest(out)
    if not recs:
        raise CliError(f"No backup found in {out} (missing or empty {MANIFEST}).")
    return recs


def parse_message(path):
    return message_from_bytes(Path(path).read_bytes(), policy=policy.default)


def make_mbox(out):
    """Bundle the saved .eml files into one .mbox per folder (for Thunderbird etc.)."""
    root = Path(out) / "mail"
    dest = Path(out) / "mbox"
    if not root.exists():
        raise CliError(f"No backup found in {root}")
    files = []
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
        log(f"  {n:>6}  {target}")
        files.append({"folder": rel.as_posix(), "messages": n, "file": str(target.resolve())})
    return {"ok": True, "command": "mbox", "out": str(dest.resolve()), "files": files}


def attachment_parts(msg, include_inline=False):
    """Yield (filename, bytes) for each attachment of a parsed message."""
    n = 0
    for part in msg.iter_attachments():
        n += 1
        inline = part.get_content_disposition() != "attachment"
        if inline and part.get("Content-ID") and not include_inline:
            continue
        if part.get_content_type() == "message/rfc822":
            inner = part.get_payload(0)
            data = inner.as_bytes()
            name = part.get_filename() or f"{safe_name(inner.get('Subject', ''), 60)}.eml"
        else:
            data = part.get_payload(decode=True)
            name = part.get_filename() or f"attachment-{n}"
        if data is None:
            continue
        yield safe_name(name, 120), data


def unique_path(path):
    if not path.exists():
        return path
    for i in range(2, 1000):
        cand = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not cand.exists():
            return cand
    return path


def extract_attachments(out, include_inline=False):
    out = Path(out)
    dest_root = out / "attachments"
    saved = skipped = errors = 0
    for rec in load_backup(out):
        eml = out / rec["path"]
        target = dest_root / Path(rec["path"]).relative_to("mail").with_suffix("")
        if target.exists():
            skipped += 1
            continue
        try:
            parts = list(attachment_parts(parse_message(eml), include_inline))
        except Exception as e:
            log(f"  ! {eml.name}: {e}")
            errors += 1
            continue
        if not parts:
            continue
        target.mkdir(parents=True, exist_ok=True)
        for name, data in parts:
            unique_path(target / name).write_bytes(data)
            saved += 1
    log(f"Saved {saved} attachments to {dest_root}"
        + (f" ({skipped} emails already extracted)" if skipped else ""))
    return {"ok": errors == 0, "command": "extract", "out": str(dest_root.resolve()),
            "attachments": saved, "emails_skipped": skipped, "errors": errors}


def backup_stats(out, top=10):
    recs = load_backup(out)
    by_folder = defaultdict(lambda: [0, 0])
    by_year = Counter()
    senders = Counter()
    for r in recs:
        by_folder[r["folder"]][0] += 1
        by_folder[r["folder"]][1] += r.get("size", 0)
        by_year[r["date"][:4] or "unknown"] += 1
        addr = parseaddr(r.get("from", ""))[1].lower()
        if addr:
            senders[addr] += 1
    total_size = sum(r.get("size", 0) for r in recs)
    result = {
        "ok": True, "command": "stats", "out": str(Path(out).resolve()),
        "messages": len(recs), "bytes": total_size,
        "folders": [{"folder": f, "messages": c, "bytes": s} for f, (c, s) in sorted(by_folder.items())],
        "by_year": dict(sorted(by_year.items())),
        "top_senders": [{"from": a, "messages": c} for a, c in senders.most_common(top)],
    }
    log(f"{len(recs)} messages, {human_size(total_size)}\n")
    log("Folders:")
    for f in result["folders"]:
        log(f"  {f['messages']:>7}  {human_size(f['bytes']):>9}  {f['folder']}")
    log("\nBy year:")
    for y, c in result["by_year"].items():
        log(f"  {y}  {c}")
    log("\nTop senders:")
    for s in result["top_senders"]:
        log(f"  {s['messages']:>6}  {s['from']}")
    return result


def verify_backup(out, fix=False):
    out = Path(out)
    recs = load_backup(out)
    problems = []
    for r in recs:
        p = out / r["path"]
        if not p.exists():
            problems.append({"path": r["path"], "problem": "missing"})
        elif p.stat().st_size != r.get("size", -1):
            problems.append({"path": r["path"], "problem": "size mismatch"})
    for p in problems:
        log(f"  {p['problem']:<14} {p['path']}")
    if problems and fix:
        bad = {p["path"] for p in problems}
        write_manifest(out, [r for r in recs if r["path"] not in bad])
        log(f"Removed {len(bad)} entries from the manifest; the next backup run re-downloads them.")
    elif problems:
        log(f"{len(problems)} problem(s). Run 'verify --fix' then 'backup' to re-download them.")
    else:
        log(f"All {len(recs)} messages present and intact.")
    return {"ok": not problems or fix, "command": "verify", "out": str(out.resolve()),
            "messages": len(recs), "problems": problems, "fixed": bool(problems and fix)}


# ---------------------------------------------------------------- html viewer

VIEWER_CSS = """
:root{--bg:#fff;--fg:#1b1b1b;--muted:#666;--line:#e3e3e3;--hover:#f3f6fb;--accent:#2b6cc4}
@media (prefers-color-scheme:dark){:root{--bg:#17181a;--fg:#e8e8e8;--muted:#9a9a9a;--line:#2e3033;--hover:#22262c;--accent:#7fb0ff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}
header{position:sticky;top:0;background:var(--bg);border-bottom:1px solid var(--line);padding:12px 16px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
h1{font-size:17px;margin:0 12px 0 0}input,select{font:inherit;padding:6px 9px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg)}
input{flex:1;min-width:180px}.muted{color:var(--muted)}a{color:var(--accent)}
table{width:100%;border-collapse:collapse}td,th{padding:7px 16px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{font-weight:600;white-space:nowrap}#dh{cursor:pointer;user-select:none}tr.r:hover{background:var(--hover);cursor:pointer}
td.d{white-space:nowrap;color:var(--muted)}td.f{color:var(--muted)}.more{display:block;margin:16px auto;padding:8px 16px}
.msg{max-width:1000px;margin:0 auto;padding:16px}.msg h2{margin:8px 0 12px;font-size:20px}
.hdr td{padding:3px 12px 3px 0;border:0}.hdr td:first-child{color:var(--muted)}
.att{margin:12px 0;display:flex;flex-wrap:wrap;gap:8px}.att a{border:1px solid var(--line);border-radius:6px;padding:4px 9px;text-decoration:none}
iframe{width:100%;height:75vh;border:1px solid var(--line);border-radius:8px;background:#fff}
@media (max-width:640px){td.f,th.f{display:none}td,th{padding:7px 8px}}
"""

VIEWER_JS = """
const rows=DATA,q=document.getElementById('q'),sel=document.getElementById('folder'),tb=document.getElementById('tb'),cnt=document.getElementById('count'),more=document.getElementById('more');
[...new Set(rows.map(r=>r.f))].sort().forEach(f=>sel.add(new Option(f,f)));
let shown=0,list=[],asc=false;
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function apply(){const terms=q.value.toLowerCase().split(/\\s+/).filter(Boolean),f=sel.value;
 list=rows.filter(r=>(!f||r.f===f)&&terms.every(t=>(r.s+' '+r.fr+' '+r.f).toLowerCase().includes(t)));
 list.sort((a,b)=>asc?a.d.localeCompare(b.d):b.d.localeCompare(a.d));tb.innerHTML='';shown=0;render();}
function render(){const next=list.slice(shown,shown+300);
 tb.insertAdjacentHTML('beforeend',next.map(r=>`<tr class=r data-p="${esc(r.p)}"><td class=d>${esc(r.d.slice(0,16).replace('T',' '))}</td><td>${esc(r.fr)}</td><td><a href="${esc(r.p)}">${esc(r.s||'(no subject)')}</a>${r.a?' <span class=muted>📎'+r.a+'</span>':''}</td><td class=f>${esc(r.f)}</td></tr>`).join(''));
 shown+=next.length;cnt.textContent=list.length+' of '+rows.length+' emails';more.style.display=shown<list.length?'block':'none';}
tb.addEventListener('click',e=>{const tr=e.target.closest('tr');if(tr&&e.target.tagName!=='A')location.href=tr.dataset.p;});
q.addEventListener('input',apply);sel.addEventListener('change',apply);more.addEventListener('click',render);
document.getElementById('dh').addEventListener('click',()=>{asc=!asc;apply();});
apply();
"""


def _body_html(msg):
    """The message body as HTML, with cid: images inlined as data: URIs."""
    body = msg.get_body(preferencelist=("html", "plain"))
    if body is None:
        return "<p><i>(no text body)</i></p>"
    content = body.get_content()
    if body.get_content_subtype() != "html":
        return ('<pre style="white-space:pre-wrap;font:14px/1.5 system-ui,sans-serif">'
                + html.escape(content) + "</pre>")
    cids = {}
    for part in msg.walk():
        cid = part.get("Content-ID")
        if cid and part.get_content_maintype() == "image":
            data = part.get_payload(decode=True) or b""
            cids[cid.strip("<> ")] = (f"data:{part.get_content_type()};base64,"
                                      + base64.b64encode(data).decode())
    if cids:
        content = re.sub(r"cid:([^\"'\s)>]+)", lambda m: cids.get(m.group(1), m.group(0)), content)
    return content


def build_viewer(out):
    """Write viewer/index.html (searchable list) and one page per message."""
    out = Path(out)
    site = out / "viewer"
    (site / "m").mkdir(parents=True, exist_ok=True)
    rows, errors = [], 0
    recs = sorted(load_backup(out), key=lambda r: r["date"], reverse=True)
    for i, rec in enumerate(recs, 1):
        page_id = hashlib.sha1(rec["key"].encode()).hexdigest()[:16]
        eml_rel = "../../" + quote(rec["path"])
        attachments = []
        try:
            msg = parse_message(out / rec["path"])
            body = _body_html(msg)
            files_dir = site / "files" / page_id
            for name, data in attachment_parts(msg):
                files_dir.mkdir(parents=True, exist_ok=True)
                target = files_dir / name
                if not target.exists():
                    target.write_bytes(data)
                attachments.append((name, len(data)))
            to, cc = msg.get("To", ""), msg.get("Cc", "")
        except Exception as e:
            errors += 1
            body = f"<p>Could not display this email ({html.escape(str(e))}). Open the original .eml instead.</p>"
            to = cc = ""
        doc = ('<!doctype html><meta charset="utf-8"><base target="_blank">'
               '<meta name="viewport" content="width=device-width,initial-scale=1">' + body)
        head = "".join(
            f"<tr><td>{k}</td><td>{html.escape(str(v))}</td></tr>"
            for k, v in (("From", rec["from"]), ("To", to), ("Cc", cc),
                         ("Date", rec["date"].replace("T", " ")[:19]), ("Folder", rec["folder"])) if v)
        atts = "".join(
            f'<a href="../files/{page_id}/{quote(n)}" download>📎 {html.escape(n)} '
            f'<span class="muted">{human_size(s)}</span></a>' for n, s in attachments)
        title = html.escape(rec["subject"] or "(no subject)")
        page = (f'<!doctype html><html lang="en"><meta charset="utf-8">'
                f'<meta name="viewport" content="width=device-width,initial-scale=1">'
                f"<title>{title}</title><style>{VIEWER_CSS}</style>"
                f'<div class="msg"><a href="../index.html">← All emails</a> · '
                f'<a href="{eml_rel}">Original .eml</a>'
                f"<h2>{title}</h2>"
                f'<table class="hdr">{head}</table><div class="att">{atts}</div>'
                f'<iframe sandbox="allow-popups allow-popups-to-escape-sandbox" '
                f'srcdoc="{html.escape(doc, quote=True)}"></iframe></div></html>')
        (site / "m" / f"{page_id}.html").write_text(page, encoding="utf-8")
        rows.append({"d": rec["date"], "f": rec["folder"], "fr": rec["from"], "s": rec["subject"],
                     "p": f"m/{page_id}.html", "a": len(attachments)})
        if i % 200 == 0:
            log(f"  {i}/{len(recs)} pages")

    data = json.dumps(rows, ensure_ascii=False).replace("</", "<\\/")
    index = (f'<!doctype html><html lang="en"><meta charset="utf-8">'
             f'<meta name="viewport" content="width=device-width,initial-scale=1">'
             f"<title>Email archive</title><style>{VIEWER_CSS}</style>"
             f'<header><h1>Email archive</h1><input id="q" type="search" placeholder="Search subject, sender, folder…" autofocus>'
             f'<select id="folder"><option value="">All folders</option></select><span id="count" class="muted"></span></header>'
             f'<table><thead><tr><th id="dh">Date ↕</th><th>From</th><th>Subject</th><th class="f">Folder</th></tr></thead>'
             f'<tbody id="tb"></tbody></table><button id="more" class="more">Show more</button>'
             f"<script>const DATA={data};{VIEWER_JS}</script></html>")
    (site / "index.html").write_text(index, encoding="utf-8")
    log(f"Viewer ready: {(site / 'index.html').resolve()}")
    return {"ok": errors == 0, "command": "html", "index": str((site / "index.html").resolve()),
            "messages": len(rows), "errors": errors}


# --------------------------------------------------------------------------
# guided mode (what the double-click launchers run)
# --------------------------------------------------------------------------

def server_from_input(text):
    """Accept 'mail.x.edu', 'https://mail.x.edu/owa/...' etc. and return the host."""
    text = text.strip()
    m = re.match(r"^(?:[a-z]+://)?([^/\s:]+)", text, re.I)
    return m.group(1) if m else text


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or default


def yes(prompt, default=True):
    value = input(f"{prompt} [{'Y/n' if default else 'y/N'}]: ").strip().lower()
    return default if not value else value.startswith("y")


def backup_defaults(**kw):
    opts = argparse.Namespace(
        email=None, user=None, server=os.environ.get("GUC_SERVER", DEFAULT_SERVER),
        auth="auto", autodiscover=False, insecure=False, imap_port=993, method=None,
        no_archive=False, out=None, folder=None, exclude=None, mail_only=False,
        since=None, until=None, batch=20, workers=4, dry_run=False)
    for k, v in kw.items():
        setattr(opts, k, v)
    return opts


def wizard():
    print("=" * 64)
    print("  University mailbox backup")
    print("  Saves every email (with attachments) from every folder to this computer.")
    print("  Your password is only sent to your university's mail server.")
    print("=" * 64 + "\n")

    email = ""
    while "@" not in email:
        email = ask("Your university email address")
    opts = backup_defaults(email=email, out=default_out(email))

    method = handle = None
    password = os.environ.get("GUC_PASSWORD")
    tries = 0
    while handle is None:
        password = password or getpass.getpass("Password (it will not show while you type): ")
        print("\nConnecting...")
        try:
            method, handle = connect(opts, password, say=print)
        except AuthFailed:
            password = None
            tries += 1
            print("\nThe server rejected that email/password.")
            if tries >= 3:
                if yes("Does your login use a different username than your email "
                       "(e.g. GUC\\first.last)?", default=False):
                    opts.user = ask("Username")
                elif not yes("Try again?"):
                    return EXIT_AUTH
        except CannotConnect as e:
            print(f"\nCould not reach the mail server {opts.server}.\n  ({e})")
            print("Open your webmail in a browser and copy the address from the address bar.")
            new = ask("Webmail address (leave empty to quit)")
            if not new:
                return EXIT_CONNECT
            opts.server = server_from_input(new)

    print(f"\nConnected ({method.upper()}). Backup folder: {Path(opts.out).resolve()}\n")
    if method == "ews":
        rows = ews_summary(handle, opts)
        for r in rows:
            print(f"  {r['items']:>7}  {r['folder']}")
        print(f"  {sum(r['items'] for r in rows):>7}  items in total\n")
    input("Press Enter to start downloading (you can close the window any time and "
          "run it again later to continue)...")

    result = run_backup(method, handle, opts)

    target = Path(opts.out).resolve()
    if yes("\nMake a browsable copy you can open in any web browser?"):
        target = Path(build_viewer(opts.out)["index"])
    if yes("Also create .mbox files (for importing into Thunderbird/Gmail)?", default=False):
        make_mbox(opts.out)

    print(f"\nYour emails are in: {Path(opts.out).resolve()}")
    print("index.csv lists everything; viewer/index.html lets you search and read them.")
    open_path(str(target))
    return EXIT_OK if not result["failed"] else EXIT_PARTIAL


# --------------------------------------------------------------------------
# command line (for scripting, agents and advanced use)
# --------------------------------------------------------------------------

def get_password(opts):
    if getattr(opts, "password_stdin", False):
        pw = sys.stdin.readline().rstrip("\r\n")
        if not pw:
            raise CliError("--password-stdin given but nothing was read from stdin")
        return pw
    if os.environ.get("GUC_PASSWORD"):
        return os.environ["GUC_PASSWORD"]
    if sys.stdin.isatty():
        return getpass.getpass(f"Password for {opts.user or opts.email}: ")
    raise CliError("No password: set GUC_PASSWORD or pipe it with --password-stdin")


def cli_connect(opts):
    try:
        return connect(opts, get_password(opts))
    except AuthFailed:
        raise CliError("Login rejected: wrong password, or try --user 'DOMAIN\\username'.", EXIT_AUTH)
    except CannotConnect as e:
        raise CliError(f"Could not connect: {e}. Check --server (the host in your webmail "
                       "address bar), or try --insecure / --autodiscover.", EXIT_CONNECT)


def cmd_probe(opts):
    method, handle = cli_connect(opts)
    log(f"Connected via {method.upper()}.")
    if method == "ews":
        rows = ews_summary(handle, opts)
        for r in rows:
            log(f"  {r['items']:>7}  {r['folder']}")
        log(f"  {sum(r['items'] for r in rows):>7}  items in total")
    else:
        rows = [{"folder": p, "items": None} for _, p in imap_folders(handle)]
        for r in rows:
            log(f"  {r['folder']}")
        handle.logout()
    return {"ok": True, "command": "probe", "method": method, "folders": rows,
            "items": sum(r["items"] or 0 for r in rows)}


def cmd_backup(opts):
    method, handle = cli_connect(opts)
    log(f"Connected via {method.upper()}. Saving to {Path(opts.out).resolve()}")
    result = run_backup(method, handle, opts)
    if opts.html and not opts.dry_run:
        result["viewer"] = build_viewer(opts.out)["index"]
    if opts.mbox and not opts.dry_run:
        make_mbox(opts.out)
    return result


def cmd_html(opts):
    result = build_viewer(resolve_out(opts))
    if opts.open:
        open_path(result["index"])
    return result


def build_parser():
    p = argparse.ArgumentParser(
        prog="guc_mail_backup.py",
        description="Back up an Exchange/OWA mailbox to local .eml files. "
                    "Run with no arguments for a guided, question-by-question mode.",
        epilog="Exit codes: 0 ok, 1 error, 2 some messages failed, 3 login rejected, "
               "4 cannot reach server, 130 interrupted.")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    output = argparse.ArgumentParser(add_help=False)
    output.add_argument("--json", action="store_true",
                        help="print a JSON result on stdout (progress stays on stderr)")
    output.add_argument("-q", "--quiet", action="store_true", help="no progress output")

    conn = argparse.ArgumentParser(add_help=False)
    g = conn.add_argument_group("connection")
    g.add_argument("--email", required=True, help="your mailbox address")
    g.add_argument("--user", help="login name if different from --email, e.g. 'GUC\\\\first.last'")
    g.add_argument("--password-stdin", action="store_true",
                   help="read the password from stdin (else $GUC_PASSWORD, else prompt)")
    g.add_argument("--server", default=os.environ.get("GUC_SERVER", DEFAULT_SERVER),
                   help=f"mail server host from your webmail URL (default {DEFAULT_SERVER}, or $GUC_SERVER)")
    g.add_argument("--method", choices=["ews", "imap"], help="force a method (default: try EWS, then IMAP)")
    g.add_argument("--auth", choices=["auto", "ntlm", "basic"], default="auto")
    g.add_argument("--autodiscover", action="store_true", help="find the EWS server automatically")
    g.add_argument("--insecure", action="store_true", help="skip TLS certificate verification")
    g.add_argument("--imap-port", type=int, default=993)
    g.add_argument("--no-archive", action="store_true", help="skip the Online Archive mailbox")

    local = argparse.ArgumentParser(add_help=False)
    local.add_argument("--out", help=f"backup folder (default: the only one under {DEFAULT_ROOT}/)")
    local.add_argument("--email", help=f"pick the backup of this address under {DEFAULT_ROOT}/")

    sp = sub.add_parser("probe", parents=[conn, output], help="check login and count items per folder")
    sp.set_defaults(func=cmd_probe)

    sp = sub.add_parser("backup", parents=[conn, output], help="download every message (resumable)")
    g = sp.add_argument_group("what to back up")
    g.add_argument("--folder", action="append", metavar="PATH",
                   help="only these folders and their subfolders; repeatable; wildcards ok, e.g. 'Inbox'")
    g.add_argument("--exclude", action="append", metavar="PATH",
                   help="skip these folders; repeatable, e.g. 'Junk E-Mail'")
    g.add_argument("--mail-only", action="store_true", help="skip calendar, contacts, tasks and notes folders")
    g.add_argument("--since", metavar="DATE", help="only messages received on/after DATE (YYYY-MM-DD)")
    g.add_argument("--until", metavar="DATE", help="only messages received on/before DATE (YYYY-MM-DD)")
    g.add_argument("--dry-run", action="store_true", help="show what would be downloaded, download nothing")
    g = sp.add_argument_group("output and speed")
    g.add_argument("--out", help=f"backup folder (default {DEFAULT_ROOT}/<email>)")
    g.add_argument("--workers", type=int, default=4, help="parallel EWS downloads (default 4)")
    g.add_argument("--batch", type=int, default=20, help="messages per EWS request (default 20)")
    g.add_argument("--html", action="store_true", help="also build the offline viewer afterwards")
    g.add_argument("--mbox", action="store_true", help="also build .mbox files afterwards")
    sp.set_defaults(func=cmd_backup)

    sp = sub.add_parser("html", parents=[local, output],
                        help="build a searchable offline viewer (viewer/index.html)")
    sp.add_argument("--open", action="store_true", help="open it in the browser when done")
    sp.set_defaults(func=cmd_html)

    sp = sub.add_parser("mbox", parents=[local, output], help="pack each folder into an .mbox file")
    sp.set_defaults(func=lambda o: make_mbox(resolve_out(o)))

    sp = sub.add_parser("extract", parents=[local, output], help="save all attachments as normal files")
    sp.add_argument("--include-inline", action="store_true", help="also save inline images (logos, signatures)")
    sp.set_defaults(func=lambda o: extract_attachments(resolve_out(o), o.include_inline))

    sp = sub.add_parser("stats", parents=[local, output], help="counts and sizes per folder, year and sender")
    sp.add_argument("--top", type=int, default=10, help="how many top senders to list")
    sp.set_defaults(func=lambda o: backup_stats(resolve_out(o), o.top))

    sp = sub.add_parser("verify", parents=[local, output], help="check every saved file is present and intact")
    sp.add_argument("--fix", action="store_true", help="forget broken entries so the next backup re-downloads them")
    sp.set_defaults(func=lambda o: verify_backup(resolve_out(o), o.fix))
    return p


def main(argv=None):
    global QUIET
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        # no arguments (e.g. double-clicked): guided mode
        try:
            code = wizard()
        except (KeyboardInterrupt, EOFError):
            print("\nStopped. Progress is saved; run it again to continue.")
            code = EXIT_INTERRUPTED
        except CliError as e:
            print(f"\n{e}")
            code = e.code
        if sys.stdin.isatty():
            try:
                input("\nPress Enter to close.")
            except (KeyboardInterrupt, EOFError):
                pass
        return code

    opts = build_parser().parse_args(argv)
    QUIET = opts.quiet
    try:
        if opts.cmd in ("probe", "backup"):
            opts = backup_defaults(**vars(opts))
            opts.since = parse_date(opts.since)
            opts.until = parse_date(opts.until, end=True)
            opts.out = opts.out or default_out(opts.email)
        result = opts.func(opts)
        code = EXIT_OK if result.get("ok", True) else EXIT_PARTIAL
    except CliError as e:
        result, code = {"ok": False, "command": opts.cmd, "error": str(e)}, e.code
        print(f"Error: {e}", file=sys.stderr)
    except KeyboardInterrupt:
        result, code = {"ok": False, "command": opts.cmd, "error": "interrupted"}, EXIT_INTERRUPTED
        print("\nInterrupted - progress is saved, re-run the same command to continue.", file=sys.stderr)
    result["exit_code"] = code
    if opts.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return code


if __name__ == "__main__":
    sys.exit(main())
