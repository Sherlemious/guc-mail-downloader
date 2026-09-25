"""Offline tests: fake EWS/IMAP servers, no network or real mailbox needed.

Run:  python -m unittest discover -s tests -v
"""

import contextlib
import io
import json
import os
import sys
import tempfile
import types
import unittest
import unittest.mock
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import guc_mail_backup as g  # noqa: E402


def make_email(subject, sender="Prof <prof@guc.edu.eg>", date="Mon, 21 Sep 2026 15:25:00 +0200",
               attachment=None, inline_image=False):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = "me@student.guc.edu.eg"
    msg["Date"] = date
    msg.set_content("plain body")
    if inline_image:
        msg.add_alternative('<p>hello <img src="cid:logo123"></p>', subtype="html")
        msg.get_payload()[1].add_related(b"PNGDATA", "image", "png", cid="<logo123>")
    else:
        msg.add_alternative("<p>hello <b>html</b></p>", subtype="html")
    if attachment:
        msg.add_attachment(attachment[1], maintype="application", subtype="pdf", filename=attachment[0])
    return msg.as_bytes()


# ------------------------------------------------------------------ fakes

class FakeItem:
    def __init__(self, id, received):
        self.id, self.changekey, self.datetime_received = id, "ck", received


class FakeFolder:
    def __init__(self, absolute, items, folder_class="IPF.Note", children=()):
        self.absolute, self._items, self.folder_class = absolute, items, folder_class
        self.total_count = len(items)
        self.children = list(children)

    def all(self):
        items = self._items
        return types.SimpleNamespace(only=lambda *a: list(items))

    def walk(self):
        for c in self.children:
            yield c
            yield from c.walk()


class FakeAccount:
    """Messages keyed by id; ids in `broken` fail every time, ids in `flaky` fail in batches only."""

    def __init__(self, folders, mime, broken=(), flaky=()):
        self.msg_folder_root = FakeFolder("/root/Top", [], children=folders)
        self.mime, self.broken, self.flaky = mime, set(broken), set(flaky)
        self.primary_smtp_address = "me@x"

    @property
    def archive_msg_folder_root(self):
        raise RuntimeError("no archive")

    def fetch(self, ids, only_fields):
        for i in ids:
            if i.id in self.broken or (i.id in self.flaky and len(ids) > 1):
                yield RuntimeError(f"cannot fetch {i.id}")
            else:
                yield types.SimpleNamespace(mime_content=self.mime[i.id])


def dt(day):
    return datetime(2026, 9, day, 12, tzinfo=timezone.utc)


def fake_mailbox(**kw):
    inbox = [FakeItem(f"in{d}", dt(d)) for d in (1, 2, 3)]
    sent = [FakeItem("sent1", dt(4))]
    mime = {i.id: make_email(f"Subject {i.id}") for i in inbox + sent}
    mime["in1"] = make_email("Has attachment", attachment=("report.pdf", b"%PDF-1.4 data"))
    mime["in2"] = make_email("Has logo", inline_image=True)
    folders = [
        FakeFolder("/root/Top/Inbox", inbox, children=[FakeFolder("/root/Top/Inbox/Memoirs", [])]),
        FakeFolder("/root/Top/Sent Items", sent),
        FakeFolder("/root/Top/Calendar", [FakeItem("cal1", dt(5))], folder_class="IPF.Appointment"),
    ]
    mime["cal1"] = make_email("Meeting")
    return FakeAccount(folders, mime, **kw)


def opts_for(out, **kw):
    return g.backup_defaults(email="me@x", out=str(out), **kw)


def quiet():
    return contextlib.redirect_stderr(io.StringIO())


class TempDirTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.out = self.dir / "backup"

    def tearDown(self):
        self._tmp.cleanup()

    def backup(self, account=None, **kw):
        with quiet():
            return g.ews_backup(account or fake_mailbox(), opts_for(self.out, **kw))


# ------------------------------------------------------------------ tests

class HelperTests(unittest.TestCase):
    def test_safe_name(self):
        self.assertEqual(g.safe_name('a/b:c*?"<>|d'), "a_b_c_d")
        self.assertEqual(g.safe_name("   "), "no-subject")
        self.assertEqual(len(g.safe_name("x" * 500)), 80)

    def test_parse_date_until_is_inclusive(self):
        until = g.parse_date("2026-09-02", end=True)
        opts = types.SimpleNamespace(since=g.parse_date("2026-09-02"), until=until)
        self.assertTrue(g.in_range(datetime(2026, 9, 2, 23, 0).astimezone(), opts))
        self.assertFalse(g.in_range(datetime(2026, 9, 3, 0, 1).astimezone(), opts))
        self.assertFalse(g.in_range(datetime(2026, 9, 1, 23, 0).astimezone(), opts))
        with self.assertRaises(g.CliError):
            g.parse_date("yesterday")

    def test_folder_selection(self):
        o = g.backup_defaults(folder=["Inbox"], exclude=["Inbox/Spam*"])
        self.assertTrue(g.folder_selected("Inbox", o))
        self.assertTrue(g.folder_selected("inbox/Memoirs", o))
        self.assertFalse(g.folder_selected("Inbox2", o))
        self.assertFalse(g.folder_selected("Inbox/Spammy", o))
        o = g.backup_defaults(mail_only=True)
        self.assertFalse(g.folder_selected("Calendar", o, "IPF.Appointment"))
        self.assertTrue(g.folder_selected("Inbox", o, "IPF.Note"))

    def test_imap_folder_decoding(self):
        self.assertEqual(g.imap_decode_folder("&BCcENQRABD0EPgQyBDgEOgQ4-/Sub&-x"), "Черновики/Sub&x")

    def test_server_from_input(self):
        self.assertEqual(g.server_from_input("https://mail.guc.edu.eg/owa/#path=/mail"), "mail.guc.edu.eg")
        self.assertEqual(g.server_from_input("mail.x.edu"), "mail.x.edu")


class EwsBackupTests(TempDirTest):
    def test_full_backup_then_resume(self):
        r = self.backup(workers=3, batch=2)
        self.assertEqual((r["new"], r["had"], r["failed"]), (5, 0, 0))
        self.assertTrue(r["ok"])
        files = sorted(p.relative_to(self.out).as_posix() for p in self.out.rglob("*.eml"))
        self.assertEqual(len(files), 5)
        self.assertTrue(any(f.startswith("mail/Inbox/2026-09-21_") and "Has attachment" in f for f in files))
        self.assertTrue((self.out / "index.csv").exists())

        r = self.backup()
        self.assertEqual((r["new"], r["had"]), (0, 5))

    def test_failed_items_are_retried_next_run(self):
        acc = fake_mailbox(broken={"in3"}, flaky={"in2"})
        r = self.backup(acc, batch=10)
        self.assertEqual((r["new"], r["failed"]), (4, 1))  # flaky one recovered by single retry
        self.assertFalse(r["ok"])
        acc.broken.clear()
        r = self.backup(acc)
        self.assertEqual((r["new"], r["had"], r["failed"]), (1, 4, 0))

    def test_filters_and_dry_run(self):
        r = self.backup(since=g.parse_date("2026-09-02"), until=g.parse_date("2026-09-03", end=True),
                        exclude=["Sent Items"], mail_only=True, dry_run=True)
        self.assertEqual(r["would_download"], 2)
        self.assertFalse(list(self.out.rglob("*.eml")))
        r = self.backup(folder=["Sent Items"])
        self.assertEqual(r["new"], 1)


class ImapBackupTests(TempDirTest):
    def test_imap_backup(self):
        mails = {b"1": make_email("imap one"), b"2": make_email("imap two")}

        class Conn:
            def list(self):
                return "OK", [b'(\\HasNoChildren) "/" INBOX', b'(\\Noselect) "/" Root']

            def select(self, name, readonly):
                return "OK", [b"2"]

            def response(self, what):
                return what, [b"77"]

            def uid(self, cmd, *args):
                if cmd == "search":
                    self.criteria = args[1:]
                    return "OK", [b"1 2"]
                return "OK", [(b"1 (BODY[] {10}", mails[args[0]]), b")"]

            def logout(self):
                pass

        conn = Conn()
        with quiet():
            r = g.imap_backup(conn, opts_for(self.out, since=g.parse_date("2026-01-05")))
        self.assertEqual(r["new"], 2)
        self.assertEqual(conn.criteria, ("SINCE", "05-Jan-2026"))
        with quiet():
            r = g.imap_backup(Conn(), opts_for(self.out))
        self.assertEqual((r["new"], r["had"]), (0, 2))


class ConnectTests(unittest.TestCase):
    def run_connect(self, ews_errors, imap_error):
        calls = []

        def fake_ews(opts, pw, auth=None, autodiscover=False):
            calls.append(auth or ("autodiscover" if autodiscover else "auto"))
            err = ews_errors.get(calls[-1])
            if err:
                raise err
            return "ACCOUNT"

        def fake_imap(opts, pw):
            if imap_error:
                raise imap_error
            return "IMAP"

        orig = g.ews_try, g.imap_connect
        g.ews_try, g.imap_connect = fake_ews, fake_imap
        try:
            return g.connect(g.backup_defaults(email="a@b"), "pw", say=lambda *a: None), calls
        finally:
            g.ews_try, g.imap_connect = orig

    def test_falls_back_to_ntlm(self):
        result, calls = self.run_connect({"auto": ConnectionError("x")}, None)
        self.assertEqual(result, ("ews", "ACCOUNT"))
        self.assertEqual(calls, ["auto", "ntlm"])

    def test_falls_back_to_imap(self):
        err = ConnectionError("down")
        result, _ = self.run_connect({"auto": err, "ntlm": err, "basic": err, "autodiscover": err}, None)
        self.assertEqual(result, ("imap", "IMAP"))

    def test_wrong_password(self):
        err = RuntimeError("401 Unauthorized")
        with self.assertRaises(g.AuthFailed):
            self.run_connect({"auto": err, "ntlm": err, "basic": err, "autodiscover": ConnectionError()},
                             ConnectionRefusedError())

    def test_unreachable(self):
        err = ConnectionError("dns")
        with self.assertRaises(g.CannotConnect):
            self.run_connect({"auto": err, "ntlm": err, "basic": err, "autodiscover": err}, err)


class LocalCommandTests(TempDirTest):
    def setUp(self):
        super().setUp()
        self.backup()

    def cli(self, *args):
        out = io.StringIO()
        with contextlib.redirect_stdout(out), quiet():
            code = g.main(list(args) + ["--out", str(self.out), "--json"])
        return code, json.loads(out.getvalue())

    def test_html_viewer(self):
        code, r = self.cli("html")
        self.assertEqual(code, 0)
        self.assertEqual(r["messages"], 5)
        index = Path(r["index"]).read_text(encoding="utf-8")
        self.assertIn("Has attachment", index)
        pages = list((self.out / "viewer" / "m").glob("*.html"))
        self.assertEqual(len(pages), 5)
        logo_page = next(p for p in pages if "Has logo" in p.read_text(encoding="utf-8"))
        self.assertIn("data:image/png;base64,", logo_page.read_text(encoding="utf-8"))
        self.assertTrue(list((self.out / "viewer" / "files").rglob("report.pdf")))

    def test_extract(self):
        code, r = self.cli("extract")
        self.assertEqual((code, r["attachments"]), (0, 1))
        pdf = next(self.out.rglob("attachments/**/report.pdf"))
        self.assertEqual(pdf.read_bytes(), b"%PDF-1.4 data")
        code, r = self.cli("extract")
        self.assertEqual(r["attachments"], 0)

    def test_stats(self):
        code, r = self.cli("stats")
        self.assertEqual(r["messages"], 5)
        self.assertEqual(r["top_senders"][0], {"from": "prof@guc.edu.eg", "messages": 5})
        self.assertEqual({f["folder"] for f in r["folders"]}, {"Inbox", "Sent Items", "Calendar"})

    def test_verify_and_fix(self):
        victim = next(self.out.rglob("*.eml"))
        victim.unlink()
        code, r = self.cli("verify")
        self.assertEqual((code, len(r["problems"])), (g.EXIT_PARTIAL, 1))
        code, r = self.cli("verify", "--fix")
        self.assertEqual(code, 0)
        r = self.backup()
        self.assertEqual((r["new"], r["had"]), (1, 4))

    def test_mbox(self):
        code, r = self.cli("mbox")
        self.assertEqual(code, 0)
        self.assertEqual(sum(f["messages"] for f in r["files"]), 5)


class CliTests(TempDirTest):
    def test_missing_password_is_an_error_not_a_hang(self):
        out = io.StringIO()
        env = {k: v for k, v in os.environ.items() if k != "GUC_PASSWORD"}
        with unittest.mock.patch.dict(os.environ, env, clear=True), \
                unittest.mock.patch.object(sys, "stdin", io.StringIO("")), \
                contextlib.redirect_stdout(out), quiet():
            code = g.main(["backup", "--email", "a@b", "--out", str(self.out), "--json"])
        self.assertEqual(code, g.EXIT_ERROR)
        self.assertIn("GUC_PASSWORD", json.loads(out.getvalue())["error"])

    def test_auth_failure_exit_code(self):
        def boom(opts, pw, say=None):
            raise g.AuthFailed("401")
        with unittest.mock.patch.object(g, "connect", boom), \
                unittest.mock.patch.dict(os.environ, {"GUC_PASSWORD": "x"}), quiet():
            code = g.main(["probe", "--email", "a@b"])
        self.assertEqual(code, g.EXIT_AUTH)

    def test_backup_command_end_to_end(self):
        out = io.StringIO()
        with unittest.mock.patch.object(g, "connect", lambda o, p, say=None: ("ews", fake_mailbox())), \
                unittest.mock.patch.dict(os.environ, {"GUC_PASSWORD": "x"}), \
                contextlib.redirect_stdout(out), quiet():
            code = g.main(["backup", "--email", "a@b", "--out", str(self.out), "--json",
                           "--exclude", "Calendar", "--html"])
        r = json.loads(out.getvalue())
        self.assertEqual((code, r["new"], r["exit_code"]), (0, 4, 0))
        self.assertTrue(Path(r["viewer"]).exists())

    def test_resolve_out_without_backups(self):
        cwd = os.getcwd()
        os.chdir(self.dir)
        try:
            with self.assertRaises(g.CliError):
                g.resolve_out(types.SimpleNamespace(out=None, email=None))
        finally:
            os.chdir(cwd)


if __name__ == "__main__":
    unittest.main()
