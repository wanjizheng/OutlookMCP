import json
import os
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from outlook_mcp import attachments, bridge, cli, rag

ROOT = Path(__file__).resolve().parents[1]


def pdf(path, encrypted=False, blank=False):
    from pypdf import PdfWriter
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

    writer = PdfWriter()
    page = writer.add_blank_page(width=200, height=200)
    if not blank:
        font = DictionaryObject(
            {
                NameObject("/Type"): NameObject("/Font"),
                NameObject("/Subtype"): NameObject("/Type1"),
                NameObject("/BaseFont"): NameObject("/Helvetica"),
            }
        )
        page[NameObject("/Resources")] = DictionaryObject(
            {
                NameObject("/Font"): DictionaryObject(
                    {NameObject("/F1"): writer._add_object(font)}
                )
            }
        )
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 12 Tf 20 100 Td (Attachment needle research) Tj ET")
        page[NameObject("/Contents")] = writer._add_object(stream)
    if encrypted:
        writer.encrypt("secret")
    writer.write(path)


class AttachmentTests(unittest.TestCase):
    def test_default_sync_indexes_attachments_when_body_unchanged(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            row = {
                "EntryID": "one",
                "Body": "same",
                "Attachments": [{"FileName": "report.txt"}],
            }
            rag.initialize()
            with rag.connect(False) as c:
                rag.ingest([row], c)
            records = [
                {
                    "index": 1,
                    "name": "report.txt",
                    "status": "ok",
                    "text": "syncattachmentneedle",
                }
            ]
            with (
                mock.patch.object(
                    bridge,
                    "list_mail",
                    return_value={"items": [row], "next_cursor": "next"},
                ),
                mock.patch.object(attachments, "collect", return_value=records),
            ):
                report = cli._sync(["Inbox"], 0, 1)
            self.assertEqual(report["Inbox"]["attachments"]["ok"], 1)
            self.assertTrue(rag.search("syncattachmentneedle", mode="bm25")["items"])
            with rag.connect() as c:
                self.assertEqual(
                    c.execute(
                        "SELECT value FROM meta WHERE key='windows_cursor:inbox'"
                    ).fetchone()[0],
                    "next",
                )

    def test_supported_formats_extract_locally(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            root = Path(tmp)
            for ext in ("txt", "md", "csv"):
                p = root / ("file." + ext)
                p.write_text("附件 needle research", encoding="utf-8-sig")
                self.assertIn("附件 needle", attachments.extract_file(p)["text"])
            parts = {
                "docx": {
                    "word/document.xml": "<document><p><t>Word needle</t></p></document>"
                },
                "pptx": {"ppt/slides/slide1.xml": "<slide><t>Slide needle</t></slide>"},
                "xlsx": {
                    "xl/sharedStrings.xml": "<sst><si><t>Sheet needle</t></si></sst>",
                    "xl/worksheets/sheet1.xml": '<worksheet><c t="s"><v>0</v></c><c><v>42</v></c></worksheet>',
                },
            }
            for ext, files in parts.items():
                p = root / ("file." + ext)
                with zipfile.ZipFile(p, "w") as z:
                    for name, text in files.items():
                        z.writestr(name, text)
                self.assertIn("needle", attachments.extract_file(p)["text"])
            p = root / "file.pdf"
            pdf(p)
            self.assertIn("Attachment needle", attachments.extract_worker(p)["text"])

    def test_encrypted_scanned_corrupt_and_timeout_are_distinct(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            p = Path(tmp) / "file.pdf"
            pdf(p, encrypted=True)
            self.assertEqual(attachments.extract_file(p)["status"], "encrypted")
            pdf(p, blank=True)
            self.assertEqual(attachments.extract_file(p)["status"], "no_text")
            p.write_bytes(b"invalid PDF")
            self.assertEqual(attachments.extract_worker(p)["status"], "parse_error")
            with mock.patch(
                "subprocess.run", side_effect=subprocess.TimeoutExpired("worker", 20)
            ):
                self.assertEqual(attachments.extract_worker(p)["status"], "timeout")

    def test_size_text_and_expansion_limits(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            p = Path(tmp) / "file.txt"
            p.write_text("x" * 20000)
            out = attachments.extract_file(p)
            self.assertTrue(out["truncated"])
            self.assertEqual(len(out["text"]), 12000)
            with p.open("wb") as f:
                f.truncate(attachments.MAX_FILE + 1)
            self.assertEqual(attachments.extract_file(p)["status"], "size_limit")
            p = Path(tmp) / "file.docx"
            with zipfile.ZipFile(p, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.writestr("word/document.xml", "x" * (21 * 1024 * 1024))
            self.assertEqual(
                attachments.extract_file(p)["status"], "expanded_size_limit"
            )
            from pypdf import PdfWriter

            p = Path(tmp) / "many-pages.pdf"
            writer = PdfWriter()
            for _ in range(201):
                writer.add_blank_page(width=20, height=20)
            writer.write(p)
            self.assertTrue(attachments.extract_file(p)["truncated"])

    def test_xml_entities_never_expand(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            p = Path(tmp) / "file.docx"
            with zipfile.ZipFile(p, "w") as z:
                z.writestr(
                    "word/document.xml",
                    '<!DOCTYPE d [<!ENTITY x SYSTEM "file:///private">]><document><t>&x;</t></document>',
                )
            self.assertEqual(attachments.extract_worker(p)["status"], "parse_error")

    def test_unchanged_body_attachment_backfill_update_and_optout(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            rag.initialize()
            row = {"EntryID": "one", "Subject": "Subject", "Body": "body"}
            with rag.connect(False) as c:
                rag.ingest([row], c)
                row["_AttachmentRecords"] = [
                    {
                        "index": 1,
                        "name": "report.txt",
                        "status": "ok",
                        "text": "uniqueneedle",
                        "sha256": "a",
                    }
                ]
                rag.ingest([row], c)
            self.assertEqual(
                rag.search("uniqueneedle", mode="bm25")["items"][0]["entry_id"], "one"
            )
            with rag.connect(False) as c:
                cid = c.execute("SELECT chunk_id FROM email_chunks").fetchone()[0]
                rag.ingest([row], c)
                self.assertEqual(
                    c.execute("SELECT chunk_id FROM email_chunks").fetchone()[0], cid
                )
                row.pop("_AttachmentRecords")
                row["Body"] = "updated body"
                rag.ingest([row], c)
            self.assertTrue(rag.search("uniqueneedle", mode="bm25")["items"])
            row["_AttachmentRecords"] = []
            with rag.connect(False) as c:
                rag.ingest([row], c)
            self.assertFalse(rag.search("uniqueneedle", mode="bm25")["items"])

    def test_failed_extraction_never_commits_page_or_cursor(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            page = {
                "items": [{"EntryID": "one", "Body": "body", "Attachments": [{}]}],
                "next_cursor": "next",
            }
            with (
                mock.patch.object(bridge, "list_mail", return_value=page),
                mock.patch.object(
                    attachments, "collect", return_value=[{"status": "timeout"}]
                ),
            ):
                with self.assertRaisesRegex(RuntimeError, "not committed"):
                    cli._sync(["Inbox"], 0, 1)
            with rag.connect() as c:
                self.assertEqual(
                    c.execute("SELECT count(*) FROM emails").fetchone()[0], 0
                )
                self.assertIsNone(
                    c.execute(
                        "SELECT value FROM meta WHERE key='windows_cursor:inbox'"
                    ).fetchone()
                )

    @unittest.skipUnless(os.name == "nt", "Windows ACL/export behavior")
    def test_export_names_limits_and_no_mailbox_writes(self):
        subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-File",
                str(ROOT / "tests/attachments.ps1"),
            ],
            check=True,
        )

    @unittest.skipUnless(os.name == "nt", "Windows temporary ACL")
    def test_raw_files_cleaned_even_on_parse_failure(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:

            def export(args, **kwargs):
                req = json.loads(kwargs["input"])
                p = Path(req["directory"]) / "001.txt"
                p.write_text("fixture")
                return subprocess.CompletedProcess(
                    args,
                    0,
                    json.dumps(
                        [
                            {
                                "index": 1,
                                "name": "../../evil.txt",
                                "status": "exported",
                                "path": str(p),
                            }
                        ]
                    ).encode(),
                )

            with (
                mock.patch("subprocess.run", side_effect=export),
                mock.patch.object(
                    attachments,
                    "extract_worker",
                    return_value={"status": "parse_error", "text": ""},
                ),
            ):
                records = attachments.collect(
                    {"EntryID": "one", "Attachments": [{}]}, tmp
                )
            self.assertEqual(records[0]["status"], "parse_error")
            self.assertEqual(list((Path(tmp) / "attachment-tmp").iterdir()), [])


if __name__ == "__main__":
    unittest.main()
