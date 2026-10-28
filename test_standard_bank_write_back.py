"""Business-rule checks for the matter gate before file-note retrieval."""

import io
import csv
import datetime as dt
import ftplib
import json
import os
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

import standard_bank_write_back as write_back


class MatterEligibilityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.matter = {
            "recordid": "565398",
            "archivestatus": "0",
            "clientid": "155128",
            "mattertypeid": "4",
            "theirref": "00335458874",
        }

    def test_sql_predicate_values(self) -> None:
        self.assertEqual(
            write_back.ALLOWED_CLIENT_IDS,
            {"150307", "334695", "155128", "334565", "209250", "334568",
             "283850", "334567", "267742", "334569"},
        )
        self.assertEqual(write_back.matter_eligibility_reasons(self.matter), [])
        self.assertEqual(write_back.matter_eligibility_reasons({**self.matter, "archivestatus": None}), [])
        self.assertEqual(write_back.matter_eligibility_reasons({**self.matter, "archivestatus": 0}), [])

        for changes in (
            {"archivestatus": "1"},
            {"archivestatus": ""},
            {"clientid": "999999"},
            {"mattertypeid": "3"},
        ):
            with self.subTest(changes=changes):
                self.assertTrue(write_back.matter_eligibility_reasons({**self.matter, **changes}))
        without_status = {key: value for key, value in self.matter.items() if key != "archivestatus"}
        self.assertTrue(write_back.matter_eligibility_reasons(without_status))

    def test_excluded_matter_does_not_request_file_notes(self) -> None:
        matter = {**self.matter, "archivestatus": "1"}
        output = io.StringIO()
        with (
            patch.dict(os.environ, {"LEGALSUITE_API_KEY": "test-key"}),
            patch.object(write_back, "load_env_file"),
            patch.object(write_back, "fetch_matter", return_value=matter) as fetch_matter,
            patch.object(write_back, "fetch_filenotes") as fetch_filenotes,
            redirect_stdout(output),
        ):
            exit_code = write_back.main(["565398"])

        self.assertEqual(exit_code, 0)
        fetch_matter.assert_called_once()
        fetch_filenotes.assert_not_called()
        result = json.loads(output.getvalue())
        self.assertFalse(result["matter_eligibility"]["eligible"])
        self.assertTrue(result["matter_eligibility"]["reasons"])
        self.assertEqual(result["matching_count"], 0)
        self.assertEqual(result["data"], [])

    def test_eligible_matter_is_fetched_before_file_notes(self) -> None:
        calls: list[str] = []

        def fetch_matter(*_args):
            calls.append("matter")
            return self.matter

        def fetch_filenotes(*_args):
            calls.append("filenote")
            return [{"recordid": "1", "matterid": "565398", "stagecode": "MT300",
                     "description": "ACKNOWLEDGEMENT", "formatteddatetime": "29 Jul 2026 14:37:20"}]

        output = io.StringIO()
        with (
            patch.dict(os.environ, {"LEGALSUITE_API_KEY": "test-key"}),
            patch.object(write_back, "load_env_file"),
            patch.object(write_back, "fetch_matter", side_effect=fetch_matter),
            patch.object(write_back, "fetch_filenotes", side_effect=fetch_filenotes),
            redirect_stdout(output),
        ):
            exit_code = write_back.main(["565398"])

        self.assertEqual(exit_code, 0)
        self.assertEqual(calls, ["matter", "filenote"])
        result = json.loads(output.getvalue())
        self.assertTrue(result["matter_eligibility"]["eligible"])
        self.assertEqual(result["matching_count"], 1)
        self.assertEqual(result["data"][0]["message_data"]["account_number"], "00335458874")


class HandoverBatchTests(unittest.TestCase):
    def _workbook(self, directory: Path) -> Path:
        path = directory / "Standard_Bank_Panel_L_Handover_20260814.xlsx"
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        sheet = (
            f'<worksheet xmlns="{namespace}"><sheetData>'
            '<row r="1"><c r="C1" t="inlineStr"><is><t>Client Code</t></is></c>'
            '<c r="D1" t="inlineStr"><is><t>Reference</t></is></c>'
            '<c r="J1" t="inlineStr"><is><t>Town/City</t></is></c>'
            '<c r="P1" t="inlineStr"><is><t>Town/City</t></is></c></row>'
            '<row r="2"><c r="C2" t="inlineStr"><is><t>STA387</t></is></c>'
            '<c r="D2" t="inlineStr"><is><t>\'00012345678</t></is></c>'
            '<c r="J2" t="inlineStr"><is><t>TEST CITY</t></is></c>'
            '<c r="P2" t="inlineStr"><is><t>OTHER CITY</t></is></c></row>'
            '</sheetData></worksheet>'
        )
        with zipfile.ZipFile(path, "w") as workbook:
            workbook.writestr("xl/worksheets/sheet1.xml", sheet)
        return path

    def _matter(self) -> dict:
        return {
            "recordid": "600001", "archivestatus": "0", "clientid": "150307",
            "mattertypeid": "4", "theirref": "00012345678", "fileref": "STA387/0001",
            "formatteddateinstructed": "14 Aug 2026",
            "internalcomment": "Imported from handover file on 14 August 2026",
        }

    def _note(self) -> dict:
        return {
            "recordid": "700001", "matterid": "600001", "stagecode": "MT300",
            "description": "ACKNOWLEDGEMENT", "formatteddatetime": "14 Aug 2026 10:30:00",
        }

    def test_previous_month_and_handover_reference(self) -> None:
        self.assertEqual(write_back._month_to_scan("previous", dt.date(2026, 9, 21)), (2026, 8))
        with tempfile.TemporaryDirectory() as directory:
            path = self._workbook(Path(directory))
            accounts = write_back.read_handover_accounts(path)
        self.assertEqual(len(accounts), 1)
        self.assertEqual(accounts[0].client_id, "150307")
        self.assertEqual(accounts[0].reference, "00012345678")
        self.assertEqual(accounts[0].physical_city, "TEST CITY")

    def test_ftp_discovery_and_older_pending_file(self) -> None:
        class FakeFTP:
            def nlst(self, directory):
                return {
                    "SBSA/Panel L/Handover_APT_LSW/Aug 2026": [
                        "Standard_Bank_Panel_L_Handover_20260814.xlsx",
                        "Standard_Bank_Panel_L_Handover_20260820.xlsx",
                        "Standard_Bank_Panel_L_Handover_20260820 (copy).xlsx",
                    ],
                    "SBSA/Debt Review/Debt_Review_Handover_APT_LWS": [
                        "Standard_Bank_Panel_L_Handover_20260819_DR.xlsx",
                    ],
                }.get(directory, [])

        sources = write_back.discover_ftp_handover_sources(FakeFTP(), 2026, 8)
        self.assertEqual(len(sources), 3)
        self.assertEqual(len(write_back.newest_sources(sources)), 2)
        pending = write_back._pending_ftp_sources({
            "ftp:SBSA/Panel L/Handover_APT_LSW/Jul 2026/Standard_Bank_Panel_L_Handover_20260731.xlsx": {
                "complete": False,
            },
            sources[0].identity: {"complete": True},
        })
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0].handover_date, dt.date(2026, 7, 31))

        class EmptyFTP:
            def nlst(self, *_args):
                raise ftplib.error_perm("550 no files")

            def pwd(self):
                return "/"

            def cwd(self, _directory):
                return None

        self.assertEqual(write_back._ftp_names(EmptyFTP(), "empty/handover"), [])

    def test_completed_file_is_skipped_on_repeat_run(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = self._workbook(root)
            output_dir = root / "csv"
            lookup = [self._matter()]
            notes = [self._note(), {**self._note(), "recordid": "700002", "stagecode": "MT326",
                                    "description": "UNMAPPED STAGE"}]
            with (
                patch.dict(os.environ, {"LEGALSUITE_API_KEY": "test-key"}),
                patch.object(write_back, "load_env_file"),
                patch.object(write_back, "fetch_matters_by_reference", return_value=lookup) as fetch_matters,
                patch.object(write_back, "fetch_filenotes", return_value=notes) as fetch_notes,
                redirect_stdout(io.StringIO()),
            ):
                args = ["--handover-file", str(workbook), "--csv-dir", str(output_dir)]
                self.assertEqual(write_back.main(args), 0)
                self.assertEqual(write_back.main(args), 0)
            self.assertEqual(fetch_matters.call_count, 1)
            self.assertEqual(fetch_notes.call_count, 1)
            csv_dir = output_dir / "2026-08"
            csv_path = next(csv_dir.glob("300_Acknowledgement_*.csv"))
            with csv_path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(len(rows), 1)
            self.assertEqual(
                list(rows[0]),
                [*write_back.BANK_ENVELOPE_HEADERS, "Account Number", "CACS Supplier"],
            )
            self.assertEqual(rows[0]["suite"], "316")
            self.assertEqual(rows[0]["messageID"], "300")
            self.assertEqual(rows[0]["date"], "20260814")
            self.assertEqual(rows[0]["Account Number"], "00012345678")
            self.assertEqual(rows[0]["CACS Supplier"], "")
            self.assertEqual(list(csv_dir.glob("326_*.csv")), [])
            state = json.loads((output_dir / "processed_handover_files.json").read_text())
            self.assertEqual(len(next(iter(state.values()))["csv_paths"]), 1)

    def test_timestamped_filename_and_mt_folder_ftp_upload(self) -> None:
        generated_at = dt.datetime(2026, 10, 1, 6, 7, 8, tzinfo=dt.timezone.utc)
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        filename = write_back.timestamped_csv_name("MT321", mappings, generated_at)
        self.assertEqual(filename, "321_Defended_Matter_20261001_060708.csv")

        class FakeFTP:
            def __init__(self):
                self.current = "/"
                self.directories = {"/", "/bank"}
                self.uploaded = []

            def pwd(self):
                return self.current

            def cwd(self, value):
                if value == "/":
                    self.current = "/"
                    return
                if value.startswith("/"):
                    target = value.rstrip("/") or "/"
                else:
                    target = (self.current.rstrip("/") + "/" + value).rstrip("/")
                if target not in self.directories:
                    raise ftplib.error_perm("550 missing")
                self.current = target

            def mkd(self, value):
                target = (self.current.rstrip("/") + "/" + value).rstrip("/")
                self.directories.add(target)

            def storbinary(self, command, handle):
                self.uploaded.append((command, handle.read()))

            def rename(self, source, destination):
                self.renamed = (source, destination)

            def delete(self, _path):
                return None

        with tempfile.TemporaryDirectory() as directory:
            csv_path = Path(directory) / filename
            csv_path.write_bytes(b"header\nvalue\n")
            ftp = FakeFTP()
            remote_path = write_back.upload_csv_to_ftp(
                ftp, csv_path, write_back.FTP_WRITE_BACK_DIR, "MT321"
            )

        expected_dir = "/LSW TO APT/SBSA Panel Write Back Data/321"
        self.assertIn(expected_dir, ftp.directories)
        self.assertEqual(remote_path, f"{expected_dir}/{filename}")
        self.assertEqual(ftp.uploaded[0][0], f"STOR {remote_path}.uploading")
        self.assertEqual(ftp.renamed, (f"{remote_path}.uploading", remote_path))

    def test_upload_summary_email_is_sent_to_helpdesk(self) -> None:
        sent = []

        class FakeSMTP:
            def __init__(self, host, port, timeout):
                self.connection = (host, port, timeout)

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def send_message(self, message, from_addr, to_addrs):
                sent.append((message, from_addr, to_addrs))

        source = write_back.HandoverSource(
            "ftp:source.xlsx", "source.xlsx", dt.date(2026, 9, 30), "panel_l"
        )
        counts = {"accounts": 2, "new_matters": 2, "file_notes": 3,
                  "excluded": 0, "not_new": 0, "pending": 0,
                  "stages_compared": 4, "stages_new": 1, "stages_skipped": 3}
        uploads = [{"mt_id": "MT303", "rows": 3,
                    "remote_path": "/LSW TO APT/SBSA Panel Write Back Data/303/file.csv"}]
        environment = {
            "MAIL_HOST": "mail.example.test",
            "MAIL_PORT": "25",
            "MAIL_FROM_ADDRESS": "automation@example.test",
            "MAIL_AUTH_MODE": "none",
            "MAIL_ENCRYPTION": "none",
        }
        with (
            patch.dict(os.environ, environment, clear=False),
            patch.object(write_back.smtplib, "SMTP", FakeSMTP),
        ):
            self.assertTrue(write_back.send_ftp_upload_summary_email(
                source, counts, uploads,
                dt.datetime(2026, 10, 1, 6, 0, tzinfo=dt.timezone.utc),
            ))

        self.assertEqual(sent[0][2], ("helpdesk@iconis.co.za", "dev@iconis.co.za"))
        self.assertIn("SBSA Panel Write Back Data/303/file.csv", sent[0][0].get_content())
        self.assertIn("1. Found 3 file notes", sent[0][0].get_content())
        self.assertIn("2. MT IDs compared against latest FTP CSV: 4", sent[0][0].get_content())
        self.assertIn("5. Duplicate CSVs generated or uploaded: No",
                      sent[0][0].get_content())

        no_upload_counts = {**counts, "file_notes": 0, "stages_new": 0}
        with (
            patch.dict(os.environ, environment, clear=False),
            patch.object(write_back.smtplib, "SMTP", FakeSMTP),
        ):
            self.assertTrue(write_back.send_ftp_upload_summary_email(
                source, no_upload_counts, [],
                dt.datetime(2026, 10, 2, 6, 0, tzinfo=dt.timezone.utc),
            ))
        self.assertIn("no-upload notification", sent[1][0]["Subject"])
        self.assertIn("No new file notes were found", sent[1][0].get_content())
        self.assertIn("no CSV was generated or uploaded", sent[1][0].get_content())

    def test_stage_headers_include_only_mapped_fields_and_disambiguate_duplicates(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        self.assertEqual(write_back._stage_headers("MT300", mappings),
                         [*write_back.BANK_ENVELOPE_HEADERS, "Account Number", "CACS Supplier"])
        headers = write_back._stage_headers("MT304", mappings)
        self.assertIn("Fee Amount [K]", headers)
        self.assertIn("Fee Amount [M]", headers)
        self.assertEqual(len(headers), len(set(headers)))
        self.assertNotIn("Field C Name", headers)

    def test_mt321_headers_match_supplied_bank_sample(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        with Path("321_Defended_Matter_20260731.csv").open(encoding="utf-8-sig", newline="") as handle:
            sample_headers = next(csv.reader(handle))
        self.assertEqual(write_back._stage_headers("MT321", mappings), [*sample_headers, "CACS Supplier"])

    def test_reviewed_mapping_ids_and_field_corrections(self) -> None:
        mappings, warnings = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        self.assertEqual(warnings, [])
        self.assertEqual(set(mappings), {
            "MT300", "MT303", "MT304", "MT321", "MT327", "MT328", "MT329",
            "MT334", "MT335", "MT336", "MT347", "MT348", "MT349", "MT350",
            "MT351", "MT354", "MT355", "MT397", "MT405", "MT409", "MT415",
            "MT416", "MT421", "MT428", "MT430",
        })
        rejected = {"MT502", "MT330", "MT326", "MT356", "MT386", "MT387", "MT500"}
        self.assertTrue(rejected.isdisjoint(mappings))
        self.assertTrue(all(write_back._stage_headers(mt_id, mappings)[-1] == "CACS Supplier"
                            for mt_id in mappings))

        matter = {
            **self._matter(),
            "theirref": "'00123456789",
            "magcourtheldat": "DURBAN MAGISTRATES COURT",
            "magcourtdistrict": "DURBAN",
        }
        mt303 = write_back.add_message_data(
            {**self._note(), "stagecode": "MT303"}, matter, mappings
        )
        mt303_fields = {field["column"]: field for field in mt303["message_data"]["fields"]}
        self.assertEqual(mt303_fields["F"]["value"], "L")

        mt334 = write_back.add_message_data(
            {**self._note(), "stagecode": "MT334"}, matter, mappings
        )
        mt334_row = write_back._stage_row(mt334, mappings)
        self.assertEqual(mt334_row["Court Name / Place"], "DURBAN MAGISTRATES COURT")
        self.assertEqual(mt334_row["Account Number"], "00123456789")
        self.assertEqual(mt334_row["CACS Supplier"], "")

    def test_csv_uses_sample_compatible_utf8_without_bom(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "321_Defended_Matter_20261001_060708.csv"
            write_back.write_csv(path, [{"suite": "316"}], ["suite"])
            content = path.read_bytes()
        self.assertFalse(content.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(content, b"suite\r\n316\r\n")

    def test_user_ignored_mt335_and_mt351_fields_are_excluded(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        mt335_headers = write_back._stage_headers("MT335", mappings)
        self.assertNotIn("Judgment Amount", mt335_headers)
        self.assertNotIn("Additional Amount", mt335_headers)
        self.assertNotIn("Additional Interest Rate", mt335_headers)
        mt351_headers = write_back._stage_headers("MT351", mappings)
        self.assertNotIn("Attached Goods Value", mt351_headers)

        mt335 = write_back.add_message_data(
            {**self._note(), "stagecode": "MT335"}, self._matter(), mappings
        )
        mt351 = write_back.add_message_data(
            {**self._note(), "stagecode": "MT351"}, self._matter(), mappings
        )
        self.assertTrue(
            {"J", "L", "M"}.isdisjoint(
                field["column"] for field in mt335["message_data"]["fields"]
            )
        )
        self.assertTrue(
            {"F"}.isdisjoint(
                field["column"] for field in mt351["message_data"]["fields"]
            )
        )

    def test_text_field_uses_file_note_description(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        note = {
            **self._note(), "stagecode": "MT355",
            "description": "  EVENT   DESCRIPTION\r\n\r\nSecond   line  ",
        }
        enriched = write_back.add_message_data(note, self._matter(), mappings)
        fields = {field["name"]: field for field in enriched["message_data"]["fields"]}
        self.assertEqual(fields["Text"]["value"], "EVENT DESCRIPTION Second line")
        self.assertEqual(fields["Text"]["source"], "FileNote.Description")

    def test_mt329_each_note_uses_its_own_description_for_both_fields(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        notes = [
            {**self._note(), "recordid": "1", "stagecode": "MT329", "description": "FIRST NOTE"},
            {**self._note(), "recordid": "2", "stagecode": "MT329", "description": "SECOND NOTE"},
        ]
        rows = []
        for note in notes:
            enriched = write_back.add_message_data(note, self._matter(), mappings)
            rows.append(write_back._stage_row(enriched, mappings))
        self.assertEqual(
            [(row["Description 1"], row["Description2"]) for row in rows],
            [("FIRST NOTE", "FIRST NOTE"), ("SECOND NOTE", "SECOND NOTE")],
        )

    def test_mt421_each_note_uses_its_own_description_for_both_pla_comments(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        notes = [
            {**self._note(), "recordid": "1", "stagecode": "MT421", "description": "FIRST PLA NOTE"},
            {**self._note(), "recordid": "2", "stagecode": "MT421", "description": "SECOND PLA NOTE"},
        ]
        rows = []
        for note in notes:
            enriched = write_back.add_message_data(note, self._matter(), mappings)
            rows.append(write_back._stage_row(enriched, mappings))
        self.assertEqual(
            [(row["PLA Comment 1"], row["PLA Comment 2"]) for row in rows],
            [("FIRST PLA NOTE", "FIRST PLA NOTE"), ("SECOND PLA NOTE", "SECOND PLA NOTE")],
        )

    def test_mt327_summons_address_uses_debtor_parlang_physical_lines(self) -> None:
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        note = {**self._note(), "stagecode": "MT327"}
        debtor_party = {
            "party": {
                "physicalline1": "PARTY LINE 1",
                "physicalline2": "PARTY LINE 2",
                "physicalline3": "PARTY LINE 3",
            },
            "parlang": {
                "physicalline1": "PARLANG LINE 1",
                "physicalline2": "PARLANG LINE 2",
                "physicalline3": "PARLANG LINE 3",
            },
        }
        enriched = write_back.add_message_data(note, self._matter(), mappings, debtor_party)
        fields = {field["column"]: field for field in enriched["message_data"]["fields"]}
        self.assertEqual(
            [fields[column]["value"] for column in ("E", "F", "G")],
            ["PARLANG LINE 1", "PARLANG LINE 2", "PARLANG LINE 3"],
        )
        self.assertEqual(
            [fields[column]["source"] for column in ("E", "F", "G")],
            ["ParLang.physicalline1", "ParLang.physicalline2", "ParLang.physicalline3"],
        )

    def test_mt397_debtor_party_uses_matparty_role_and_parlang(self) -> None:
        calls: list[tuple[str, object]] = []

        def records(table, where, *_args):
            calls.append((table, where))
            return {
                "matparty": [{"matterid": "600001", "roleid": "103", "partyid": "42"}],
                "party": [{"recordid": "42", "partytypeid": "1",
                           "identitynumber": "TEST-ID", "name": "TEST PERSON"}],
                "parlang": [{"partyid": "42", "languageid": "1",
                             "physicalline1": "LINE 1", "physicalline2": "LINE 2",
                             "physicalline3": "LINE 3", "postalcode": "TEST CODE"}],
            }[table]

        with patch.object(write_back, "_fetch_records", side_effect=records):
            party = write_back.fetch_debtor_party("600001", "test-key")
        self.assertEqual([call[0] for call in calls], ["matparty", "party", "parlang"])
        self.assertEqual(calls[0][1], ["MatParty.MatterID,=,600001", "MatParty.RoleID,=,103"])
        self.assertEqual(calls[2][1], ["ParLang.PartyID,=,42", "ParLang.LanguageID,=,1"])

        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        note = {**self._note(), "stagecode": "MT397"}
        enriched = write_back.add_message_data(note, self._matter(), mappings, party, "TEST CITY")
        fields = {field["column"]: field for field in enriched["message_data"]["fields"]}
        self.assertEqual([fields[column]["value"] for column in "CDEFGH"],
                         ["1", "TEST-ID", "TEST PERSON", "LINE 1", "LINE 2", "LINE 3"])
        self.assertEqual(fields["F"]["source"], "ParLang.physicalline1")
        self.assertEqual(fields["I"]["value"], "TEST CITY")
        self.assertEqual(fields["I"]["source"], "Handover.Town/City")
        self.assertEqual(fields["J"]["value"], "TEST CODE")
        self.assertEqual(fields["J"]["source"], "ParLang.postalcode")
        self.assertIsNone(fields["N"]["value"])
        self.assertIsNone(fields["O"]["value"])

        with tempfile.TemporaryDirectory() as directory:
            workbook = self._workbook(Path(directory))
            source = write_back.local_handover_sources([workbook])[0]
            with (
                patch.object(write_back, "fetch_matters_by_reference", return_value=[self._matter()]),
                patch.object(write_back, "fetch_filenotes", return_value=[note]),
                patch.object(write_back, "fetch_debtor_party", return_value=party) as fetch_party,
            ):
                rows_by_stage, counts = write_back.process_handover_source(
                    source, workbook, "test-key", {"MT397"}, mappings)
        fetch_party.assert_called_once_with("600001", "test-key")
        self.assertEqual(counts["file_notes"], 1)
        self.assertEqual(rows_by_stage["MT397"][0]["City"], "TEST CITY")
        self.assertEqual(rows_by_stage["MT397"][0]["Postal Code"], "TEST CODE")

    def test_mt397_ambiguous_debtor_link_is_not_used(self) -> None:
        links = [{"matterid": "600001", "roleid": "103", "partyid": party_id}
                 for party_id in ("101", "102")]
        with (
            patch.object(write_back, "_fetch_records", return_value=links) as fetch_records,
            patch("sys.stderr", new_callable=io.StringIO),
        ):
            self.assertIsNone(write_back.fetch_debtor_party("600001", "test-key"))
        fetch_records.assert_called_once()

    def test_stage_comparison_only_allows_new_rows(self) -> None:
        headers = ["messageID", "Acccount Number"]
        rows = [{"messageID": "300", "Acccount Number": "000123"}]
        with patch.object(write_back, "_latest_uploaded_stage_rows",
                          return_value=(headers, {("300", "000123")})):
            self.assertFalse(write_back._has_new_stage_rows(object(), "MT300", rows, headers))
        with patch.object(write_back, "_latest_uploaded_stage_rows",
                          return_value=(headers, {("300", "000123")})):
            rows.append({"messageID": "300", "Acccount Number": "000456"})
            self.assertTrue(write_back._has_new_stage_rows(object(), "MT300", rows, headers))

    def test_missing_matter_is_retried_on_next_run(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = self._workbook(root)
            output_dir = root / "csv"
            with (
                patch.dict(os.environ, {"LEGALSUITE_API_KEY": "test-key"}),
                patch.object(write_back, "load_env_file"),
                patch.object(write_back, "fetch_matters_by_reference", side_effect=[[], [self._matter()]]) as fetch_matters,
                patch.object(write_back, "fetch_filenotes", return_value=[self._note()]) as fetch_notes,
                redirect_stdout(io.StringIO()),
                patch("sys.stderr", new_callable=io.StringIO),
            ):
                args = ["--handover-file", str(workbook), "--csv-dir", str(output_dir)]
                self.assertEqual(write_back.main(args), 0)
                state = json.loads((output_dir / "processed_handover_files.json").read_text())
                self.assertFalse(next(iter(state.values()))["complete"])
                self.assertEqual(write_back.main(args), 0)
            self.assertEqual(fetch_matters.call_count, 2)
            self.assertEqual(fetch_notes.call_count, 1)
            state = json.loads((output_dir / "processed_handover_files.json").read_text())
            self.assertTrue(next(iter(state.values()))["complete"])


class IncrementalFileNoteTests(unittest.TestCase):
    def _matter(self) -> dict:
        return {
            "recordid": "600001", "archivestatus": "0", "clientid": "150307",
            "mattertypeid": "4", "theirref": "00012345678", "fileref": "STA387/0001",
            "formatteddateinstructed": "06 Oct 2026",
            "internalcomment": "Imported from handover file on 6 October 2026",
        }

    def _state(self) -> dict:
        return {
            "version": 1,
            "start_date": "2026-10-06",
            "matters": {
                "600001": {
                    "matter_id": "600001",
                    "client_id": "150307",
                    "account_reference": "00012345678",
                    "file_ref": "STA387/0001",
                    "date_instructed": "2026-10-06",
                    "physical_city": "TEST CITY",
                    "handover_date": "2026-10-06",
                    "handover_sources": ["ftp:test.xlsx"],
                    "sent_file_note_ids": ["700001"],
                }
            },
            "pending_batch": None,
        }

    def test_monitor_discovery_includes_cutoff_and_filters_earlier_files(self) -> None:
        october = [
            write_back.HandoverSource("old", "old.xlsx", dt.date(2026, 10, 5), "panel_l"),
            write_back.HandoverSource("cutoff", "cutoff.xlsx", dt.date(2026, 10, 6), "panel_l"),
            write_back.HandoverSource("new", "new.xlsx", dt.date(2026, 10, 19), "panel_l"),
        ]
        november = [
            write_back.HandoverSource("later", "later.xlsx", dt.date(2026, 11, 1), "panel_l"),
        ]
        with patch.object(
            write_back, "discover_ftp_handover_sources", side_effect=[october, november]
        ) as discover:
            sources = write_back.discover_ftp_handover_sources_since(
                object(), dt.date(2026, 10, 6), dt.date(2026, 11, 1)
            )
        self.assertEqual([source.identity for source in sources], ["cutoff", "new", "later"])
        self.assertEqual(discover.call_count, 2)

    def test_only_unsent_file_note_record_ids_are_exported(self) -> None:
        notes = [
            {"recordid": "700001", "matterid": "600001", "stagecode": "MT300",
             "description": "ALREADY SENT", "formatteddatetime": "06 Oct 2026 09:00:00"},
            {"recordid": "700002", "matterid": "600001", "stagecode": "MT300",
             "description": "NEW NOTE", "formatteddatetime": "19 Oct 2026 10:30:00"},
            {"recordid": "700002", "matterid": "600001", "stagecode": "MT300",
             "description": "DUPLICATE API ROW", "formatteddatetime": "19 Oct 2026 10:30:00"},
        ]
        mappings, _ = write_back.load_message_mappings(write_back.DEFAULT_MAPPINGS_FILE)
        with (
            patch.object(write_back, "fetch_matter", return_value=self._matter()),
            patch.object(write_back, "fetch_filenotes", return_value=notes),
        ):
            rows, note_ids, counts = write_back._collect_unsent_file_notes(
                self._state(), "test-key", {"MT300"}, mappings
            )
        self.assertEqual(counts["file_notes"], 1)
        self.assertEqual(note_ids, {"600001": ["700002"]})
        self.assertEqual(len(rows["MT300"]), 1)
        self.assertEqual(rows["MT300"][0]["messageID"], "300")

    def test_state_backup_can_restore_a_missing_local_ledger(self) -> None:
        class MemoryFTP:
            def __init__(self):
                self.files = {}

            def storbinary(self, command, handle):
                self.files[command.removeprefix("STOR ")] = handle.read()

            def rename(self, source, destination):
                self.files[destination] = self.files.pop(source)

            def nlst(self, directory):
                return [path for path in self.files if path.startswith(directory.rstrip("/") + "/")]

            def retrbinary(self, command, callback):
                callback(self.files[command.removeprefix("RETR ")])

            def delete(self, path):
                self.files.pop(path, None)

        ftp = MemoryFTP()
        state = self._state()
        with patch.object(write_back, "ensure_ftp_directory"):
            remote_path = write_back.backup_note_state_to_ftp(ftp, state)
        self.assertTrue(remote_path.startswith(
            write_back.FTP_STATE_BACKUP_DIR + "/" + write_back.FTP_STATE_BACKUP_PREFIX
        ))
        self.assertTrue(remote_path.endswith(".json"))

        with tempfile.TemporaryDirectory() as directory:
            local_path = Path(directory) / write_back.NOTE_STATE_FILENAME
            with redirect_stdout(io.StringIO()):
                restored = write_back._restore_note_state_from_ftp(ftp, local_path)
            self.assertEqual(restored, state)
            self.assertEqual(json.loads(local_path.read_text(encoding="utf-8")), state)

    def test_no_unseen_notes_sends_status_without_csv_or_upload(self) -> None:
        class FakeFTP:
            def quit(self):
                return None

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            args = write_back.argparse.Namespace(
                csv_dir=root / "output", download_dir=root / "downloads", ftp_timeout=30,
                latest_only=False,
            )
            discovery_counts = {
                "handover_files": 1, "accounts": 1, "registered": 0,
                "pending": 0, "excluded": 0, "not_new": 0,
            }
            with (
                patch.dict(os.environ, {"FTP_HOST": "test", "FTP_USER": "test", "FTP_PASS": "test"}),
                patch.object(write_back, "connect_ftp", return_value=FakeFTP()),
                patch.object(write_back, "_load_or_restore_note_state",
                             return_value=write_back._default_note_state()),
                patch.object(
                    write_back, "backup_note_state_to_ftp",
                    return_value="/backup/standard_bank_write_back_state_test.json",
                ),
                patch.object(write_back, "discover_ftp_handover_sources_since", return_value=[]),
                patch.object(write_back, "_register_handover_matters", return_value=discovery_counts),
                patch.object(
                    write_back, "_collect_unsent_file_notes",
                    return_value=({}, {}, {"new_matters": 1, "file_notes": 0, "excluded": 0,
                                           "not_new": 0, "pending": 0}),
                ),
                patch.object(write_back, "write_csv") as write_csv,
                patch.object(write_back, "upload_csv_to_ftp") as upload,
                patch.object(write_back, "send_ftp_upload_summary_email", return_value=True) as email,
                redirect_stdout(io.StringIO()),
            ):
                self.assertEqual(
                    write_back.run_incremental_handover_batch(args, "test-key", {"MT300"}, {}), 0
                )
            write_csv.assert_not_called()
            upload.assert_not_called()
            email.assert_called_once()
            self.assertEqual(email.call_args.args[2], [])
            self.assertTrue((args.csv_dir / write_back.NOTE_STATE_FILENAME).is_file())

    def test_pending_delivery_commits_note_ids_after_upload_and_email(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            csv_path = root / "300_Acknowledgement_20261019_060000.csv"
            csv_path.write_text("messageID\r\n300\r\n", encoding="utf-8")
            state = self._state()
            state["pending_batch"] = {
                "created_at": "2026-10-19T06:00:00+02:00",
                "files": [{
                    "mt_id": "MT300", "rows": 1, "local_path": str(csv_path),
                    "remote_path": write_back._remote_csv_path(csv_path, "MT300"),
                    "uploaded": False,
                }],
                "note_ids_by_matter": {"600001": ["700002"]},
                "counts": {"file_notes": 1},
            }
            state_path = root / write_back.NOTE_STATE_FILENAME
            with (
                patch.object(write_back, "_ftp_file_exists", return_value=False),
                patch.object(
                    write_back, "upload_csv_to_ftp",
                    return_value=state["pending_batch"]["files"][0]["remote_path"],
                ) as upload,
                patch.object(write_back, "send_ftp_upload_summary_email", return_value=True) as email,
                patch.object(write_back, "backup_note_state_to_ftp",
                             return_value="/backup/standard_bank_write_back_state_test.json"),
            ):
                write_back._finish_pending_batch(object(), state, state_path)
            upload.assert_called_once()
            email.assert_called_once()
            self.assertIsNone(state["pending_batch"])
            self.assertEqual(state["matters"]["600001"]["sent_file_note_ids"], ["700001", "700002"])
            persisted = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertIsNone(persisted["pending_batch"])

    def test_pending_csv_is_rebuilt_after_local_box_loss(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            missing_path = root / "lost" / "300_Acknowledgement_20261019_060000.csv"
            state = self._state()
            remote_path = write_back._remote_csv_path(missing_path, "MT300")
            state["pending_batch"] = {
                "created_at": "2026-10-19T06:00:00+02:00",
                "files": [{
                    "mt_id": "MT300", "row_count": 1,
                    "headers": ["messageID", "Acccount Number"],
                    "data": [{"messageID": "300", "Acccount Number": "00012345678"}],
                    "local_path": str(missing_path), "remote_path": remote_path,
                    "uploaded": False,
                }],
                "note_ids_by_matter": {"600001": ["700002"]},
                "counts": {"file_notes": 1},
            }
            state_path = root / write_back.NOTE_STATE_FILENAME
            with (
                patch.object(write_back, "_ftp_file_exists", return_value=False),
                patch.object(write_back, "upload_csv_to_ftp", return_value=remote_path) as upload,
                patch.object(write_back, "send_ftp_upload_summary_email", return_value=True),
                patch.object(write_back, "backup_note_state_to_ftp",
                             return_value="/backup/standard_bank_write_back_state_test.json"),
                redirect_stdout(io.StringIO()),
            ):
                write_back._finish_pending_batch(object(), state, state_path)
            rebuilt_path = upload.call_args.args[1]
            self.assertTrue(rebuilt_path.is_file())
            with rebuilt_path.open(encoding="utf-8", newline="") as handle:
                self.assertEqual(list(csv.DictReader(handle))[0]["Acccount Number"], "00012345678")
            self.assertIsNone(state["pending_batch"])

    def test_email_failure_leaves_note_ids_pending_for_retry(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            csv_path = root / "300_Acknowledgement_20261019_060000.csv"
            csv_path.write_text("messageID\r\n300\r\n", encoding="utf-8")
            state = self._state()
            remote_path = write_back._remote_csv_path(csv_path, "MT300")
            state["pending_batch"] = {
                "created_at": "2026-10-19T06:00:00+02:00",
                "files": [{"mt_id": "MT300", "rows": 1, "local_path": str(csv_path),
                           "remote_path": remote_path, "uploaded": False}],
                "note_ids_by_matter": {"600001": ["700002"]},
                "counts": {"file_notes": 1},
            }
            state_path = root / write_back.NOTE_STATE_FILENAME
            with (
                patch.object(write_back, "_ftp_file_exists", return_value=False),
                patch.object(write_back, "upload_csv_to_ftp", return_value=remote_path),
                patch.object(write_back, "send_ftp_upload_summary_email", return_value=False),
                patch.object(write_back, "backup_note_state_to_ftp",
                             return_value="/backup/standard_bank_write_back_state_test.json"),
            ):
                with self.assertRaisesRegex(RuntimeError, "summary email"):
                    write_back._finish_pending_batch(object(), state, state_path)
            self.assertIsNotNone(state["pending_batch"])
            self.assertTrue(state["pending_batch"]["files"][0]["uploaded"])
            self.assertEqual(state["matters"]["600001"]["sent_file_note_ids"], ["700001"])


if __name__ == "__main__":
    unittest.main()
