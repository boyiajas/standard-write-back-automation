#!/usr/bin/env python3
"""Export file notes for newly created Standard Bank handover matters.

With no MatterID, scan previous-month FTP handovers and write CSVs. An explicit
MatterID runs the original JSON lookup. The script reads LegalSuite only;
fields without a confirmed source remain empty.
"""

import argparse
import csv
import datetime as dt
import ftplib
import hashlib
import json
import os
import re
import smtplib
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from urllib import error, parse, request
from zoneinfo import ZoneInfo

from env_config import load_env_file


DEFAULT_STAGES_FILE = Path(__file__).resolve().parent / "SBSA Panel L LegalSuite Stages.xlsx"
DEFAULT_MAPPINGS_FILE = Path(__file__).resolve().parent / "SD-E4 Message ID Mappings.xlsx"
DEFAULT_DOWNLOAD_DIR = Path(__file__).resolve().parent / "downloads_sbsa"
DEFAULT_CSV_DIR = Path(__file__).resolve().parent / "output_sbsa"
EXPORT_FORMAT_VERSION = 12
FTP_WRITE_BACK_DIR = "/LSW TO APT/SBSA Panel Write Back Data"
WRITE_BACK_REPORT_TO = (
    "helpdesk@iconis.co.za",
    "dev@iconis.co.za",
)
BANK_SUITE_CODE = "316"
BANK_ENVELOPE_HEADERS = (
    "suite", "lawRef", "messageID", "sender", "senderUser", "recipient", "date", "dateTime",
)
BANK_STAGE_FILENAME_OVERRIDES = {
    "MT321": "Defended_Matter",
}
BANK_STAGE_HEADER_OVERRIDES = {
    "MT321": (
        "CourtActionCasNo",
        "CourtActioncourtDate",
        "CourtActionamount",
        "DefendedDeclarationFiledDate",
        "DefendedPleaEnteredDate",
        "DefendedDeclarationReceived",
        "DefendedDeclarationReply",
        "DefendedNoticeOfBarIssued",
        "DefendedNoticeOfBarIssuedDate",
        "DefendedCounterClaimAmount",
        "DefendedTrialRequestDate",
        "DefendedPreTrialDate",
        "DefendedTrialDateSetDate",
        "DefendedTrialVerdictDate",
        "DefendedrescissionApplicationRec",
        "DefendedrescissionReason",
        "DefendedrescissionRecDate",
        "DefendedrescissionOpposedStatus",
        "DefendedopposingAffidavitIssued",
        "DefendedrescissionGranted",
        "Acccount Number",
    ),
}
API_BASE = "https://api.legalsuite.net"
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RETRYABLE_STATUSES = {429, 500, 502, 503, 504}
LEGALSUITE_DATE_OFFSET = 36161
EXCEL_EPOCH = dt.datetime(1899, 12, 30)
ALLOWED_CLIENT_IDS = frozenset({
    "150307", "334695", "155128", "334565", "209250",
    "334568", "283850", "334567", "267742", "334569",
})
REQUIRED_MATTER_TYPE_ID = "4"
IGNORED_MAPPING_FIELDS = frozenset({
    ("MT335", "J"),  # Judgment Amount
    ("MT335", "L"),  # Additional Amount
    ("MT335", "M"),  # Additional Interest Rate
    ("MT351", "F"),  # Attached Goods Value
})
CLIENT_CODE_MAP = {
    "STA387": "150307", "DR387": "334695", "STD9": "155128", "DRR9": "334565",
    "STA482": "209250", "DR482": "334568", "STA822": "283850", "DR822": "334567",
    "STA614": "267742", "DR614": "334569",
}
HANDOVER_FILENAME = re.compile(r"^Standard_Bank_Panel_L_Handover_(\d{8})(_DR)?\.xlsx$", re.IGNORECASE)
PANEL_HANDOVER_DIR = "SBSA/Panel L/Handover_APT_LSW"
DEBT_REVIEW_HANDOVER_DIRS = (
    "SBSA/Debt Review/Debt_Review_Handover_APT_LWS",
    "SBSA/Debt Review/Debt_Review_ Handover_APT_LWS",
)
ENGLISH_MONTH_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


@dataclass(frozen=True)
class HandoverSource:
    identity: str
    filename: str
    handover_date: dt.date
    kind: str
    remote_dir: str | None = None
    local_path: Path | None = None


@dataclass(frozen=True)
class HandoverAccount:
    row_number: int
    client_code: str
    client_id: str | None
    reference: str
    physical_city: str | None = None


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    kind = cell.get("t")
    if kind == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{SPREADSHEET_NS}}}t"))
    value = cell.find(f"{{{SPREADSHEET_NS}}}v")
    if value is None or value.text is None:
        return ""
    return shared_strings[int(value.text)] if kind == "s" else value.text


def read_workbook_rows(workbook_path: Path, first_sheet_only: bool = False) -> list[tuple[int, dict[str, str]]]:
    """Read cell text from simple XLSX sheets without requiring openpyxl."""
    if not workbook_path.is_file():
        raise ValueError(f"Workbook not found: {workbook_path}")

    try:
        with zipfile.ZipFile(workbook_path) as workbook:
            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in workbook.namelist():
                root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
                shared_strings = [
                    "".join(node.text or "" for node in item.iter(f"{{{SPREADSHEET_NS}}}t"))
                    for item in root.findall(f"{{{SPREADSHEET_NS}}}si")
                ]

            sheet_paths = sorted(
                name for name in workbook.namelist()
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
            )
            if first_sheet_only:
                sheet_paths = sheet_paths[:1]
            rows: list[tuple[int, dict[str, str]]] = []
            for sheet_path in sheet_paths:
                sheet = ET.fromstring(workbook.read(sheet_path))
                for row in sheet.findall(f".//{{{SPREADSHEET_NS}}}sheetData/{{{SPREADSHEET_NS}}}row"):
                    cells = {
                        "".join(char for char in cell.get("r", "") if char.isalpha()):
                        _cell_text(cell, shared_strings).strip()
                        for cell in row.findall(f"{{{SPREADSHEET_NS}}}c")
                    }
                    rows.append((int(row.get("r", "0")), cells))
    except (zipfile.BadZipFile, ET.ParseError, IndexError, ValueError) as exc:
        raise ValueError(f"Could not read {workbook_path}: {exc}") from exc

    return rows


def load_stage_codes(workbook_path: Path) -> set[str]:
    """Read every nonempty Code cell below a Code header in the workbook."""
    codes: set[str] = set()
    code_column: str | None = None
    for _, cells in read_workbook_rows(workbook_path):
        if code_column is None:
            code_column = next(
                (column for column, value in cells.items() if value.casefold() == "code"),
                None,
            )
            continue
        code = cells.get(code_column, "").upper()
        if code:
            codes.add(code)

    if code_column is None:
        raise ValueError(f"No Code column found in {workbook_path}")
    if not codes:
        raise ValueError(f"The Code column has no stage codes in {workbook_path}")
    return codes


def _column_order(column: str) -> int:
    number = 0
    for character in column:
        number = number * 26 + ord(character) - ord("A") + 1
    return number


def load_message_mappings(workbook_path: Path) -> tuple[dict[str, dict], list[str]]:
    """Read the message name and ordered field labels for each numeric MT ID."""
    mappings: dict[str, dict] = {}
    warnings: list[str] = []
    for row_number, cells in read_workbook_rows(workbook_path):
        raw_id = cells.get("A", "")
        description = cells.get("B", "")
        if not raw_id.isdecimal() or not description:
            continue
        mt_id = f"MT{int(raw_id)}"
        # Row 50 in the supplied workbook duplicates MT349, while the stages
        # workbook names Writ Requested as MT350. Keep the source visible.
        if mt_id == "MT349" and description.casefold() == "writ requested date":
            mt_id = "MT350"
            warnings.append(
                f"Row {row_number}: Writ Requested Date is MT349 in the mapping "
                "workbook; using MT350 from the stages workbook."
            )
        if mt_id in mappings:
            raise ValueError(f"Duplicate message mapping for {mt_id} in {workbook_path}")
        fields = [
            {"column": column, "name": name}
            for column, name in sorted(cells.items(), key=lambda item: _column_order(item[0]))
            if _column_order(column) > 2 and name
        ]
        mappings[mt_id] = {
            "file_description": description,
            "source_row": row_number,
            "fields": fields,
        }
    if not mappings:
        raise ValueError(f"No MT message mappings found in {workbook_path}")
    return mappings, warnings


def _fetch_records(table: str, where: str | list[str], api_key: str, timeout: int = 30) -> list[dict]:
    clauses = [where] if isinstance(where, str) else where
    form = parse.urlencode([("where[]", clause) for clause in clauses]).encode("utf-8")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    api_request = request.Request(f"{API_BASE}/{table}/get", data=form, headers=headers, method="POST")

    for attempt in range(3):
        try:
            with request.urlopen(api_request, timeout=timeout) as response:
                payload = json.load(response)
            break
        except error.HTTPError as exc:
            if exc.code not in RETRYABLE_STATUSES or attempt == 2:
                raise RuntimeError(f"LegalSuite returned HTTP {exc.code}") from exc
        except (error.URLError, TimeoutError) as exc:
            if attempt == 2:
                raise RuntimeError(f"Could not reach LegalSuite: {exc}") from exc
        if attempt < 2:
            time.sleep(2 * (attempt + 1))

    if not isinstance(payload, dict):
        raise ValueError("LegalSuite returned an unexpected response")
    if payload.get("errors"):
        raise ValueError(f"LegalSuite reported an error: {payload['errors']}")
    notes = payload.get("data")
    if not isinstance(notes, list) or any(not isinstance(note, dict) for note in notes):
        raise ValueError(f"LegalSuite response does not contain a {table} list")
    return notes


def fetch_filenotes(matter_id: str, api_key: str, timeout: int = 30) -> list[dict]:
    return _fetch_records("filenote", f"FileNote.MatterID,=,{matter_id}", api_key, timeout)


def fetch_matter(matter_id: str, api_key: str, timeout: int = 30) -> dict:
    matters = _fetch_records("matter", f"Matter.RecordID,=,{matter_id}", api_key, timeout)
    matches = [matter for matter in matters if str(matter.get("recordid")) == matter_id]
    if len(matches) != 1:
        raise ValueError(f"Expected one matter with RecordID {matter_id}; found {len(matches)}")
    return matches[0]


def fetch_matters_by_reference(client_id: str, reference: str, api_key: str, timeout: int = 30) -> list[dict]:
    matters = _fetch_records(
        "matter",
        [f"Matter.ClientID,=,{client_id}", f"Matter.TheirRef,=,{reference}"],
        api_key,
        timeout,
    )
    return [
        matter for matter in matters
        if str(matter.get("clientid") or "").strip() == client_id
        and normalize_reference(matter.get("theirref")) == reference
    ]


def fetch_debtor_party(matter_id: str, api_key: str, timeout: int = 30) -> dict | None:
    """Find the single debtor linked to a matter, then read its party details."""
    links = _fetch_records(
        "matparty",
        [f"MatParty.MatterID,=,{matter_id}", "MatParty.RoleID,=,103"],
        api_key, timeout,
    )
    party_ids = {
        str(link.get("partyid") or "").strip()
        for link in links
        if str(link.get("matterid") or "").strip() == matter_id
        and str(link.get("roleid") or "").strip() == "103"
        and str(link.get("partyid") or "").strip().isdecimal()
    }
    if len(party_ids) != 1:
        print(f"Matter {matter_id}: expected one debtor PartyID from MatParty role 103; found {len(party_ids)}",
              file=sys.stderr)
        return None
    party_id = party_ids.pop()
    parties = _fetch_records("party", f"Party.RecordID,=,{party_id}", api_key, timeout)
    parties = [party for party in parties if str(party.get("recordid") or "").strip() == party_id]
    if len(parties) != 1:
        print(f"Matter {matter_id}: expected one Party record; found {len(parties)}", file=sys.stderr)
        return None
    languages = _fetch_records(
        "parlang",
        [f"ParLang.PartyID,=,{party_id}", "ParLang.LanguageID,=,1"],
        api_key, timeout,
    )
    languages = [row for row in languages if str(row.get("partyid") or "").strip() == party_id
                 and str(row.get("languageid") or "").strip() == "1"]
    if len(languages) != 1:
        print(f"Matter {matter_id}: expected one English ParLang record; found {len(languages)}",
              file=sys.stderr)
    return {"party": parties[0], "parlang": languages[0] if len(languages) == 1 else {}}


def matter_eligibility_reasons(matter: dict) -> list[str]:
    """Apply the Standard Bank matter predicate before requesting file notes."""
    reasons: list[str] = []
    if "archivestatus" not in matter:
        reasons.append("Matter.ArchiveStatus was not returned")
    elif matter["archivestatus"] is not None and str(matter["archivestatus"]).strip() != "0":
        reasons.append(f"Matter.ArchiveStatus must be 0 or NULL (got {matter['archivestatus']!r})")

    client_id = str(matter.get("clientid") or "").strip()
    if client_id not in ALLOWED_CLIENT_IDS:
        reasons.append(f"Matter.ClientID is not in the allowed list (got {client_id or 'missing'})")

    matter_type_id = str(matter.get("mattertypeid") or "").strip()
    if matter_type_id != REQUIRED_MATTER_TYPE_ID:
        reasons.append(f"Matter.MatterTypeID must be 4 (got {matter_type_id or 'missing'})")
    return reasons


def matching_filenotes(notes: list[dict], matter_id: str, stage_codes: set[str]) -> list[dict]:
    return [
        note for note in notes
        if str(note.get("matterid", "")).strip() == matter_id
        and str(note.get("stagecode") or "").strip().upper() in stage_codes
    ]


def _note_date(note: dict) -> dt.date | None:
    date_text = str(note.get("formatteddate") or "").strip()
    if date_text:
        try:
            return dt.datetime.strptime(date_text, "%d %b %Y").date()
        except ValueError:
            pass
    for key in ("formatteddatetime", "datetime"):
        timestamp_text = str(note.get(key) or "").strip()
        for pattern in ("%d %b %Y %H:%M:%S", "%d %b %Y %H:%M"):
            try:
                return dt.datetime.strptime(timestamp_text, pattern).date()
            except ValueError:
                pass
    raw_date = str(note.get("date") or "").strip()
    if raw_date.isdecimal() and int(raw_date) > LEGALSUITE_DATE_OFFSET:
        return (EXCEL_EPOCH + dt.timedelta(days=int(raw_date) - LEGALSUITE_DATE_OFFSET)).date()
    return None


def _note_datetime(note: dict) -> dt.datetime | None:
    for key in ("formatteddatetime", "datetime"):
        value = str(note.get(key) or "").strip()
        for date_format in ("%d %b %Y %H:%M:%S", "%d %b %Y %H:%M"):
            try:
                return dt.datetime.strptime(value, date_format)
            except ValueError:
                pass
    date_value = _note_date(note)
    time_text = str(note.get("formattedtime") or "").strip()
    if date_value and time_text:
        try:
            return dt.datetime.combine(date_value, dt.time.fromisoformat(time_text))
        except ValueError:
            pass
    raw_time = str(note.get("time") or "").strip()
    if date_value and raw_time.isdecimal() and int(raw_time) > 0:
        seconds = int(raw_time) // 100  # LegalSuite stores centiseconds after midnight.
        if 0 <= seconds < 86400:
            return dt.datetime.combine(date_value, dt.time()) + dt.timedelta(seconds=seconds)
    return None


def _first_party_value(party: dict, parlang: dict, key: str, prefer_parlang: bool = False) -> tuple[object, str | None]:
    candidates = (("ParLang", parlang), ("Party", party)) if prefer_parlang else (("Party", party), ("ParLang", parlang))
    for table, record in candidates:
        value = record.get(key)
        if value is not None and str(value).strip():
            return value, f"{table}.{key}"
    return None, None


def _mt397_debtor_field(field: dict, debtor_party: dict | None,
                        handover_city: str | None) -> tuple[object, str | None]:
    expected = {
        "C": "party type", "D": "id no.", "E": "party name", "F": "address line 1",
        "G": "address line 2", "H": "address line 3", "I": "city", "J": "postal code",
    }
    column = field["column"]
    if field["name"].strip().casefold() != expected.get(column):
        return None, None
    if column == "I":
        return (handover_city, "Handover.Town/City") if handover_city else (None, None)
    if debtor_party is None:
        return None, None
    party, parlang = debtor_party["party"], debtor_party["parlang"]
    if column == "C":
        return _first_party_value(party, parlang, "partytypeid")
    if column == "D":
        return _first_party_value(party, parlang, "identitynumber")
    if column == "E":
        value, source = _first_party_value(party, parlang, "name")
        if value is not None:
            return value, source
        return _first_party_value(party, parlang, "fullname", prefer_parlang=True)
    address_line = {"F": 1, "G": 2, "H": 3}.get(column)
    if address_line is not None:
        return _first_party_value(party, parlang, f"physicalline{address_line}", prefer_parlang=True)
    if column == "J":
        return _first_party_value(party, parlang, "postalcode", prefer_parlang=True)
    return None, None


def _field_value(field: dict, mt_id: str, matter: dict, note: dict,
                 note_date: dt.date | None, timestamp: dt.datetime | None,
                 debtor_party: dict | None = None,
                 handover_city: str | None = None) -> tuple[object, str | None]:
    if mt_id == "MT397" and field["column"] in {"C", "D", "E", "F", "G", "H", "I", "J"}:
        return _mt397_debtor_field(field, debtor_party, handover_city)
    if mt_id == "MT327":
        expected = {
            "E": "summons address line 1",
            "F": "summons address line 2",
            "G": "summons address line 3",
        }
        column = field["column"]
        if field["name"].strip().casefold() == expected.get(column) and debtor_party is not None:
            party, parlang = debtor_party["party"], debtor_party["parlang"]
            address_line = {"E": 1, "F": 2, "G": 3}[column]
            return _first_party_value(
                party, parlang, f"physicalline{address_line}", prefer_parlang=True
            )
    name = field["name"].strip().casefold()
    if mt_id == "MT329" and (
        (field["column"] == "E" and name == "description 1")
        or (field["column"] == "F" and name == "description2")
    ):
        value = note.get("description")
        return (value, "FileNote.Description") if value else (None, None)
    if mt_id == "MT421" and (
        (field["column"] == "G" and name == "pla comment 1")
        or (field["column"] == "H" and name == "pla comment 2")
    ):
        value = note.get("description")
        return (value, "FileNote.Description") if value else (None, None)
    if name == "date":
        return (note_date.strftime("%Y%m%d"), "FileNote.Date") if note_date else (None, None)
    if name == "datetime":
        return (timestamp.strftime("%Y%m%d%H%M%S"), "FileNote.Date/Time") if timestamp else (None, None)
    if name == "account number":
        value = matter.get("theirref")
        return (value, "Matter.TheirRef") if value else (None, None)
    if name == "case number":
        value = matter.get("casenumber")
        return (value, "Matter.CaseNumber") if value else (None, None)
    if name == "text":
        value = note.get("description")
        return (value, "FileNote.Description") if value else (None, None)
    return None, None


def add_message_data(note: dict, matter: dict, mappings: dict[str, dict],
                     debtor_party: dict | None = None, handover_city: str | None = None) -> dict:
    mt_id = str(note.get("stagecode") or "").strip().upper()
    mapping = mappings.get(mt_id)
    note_date = _note_date(note)
    timestamp = _note_datetime(note)
    fields = []
    if mapping:
        for field in mapping["fields"]:
            if (mt_id, field["column"]) in IGNORED_MAPPING_FIELDS:
                continue
            value, source = _field_value(field, mt_id, matter, note, note_date, timestamp,
                                         debtor_party, handover_city)
            fields.append({**field, "value": value, "source": source})
    message_data = {
        "mt_id": mt_id,
        "mapping_status": "mapped" if mapping else "not_in_mapping_workbook",
        "mapping_row": mapping["source_row"] if mapping else None,
        "file_description": mapping["file_description"] if mapping else note.get("description"),
        "file_description_source": "message mapping workbook" if mapping else "FileNote.Description",
        "file_note_description": note.get("description"),
        "date": note_date.strftime("%Y%m%d") if note_date else None,
        "datetime": timestamp.strftime("%Y%m%d%H%M%S") if timestamp else None,
        "account_number": matter.get("theirref") or None,
        "fields": fields,
        "unresolved_fields": [
            {"column": field["column"], "name": field["name"], "reason": "No confirmed LegalSuite source or value"}
            for field in fields if field["value"] is None
        ],
    }
    return {**note, "message_data": message_data}


def normalize_reference(value: object) -> str:
    text = str(value or "").strip()
    return text[1:].strip() if text.startswith("'") else text


def _normal_header(value: str) -> str:
    return "".join(character for character in value.casefold() if character.isalnum())


def read_handover_accounts(path: Path) -> list[HandoverAccount]:
    rows = read_workbook_rows(path, first_sheet_only=True)
    header_index = next(
        (index for index, (_, cells) in enumerate(rows[:10])
         if {"clientcode", "reference"}.issubset({_normal_header(value) for value in cells.values()})),
        None,
    )
    if header_index is None:
        raise ValueError(f"Client Code and Reference headers not found in {path}")
    headers = {
        _normal_header(value): column
        for column, value in rows[header_index][1].items()
    }
    physical_city_column = next(
        (column for column, value in sorted(rows[header_index][1].items(),
                                            key=lambda item: _column_order(item[0]))
         if _normal_header(value) == "towncity"),
        None,
    )
    accounts: list[HandoverAccount] = []
    seen: set[tuple[str, str]] = set()
    for row_number, cells in rows[header_index + 1:]:
        client_code = cells.get(headers["clientcode"], "").strip().upper()
        reference = normalize_reference(cells.get(headers["reference"], ""))
        if not client_code and not reference:
            continue
        key = (client_code, reference)
        if key in seen:
            continue
        seen.add(key)
        city = cells.get(physical_city_column, "").strip() if physical_city_column else ""
        accounts.append(HandoverAccount(row_number, client_code, CLIENT_CODE_MAP.get(client_code),
                                        reference, city or None))
    return accounts


def _matter_instructed_date(matter: dict) -> dt.date | None:
    formatted = str(matter.get("formatteddateinstructed") or "").strip()
    for pattern in ("%d %b %Y", "%d %B %Y"):
        try:
            return dt.datetime.strptime(formatted, pattern).date()
        except ValueError:
            pass
    raw = str(matter.get("dateinstructed") or "").strip()
    if raw.isdecimal() and int(raw) > LEGALSUITE_DATE_OFFSET:
        return (EXCEL_EPOCH + dt.timedelta(days=int(raw) - LEGALSUITE_DATE_OFFSET)).date()
    return None


def _is_new_handover_matter(matter: dict, handover_date: dt.date) -> bool:
    marker = f"Imported from handover file on {handover_date.strftime('%d %B %Y')}"
    return (
        _matter_instructed_date(matter) == handover_date
        and str(matter.get("internalcomment") or "").strip().casefold().startswith(marker.casefold())
    )


def _month_to_scan(value: str, today: dt.date | None = None) -> tuple[int, int]:
    today = today or dt.datetime.now(ZoneInfo("Africa/Johannesburg")).date()
    if value == "current":
        return today.year, today.month
    if value == "previous":
        last_month = dt.date(today.year, today.month, 1) - dt.timedelta(days=1)
        return last_month.year, last_month.month
    if not re.fullmatch(r"\d{4}-\d{2}", value):
        raise ValueError("--month must be previous, current, or YYYY-MM")
    year, month = (int(part) for part in value.split("-"))
    dt.date(year, month, 1)  # Validate the month.
    return year, month


def _source_from_name(filename: str, kind: str, identity: str, remote_dir: str | None = None,
                      local_path: Path | None = None) -> HandoverSource | None:
    match = HANDOVER_FILENAME.fullmatch(filename)
    if not match or (match.group(2) is not None) != (kind == "debt_review"):
        return None
    try:
        handover_date = dt.datetime.strptime(match.group(1), "%Y%m%d").date()
    except ValueError:
        return None
    return HandoverSource(identity, filename, handover_date, kind, remote_dir, local_path)


def _ftp_names(ftp: ftplib.FTP, remote_dir: str) -> list[str]:
    try:
        listing = ftp.nlst(remote_dir)
    except ftplib.error_perm:
        original_dir = ftp.pwd()
        try:
            ftp.cwd(remote_dir)
        except ftplib.error_perm:
            return []
        try:
            try:
                listing = ftp.nlst()
            except ftplib.error_perm:
                return []
        finally:
            ftp.cwd(original_dir)
    return [item.rsplit("/", 1)[-1] for item in listing]


def connect_ftp(timeout: int) -> ftplib.FTP:
    ftp = ftplib.FTP(os.environ["FTP_HOST"], timeout=timeout)
    ftp.login(os.environ["FTP_USER"], os.environ["FTP_PASS"])
    ftp.set_pasv(True)
    return ftp


def discover_ftp_handover_sources(ftp: ftplib.FTP, year: int, month: int) -> list[HandoverSource]:
    panel_dir = f"{PANEL_HANDOVER_DIR}/{ENGLISH_MONTH_ABBR[month - 1]} {year}"
    locations = [(panel_dir, "panel_l")]
    locations.extend((directory, "debt_review") for directory in DEBT_REVIEW_HANDOVER_DIRS)
    sources: list[HandoverSource] = []
    seen: set[tuple[str, str]] = set()
    for remote_dir, kind in locations:
        for filename in _ftp_names(ftp, remote_dir):
            source = _source_from_name(filename, kind, f"ftp:{remote_dir}/{filename}", remote_dir)
            if source is None or (source.handover_date.year, source.handover_date.month) != (year, month):
                continue
            key = (kind, filename.casefold())
            if key not in seen:
                seen.add(key)
                sources.append(source)
    return sorted(sources, key=lambda source: (source.handover_date, source.kind, source.filename))


def local_handover_sources(paths: list[Path]) -> list[HandoverSource]:
    sources: list[HandoverSource] = []
    for path in paths:
        path = path.resolve()
        if not path.is_file():
            raise ValueError(f"Handover workbook not found: {path}")
        kind = "debt_review" if path.stem.upper().endswith("_DR") else "panel_l"
        source = _source_from_name(path.name, kind, f"local:{path}", local_path=path)
        if source is None:
            raise ValueError(f"Unrecognized handover filename: {path.name}")
        sources.append(source)
    return sorted(sources, key=lambda source: (source.handover_date, source.kind, source.filename))


def newest_sources(sources: list[HandoverSource]) -> list[HandoverSource]:
    latest_by_kind = {
        kind: max(source.handover_date for source in sources if source.kind == kind)
        for kind in {source.kind for source in sources}
    }
    return [source for source in sources if source.handover_date == latest_by_kind[source.kind]]


def download_handover_source(ftp: ftplib.FTP | None, source: HandoverSource, download_dir: Path) -> Path:
    if source.local_path is not None:
        return source.local_path
    if ftp is None or source.remote_dir is None:
        raise ValueError("FTP connection is required for a remote handover file")
    destination = download_dir / source.remote_dir / source.filename
    destination.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(dir=destination.parent, prefix=f".{source.filename}.", delete=False) as handle:
            temporary_path = Path(handle.name)
            ftp.retrbinary(f"RETR {source.remote_dir}/{source.filename}", handle.write)
        os.replace(temporary_path, destination)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return destination


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _stage_headers(mt_id: str, mappings: dict[str, dict]) -> list[str]:
    """Return the Standard Bank envelope and ordered MT-specific fields."""
    mapping = mappings.get(mt_id)
    if mapping is None:
        return [*BANK_ENVELOPE_HEADERS, "Acccount Number"]
    fields = [
        field for field in mapping["fields"]
        if field["column"] not in {"C", "D"}
        and (mt_id, field["column"]) not in IGNORED_MAPPING_FIELDS
    ]
    override = BANK_STAGE_HEADER_OVERRIDES.get(mt_id)
    if override is not None:
        if len(override) != len(fields):
            raise ValueError(
                f"{mt_id} bank header template has {len(override)} fields; "
                f"the mapping workbook requires {len(fields)}"
            )
        return [*BANK_ENVELOPE_HEADERS, *override]
    names = [field["name"] for field in fields]
    duplicates = {name for name in names if names.count(name) > 1}
    headers = list(BANK_ENVELOPE_HEADERS)
    for field in fields:
        name = field["name"]
        if name == "Account Number":
            headers.append("Acccount Number")
        else:
            headers.append("{} [{}]".format(name, field["column"]) if name in duplicates else name)
    return headers

def _stage_row(note: dict, mappings: dict[str, dict]) -> dict[str, object]:
    message = note["message_data"]
    mt_id = message["mt_id"]
    headers = _stage_headers(mt_id, mappings)
    envelope = [BANK_SUITE_CODE, "", mt_id.removeprefix("MT"), "", "", "",
                message["date"], message["datetime"]]
    if mt_id not in mappings:
        values = [message["account_number"]]
    else:
        values = [
            field["value"] for field in message["fields"]
            if field["column"] not in {"C", "D"}
        ]
    return dict(zip(headers, [*envelope, *values]))


def _safe_csv_value(value: object) -> object:
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def write_csv(path: Path, rows: list[dict[str, object]], headers: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", newline="", encoding="utf-8", dir=path.parent,
            prefix=f".{path.name}.", delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow({key: _safe_csv_value(value) for key, value in row.items()})
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def timestamped_csv_name(mt_id: str, mappings: dict[str, dict], generated_at: dt.datetime) -> str:
    """Build the bank filename using numeric MT ID, description, date, and time."""
    if re.fullmatch(r"MT\d+", mt_id) is None:
        raise ValueError(f"Invalid MT ID for FTP folder: {mt_id!r}")
    mapping = mappings.get(mt_id)
    description = BANK_STAGE_FILENAME_OVERRIDES.get(mt_id)
    if description is None:
        description = re.sub(r"[^A-Za-z0-9]+", "_", str(
            mapping["file_description"] if mapping else "Write_Back"
        )).strip("_") or "Write_Back"
    return f"{mt_id.removeprefix('MT')}_{description}_{generated_at.strftime('%Y%m%d_%H%M%S')}.csv"


def ensure_ftp_directory(ftp: ftplib.FTP, remote_dir: str) -> None:
    """Create a nested FTP directory when needed, preserving the current directory."""
    normalized = remote_dir.strip().replace("\\", "/")
    if not normalized or normalized == ".":
        return

    original_dir = ftp.pwd()
    try:
        if normalized.startswith("/"):
            ftp.cwd("/")
        for part in (part for part in normalized.split("/") if part and part != "."):
            try:
                ftp.cwd(part)
            except ftplib.error_perm:
                ftp.mkd(part)
                ftp.cwd(part)
    finally:
        ftp.cwd(original_dir)


def upload_csv_to_ftp(ftp: ftplib.FTP, local_path: Path,
                      remote_base_dir: str, mt_id: str) -> str:
    """Upload a CSV atomically into its MT ID directory and return its remote path."""
    if re.fullmatch(r"MT\d+", mt_id) is None:
        raise ValueError(f"Invalid MT ID for FTP folder: {mt_id!r}")
    stage_folder = mt_id.removeprefix("MT")
    remote_dir = f"{remote_base_dir.rstrip('/')}/{stage_folder}"
    ensure_ftp_directory(ftp, remote_dir)
    remote_path = f"{remote_dir}/{local_path.name}"
    temporary_remote_path = f"{remote_path}.uploading"
    with local_path.open("rb") as handle:
        ftp.storbinary(f"STOR {temporary_remote_path}", handle)
    try:
        ftp.rename(temporary_remote_path, remote_path)
    except ftplib.all_errors:
        try:
            ftp.delete(temporary_remote_path)
        except ftplib.all_errors:
            pass
        raise
    return remote_path


def send_ftp_upload_summary_email(source: HandoverSource, counts: dict[str, int],
                                  uploads: list[dict[str, object]],
                                  completed_at: dt.datetime) -> bool:
    """Email the helpdesk only after every CSV for one handover has uploaded."""
    smtp_host = os.getenv("MAIL_HOST", os.getenv("SMTP_HOST", "")).strip()
    smtp_port = int(os.getenv("MAIL_PORT", os.getenv("SMTP_PORT", "587")).strip() or "587")
    smtp_user = os.getenv("MAIL_USERNAME", os.getenv("SMTP_USER", "")).strip()
    smtp_pass = os.getenv("MAIL_PASSWORD", os.getenv("SMTP_PASS", "")).strip()
    smtp_from = os.getenv("MAIL_FROM_ADDRESS", os.getenv("SMTP_FROM", smtp_user)).strip()
    auth_mode = os.getenv("MAIL_AUTH_MODE", os.getenv("SMTP_AUTH_MODE", "login")).strip().lower()
    use_auth = auth_mode not in {"none", "noauth", "false", "0", "no"}
    encryption = os.getenv("MAIL_ENCRYPTION", "").strip().lower()
    use_tls = encryption not in {"", "null", "none", "false", "0", "no"} if "MAIL_ENCRYPTION" in os.environ else (
        os.getenv("SMTP_USE_TLS", "true").strip().lower() not in {"0", "false", "no"}
    )
    missing = [name for name, value in (("MAIL_HOST", smtp_host), ("MAIL_FROM_ADDRESS", smtp_from)) if not value]
    if missing:
        print(f"Write-back summary email skipped: missing settings: {', '.join(missing)}", file=sys.stderr)
        return False

    recipients = WRITE_BACK_REPORT_TO
    message = EmailMessage()
    message["Subject"] = f"Standard Bank write-back FTP upload summary -- {completed_at.strftime('%Y/%m/%d %H:%M:%S')}"
    message["From"] = smtp_from
    message["To"] = ", ".join(recipients)
    upload_lines = [
        f"- {item['mt_id']}: {item['rows']} row(s) -> {item['remote_path']}"
        for item in uploads
    ] or ["- No stage CSVs were generated."]
    message.set_content(
        "Good Day,\n\n"
        "The Standard Bank LegalSuite write-back CSV upload has completed.\n\n"
        f"Source handover: {source.filename}\n"
        f"Handover date: {source.handover_date.isoformat()}\n"
        f"Completed at: {completed_at.isoformat()}\n"
        f"Accounts: {counts['accounts']}\n"
        f"New matters: {counts['new_matters']}\n"
        f"File notes exported: {counts['file_notes']}\n"
        f"Stage files uploaded: {len(uploads)}\n"
        f"Excluded accounts: {counts['excluded']}\n"
        f"Not-new accounts: {counts['not_new']}\n"
        f"Pending accounts: {counts['pending']}\n\n"
        "Uploaded files:\n" + "\n".join(upload_lines) + "\n\n"
        "Kind Regards,\n"
    )

    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            with smtplib.SMTP(smtp_host, smtp_port, timeout=60) as server:
                if use_tls:
                    server.starttls()
                if use_auth and smtp_user:
                    server.login(smtp_user, smtp_pass)
                server.send_message(message, from_addr=smtp_from, to_addrs=recipients)
            return True
        except (OSError, smtplib.SMTPException) as exc:
            last_error = exc
            print(f"Write-back summary email attempt {attempt}/3 failed: {exc}", file=sys.stderr)
    print(f"Write-back summary email failed: {last_error}", file=sys.stderr)
    return False


def process_handover_source(source: HandoverSource, path: Path, api_key: str,
                            stage_codes: set[str], mappings: dict[str, dict]) -> tuple[dict[str, list[dict[str, object]]], dict[str, int]]:
    accounts = read_handover_accounts(path)
    rows_by_stage: dict[str, list[dict[str, object]]] = {}
    counts = {"accounts": len(accounts), "new_matters": 0, "excluded": 0,
              "not_new": 0, "pending": 0, "file_notes": 0}
    for account in accounts:
        if account.client_id is None or not account.reference:
            counts["pending"] += 1
            print(f"Handover row {account.row_number}: missing mapped client code or reference", file=sys.stderr)
            continue
        matters = fetch_matters_by_reference(account.client_id, account.reference, api_key)
        if not matters:
            counts["pending"] += 1
            print(f"Handover row {account.row_number}: matter not found yet", file=sys.stderr)
            continue
        newly_created = [matter for matter in matters if _is_new_handover_matter(matter, source.handover_date)]
        if not newly_created:
            counts["not_new"] += 1
            continue
        if len(newly_created) != 1:
            counts["pending"] += 1
            print(f"Handover row {account.row_number}: multiple new matters matched", file=sys.stderr)
            continue
        matter = newly_created[0]
        reasons = matter_eligibility_reasons(matter)
        if reasons:
            counts["excluded"] += 1
            continue
        matter_id = str(matter.get("recordid") or "")
        if not matter_id.isdecimal():
            raise ValueError(f"Handover row {account.row_number}: matched matter has no RecordID")
        counts["new_matters"] += 1
        notes = matching_filenotes(fetch_filenotes(matter_id, api_key), matter_id, stage_codes)
        notes.sort(key=lambda note: (str(note.get("date") or ""), str(note.get("time") or ""),
                                     str(note.get("recordid") or "")))
        needs_debtor_party = any(
            str(note.get("stagecode") or "").strip().upper() in {"MT327", "MT397"}
            for note in notes
        )
        debtor_party = fetch_debtor_party(matter_id, api_key) if needs_debtor_party else None
        if needs_debtor_party and debtor_party is None:
            counts["pending"] += 1
        for note in notes:
            enriched = add_message_data(note, matter, mappings, debtor_party, account.physical_city)
            mt_id = enriched["message_data"]["mt_id"]
            rows_by_stage.setdefault(mt_id, []).append(_stage_row(enriched, mappings))
            counts["file_notes"] += 1
    return rows_by_stage, counts


def _load_state(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    state = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(state, dict) or any(not isinstance(value, dict) for value in state.values()):
        raise ValueError(f"Invalid processing state: {path}")
    return state


def _save_state(path: Path, state: dict[str, dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         prefix=".state.", delete=False) as handle:
            temporary_path = Path(handle.name)
            json.dump(state, handle, indent=2, sort_keys=True)
            handle.write("\n")
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def _pending_ftp_sources(state: dict[str, dict]) -> list[HandoverSource]:
    sources: list[HandoverSource] = []
    for identity, record in state.items():
        if record.get("complete") or not identity.startswith("ftp:"):
            continue
        remote_path = identity[4:]
        if "/" not in remote_path:
            continue
        remote_dir, filename = remote_path.rsplit("/", 1)
        if remote_dir.startswith(PANEL_HANDOVER_DIR + "/"):
            kind = "panel_l"
        elif remote_dir in DEBT_REVIEW_HANDOVER_DIRS:
            kind = "debt_review"
        else:
            continue
        source = _source_from_name(filename, kind, identity, remote_dir)
        if source is not None:
            sources.append(source)
    return sources


def run_handover_batch(args: argparse.Namespace, api_key: str, stage_codes: set[str],
                       mappings: dict[str, dict]) -> int:
    year, month = _month_to_scan(args.month)
    state_path = args.csv_dir / "processed_handover_files.json"
    state = _load_state(state_path)
    ftp: ftplib.FTP | None = None
    try:
        if args.handover_file:
            sources = local_handover_sources(args.handover_file)
        else:
            missing = [
                name for name in ("FTP_HOST", "FTP_USER", "FTP_PASS")
                if not os.getenv(name, "").strip()
            ]
            if missing:
                raise ValueError(f"FTP settings missing: {', '.join(missing)}")
            ftp = connect_ftp(args.ftp_timeout)
            sources = discover_ftp_handover_sources(ftp, year, month)
            if args.month == "previous" and not args.latest_only:
                known = {source.identity for source in sources}
                sources.extend(source for source in _pending_ftp_sources(state) if source.identity not in known)
                sources.sort(key=lambda source: (source.handover_date, source.kind, source.filename))
        if args.latest_only:
            sources = newest_sources(sources)
        if not sources:
            print(f"No handover files found for {year:04d}-{month:02d}")
            return 0

        config_hashes = {"stages_sha256": _sha256(args.stages_file),
                         "mappings_sha256": _sha256(args.mappings_file)}
        processed = skipped = pending = 0
        for source in sources:
            local_path = download_handover_source(ftp, source, args.download_dir)
            source_hash = _sha256(local_path)
            csv_dir = args.csv_dir / source.handover_date.strftime("%Y-%m")
            previous = state.get(source.identity, {})
            if (not args.force and previous.get("complete") and previous.get("source_sha256") == source_hash
                    and all(previous.get(key) == value for key, value in config_hashes.items())
                    and previous.get("export_format_version") == EXPORT_FORMAT_VERSION
                    and isinstance(previous.get("csv_paths"), list)
                    and all(Path(path).is_file() for path in previous["csv_paths"])):
                print(f"Already processed: {source.identity}")
                skipped += 1
                continue
            rows_by_stage, counts = process_handover_source(source, local_path, api_key, stage_codes, mappings)
            generated_at = dt.datetime.now(ZoneInfo("Africa/Johannesburg"))
            csv_paths: list[str] = []
            generated_files: list[tuple[str, Path, int]] = []
            for mt_id, rows in sorted(rows_by_stage.items()):
                csv_path = csv_dir / timestamped_csv_name(mt_id, mappings, generated_at)
                write_csv(csv_path, rows, _stage_headers(mt_id, mappings))
                csv_paths.append(str(csv_path.resolve()))
                generated_files.append((mt_id, csv_path, len(rows)))
                print(f"CSV: {csv_path} | stage={mt_id} rows={len(rows)}")

            uploads: list[dict[str, object]] = []
            notification_sent = False
            if source.local_path is None and generated_files:
                if ftp is None:
                    raise RuntimeError("FTP connection is unavailable for write-back upload")
                try:
                    ftp.voidcmd("NOOP")
                except ftplib.all_errors:
                    try:
                        ftp.close()
                    except ftplib.all_errors:
                        pass
                    ftp = connect_ftp(args.ftp_timeout)
                for mt_id, csv_path, row_count in generated_files:
                    remote_path = upload_csv_to_ftp(ftp, csv_path, FTP_WRITE_BACK_DIR, mt_id)
                    uploads.append({"mt_id": mt_id, "rows": row_count, "remote_path": remote_path})
                    print(f"FTP upload: {remote_path} | stage={mt_id} rows={row_count}")
                notification_sent = send_ftp_upload_summary_email(
                    source, counts, uploads, dt.datetime.now(ZoneInfo("Africa/Johannesburg"))
                )
                if not notification_sent:
                    raise RuntimeError("CSV files uploaded, but the helpdesk summary email could not be sent")
            complete = counts["pending"] == 0
            state[source.identity] = {
                "source_sha256": source_hash,
                **config_hashes,
                "export_format_version": EXPORT_FORMAT_VERSION,
                "complete": complete,
                "pending_accounts": counts["pending"],
                "csv_paths": csv_paths,
                "ftp_paths": [str(item["remote_path"]) for item in uploads],
                "notification_sent": notification_sent,
                "processed_at": dt.datetime.now(ZoneInfo("Africa/Johannesburg")).isoformat(),
            }
            _save_state(state_path, state)
            processed += 1
            pending += counts["pending"]
            print(f"Handover: {source.identity} | stage_files={len(csv_paths)} accounts={counts['accounts']} "
                  f"new_matters={counts['new_matters']} file_notes={counts['file_notes']} excluded={counts['excluded']} "
                  f"not_new={counts['not_new']} pending={counts['pending']}")
        print(f"Handover files processed={processed} already_processed={skipped} pending_accounts={pending}")
        return 0
    finally:
        if ftp is not None:
            try:
                ftp.quit()
            except ftplib.all_errors:
                ftp.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("matter_id", nargs="?", help="Single-matter JSON mode; omit to process handover files")
    parser.add_argument("--stages-file", type=Path, default=DEFAULT_STAGES_FILE)
    parser.add_argument("--mappings-file", type=Path, default=DEFAULT_MAPPINGS_FILE)
    parser.add_argument("--output", type=Path, help="JSON output path in single-matter mode")
    parser.add_argument("--month", default="previous", help="Handover month: previous (default), current, or YYYY-MM")
    parser.add_argument("--handover-file", action="append", type=Path,
                        help="Use a local handover workbook instead of FTP; repeat for multiple files")
    parser.add_argument("--latest-only", action="store_true", help="Process only the newest dated file in each handover folder")
    parser.add_argument("--download-dir", type=Path, default=DEFAULT_DOWNLOAD_DIR)
    parser.add_argument("--csv-dir", type=Path, default=DEFAULT_CSV_DIR)
    parser.add_argument("--ftp-timeout", type=int, default=30)
    parser.add_argument("--force", action="store_true", help="Reprocess files marked complete in the state file")
    args = parser.parse_args(argv)

    if args.matter_id is not None and not args.matter_id.isdecimal():
        parser.error("matter_id must contain digits only")
    if args.matter_id is None and args.output is not None:
        parser.error("--output is for single-matter JSON mode; use --csv-dir for handover CSVs")
    if args.matter_id is not None and args.handover_file:
        parser.error("Use either matter_id or --handover-file")
    if args.ftp_timeout <= 0:
        parser.error("--ftp-timeout must be greater than zero")

    load_env_file()
    api_key = os.getenv("LEGALSUITE_API_KEY", "").strip()
    if not api_key:
        parser.error("LEGALSUITE_API_KEY is required in the environment or .env file")

    try:
        stage_codes = load_stage_codes(args.stages_file)
        mappings, mapping_warnings = load_message_mappings(args.mappings_file)
        if args.matter_id is None:
            return run_handover_batch(args, api_key, stage_codes, mappings)
        matter = fetch_matter(args.matter_id, api_key)
        exclusion_reasons = matter_eligibility_reasons(matter)
        notes = fetch_filenotes(args.matter_id, api_key) if not exclusion_reasons else []
        matches = matching_filenotes(notes, args.matter_id, stage_codes)
        debtor_party = fetch_debtor_party(args.matter_id, api_key) if any(
            str(note.get("stagecode") or "").strip().upper() in {"MT327", "MT397"}
            for note in matches
        ) else None
        data = [add_message_data(note, matter, mappings, debtor_party) for note in matches]
        result = {
            "matter_id": args.matter_id,
            "matter_eligibility": {"eligible": not exclusion_reasons, "reasons": exclusion_reasons},
            "matching_count": len(data),
            "mapping_warnings": mapping_warnings,
            "unmapped_stage_codes": sorted(stage_codes - mappings.keys()),
            "data": data,
        }
        formatted = json.dumps(result, indent=2, ensure_ascii=False) + "\n"
        if args.output:
            args.output.write_text(formatted, encoding="utf-8")
            print(f"Wrote {len(matches)} matching file notes to {args.output}", file=sys.stderr)
        else:
            sys.stdout.write(formatted)
        return 0
    except (OSError, ValueError, RuntimeError, ftplib.Error, EOFError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
