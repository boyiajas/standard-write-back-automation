# Standard Bank handover file-note export

`standard_bank_write_back.py` reads LegalSuite and the Standard Bank FTP server, creates per-stage CSV files, uploads production batch files to FTP, and emails a completion summary to the helpdesk. It does not update LegalSuite.

See [OPERATIONS.md](OPERATIONS.md) for the complete workflow, configuration reference, manual commands, file-note state behavior, troubleshooting guidance, and cron deployment example.

## Server deployment

The automation uses only the Python 3 standard library and requires Python 3.9 or newer. Clone the repository, create the private runtime configuration from the committed template, and lock down its permissions:

```bash
git clone https://github.com/boyiajas/standard-write-back-automation.git
cd standard-write-back-automation
cp .env.example .env
chmod 600 .env
```

Edit `.env` with the production FTP, LegalSuite API, and SMTP credentials. The populated `.env`, downloaded handovers, generated exports, state files, caches, and logs are deliberately excluded from Git.

Verify the deployment before enabling cron:

```bash
python3 -m py_compile standard_bank_write_back.py
python3 -m unittest test_standard_bank_write_back.py
```

## Daily run

```bash
python3 standard_bank_write_back.py
```

With no arguments, the script checks every Panel L and Debt Review handover dated from **6 October 2026 through today** in Africa/Johannesburg time. It registers matters created from those handovers, refreshes each registered matter, and exports only allowed LegalSuite file notes whose `FileNote.RecordID` has not already completed delivery. Each MT stage gets a separate CSV. If no new file note exists, the run creates no CSV and uploads nothing, but sends a daily no-upload notification. The script does not schedule itself; configure the daily schedule separately in the Linux crontab.

The script needs `LEGALSUITE_API_KEY`, `FTP_HOST`, `FTP_USER`, and `FTP_PASS` in the environment or local `.env`. The write-back destination is fixed in code as `/LSW TO APT/SBSA Panel Write Back Data`. SMTP uses the existing `MAIL_HOST`, `MAIL_PORT`, `MAIL_USERNAME`, `MAIL_PASSWORD`, `MAIL_AUTH_MODE`, `MAIL_ENCRYPTION`, and `MAIL_FROM_ADDRESS` settings. Upload summaries are always sent to `helpdesk@iconis.co.za` and `dev@iconis.co.za`. It uses only the Python standard library. A daily crontab entry can use an absolute script path; the script finds `.env` and its workbooks relative to itself:

```cron
0 6 * * * cd /opt/standard-write-back-automation && /usr/bin/python3 standard_bank_write_back.py >> /var/log/standard-write-back-automation.log 2>&1
```

## Test and backfill options

```bash
python3 standard_bank_write_back.py --latest-only
python3 standard_bank_write_back.py --month 2026-08
python3 standard_bank_write_back.py --month current
python3 standard_bank_write_back.py --handover-file /path/to/Standard_Bank_Panel_L_Handover_20260814.xlsx
python3 standard_bank_write_back.py 565398
```

`--latest-only` limits discovery to the newest dated file in each handover category. An explicit `--month current` or `--month YYYY-MM` invokes the legacy month backfill flow. `--handover-file` works without FTP, remains local-only, and can be repeated. A positional MatterID keeps the older single-matter JSON mode. `--force` applies to the legacy backfill flow. `--csv-dir` and `--download-dir` change the default output and download locations.

## Files produced

- `downloads_sbsa/` holds cached FTP workbooks.
- `output_sbsa/YYYY-MM/<numeric MT ID>_<description>_YYYYMMDD_HHMMSS.csv` has one row per file note for that stage, for example `321_Defended_Matter_20261001_060708.csv`. Files use the bank envelope `suite, lawRef, messageID, sender, senderUser, recipient, date, dateTime`, followed by that MT ID's fields from `SD-E4 Message ID Mappings.xlsx` in workbook order. `suite` is `316`, `messageID` is the numeric MT ID, and unavailable envelope values remain blank. MT321 exactly follows the supplied `321_Defended_Matter_20260731.csv` header layout, including the bank's legacy `Acccount Number` spelling. Repeated generic mapping labels include their workbook column letter, such as `Fee Amount [K]` and `Fee Amount [M]`.
- Production FTP runs upload each CSV to `/LSW TO APT/SBSA Panel Write Back Data/<numeric MT ID>/<timestamped filename>`, for example `/LSW TO APT/SBSA Panel Write Back Data/321/...csv`. Missing folders are created. Uploads use a temporary `.uploading` name and are renamed only after transfer completes.
- The daily flow de-duplicates by the immutable numeric `FileNote.RecordID` recorded in `output_sbsa/write_back_note_state.json`. It does not resend a note because another field on its existing record changed.
- Every state change is also uploaded atomically to `/LSW TO APT/SBSA Write Back Feedback Files/Automation State/` as a timestamped `standard_bank_write_back_state_*.json` snapshot. The newest valid snapshot is restored automatically when the local state file is missing. The newest 100 snapshots are retained.
- The helpdesk receives a summary after all stage files are uploaded. Note IDs are committed to the sent ledger only after every upload and that email succeed. An interrupted delivery remains pending; the next run reuses the same local files, skips an exact remote file that is already present, and resumes the remaining delivery work.
- A successful daily check with no unseen notes emails the same recipients with a clear no-upload result. Failure to send that status makes the run fail so the next execution retries it.
- A stage with no mapping row gets the bank envelope plus `Acccount Number` as its fallback layout. Blank cells indicate fields with no confirmed LegalSuite value.
- `output_sbsa/write_back_note_state.json` records registered matters, delivered note IDs, and any pending delivery. A pending snapshot includes the CSV headers and rows so a replacement box can rebuild a missing pending CSV. `output_sbsa/processed_handover_files.json` remains the state file for local and explicit month backfill runs.

The export is based on the handover date in the filename. A matter must have the same `DateInstructed` and the `Imported from handover file on <date>` comment set by `ftp_download_today.py`. It must also meet the archive status, client ID, and matter type rule documented in [AGENTS.md](AGENTS.md). Fields without a confirmed LegalSuite source remain blank in the CSV.

## Verification

```bash
python3 -m py_compile standard_bank_write_back.py
python3 -m unittest test_standard_bank_write_back.py
```

## Missing LegalSuite field sources

[LegalSuite_Field_Source_Checklist.csv](LegalSuite_Field_Source_Checklist.csv) tracks every mapped field and its source status. [LegalSuite_Missing_Field_Sources.csv](LegalSuite_Missing_Field_Sources.csv) contains only the 107 fields still needing Postman research, and [LegalSuite_Missing_Field_Sources.md](LegalSuite_Missing_Field_Sources.md) groups them by MT stage. [LegalSuite_Field_Source_Gaps.md](LegalSuite_Field_Source_Gaps.md) explains the MT397 debtor-party lookup and other mapping gaps.
