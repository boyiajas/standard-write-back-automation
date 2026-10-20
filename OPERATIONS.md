# Standard Bank LegalSuite Write-Back Automation

This document explains what `standard_bank_write_back.py` does, how it is configured, how it identifies a new file note, and how to run it manually or from cron.

## Purpose

The script reads Standard Bank handover workbooks from the configured FTP server, finds the corresponding newly created LegalSuite matters, retrieves matching LegalSuite file notes, maps those notes into Standard Bank MT-stage CSV files, and uploads the files back to the FTP write-back directory.

It also sends a completion summary to the helpdesk after successful uploads. If a required operation fails, it attempts to send a failure notification and exits with status `1`.

The script does not modify LegalSuite records.

## Runtime requirements

- Python 3.9 or newer.
- Network access to `api.legalsuite.net`, the configured FTP server, and the SMTP server.
- The two mapping workbooks in the same directory as the script:
  - `SBSA Panel L LegalSuite Stages.xlsx`
  - `SD-E4 Message ID Mappings.xlsx`
- A populated private `.env` file. Start from `.env.example` and never commit credentials.

Only Python standard-library modules are used; no `pip install` step is required.

## Configuration

Required `.env` values:

```dotenv
FTP_HOST=...
FTP_USER=...
FTP_PASS=...
LEGALSUITE_API_KEY=...
MAIL_HOST=...
MAIL_PORT=587
MAIL_USERNAME=...
MAIL_PASSWORD=...
MAIL_AUTH_MODE=login
MAIL_ENCRYPTION=tls
MAIL_FROM_ADDRESS=...
```

The write-back directory is fixed in the script as:

`/LSW TO APT/SBSA Panel Write Back Data`

Each MT ID is stored in its numeric subfolder, for example `.../321/`.

Incremental state snapshots are stored separately under:

`/LSW TO APT/SBSA Write Back Feedback Files/Automation State/`

They use timestamped names such as `standard_bank_write_back_state_20261020_060000000000.json`.

Completion and failure messages are sent to:

- `helpdesk@iconis.co.za`
- `dev@iconis.co.za`

The SMTP sender must be authorized to relay to those recipients. A Mimecast `451 Open relay not allowed` response means the CSV/FTP work may have completed, but the email notification was rejected and the process returns a failure status.

## Normal workflow

1. Load `.env`, stage definitions, and field mappings.
2. Use Africa/Johannesburg time and list Panel L and Debt Review handovers dated from 6 October 2026 through today.
3. Download each workbook into the local `downloads_sbsa/` cache.
4. Read client code and account/reference values.
5. Find the matching LegalSuite matter by client and reference.
6. Require the matter to be newly created from the matching handover and satisfy the configured eligibility rules.
7. Register the matter by storing its LegalSuite MatterID, FileRef, account reference, handover source, and physical city.
8. Retrieve matching file notes and exclude each numeric `FileNote.RecordID` already in the sent ledger.
9. Enrich the remaining rows using the confirmed LegalSuite field sources and mapping workbook, then group them by MT stage code.
10. If every registered matter has no unseen note, create no CSV, upload nothing, and send no completion email.
11. Otherwise generate one timestamped CSV per stage and upload it atomically through a temporary `.uploading` filename.
12. Send the completion summary after uploads, then commit the delivered note IDs to state.
13. Save each state transition locally and as an atomic timestamped FTP snapshot. Retain the newest 100 snapshots.

The numeric LegalSuite file-note RecordID prevents the same note from being uploaded repeatedly. Existing note edits are not treated as new notes. If the local state is missing, the newest valid FTP snapshot is restored before handover discovery. If upload or email delivery is interrupted, the pending batch stays in state and is completed before the next discovery pass. Pending snapshots contain the CSV headers and rows needed to rebuild a missing local pending file after box loss.

## CSV format and naming

Files are written beneath `output_sbsa/YYYY-MM/` using:

`<numeric MT ID>_<description>_YYYYMMDD_HHMMSS.csv`

For example:

`321_Defended_Matter_20261008_060000.csv`

The bank envelope begins with `suite, lawRef, messageID, sender, senderUser, recipient, date, dateTime`; `suite` is `316` and `messageID` is the numeric MT ID. MT321 follows the supplied sample layout exactly, including the legacy `Acccount Number` header. CSVs are plain UTF-8 with CRLF records.

## Manual commands

From the repository directory:

```bash
python3 standard_bank_write_back.py
python3 standard_bank_write_back.py --latest-only
python3 standard_bank_write_back.py --month 2026-08
python3 standard_bank_write_back.py --month current
python3 standard_bank_write_back.py --handover-file /path/to/Standard_Bank_Panel_L_Handover_20260814.xlsx
python3 standard_bank_write_back.py 565398
```

The positional MatterID command runs the original single-matter JSON mode. `--handover-file` is local-only. Explicit `--month current` and `--month YYYY-MM` commands run the legacy month backfill path, which retains its source and FTP row comparison behavior. Use `--force` only with that legacy flow when a completed source must be rebuilt.

## Verification

```bash
python3 -m py_compile standard_bank_write_back.py env_config.py
python3 -m unittest test_standard_bank_write_back.py
```

The current test suite covers mapping rules, MT321 sample compatibility, FTP paths, state retries, duplicate detection, and notification behavior.

## Cron deployment

Create the log directory and make sure the server `.env` has mode `600`:

```bash
mkdir -p /home/pajakaiye/logs
chmod 600 /legalsuite-automation/daily-run-script/standardbank-legalsuite/.env
```

Example cron entry, following the requested deployment style:

```cron
30 5 * * * bash -lc 'cd /legalsuite-automation/daily-run-script/standardbank-legalsuite && source ~/pyenv/bin/activate && python3 standard_bank_write_back.py' >> /home/pajakaiye/logs/standard_bank_write_back.log 2>&1
```

This runs daily at 05:30 in the server's configured timezone. Confirm the timezone with `timedatectl`; the script itself uses Africa/Johannesburg when selecting dates and naming files.

Before enabling the cron entry, run the same command manually and verify the generated output, FTP destination, and SMTP result. Do not use `--force` in cron.

## Runtime directories and state

- `downloads_sbsa/`: cached source handover workbooks; ignored by Git.
- `output_sbsa/YYYY-MM/`: generated CSVs; ignored by Git.
- `output_sbsa/write_back_note_state.json`: registered matters, sent file-note IDs, and an interrupted delivery batch; ignored by Git.
- `/LSW TO APT/SBSA Write Back Feedback Files/Automation State/standard_bank_write_back_state_*.json`: remote recovery snapshots; newest 100 retained.
- `output_sbsa/processed_handover_files.json`: legacy local and explicit month backfill state; ignored by Git.
- `.env`: private credentials; ignored by Git.

If a matter is not found yet, handover discovery retries it on a later run. If email is rejected, FTP files may already exist, but their note IDs remain uncommitted. The next run checks for those exact remote filenames and resumes the same pending batch without uploading the same file again.
