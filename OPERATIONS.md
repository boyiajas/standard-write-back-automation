# Standard Bank LegalSuite Write-Back Automation

This document explains what `standard_bank_write_back.py` does, how it is configured, how it decides whether to upload a CSV, and how to run it manually or from cron.

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

Completion and failure messages are sent to:

- `helpdesk@iconis.co.za`
- `dev@iconis.co.za`

The SMTP sender must be authorized to relay to those recipients. A Mimecast `451 Open relay not allowed` response means the CSV/FTP work may have completed, but the email notification was rejected and the process returns a failure status.

## Normal workflow

1. Load `.env`, stage definitions, and field mappings.
2. Select the previous calendar month by default using Africa/Johannesburg time.
3. List eligible Panel L and Debt Review handover workbooks on FTP.
4. Download each workbook into the local `downloads_sbsa/` cache.
5. Read client code and account/reference values.
6. Find the matching LegalSuite matter by client and reference.
7. Require the matter to be newly created from the matching handover and satisfy the configured eligibility rules.
8. Retrieve matching file notes and group them by MT stage code.
9. Enrich each row using the confirmed LegalSuite field sources and mapping workbook.
10. For every MT ID, inspect the newest completed CSV already stored in that MT ID's FTP folder.
11. Skip the MT ID if its current rows contain no new file-note rows and its headers are unchanged.
12. Otherwise generate a timestamped CSV and upload it atomically through a temporary `.uploading` filename.
13. Send the completion summary after uploads. State is marked complete only after the required workflow succeeds.

The FTP comparison prevents the same file notes from being uploaded repeatedly on daily runs. A new MT ID, a new file note, or a header change is sufficient to trigger a new CSV upload.

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

The positional MatterID command runs the original single-matter JSON mode. `--handover-file` is local-only and therefore does not perform the FTP previous-CSV comparison or upload. Use `--force` only when a completed source must be rebuilt.

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
- `output_sbsa/processed_handover_files.json`: source hashes, completion state, generated paths, FTP paths, and pending counts; ignored by Git.
- `.env`: private credentials; ignored by Git.

If a matter is not found yet, the source remains pending and is retried on a later run. If email is rejected, FTP files may already exist but the source is intentionally not marked fully complete.
