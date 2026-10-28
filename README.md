# AccessStaffSync

**A practical Microsoft Access + VBA prototype for synchronizing WhatsApp staff updates from Google Drive–synced Excel files.**

![VBA](https://img.shields.io/badge/Microsoft_Access-VBA-A4373A?style=flat-square) ![Platform](https://img.shields.io/badge/Windows-Access_%2B_Excel-0078D4?style=flat-square) ![Purpose](https://img.shields.io/badge/Scope-4%2C000--record_demo-007F7F?style=flat-square)

> **Status:** The source code, two Excel example batches, and cross-platform data-contract tests are included. The test model passes; **the actual VBA/Access application has NOT been executed on Windows in this environment**. The installer builds the runnable `.accdb` on a Windows machine with Microsoft Access and Excel. This is a demonstration using wholly synthetic records—not a claim of integration with the client's original database.

## What it does

```text
WhatsApp assistant
       │ creates uniquely named .xlsx files with stable EventIDs
       ▼
Google Drive synced folder on Windows PC
       │ checks files have settled (20 s by default)
       ▼
Microsoft Access: frmSyncTimer (60-second OnTimer)
       │ exclusive local runner lock
       ▼
Late-bound Excel reader → strict headers, UTC timestamps, row validation
       │
       ├─ Invalid or unknown staff → REJECTED (reported; never inserted)
       ├─ Replayed EventID          → DUPLICATE (not applied)
       ├─ Older field timestamp     → STALE (not overwritten)
       └─ Valid event               → DAO transaction
                                     ├─ Update existing tblStaff row
                                     └─ Insert tblImportEvents row (unique EventID)
                                             │
                                             ▼
                                tblImportFiles completion ledger
                                             │
                                        Archive .xlsx + write receipt
                                        (retries after interrupted output)
```

The importer is designed to be **idempotent**: if a batch is retried after a lock or interruption, already-committed events are recognized by their unique `EventID`. Each staff row + its event ledger insert occur inside the same DAO transaction. New records are **never created** from inbound update files.

## 5-minute Windows demo

**Prerequisites:** Windows 10/11, desktop Microsoft Access, desktop Microsoft Excel, PowerShell 5.1. For the installer only, enable Access → File → Options → Trust Center → Trust Center Settings → Macro Settings → **Trust access to the VBA project object model**. Enable content only for this *trusted, local demonstration*; disable programmatic VBA access after setup. You may need to unblock extracted files / use a Trusted Location.

Open PowerShell from the extracted project folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup\Install-Demo.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup\Run-Demo.ps1
```

The first script:

1. Creates `demo/StaffSyncDemo.accdb`, **refusing to overwrite any existing DB**.
2. Imports `vba/modStaffSync.bas`, creates the Access schema, seeds exactly **4,000 synthetic staff records**.
3. Creates an unbound `frmSyncTimer` form with `OnTimer = =SyncTick()` and `TimerInterval = 60000`, registered as the Access startup form.
4. Sets up `demo/inbox`, `demo/archive`, and `demo/receipts`.

The second script copies two realistic example Excel batches into the inbox, invokes the importer twice, and prints the resulting staff records, including invalid, stale, and duplicate-update cases. Then open `demo/StaffSyncDemo.accdb` in Access to inspect `tblStaff`, `tblImportEvents`, and `tblImportFiles`.

> **Important:** Close the demo database in Access before running `Run-Demo.ps1`, to avoid confusing simultaneous process behavior. The sample inbox uses `ReadyDelaySeconds=0` exclusively for immediate demonstration. Production must use a nonzero settle delay.

### Expected results from the two sample batches

| File | Rows | Applied | Stale | Duplicate | Rejected | Receipt |
|---|---:|---:|---:|---:|---:|---|
| `updates_001.xlsx` | 6 | 4 | 0 | 1 | 1 | PARTIAL |
| `updates_002.xlsx` | 7 | 3 | 1 | 1 | 2 | PARTIAL |

Both are intentionally `PARTIAL` because they contain bad or unrecognized rows; valid updates still commit. Change row data or use `updates_template.xlsx` to simulate a clean `SUCCESS` import.

**Result example (staff 000042):** `JobOfferStatus = Accepted`, `Availability = Available next week`, `Address = Carrer de Mallorca 10, Barcelona`, `MessagingConsent = True`. The older availability event from batch 2 is ignored, and the first batch's repeated EventID is not applied twice.

## Excel update contract

The first worksheet of each `.xlsx` must contain exactly these seven columns in this order:

| EventID | StaffID | EventTimeUTC | JobOfferStatus | Availability | Address | MessagingConsent |
|---|---|---|---|---|---|---|
| WA-1234 | 000042 | 2026-10-27T09:00:00Z | Offered | Available | 21 Main St | YES |

- `EventID`: stable unique ID of an individual WhatsApp change, **reused verbatim on redelivery** (up to 64 characters). Use a GUID or the WhatsApp automation's event UUID.
- `StaffID`: the Access staff primary key. **Send as text** to preserve leading zeros. Never identify staff by editable address/name alone.
- `EventTimeUTC`: **text** formatted `YYYY-MM-DDTHH:MM:SSZ` in UTC. The Access demo stores incoming UTC as naive `Date/Time` values (in UTC). Two conflicting updates to the same field should have distinct timestamps.
- Remaining fields: empty = **leave unchanged**; use literal `__CLEAR__` for clearing one of the three optional text fields. Consent accepts `YES/NO`, `TRUE/FALSE`, `1/0`, `SI/SÍ` (but no blank-to-NULL consent).
- The `MessagingConsent` YES/NO flag is authoritative only for the sample import schema. **Real privacy/compliance consent rules, provenance and withdrawal processing must be verified with the client**.
- Use **unique filenames** per exported batch. Do not overwrite an old batch under the same filename. Limit exports to 10,000 update rows per file.

### Error behavior

- Unknown StaffID / invalid timestamp / malformed consent: **rejected**, counted in the receipt; no staff insertion.
- Record locked (e.g. another Access operator is editing): skip this event for now; **do not finalize the file**; retry on the next 60-second interval. Events committed in an earlier attempt are safe to replay.
- Unexpected DAO errors: leave the batch pending for retry instead of silently dropping data; investigate `staffsync.log`.
- Import parser cannot open/read the workbook: log the failure and leave the file in the inbox. Invalid headers are never silently mapped to the wrong columns.
- A completed batch is recorded before the filesystem archive/receipt steps; if the app crashes between those steps, the next tick will finish them using the durable `tblImportFiles` ledger.
- An already-processed *filename* appearing again while the original is in the archive is held for manual review (not silently deleted).

## Source layout

```text
vba/modStaffSync.bas       600+ lines: Access timer, Excel reader, record locks,
                           event transactions, audit tables, receipts, demo seed
setup/Install-Demo.ps1     Creates .accdb from scratch using Access COM
setup/Run-Demo.ps1         Windows/Access end-to-end example runner
samples/inbox/*.xlsx      Excel contract template and two demonstration batches
qa/test_contract.py       Portable input/idempotence/staleness reference tests
qa/make_workbooks.py      Recreates sample XLSX files (artifact_tool dependency)
docs/CLIENT_HANDOFF.md    Feasibility review, integration checklist and assumptions
preview/index.html        Illustrative, read-only status dashboard for presentation
```

Cross-platform contract tests (Python 3, stdlib only):

```bash
python -m unittest discover -s qa -p 'test_*.py' -v
```

These tests check the input fixtures and a reference model; they **do not compile VBA**, execute Access transactions, simulate Access record locks, or test Google Drive integration. Those require the target Windows environment.

## Working with the client's existing database

**Do not run `InitializeDemo` or `SeedDemoStaff` on the live client database.** The real `tblStaff` schema, exact staff key format, Access version/bitness, existing forms, backend split, multi-user locking settings, privacy requirements, and WhatsApp export payloads have not been provided. A proper integration starts with a **sanitized copy and a mapping review**. See [CLIENT_HANDOFF.md](docs/CLIENT_HANDOFF.md).

For a production pilot: copy the existing DB, adapt the `tblStaff` field mapping and create only the audit/config tables not already present, configure the Drive-synced inbox path with `PutSetting`, and run all lock/restart/idempotence tests. Keep the Access process open (its form timer **does not run when Access is closed, the PC sleeps, or the database is not trusted**). For unattended SLA, consider a Windows Task Scheduler startup + watchdog or a service-based queue processor.

## Design limits

This demo deliberately prioritizes safe updates to 4,000 local Access records. It is **not** a generic migration tool or a production-certified consent system. No network API calls or WhatsApp/Google Drive credentials are included. Excel desktop automation is used (and thus must be installed), and native Access lock/retry behavior has not yet been Windows-tested. The in-process 60-second form timer requires Access to be open and responsive; it is not a durable Windows service. Timestamp ordering is per field, with equal timestamps resolved by arrival order; a production feed should provide strictly ordered event versions if equal-second conflicts are possible.

**License:** MIT. All demo staff and addresses are synthetic or illustrative.
