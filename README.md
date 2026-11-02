# AD Account Hygiene Reporting

A read-only PowerShell collector and a Python report processor for turning inactive Active Directory account exports into reviewable remediation queues.

**Collect → normalize → separate account types → prioritize → human approval.**

This project creates evidence for a cleanup decision. It does **not** disable, delete, reset, or otherwise modify directory accounts. Account type and age are triage signals, not proof an account is safe to remove.

## What is included

- `Export-InactiveADUsers.ps1`: queries enabled AD users and writes a MaxPowerSoft-style, 11-column XLSX plus an audit JSON sidecar. No Excel or ImportExcel installation required.
- `ad_account_hygiene_report.py`: normalizes the export and produces account-category sheets, category/priority summaries, and a dashboard with charts.
- `sample_config.json`: fictional OU roots and generic naming rules. Copy it to ignored `private_rules.json` for local configuration.
- `samples/make_sample_report.py`: generates deterministic **synthetic** input locally. It never reads production data.
- `Test-Export-InactiveADUsers.ps1` and `tests/`: offline regression tests.

No real directory exports, employer-specific mappings, or workplace screenshots belong in this repository. Generated workbooks are ignored. Any demonstration images use fictional data.

## 1. Collect the report on Windows

Use Windows PowerShell 5.1, the RSAT `ActiveDirectory` module, network access to the relevant domain controllers, and an identity authorized to read the required attributes. Administrative group membership is not the intended prerequisite; use the least privilege your directory policy permits. Installation/signing/execution-policy changes are deliberately outside the script.

```powershell
Get-Module -ListAvailable ActiveDirectory
.\Export-InactiveADUsers.ps1 -Days 90 -OutputPath 'C:\Reports\Inactive90.xlsx' -Verbose
```

For an operational cleanup review, prefer the explicit, more conservative selection mode:

```powershell
.\Export-InactiveADUsers.ps1 -Days 90 `
    -InactivityBasis LastLogon -UseTimestampSafetyNet `
    -DayCalculation ElapsedDays `
    -OutputPath 'C:\Reports\Inactive90-Review.xlsx'
```

Default mode prioritizes compatibility with the source report format: it selects on `lastLogonTimestamp`, but displays days based on maximum `lastLogon` across discovered DCs. Consequently a selected row can show fewer than 90 days. That is not a reason to silently label it 90+ or remove it from the review data. See [collector semantics and parameters](docs/collector.md).

`-Server`, `-SearchBase`, `-Credential`, `-TimeZoneId` and other options allow controlled scope. A whole-domain run covers one domain, **not an entire forest**. Failed required DC queries stop the report rather than substituting `Never`/zero data.

## 2. Install the Python processor

Python 3.10+; Windows, macOS or Linux. The Python stage reads a local workbook and does not connect to AD.

```bash
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows PowerShell instead: .\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Some systems name the interpreter `python3`; use that name consistently when creating the environment.

## 3. Try the complete offline demonstration

From the repository root, inside the virtual environment:

```bash
python samples/make_sample_report.py
python ad_account_hygiene_report.py output/synthetic-input.xlsx output/synthetic-review.xlsx --config sample_config.json
```

Both workbooks are generated into ignored `output/`. All account names, paths, dates and counts in the demonstration are fabricated. This validates spreadsheet processing, **not live directory collection**.

## 4. Configure and process an authorized export

Copy `sample_config.json` to `private_rules.json`, then edit the copy for your environment. Keep it outside version control. Canonical roots are path prefixes such as `example.test/Users`, not distinguished names such as `OU=Users,DC=example,DC=test`.

```bash
python ad_account_hygiene_report.py /secure/Inactive90.xlsx /secure/Inactive90-Review.xlsx --config private_rules.json
```

Use distinct input and output paths. The processor preserves the input, rejects an output alias of the input, and stages the new workbook before replacement. An existing output at a different path is replaced after successful generation; keep dated report names if history matters.

## What the output means

The four primary review queues are **Workforce Accounts**, **Privileged Accounts**, **Service Accounts**, and **Other Accounts**. Naming/path rules assign provisional categories, with service-account clues considered before the broad OU classification. They are **not** a group-membership or effective-permissions audit; verify privilege and ownership separately.

`All Review Items` preserves the processed population. Primary queues partition that population; summary sheets are not extra accounts. The dashboard shows category and priority distributions, with a separate OU/account-class breakdown.

The default prioritization is an age-based review policy:

- Non-service accounts with a source `Never`/`Never Logged On` sentinel enter `Critical - Source Reports Never Logged On`.
- Non-service ages of at least 365 days enter `High`; 180–364 enter `Medium`; 90–179 enter `Low`.
- Service accounts use separate `Service Review` priorities, with explicit source-never and unknown-age variants, rather than the workforce age labels.
- Values below 90 days are labeled separately, not incorrectly reported as 90+.
- Missing, negative or invalid day values require investigation rather than being guessed as `Never`.

A label such as `Critical` is a **review priority**, not an exploitability score, incident finding or disable instruction. A new account can legitimately have no recorded logon. The defaults are not a substitute for your retention, emergency-access, HR and change-management policies.

See [review and remediation playbook](docs/remediation.md) before acting on a result.

## Safety and scope

- Read-only directory access; report generation only.
- No credentials embedded in code or example commands.
- PowerShell serializes directory values as literal strings; Python output disables automatic formula and URL conversion.
- Unknown data stays visible. Unexpected input headers fail clearly.
- `Never` is a vendor-compatible display convention, not proof of no historical authentication.
- AD attributes are not a complete lifetime authentication log. Entra/cloud-only, application-specific and cached offline activity are not established by this report.
- Successful queries do not guarantee attribute-level permissions are complete, nor an atomic snapshot across DCs.
- Operational reports and audit sidecars contain sensitive identifiers. Restrict their storage, access and retention.

## Verification

Run each PowerShell fixture suite in a **fresh process**; it deliberately mocks directory cmdlets and must not be dot-sourced into an administration session:

```powershell
powershell.exe -NoProfile -File .\Test-Export-InactiveADUsers.ps1
```

Run the Python tests:

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

The tests exercise synthetic fixtures, workbook parsing/writing, classification, boundary conditions and failure paths. Offline passes do not establish behavior against a real organization's permissions, topology or historical directory state. Pilot with an approved scope and confirm findings with authoritative owners before using the output operationally.

## Documentation

- [Collector, attribute mapping and limitations](docs/collector.md)
- [Human-reviewed remediation process](docs/remediation.md)
- [Privacy and reporting policy](SECURITY.md)

This is an independent workflow compatible with a reporting layout, not a MaxPowerSoft product or endorsement.
