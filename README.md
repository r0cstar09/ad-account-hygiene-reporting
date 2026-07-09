# AD Account Hygiene Reporting

Portfolio-safe Python project for turning a raw Active Directory inactive-user export into an actionable review workbook.

The tool was designed around MaxPowerSoft-style AD reports, but the public version is intentionally generic and configurable. Organization-specific domains, OU names, vendors, and account examples should live in a private config file that is **not committed**.

## What it does

- Detects the real header row in title-heavy AD Excel exports.
- Parses `Canonical Name` into domain/root OU/path/site fields.
- Separates:
  - Workforce accounts
  - Privileged accounts
  - Service / non-human accounts
  - Other accounts
- Flags likely service accounts by path, SAM naming convention, and device/non-human keywords.
- Normalizes `Never Logged On` and days-since-last-logon values.
- Adds priority and recommended-action fields.
- Writes a formatted Excel workbook with:
  - dashboard KPIs
  - charts
  - filters
  - conditional formatting
  - separate review tabs

## Why it matters

Inactive enabled accounts are easy wins for identity-risk reduction. This converts a raw report into something a security/IAM team can actually action: what to review first, what needs owner validation, and what should be treated carefully because it may be a service account.

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run with synthetic sample data

```bash
python samples/make_sample_report.py
python ad_account_hygiene_report.py samples/sample_inactive_enabled_users.xlsx output/sample_action_review.xlsx
```

## Run with a real export

```bash
python ad_account_hygiene_report.py "Inactive Enabled Users.xlsx" "inactive_accounts_action_review.xlsx" --config private_rules.json
```

## Private config

Copy the sample config and edit it locally:

```bash
cp sample_config.json private_rules.json
```

`private_rules.json` is ignored by git.

## Public-safety note

Do **not** publish real reports or organization-specific rules. Keep internal domains, OU names, vendors, business units, and site codes out of the repo. This project includes only synthetic examples.

## Outputs

Workbook tabs:

- `Dashboard`
- `All Review Items`
- `Workforce Accounts`
- `Privileged Accounts`
- `Service Accounts`
- `Other Accounts`
- `OU Summary`

## Screenshots for portfolio/blog

Use the synthetic sample workbook only. Recommended blog framing:

> Built a Python tool that converts raw Active Directory inactive-account exports into a prioritized remediation workbook, separating human, privileged, and service accounts to reduce disablement risk.
