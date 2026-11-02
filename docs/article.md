---
title: "Cleaning Up Inactive AD Accounts with PowerShell and Python"
description: "How I built a PowerShell and Python workflow to turn inactive Active Directory exports into separate, prioritized account-cleanup queues."
publishDate: "2026-11-02"
tags: ["Active Directory", "PowerShell", "Python", "Identity Security", "Security Automation"]
draft: false
---

Getting a list of inactive Active Directory accounts is the easy part. The work starts when you have to decide what to do with each one.

A stale employee account, an administrator account and an application service account can all land in the same export. They need different owners, different checks and different cleanup decisions. A spreadsheet sorted by “days since last logon” doesn't give you that separation.

I built this project to handle the repetitive part: collect the account data with PowerShell, then use Python to clean up the export and organize it into review queues. The result is a workbook you can work through with the people responsible for those accounts, rather than another long list to sort by hand.

[The scripts and setup instructions are on GitHub.](https://github.com/r0cstar09/ad-account-hygiene-reporting)

*The screenshots use sample data. Employer records and internal account details are not included.*

## Start with the AD export

The PowerShell script, `Export-InactiveADUsers.ps1`, queries enabled user accounts and produces an Excel report in the familiar MaxPowerSoft-style layout. It keeps the fields needed for the review:

- Display name and SAM account name
- Distinguished and canonical directory paths
- Account expiry, creation date and password-last-set date
- Last logon timestamp, days since last logon and logon count

It also keeps the merged report title, frozen headings and date formatting. The script writes the XLSX using built-in .NET functionality, so it doesn't need Excel installed or the ImportExcel module.

Run it from Windows PowerShell with the RSAT ActiveDirectory module and an account that has permission to read the required attributes:

```powershell
.\Export-InactiveADUsers.ps1 -Days 90 `
    -InactivityBasis LastLogon -UseTimestampSafetyNet `
    -DayCalculation ElapsedDays `
    -OutputPath 'C:\Reports\Inactive90-Review.xlsx'
```

![Sample AD export showing account names, logon dates and inactivity ages.](images/01-raw-export.png)

*The raw export, with some columns hidden to make the example easier to read.*

### Get the logon dates right

There are two AD attributes worth paying attention to here. `lastLogon` is maintained separately on each domain controller. Getting the latest recorded value means querying the DCs and taking the largest value.[1] `lastLogonTimestamp` is replicated, but its update interval means it can lag behind recent activity.[2]

The script's default compatibility mode selects accounts using `lastLogonTimestamp`, while the displayed age uses the latest `lastLogon`. That's why an account can appear in a timestamp-based 90-day report even when the displayed age is below 90 days.

The command above selects on the latest observed `lastLogon` and also checks for a newer replicated timestamp. It uses complete elapsed days for the displayed age. These options keep the cleanup view separate from the settings used to reproduce the original export layout and selection behavior.

The collector stops if a required DC cannot be read or a baseline account is missing from a controller's results. It also writes an audit file alongside the workbook with the scope, DCs, settings and output hash.

## Turn the export into work queues

The Python script handles the second stage. It reads the workbook, finds the report headers, normalizes the values and separates the accounts into four sheets:

- **Workforce Accounts** — reconcile these with employment status, leave and the normal joiner/mover/leaver process.
- **Privileged Accounts** — review the account's access and confirm whether it is still needed.
- **Service Accounts** — identify the application or system owner and check dependencies before making a change.
- **Other Accounts** — keep unmatched records visible instead of dropping them from the review.

The classification uses configurable directory paths and naming rules. A service-account match takes precedence over a broad workforce OU match, so an application identity doesn't stay buried in the employee queue simply because of where it was created.

Each service-account row includes the reason it was classified that way. That makes it much easier to spot a bad rule or an account that needs a second look. Actual group memberships and permissions still belong in the access review; an OU name alone doesn't tell the whole story.

```bash
python ad_account_hygiene_report.py \
    /secure/Inactive90.xlsx \
    /secure/Inactive90-Review.xlsx \
    --config private_rules.json
```

The repository includes a generic configuration file. Organization-specific OU paths and naming conventions go in `private_rules.json`, which stays out of Git.

![Dashboard showing account categories, review priorities and OU summaries using sample data.](images/02-dashboard.png)

*The dashboard gives an overview; the separate account sheets contain the rows to work through.*

## Make the priorities useful

For non-service accounts, the processor assigns these review bands:

- **Critical:** the source reports `Never` or `Never Logged On`.
- **High:** 365 days or more.
- **Medium:** 180–364 days.
- **Low:** 90–179 days.

Accounts with more recent activity get a separate below-threshold label. Missing or invalid ages go into a needs-review band rather than being treated as zero or “never.” Service accounts have their own `Service Review` priorities.

The `Never` group needs a little care. A newly created account may be waiting for its first use, and Microsoft defines a zero `lastLogon` value as unknown.[1] Creation date and account ownership help distinguish those cases from accounts that have simply been forgotten.

Logon count is useful context, but not a lifetime history. The counters are maintained per DC, have a per-controller limit and don't recover activity recorded by retired controllers.[3]

## Keep service accounts out of the employee cleanup pile

This is the separation I care about most. An employee account can be checked against HR and a manager. A service account needs someone who understands what runs under it.

![Separate service-account sheet showing sample identities, priorities and classification reasons.](images/03-service-review.png)

*Service accounts get their own queue, with the matching rule shown alongside each account.*

The next steps are straightforward: identify the owner, check services and scheduled tasks, review integrations and relevant application logs, then agree on the change. Emergency-access accounts, seasonal use and planned provisioning need their own exceptions.

The scripts stop at reporting. Account changes go through the normal approval process, with a recovery path and a follow-up check for anything that depends on the identity. That keeps the collection repeatable without turning an inactivity report into a bulk-disable command.

## Less spreadsheet work, clearer decisions

The useful output isn't just the dashboard. It's having employee accounts, privileged accounts and service identities in separate queues, with the information needed to decide what happens next.

PowerShell handles the collection. Python handles the cleanup and organization. The account owner can then spend time on the decision instead of trying to make sense of a raw export.

## Sources

[1] https://learn.microsoft.com/en-us/windows/win32/adschema/a-lastlogon
[2] https://learn.microsoft.com/en-us/windows/win32/adschema/a-lastlogontimestamp
[3] https://learn.microsoft.com/en-us/windows/win32/adschema/a-logoncount
