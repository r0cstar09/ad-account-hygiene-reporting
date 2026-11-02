# Collector: behavior, evidence and limitations

## Prerequisites and boundaries

`Export-InactiveADUsers.ps1` runs under Windows PowerShell 5.1 with the RSAT ActiveDirectory module, permitted network access and authorized reads of the required attributes. It uses the current identity unless a `PSCredential` is supplied. It never embeds credentials or requests account changes.

Only enabled user objects are selected. Service identities implemented as user objects can appear. This is not an inventory of all AD object classes, managed service account types, computers, application identities or Entra identities.

The scope is a single current domain. `-SearchBase` must be that domain naming context or an OU beneath it. The script does not traverse a forest, query cloud sign-ins, recover retired-DC history or reconstruct an historical directory snapshot.

## Output contract

The workbook has a single `Report` worksheet: domain at A1, run timestamp at K1, a merged A2:K2 report title, headers on row 3, records beginning on row 4 and frozen headings at A4. Data cells are text, including apparent numeric/date fields; K1 is an Excel date. Fonts, colors, column widths and the title layout follow the reference export style. Equivalent XML serialization does not mean a byte-identical vendor file.

Column mappings, in order:

1. **Disabled** — blank for these enabled-account rows; disabled accounts are not exported.
2. **Display Name** — `displayName`.
3. **SAM Account Name** — `sAMAccountName`.
4. **Distinguished Name** — `distinguishedName`.
5. **Canonical Name** — AD's `canonicalName` property.
6. **Account Expires On** — `accountExpires`; zero/maximum integer is displayed as `Never`.
7. **Password Last Set** — `pwdLastSet`; unset values are blank.
8. **When Created** — `whenCreated`.
9. **Last Logon Timestamp** — source-DC `lastLogonTimestamp`, independently displayed.
10. **Days Since Last Logon** — by default, days from the greatest observed `lastLogon` across discovered DCs.
11. **Logon Count** — by default, the sum of observed per-DC `logonCount` values.

## Why selection and displayed age differ

The default membership rule is enabled AND (`lastLogonTimestamp` is absent/zero OR it is at or before the captured run clock minus `Days`). The comparison is inclusive and uses unrounded UTC FILETIME values. A missing timestamp can include a newly created account; account age is not an additional default exclusion.

The displayed age separately uses maximum `lastLogon`. Microsoft documents `lastLogon` as non-replicated and recommends retrieving every DC's value and taking the largest.[1] `lastLogonTimestamp` updates subject to `msDS-LogonTimeSyncInterval`, so it is not a real-time record of every logon.[2]

A row can therefore meet a timestamp-based inactivity filter yet show a more recent displayed last-logon age. The Python stage keeps those rows visible in a below-threshold review band. This compatibility default is a documented inference from the report's behavior, not a claim to know the proprietary product's historical configuration or hidden implementation.

For cleanup review rather than layout/selection compatibility:

```powershell
.\Export-InactiveADUsers.ps1 -Days 90 `
    -InactivityBasis LastLogon -UseTimestampSafetyNet `
    -DayCalculation ElapsedDays `
    -OutputPath 'C:\Reports\Inactive90-Review.xlsx'
```

This filters on maximum `lastLogon`, using a newer replicated timestamp as a selection safety net. It intentionally changes membership. The safety net does not silently rewrite either displayed attribute.

## Collection mechanics

- Discover and pin a source DC for the baseline enabled-user query and replicated/base attributes.
- Use one paged user query per required DC, not an individual query for each user on every DC.
- Join results using `objectGUID`, not display name or SAM account name.
- When `LastLogon` is needed or count mode is `Sum`, query every discovered DC, including discovered read-only controllers.
- Reject empty/incomplete discovery, duplicate GUIDs, failed required queries or missing baseline objects. Resolve replication, scope or read issues and rerun instead of ignoring a DC.
- Apply selection after aggregation. The scan is not an atomic directory snapshot; account state can change during collection.

A successful query does not prove every requested attribute was readable. Zero/missing values require interpretation, not a confident assumption that an account has never authenticated. Microsoft defines zero `lastLogon` as unknown; the source-compatible `Never` label should be read in that light.[1]

`logonCount` is also non-replicated. Summing current DCs omits retired-controller history, and the per-DC counter has a documented upper limit. It is not a reliable lifetime authentication audit trail.[3]

## Important parameters

- `-Days 90`: inactivity cutoff; supported range is 1–36500.
- `-Server dc01.example.test`: source discovery connection.
- `-SearchBase 'OU=Users,DC=example,DC=test'`: approved within-domain scope.
- `-Credential $cred`: obtain `$cred` with `Get-Credential` in your own authorized interactive environment; never put a password in code or command arguments.
- `-InactivityBasis LastLogonTimestamp|LastLogon`: row-selection attribute.
- `-DaysSinceBasis LastLogon|LastLogonTimestamp`: displayed-age attribute.
- `-LogonCountMode Sum|SourceDC`: aggregation scope for the count.
- `-UseTimestampSafetyNet`: when selection uses `LastLogon`, use a newer timestamp for selection only.
- `-DayCalculation NearestDays|CalendarDays|ElapsedDays`: display rounding, not cutoff rounding. Default nearest-day arithmetic uses explicit midpoint-to-even rounding.
- `-TimeZoneId UTC`: display zone; default is the local machine zone. Windows also accepts names such as `Eastern Standard Time`.
- `-AsOf`: fixed comparison clock for testing. It does **not** retrieve historical AD state.
- `-SortBySamAccountName`: optional deterministic sort; baseline enumeration order is otherwise retained.
- `-Force`: permit replacement of an existing output/audit target.

Selecting timestamp for both membership and displayed days, with `SourceDC` count mode, avoids the cross-DC scan. That faster mode is not equivalent to the default report.

## Files, failure handling and safety

The collector writes the XLSX to a same-directory temporary file and commits it after successful completion. A generation failure preserves the existing workbook. Directory strings are explicitly serialized as literal text; formula-like names are not Excel formulas. Invalid XML text and overlong strings fail rather than being silently corrupted.

An adjacent `.xlsx.audit.json` records run time, domain/scope, controllers, semantic options, result counts, inconsistencies and the workbook SHA-256. Treat it as confidential too. The audit is written **after** workbook commit: these are not an atomic pair. If the audit write fails, the command reports an error and a valid new workbook may already exist. Resolve that failure before treating the run as complete.

The script does not change execution policy, install RSAT, parallelize directory queries or modify accounts. Use an approved reporting window and pilot scope for large domains.

## Running the tests

The test suite covers aggregation, date boundaries, empty reports, text handling and file replacement. Run it in a fresh PowerShell process: it mocks `Import-Module` and the AD cmdlets, so that session should not be reused for directory administration.

## Sources

[1] https://learn.microsoft.com/en-us/windows/win32/adschema/a-lastlogon
[2] https://learn.microsoft.com/en-us/windows/win32/adschema/a-lastlogontimestamp
[3] https://learn.microsoft.com/en-us/windows/win32/adschema/a-logoncount
