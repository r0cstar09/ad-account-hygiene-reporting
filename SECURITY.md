# Security and privacy

Do not attach real AD exports, processed workbooks, audit JSON, private classification rules, employee identifiers or internal paths to issues, pull requests or releases. Reproduce bugs with `samples/make_sample_report.py` or a minimal fictional fixture.

The repository ignores common data formats, but `.gitignore` is not a confidentiality control and does not remove files already tracked. Inspect staged contents and history before publishing; use an explicit source-file allowlist.

Only use the collector with authorization. Review scripts, follow local execution/signing policies and apply least-privilege read access. Protect generated reports and audit sidecars as sensitive directory inventories.

Classification is heuristic and priorities are review bands. Neither script performs account remediation. Confirm ownership, privilege, dependencies and an approved recovery plan before changing any account.

Never paste credentials into an issue. For a suspected sensitive disclosure, contact the repository owner privately rather than creating a public proof-of-concept containing data.
