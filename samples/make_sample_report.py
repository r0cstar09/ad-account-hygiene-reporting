#!/usr/bin/env python3
"""Generate a synthetic MaxPowerSoft-style inactive users report for demos/tests."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "sample_inactive_enabled_users.xlsx"

columns = [
    "Disabled",
    "Display Name",
    "SAM Account Name",
    "Account Expires On",
    "Password Last Set",
    "When Created",
    "Last Logon Timestamp",
    "Days Since Last Logon",
    "Logon Count",
    "Distinguished Name",
    "Canonical Name",
]

rows = [
    [False, "Alex Rivera", "arivera", "Never", "2025-01-15", "2024-12-01", "Never", "Never Logged On", 0, "CN=Alex Rivera,OU=SiteA,OU=Accounts,OU=North,OU=Users,DC=example,DC=local", "example.local/Users/North/Accounts/SiteA/Alex Rivera"],
    [False, "Jordan Lee", "jlee", "Never", "2024-02-01", "2022-01-12", "2025-05-01", 435, 12, "CN=Jordan Lee,OU=SiteB,OU=Accounts,OU=South,OU=Users,DC=example,DC=local", "example.local/Users/South/Accounts/SiteB/Jordan Lee"],
    [False, "Privileged Ops", "admin.ops", "Never", "2023-10-01", "2021-06-10", "2025-09-15", 297, 24, "CN=Privileged Ops,OU=Admins,OU=Privileged,DC=example,DC=local", "example.local/Privileged/Admins/Operations/Privileged Ops"],
    [False, "SQL Service", "svc-sql-reporting", "Never", "2022-03-14", "2020-11-04", "2025-04-30", 436, 240, "CN=SQL Service,OU=Service Accounts,OU=Applications,DC=example,DC=local", "example.local/Applications/Service Accounts/Reporting/SQL Service"],
    [False, "Kiosk Front Desk", "kiosk-frontdesk-01", "Never", "2025-02-11", "2024-07-19", "2025-11-20", 232, 88, "CN=Kiosk Front Desk,OU=Devices,OU=Other,DC=example,DC=local", "example.local/Other/Devices/SiteC/Kiosk Front Desk"],
    [False, "Morgan Chen", "mchen", "Never", "2025-04-01", "2024-04-01", "2026-01-01", 90, 17, "CN=Morgan Chen,OU=SiteD,OU=Accounts,OU=East,OU=Users,DC=example,DC=local", "example.local/Users/East/Accounts/SiteD/Morgan Chen"],
]

# MaxPowerSoft-style title rows, then a real header row where first cell is Disabled.
title = pd.DataFrame([
    ["Inactive Enabled Users", "", "", "", "", "", "", "", "", "", ""],
    ["Generated from synthetic demo data", "", "", "", "", "", "", "", "", "", ""],
])
header = pd.DataFrame([columns])
data = pd.DataFrame(rows)
combined = pd.concat([title, header, data], ignore_index=True)
OUT.parent.mkdir(parents=True, exist_ok=True)
combined.to_excel(OUT, index=False, header=False)
print(OUT)
