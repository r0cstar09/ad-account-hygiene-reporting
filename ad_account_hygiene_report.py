#!/usr/bin/env python3
"""Create an actionable inactive-account review workbook from AD report exports.

This is a public-safe, configurable version of an internal MaxPowerSoft-style
Active Directory report cleaner. It intentionally avoids organization-specific
OU names, domains, vendors, and account examples. Put private rules in a local
JSON config that is not committed.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pandas as pd

DEFAULT_CONFIG: dict[str, Any] = {
    "privileged_roots": ["Admin", "Admins", "Administrators", "Privileged"],
    "workforce_roots": ["Users", "Workforce", "Employees", "Corporate"],
    "service_path_keywords": ["service accounts", "service account", "non-human", "application accounts"],
    "service_name_patterns": [
        r"(^|[._\\-])svc([._\\-]|$)",
        r"(^|[._\\-])service([._\\-]|$)",
        r"^svc[._\\-]",
        r"^sa[._\\-]",
        r"^sql[._\\-]",
        r"^iis[._\\-]",
        r"^app[._\\-]",
        r"^api[._\\-]",
        r"^sync[._\\-]",
    ],
    "device_keywords": ["dvr", "camera", "printer", "scanner", "kiosk", "pos", "automation", "bot", "monitor"],
    "columns": {
        "canonical_name": "Canonical Name",
        "distinguished_name": "Distinguished Name",
        "sam_account_name": "SAM Account Name",
        "display_name": "Display Name",
        "account_expires": "Account Expires On",
        "password_last_set": "Password Last Set",
        "when_created": "When Created",
        "last_logon_timestamp": "Last Logon Timestamp",
        "days_since_last_logon": "Days Since Last Logon",
        "logon_count": "Logon Count",
    },
}

OUTPUT_COLUMNS = [
    "Category", "Original Category", "Root OU", "Account Class", "Site/OU", "OU Path",
    "Display Name", "SAM Account Name", "Account Expires On", "Password Last Set",
    "When Created", "Last Logon Timestamp", "Days Since Last Logon", "Logon Count",
    "Priority", "Recommended Action", "Service Account Reason", "Distinguished Name", "Canonical Name",
]

PRIORITY_ORDER = {
    "Critical - Source Reports Never Logged On": 0,
    "High - 365+ days": 1,
    "Medium - 180+ days": 2,
    "Low - 90+ days": 3,
    "Needs Review - Unknown Logon Age": 4,
    "Below Threshold - Under 90 days": 5,
    "Service Review - Source Reports Never Logged On": 6,
    "Service Review - Unknown Logon Age": 7,
    "Service Review": 8,
}

UNKNOWN_DAYS = "Unknown"
SOURCE_NEVER = "Source Reports Never Logged On"
SOURCE_NEVER_SENTINELS = {"never", "never logged on"}


def load_config(path: Path | None) -> dict[str, Any]:
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if not path:
        return config
    user_config = json.loads(path.read_text(encoding="utf-8"))
    for key, value in user_config.items():
        if isinstance(value, dict) and isinstance(config.get(key), dict):
            config[key].update(value)
        else:
            config[key] = value
    return config


def normalize_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def parse_canonical(canonical_name: Any, config: dict[str, Any]) -> dict[str, str]:
    """Parse Canonical Name like example.local/Users/Region/Accounts/Site/Jane Doe."""
    if canonical_name is None or pd.isna(canonical_name):
        return {"domain": "", "root_ou": "", "base_category": "Other Accounts", "ou_path": "", "account_class": "", "site_ou": ""}

    parts = normalize_text(canonical_name).split("/")
    domain = parts[0] if parts else ""
    path_parts = parts[1:-1] if len(parts) > 1 else []
    root_ou = path_parts[0] if path_parts else ""
    ou_path = "/".join(path_parts)
    lower_path = [p.lower() for p in path_parts]
    root_lower = root_ou.lower()

    privileged_roots = {x.lower() for x in config.get("privileged_roots", [])}
    workforce_roots = {x.lower() for x in config.get("workforce_roots", [])}
    service_terms = [x.lower() for x in config.get("service_path_keywords", [])]

    if root_lower in privileged_roots or "administrators" in lower_path:
        base_category = "Privileged Accounts"
    elif any(term in "/".join(lower_path) for term in service_terms):
        base_category = "Service Accounts"
    elif root_lower in workforce_roots:
        base_category = "Workforce Accounts"
    else:
        base_category = "Other Accounts"

    if base_category == "Workforce Accounts" and len(path_parts) > 1:
        account_class = path_parts[1]
    elif base_category == "Privileged Accounts" and len(path_parts) > 1:
        account_class = path_parts[1]
    else:
        account_class = root_ou

    return {
        "domain": domain,
        "root_ou": root_ou,
        "base_category": base_category,
        "ou_path": ou_path,
        "account_class": account_class,
        "site_ou": path_parts[-1] if path_parts else "",
    }


def normalize_days(value: Any) -> Any:
    """Normalize source evidence without treating missing data as a never-logon assertion."""
    if value is None or pd.isna(value) or isinstance(value, bool):
        return UNKNOWN_DAYS
    text = str(value).strip()
    if not text:
        return UNKNOWN_DAYS
    if " ".join(text.casefold().split()) in SOURCE_NEVER_SENTINELS:
        return SOURCE_NEVER
    try:
        number = Decimal(text)
    except InvalidOperation:
        return UNKNOWN_DAYS
    if not number.is_finite() or number < 0 or number != number.to_integral_value():
        return UNKNOWN_DAYS
    return int(number)


def priority_and_action(days: Any, is_service: bool) -> tuple[str, str]:
    """Assign evidence-aware review bands without inferring policy or account privilege."""
    if is_service:
        if days == SOURCE_NEVER:
            return "Service Review - Source Reports Never Logged On", "Validate the source sentinel, owner, purpose, host/service dependency, and interactive-logon requirements"
        if days == UNKNOWN_DAYS:
            return "Service Review - Unknown Logon Age", "Resolve the unknown logon age, then validate owner, purpose, host/service dependency, and interactive-logon requirements"
        return "Service Review", "Validate owner, purpose, host/service dependency, and remove if no longer required"
    if days == SOURCE_NEVER:
        return "Critical - Source Reports Never Logged On", "Validate the source sentinel against authoritative logon evidence, then confirm need before disablement"
    if days == UNKNOWN_DAYS:
        return "Needs Review - Unknown Logon Age", "Resolve the missing or invalid logon age before assigning a stale-account action"
    if isinstance(days, int) and days >= 365:
        return "High - 365+ days", "Stale enabled account; verify with owner or manager before disablement"
    if isinstance(days, int) and days >= 180:
        return "Medium - 180+ days", "Review business need and prepare disablement if no exception"
    if isinstance(days, int) and days >= 90:
        return "Low - 90+ days", "Monitor, validate, or stage for normal stale-account review"
    return "Below Threshold - Under 90 days", "Below the 90-day stale-age threshold; validate report scope and monitor as appropriate"


def load_ad_report(path: Path) -> pd.DataFrame:
    """Load and validate a title-prefixed report without guessing malformed headers."""
    raw = pd.read_excel(path, header=None, dtype=object)
    if raw.empty or raw.shape[1] == 0:
        raise ValueError("Input workbook is empty; expected title rows and a report header row.")
    first_col = raw.iloc[:, 0].map(normalize_text).str.casefold()
    header_rows = raw.index[first_col.eq("disabled")]
    if len(header_rows) == 0:
        raise ValueError("Could not find the report header row. Expected first header cell to be 'Disabled'.")
    idx = int(header_rows[0])
    headers = [normalize_text(value) for value in raw.iloc[idx].tolist()]
    keep_positions: list[int] = []
    kept_headers: list[str] = []
    for position, header in enumerate(headers):
        if header:
            keep_positions.append(position)
            kept_headers.append(header)
        elif raw.iloc[idx + 1 :, position].notna().any():
            raise ValueError(f"Report header row contains a blank name for populated column {position + 1}.")
    duplicates = sorted({header for header in kept_headers if kept_headers.count(header) > 1})
    if duplicates:
        raise ValueError(f"Report header row contains duplicate column names: {duplicates}")
    if not kept_headers:
        raise ValueError("Report header row contains no named columns.")
    df = raw.iloc[idx + 1 :, keep_positions].copy()
    df.columns = kept_headers
    return df.dropna(how="all").reset_index(drop=True)


def has_keyword(text: str, keyword: Any) -> bool:
    """Match configured terms at Unicode word boundaries to avoid short substring collisions."""
    term = normalize_text(keyword)
    return bool(term and re.search(rf"(?<!\w){re.escape(term)}(?!\w)", text, flags=re.IGNORECASE))


def service_account_reason(row: pd.Series, config: dict[str, Any]) -> str:
    """Return a configurable heuristic reason, not proof of account type or privilege."""
    cols = config["columns"]
    path_values = [
        row.get(cols["canonical_name"], ""),
        row.get(cols["distinguished_name"], ""),
        row.get("ou_path", ""),
        row.get("account_class", ""),
        row.get("root_ou", ""),
    ]
    values = [
        *path_values,
        row.get(cols["sam_account_name"], ""),
        row.get(cols["display_name"], ""),
    ]
    path_haystack = " | ".join(normalize_text(v) for v in path_values)
    haystack = " | ".join(normalize_text(v).lower() for v in values)
    sam = normalize_text(row.get(cols["sam_account_name"], "")).lower()

    for term in config.get("service_path_keywords", []):
        if has_keyword(path_haystack, term):
            return f"Service-account path keyword: {term}"
    for pattern in config.get("service_name_patterns", []):
        if re.search(pattern, sam):
            return f"Service-style SAM naming convention: {pattern}"
    for term in config.get("device_keywords", []):
        if has_keyword(haystack, term):
            return f"Likely non-human/device account keyword: {term}"
    return ""


def build_clean_frame(input_file: Path, config: dict[str, Any]) -> pd.DataFrame:
    """Build the four-category review table after structural and evidence validation."""
    df = load_ad_report(input_file)
    cols = config["columns"]
    required = [cols[k] for k in ["canonical_name", "sam_account_name", "days_since_last_logon"]]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required input columns: {missing}. Found: {list(df.columns)}")
    if df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    parsed = df[cols["canonical_name"]].apply(lambda v: parse_canonical(v, config)).apply(pd.Series)
    df = pd.concat([parsed, df], axis=1)
    df["Service Account Reason"] = df.apply(lambda r: service_account_reason(r, config), axis=1)
    df["Category"] = df.apply(lambda r: "Service Accounts" if r["Service Account Reason"] or r["base_category"] == "Service Accounts" else r["base_category"], axis=1)
    df["Days Since Last Logon - Clean"] = df[cols["days_since_last_logon"]].apply(normalize_days)
    priority = df.apply(lambda r: priority_and_action(r["Days Since Last Logon - Clean"], r["Category"] == "Service Accounts"), axis=1)
    df["Priority"] = priority.apply(lambda x: x[0])
    df["Recommended Action"] = priority.apply(lambda x: x[1])

    def optional_column(name: str) -> pd.Series:
        return df[name] if name in df.columns else pd.Series("", index=df.index, dtype=object)

    clean = pd.DataFrame({
        "Category": df["Category"],
        "Original Category": df["base_category"],
        "Root OU": df["root_ou"],
        "Account Class": df["account_class"],
        "Site/OU": df["site_ou"],
        "OU Path": df["ou_path"],
        "Display Name": optional_column(cols["display_name"]),
        "SAM Account Name": optional_column(cols["sam_account_name"]),
        "Account Expires On": optional_column(cols["account_expires"]),
        "Password Last Set": optional_column(cols["password_last_set"]),
        "When Created": optional_column(cols["when_created"]),
        "Last Logon Timestamp": optional_column(cols["last_logon_timestamp"]),
        "Days Since Last Logon": df["Days Since Last Logon - Clean"],
        "Logon Count": optional_column(cols["logon_count"]),
        "Priority": df["Priority"],
        "Recommended Action": df["Recommended Action"],
        "Service Account Reason": df["Service Account Reason"],
        "Distinguished Name": optional_column(cols["distinguished_name"]),
        "Canonical Name": optional_column(cols["canonical_name"]),
    })
    clean["Priority Sort"] = clean["Priority"].map(PRIORITY_ORDER).fillna(9)
    return clean.sort_values(["Priority Sort", "Category", "Site/OU", "SAM Account Name"]).drop(columns=["Priority Sort"])


def autosize_columns(worksheet, df: pd.DataFrame, max_width: int = 52) -> None:
    for idx, col in enumerate(df.columns):
        series = df[col].map(lambda value: "" if pd.isna(value) else str(value)) if len(df) else pd.Series(dtype=str)
        width = min(max(series.map(len).max() if len(series) else 0, len(str(col))) + 2, max_width)
        worksheet.set_column(idx, idx, width)


def add_table_sheet(writer, sheet_name: str, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    safe_name = sheet_name[:31]
    df.to_excel(writer, sheet_name=safe_name, index=False)
    ws = writer.sheets[safe_name]
    ws.freeze_panes(1, 0)
    ws.set_row(0, 22, formats["header"])
    autosize_columns(ws, df)
    rows, cols = df.shape
    if rows > 0 and cols > 0:
        ws.autofilter(0, 0, rows, cols - 1)
        for col_name, width in {"Priority": 30, "Recommended Action": 58, "Service Account Reason": 55, "Distinguished Name": 70, "Canonical Name": 70}.items():
            if col_name in df.columns:
                ws.set_column(df.columns.get_loc(col_name), df.columns.get_loc(col_name), width, formats.get("wrap"))
        if "Priority" in df.columns:
            pcol = df.columns.get_loc("Priority")
            for label, fmt in [("Critical", "critical"), ("High", "high"), ("Medium", "medium"), ("Low", "low"), ("Service", "service")]:
                ws.conditional_format(1, pcol, rows, pcol, {"type": "text", "criteria": "containing", "value": label, "format": formats[fmt]})
        if "Days Since Last Logon" in df.columns:
            dcol = df.columns.get_loc("Days Since Last Logon")
            ws.conditional_format(1, dcol, rows, dcol, {"type": "data_bar", "bar_color": "#5B9BD5"})


def write_workbook(clean: pd.DataFrame, output_file: Path) -> None:
    """Atomically replace the destination only after a complete workbook closes successfully."""
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_file.name}.", suffix=".xlsx", dir=output_file.parent
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        _write_workbook_contents(clean, temporary_path)
        os.replace(temporary_path, output_file)
    finally:
        temporary_path.unlink(missing_ok=True)


def _write_workbook_contents(clean: pd.DataFrame, output_file: Path) -> None:
    """Write the dashboard and category workflow with literal-string XLSX handling."""
    category_summary = clean.groupby("Category", dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)
    priority_summary = clean.groupby("Priority", dropna=False).size().reset_index(name="Count")
    priority_summary["Priority Sort"] = priority_summary["Priority"].map(PRIORITY_ORDER).fillna(99)
    priority_summary = priority_summary.sort_values("Priority Sort").drop(columns=["Priority Sort"])
    ou_summary = clean.groupby(["Category", "Account Class", "Site/OU"], dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)
    top_ou = ou_summary.head(15).copy()
    top_ou["OU Label"] = top_ou["Category"].astype(str) + " | " + top_ou["Account Class"].astype(str) + " | " + top_ou["Site/OU"].astype(str)

    with pd.ExcelWriter(
        output_file,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm:ss",
        date_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_formulas": False, "strings_to_urls": False}},
    ) as writer:
        workbook = writer.book
        dashboard = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dashboard
        if workbook.worksheets()[0] is not dashboard:
            raise RuntimeError("Dashboard worksheet was not created first.")

        formats = {
            "title": workbook.add_format({"bold": True, "font_size": 18, "font_color": "#1F4E78"}),
            "subtitle": workbook.add_format({"font_size": 10, "font_color": "#666666"}),
            "header": workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "align": "center", "valign": "vcenter", "border": 1}),
            "section": workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "font_color": "#1F4E78", "border": 1}),
            "kpi_label": workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "align": "center"}),
            "kpi_value": workbook.add_format({"bold": True, "font_size": 16, "border": 1, "align": "center"}),
            "wrap": workbook.add_format({"text_wrap": True, "valign": "top"}),
            "critical": workbook.add_format({"bg_color": "#C00000", "font_color": "#FFFFFF", "bold": True}),
            "high": workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True}),
            "medium": workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "bold": True}),
            "low": workbook.add_format({"bg_color": "#E2F0D9", "font_color": "#375623"}),
            "service": workbook.add_format({"bg_color": "#DDEBF7", "font_color": "#1F4E78", "bold": True}),
            "note": workbook.add_format({"text_wrap": True, "valign": "top", "font_color": "#555555"}),
        }

        add_table_sheet(writer, "All Review Items", clean, formats)
        for category, sheet in [
            ("Workforce Accounts", "Workforce Accounts"),
            ("Privileged Accounts", "Privileged Accounts"),
            ("Service Accounts", "Service Accounts"),
            ("Other Accounts", "Other Accounts"),
        ]:
            add_table_sheet(writer, sheet, clean[clean["Category"].eq(category)], formats)
        add_table_sheet(writer, "OU Summary", ou_summary, formats)

        dashboard.hide_gridlines(2)
        dashboard.set_default_row(18)
        dashboard.set_column("A:A", 3)
        dashboard.set_column("B:G", 18)
        dashboard.set_column("H:H", 3)
        dashboard.set_column("I:P", 12)
        dashboard.merge_range("B2:G2", "Inactive Enabled Users Review", formats["title"])
        dashboard.merge_range("B3:G3", "Provisional account categories; human approval required before changes.", formats["subtitle"])
        dashboard.merge_range("B4:G4", "Priority and category KPIs overlap; do not add them together.", formats["subtitle"])

        kpis = [
            ("Total Review Items", len(clean)),
            ("Critical", int((clean["Priority"] == "Critical - Source Reports Never Logged On").sum())),
            ("High", int((clean["Priority"] == "High - 365+ days").sum())),
            ("Medium", int((clean["Priority"] == "Medium - 180+ days").sum())),
            ("Privileged", int((clean["Category"] == "Privileged Accounts").sum())),
            ("Service Accounts", int((clean["Category"] == "Service Accounts").sum())),
        ]
        for i, (label, value) in enumerate(kpis):
            col = 1 + i
            dashboard.write(5, col, label, formats["kpi_label"])
            dashboard.write(6, col, value, formats["kpi_value"])
        dashboard.merge_range("B9:G9", "Review guidance", formats["section"])
        dashboard.merge_range("B10:G12", "Review Critical/High non-service items first, but validate source evidence. Service accounts need owner and dependency checks. Labels are not disable instructions.", formats["note"])

        data_start = 13
        dashboard.merge_range(data_start, 1, data_start, 5, "Category", formats["header"])
        dashboard.write(data_start, 6, "Count", formats["header"])
        for r, row in enumerate(category_summary.itertuples(index=False), data_start + 1):
            dashboard.merge_range(r, 1, r, 5, row.Category)
            dashboard.write(r, 6, int(row.Count))

        priority_start = 20
        dashboard.merge_range(priority_start, 1, priority_start, 5, "Review priority", formats["header"])
        dashboard.write(priority_start, 6, "Count", formats["header"])
        for r, row in enumerate(priority_summary.itertuples(index=False), priority_start + 1):
            dashboard.merge_range(r, 1, r, 5, row.Priority)
            dashboard.write(r, 6, int(row.Count))

        ou_start = max(33, priority_start + len(priority_summary) + 3)
        dashboard.merge_range(ou_start, 1, ou_start, 5, "Top OU / account class", formats["header"])
        dashboard.write(ou_start, 6, "Count", formats["header"])
        for r, (label, count) in enumerate(
            top_ou[["OU Label", "Count"]].itertuples(index=False, name=None), ou_start + 1
        ):
            dashboard.merge_range(r, 1, r, 5, label, formats["note"])
            dashboard.set_row(r, 28)
            dashboard.write(r, 6, int(count))

        bottom_chart_row = ou_start + len(top_ou) + 3
        charts = [
            ("column", "Accounts by Category", data_start + 1, data_start + len(category_summary), "I2", "#5B9BD5", len(category_summary), 570, 310),
            ("bar", "Accounts by Priority", priority_start + 1, priority_start + len(priority_summary), "I18", "#ED7D31", len(priority_summary), 570, 380),
            ("bar", "Top OU / Account Class Buckets", ou_start + 1, ou_start + len(top_ou), f"B{bottom_chart_row + 1}", "#70AD47", len(top_ou), 1280, 350),
        ]
        for chart_type, title, start, end, cell, color, row_count, width, height in charts:
            if row_count == 0:
                continue
            chart = workbook.add_chart({"type": chart_type})
            chart.add_series({"name": title, "categories": ["Dashboard", start, 1, end, 1], "values": ["Dashboard", start, 6, end, 6], "fill": {"color": color}, "data_labels": {"value": True}})
            chart.set_title({"name": title})
            chart.set_legend({"none": True})
            chart.set_size({"width": width, "height": height})
            if chart_type == "bar":
                chart.set_y_axis({"num_font": {"size": 9}})
                chart.set_plotarea({"layout": {"x": 0.51 if cell == "I18" else 0.36, "y": 0.13, "width": 0.44 if cell == "I18" else 0.60, "height": 0.75}})
            dashboard.insert_chart(cell, chart)
        dashboard.set_landscape()
        dashboard.fit_to_pages(1, 1)
        dashboard.print_area(0, 0, bottom_chart_row + 17, 15)
        dashboard.activate()


def ensure_distinct_paths(input_file: Path, output_file: Path) -> None:
    """Reject identical, symlinked, or hard-linked paths before reading or writing report data."""
    input_resolved = input_file.expanduser().resolve(strict=False)
    output_resolved = output_file.expanduser().resolve(strict=False)
    same_file = input_resolved == output_resolved
    if not same_file and input_file.exists() and output_file.exists():
        try:
            same_file = os.path.samefile(input_file, output_file)
        except OSError as exc:
            raise ValueError(f"Could not safely compare input and output paths: {exc}") from exc
    if same_file:
        raise ValueError("Input and output paths must be different; refusing to overwrite the source workbook.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create an actionable inactive-account review workbook from AD report exports.")
    parser.add_argument("input", type=Path, help="MaxPowerSoft-style inactive enabled users Excel export")
    parser.add_argument("output", nargs="?", type=Path, help="Output .xlsx path")
    parser.add_argument("--config", type=Path, help="Optional private JSON rules file for org-specific OU/category mappings")
    args = parser.parse_args(argv)

    config = load_config(args.config)
    output = args.output or args.input.with_name(args.input.stem + "_action_review.xlsx")
    ensure_distinct_paths(args.input, output)
    clean = build_clean_frame(args.input, config)
    write_workbook(clean, output)
    print(f"review_items={len(clean)}")
    print(f"output={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
