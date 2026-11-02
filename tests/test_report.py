"""Regression tests for conservative classification and safe XLSX generation."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

import ad_account_hygiene_report as report


REPORT_COLUMNS = [
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


def make_row(**overrides: object) -> dict[str, object]:
    """Return one generic source row while letting each test expose only its relevant values."""
    row: dict[str, object] = {
        "Disabled": False,
        "Display Name": "Example User",
        "SAM Account Name": "example.user",
        "Account Expires On": "Never",
        "Password Last Set": "2025-01-01",
        "When Created": "2024-01-01",
        "Last Logon Timestamp": "2025-06-01",
        "Days Since Last Logon": 180,
        "Logon Count": 4,
        "Distinguished Name": "CN=Example User,OU=Site,OU=Users,DC=example,DC=test",
        "Canonical Name": "example.test/Users/Region/Site/Example User",
    }
    row.update(overrides)
    return row


def write_report(path: Path, rows: list[dict[str, object]], columns: list[str] | None = None) -> None:
    """Write the title rows, detected header row, and data used by the real loader."""
    selected = columns or REPORT_COLUMNS
    matrix = [
        ["Inactive Enabled Users", *([None] * (len(selected) - 1))],
        ["Synthetic test data", *([None] * (len(selected) - 1))],
        selected,
    ]
    matrix.extend([[row.get(column) for column in selected] for row in rows])
    with pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_formulas": False, "strings_to_urls": False}},
    ) as writer:
        pd.DataFrame(matrix).to_excel(writer, index=False, header=False)


def output_row_count(workbook: openpyxl.Workbook, sheet_name: str) -> int:
    """Count populated records below the one-row output header."""
    sheet = workbook[sheet_name]
    return max(sheet.max_row - 1, 0)


@pytest.mark.parametrize(
    ("source_value", "clean_value", "priority"),
    [
        (None, "Unknown", "Needs Review - Unknown Logon Age"),
        ("", "Unknown", "Needs Review - Unknown Logon Age"),
        ("not-a-day-count", "Unknown", "Needs Review - Unknown Logon Age"),
        (-1, "Unknown", "Needs Review - Unknown Logon Age"),
        (90.5, "Unknown", "Needs Review - Unknown Logon Age"),
        ("Never Logged On", "Source Reports Never Logged On", "Critical - Source Reports Never Logged On"),
        (89, 89, "Below Threshold - Under 90 days"),
        (90, 90, "Low - 90+ days"),
        (179, 179, "Low - 90+ days"),
        (180, 180, "Medium - 180+ days"),
        (364, 364, "Medium - 180+ days"),
        (365, 365, "High - 365+ days"),
    ],
)
def test_day_values_use_conservative_priority_bands(
    source_value: object, clean_value: object, priority: str
) -> None:
    """Unknown values and values below 90 days must not inherit a stale-account claim."""
    normalized = report.normalize_days(source_value)
    actual_priority, _ = report.priority_and_action(normalized, is_service=False)
    assert normalized == clean_value
    assert actual_priority == priority


def test_never_is_described_as_a_source_sentinel_not_proof() -> None:
    """The source export's Never marker must prompt validation rather than claim observed fact."""
    normalized = report.normalize_days("Never")
    priority, action = report.priority_and_action(normalized, is_service=False)
    assert normalized == "Source Reports Never Logged On"
    assert priority == "Critical - Source Reports Never Logged On"
    assert "source sentinel" in action.lower()
    assert "authoritative" in action.lower()


def test_unknown_service_age_stays_unknown_and_in_review() -> None:
    """Service routing must not convert missing age evidence into Never or a low-priority band."""
    priority, action = report.priority_and_action(report.normalize_days(None), is_service=True)
    assert priority == "Service Review - Unknown Logon Age"
    assert "logon age" in action.lower()


def test_keyword_matching_uses_boundaries_and_keeps_rules_configurable() -> None:
    """Short terms such as bot and pos must not classify Abbott or Post by substring alone."""
    config = deepcopy(report.DEFAULT_CONFIG)
    row = pd.Series(
        {
            "Canonical Name": "example.test/Users/Region/Site/Abbott Post",
            "Distinguished Name": "CN=Abbott Post,OU=Users,DC=example,DC=test",
            "SAM Account Name": "abbott.post",
            "Display Name": "Abbott Post",
            "ou_path": "Users/Region/Site",
            "account_class": "Region",
            "root_ou": "Users",
        }
    )
    assert report.service_account_reason(row, config) == ""

    bot_row = row.copy()
    bot_row["SAM Account Name"] = "night-bot"
    assert "keyword: bot" in report.service_account_reason(bot_row, config).lower()
    pos_row = row.copy()
    pos_row["Display Name"] = "POS Terminal"
    assert "keyword: pos" in report.service_account_reason(pos_row, config).lower()

    custom = deepcopy(config)
    custom["service_path_keywords"] = ["robot identities"]
    custom["service_name_patterns"] = [r"^machine_"]
    custom["device_keywords"] = []
    path_row = row.copy()
    path_row["Canonical Name"] = "example.test/Applications/Robot Identities/Site/Worker"
    assert "robot identities" in report.service_account_reason(path_row, custom).lower()
    named_row = row.copy()
    named_row["SAM Account Name"] = "machine_worker"
    assert "machine_" in report.service_account_reason(named_row, custom)


def test_empty_workbook_has_a_clear_error(tmp_path: Path) -> None:
    """A workbook with no report cells must fail before header indexing produces a cryptic error."""
    source = tmp_path / "empty.xlsx"
    openpyxl.Workbook().save(source)
    with pytest.raises(ValueError, match="empty"):
        report.load_ad_report(source)


def test_header_marker_and_required_columns_are_validated(tmp_path: Path) -> None:
    """Malformed source structure must fail with a report-specific validation message."""
    no_header = tmp_path / "no-header.xlsx"
    pd.DataFrame([["title only"]]).to_excel(no_header, index=False, header=False)
    with pytest.raises(ValueError, match="header row"):
        report.load_ad_report(no_header)

    missing_required = tmp_path / "missing-required.xlsx"
    columns = [column for column in REPORT_COLUMNS if column != "Days Since Last Logon"]
    write_report(missing_required, [make_row()], columns)
    with pytest.raises(ValueError, match="Missing required input columns"):
        report.build_clean_frame(missing_required, deepcopy(report.DEFAULT_CONFIG))


def test_duplicate_headers_are_rejected(tmp_path: Path) -> None:
    """Duplicate names are ambiguous in pandas and must not be silently selected or overwritten."""
    source = tmp_path / "duplicate-header.xlsx"
    columns = [*REPORT_COLUMNS, "Canonical Name"]
    write_report(source, [make_row()], columns)
    with pytest.raises(ValueError, match="duplicate"):
        report.load_ad_report(source)


def test_zero_data_rows_generate_a_readable_empty_workbook(tmp_path: Path) -> None:
    """A valid report with no records should still produce the standard dashboard and tabs."""
    source = tmp_path / "zero-rows.xlsx"
    output = tmp_path / "zero-rows-output.xlsx"
    write_report(source, [])

    clean = report.build_clean_frame(source, deepcopy(report.DEFAULT_CONFIG))
    assert list(clean.columns) == report.OUTPUT_COLUMNS
    assert clean.empty
    report.write_workbook(clean, output)

    workbook = openpyxl.load_workbook(output, data_only=False)
    assert workbook.sheetnames == [
        "Dashboard",
        "All Review Items",
        "Workforce Accounts",
        "Privileged Accounts",
        "Service Accounts",
        "Other Accounts",
        "OU Summary",
    ]
    assert workbook["Dashboard"]["B7"].value == 0
    assert output_row_count(workbook, "All Review Items") == 0


def test_categories_partition_every_row_and_missing_canonical_is_other(tmp_path: Path) -> None:
    """The four category tabs must form a complete, non-overlapping partition of all items."""
    source = tmp_path / "partition-source.xlsx"
    output = tmp_path / "partition-output.xlsx"
    rows = [
        make_row(**{"SAM Account Name": "worker", "Canonical Name": "example.test/Users/Region/Site/Worker"}),
        make_row(**{"SAM Account Name": "admin.ops", "Canonical Name": "example.test/Privileged/Admins/Site/Admin Ops"}),
        make_row(**{"SAM Account Name": "svc-job", "Canonical Name": "example.test/Applications/Service Accounts/Site/Job"}),
        make_row(**{"SAM Account Name": "contractor", "Canonical Name": "example.test/Contractors/Region/Site/Contractor"}),
        make_row(**{"SAM Account Name": "missing.path", "Canonical Name": None}),
        make_row(**{"Display Name": "Abbott Post", "SAM Account Name": "abbott.post", "Canonical Name": "example.test/Users/Region/Site/Abbott Post"}),
    ]
    write_report(source, rows)

    clean = report.build_clean_frame(source, deepcopy(report.DEFAULT_CONFIG))
    assert clean.loc[clean["SAM Account Name"].eq("missing.path"), "Category"].item() == "Other Accounts"
    assert clean.loc[clean["SAM Account Name"].eq("abbott.post"), "Category"].item() == "Workforce Accounts"
    assert set(clean["Category"]) <= {
        "Workforce Accounts",
        "Privileged Accounts",
        "Service Accounts",
        "Other Accounts",
    }

    report.write_workbook(clean, output)
    workbook = openpyxl.load_workbook(output, data_only=False)
    category_total = sum(
        output_row_count(workbook, name)
        for name in ["Workforce Accounts", "Privileged Accounts", "Service Accounts", "Other Accounts"]
    )
    assert output_row_count(workbook, "All Review Items") == len(rows)
    assert category_total == len(rows)


def test_generated_xlsx_preserves_unicode_xml_and_formula_like_text(tmp_path: Path) -> None:
    """Untrusted text must round-trip as text without formulas or automatic hyperlinks."""
    source = tmp_path / "hostile-text-source.xlsx"
    output = tmp_path / "hostile-text-output.xlsx"
    display_name = 'Zoë 東京 🛡 <ops>&"quoted"'
    sam_name = "=2+2"
    distinguished_name = "https://example.invalid/?a=1&b=<node>"
    write_report(
        source,
        [
            make_row(
                **{
                    "Display Name": display_name,
                    "SAM Account Name": sam_name,
                    "Distinguished Name": distinguished_name,
                    "Canonical Name": "example.test/Users/Region/Site/Formula-like User",
                    "Days Since Last Logon": 90,
                }
            )
        ],
    )

    clean = report.build_clean_frame(source, deepcopy(report.DEFAULT_CONFIG))
    report.write_workbook(clean, output)
    workbook = openpyxl.load_workbook(output, data_only=False)
    sheet = workbook["All Review Items"]
    headers = {cell.value: cell.column for cell in sheet[1]}

    display_cell = sheet.cell(2, headers["Display Name"])
    sam_cell = sheet.cell(2, headers["SAM Account Name"])
    dn_cell = sheet.cell(2, headers["Distinguished Name"])
    assert display_cell.value == display_name
    assert sam_cell.value == sam_name
    assert sam_cell.data_type == "s"
    assert dn_cell.value == distinguished_name
    assert dn_cell.data_type == "s"
    assert dn_cell.hyperlink is None


def test_same_input_and_output_path_is_rejected_without_damage(tmp_path: Path) -> None:
    """CLI path validation must protect the only copy of the source workbook."""
    source = tmp_path / "source.xlsx"
    write_report(source, [make_row()])
    before = source.read_bytes()
    with pytest.raises(ValueError, match="different"):
        report.main([str(source), str(source)])
    assert source.read_bytes() == before


def test_failed_generation_preserves_existing_output(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Only a fully closed temporary workbook may replace an existing completed output."""
    output = tmp_path / "existing.xlsx"
    previous = b"previous-completed-output"
    output.write_bytes(previous)
    clean = pd.DataFrame(columns=report.OUTPUT_COLUMNS)

    def fail_after_partial_write(_: pd.DataFrame, temporary_path: Path) -> None:
        temporary_path.write_bytes(b"partial")
        raise RuntimeError("simulated writer failure")

    monkeypatch.setattr(report, "_write_workbook_contents", fail_after_partial_write)
    with pytest.raises(RuntimeError, match="simulated"):
        report.write_workbook(clean, output)
    assert output.read_bytes() == previous
    assert list(tmp_path.glob(".existing.xlsx.*.xlsx")) == []
