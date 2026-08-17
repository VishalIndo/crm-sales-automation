from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path


ROOT = Path(__file__).resolve().parent
GENERATED_DIR = ROOT / "generated"


@dataclass(frozen=True)
class TableSpec:
    name: str
    filename: str
    headers: tuple[str, ...]
    sample_rows: tuple[tuple[str, ...], ...]


LEADS_HEADERS = (
    "LeadID",
    "CreatedDate",
    "CompanyName",
    "ContactName",
    "Phone",
    "Email",
    "LeadSource",
    "Industry",
    "AssignedTo",
    "Stage",
    "ExpectedValue",
    "Probability",
    "NextFollowUp",
    "Status",
    "LastContactDate",
    "Notes",
)

ACTIVITY_HEADERS = (
    "ActivityID",
    "LeadID",
    "ActivityDate",
    "ActivityType",
    "Summary",
    "NextAction",
    "NextFollowUp",
    "Owner",
    "Outcome",
)

QUOTE_HEADERS = (
    "QuoteID",
    "LeadID",
    "QuoteDate",
    "Amount",
    "QuoteStatus",
    "ValidUntil",
    "Notes",
)


def build_specs() -> tuple[TableSpec, ...]:
    today = date.today()
    in_three_days = today + timedelta(days=3)
    in_seven_days = today + timedelta(days=7)
    in_ten_days = today + timedelta(days=10)

    return (
        TableSpec(
            name="Leads",
            filename="leads.csv",
            headers=LEADS_HEADERS,
            sample_rows=(
                (
                    "LD-20260811-001",
                    today.isoformat(),
                    "Northwind Traders",
                    "Alicia Warren",
                    "+49 30 555 1201",
                    "alicia@northwind.example",
                    "Referral",
                    "Wholesale",
                    "Vishal",
                    "Qualified",
                    "25000",
                    "60",
                    in_three_days.isoformat(),
                    "Open",
                    today.isoformat(),
                    "Interested in quarterly supply pricing.",
                ),
                (
                    "LD-20260811-002",
                    today.isoformat(),
                    "BluePeak Services",
                    "Daniel Ruiz",
                    "+49 89 555 2044",
                    "daniel@bluepeak.example",
                    "Website",
                    "Services",
                    "Vishal",
                    "Quoted",
                    "12000",
                    "75",
                    in_seven_days.isoformat(),
                    "Open",
                    today.isoformat(),
                    "Waiting for revised commercial offer.",
                ),
            ),
        ),
        TableSpec(
            name="Activities",
            filename="activities.csv",
            headers=ACTIVITY_HEADERS,
            sample_rows=(
                (
                    "AC-20260811-001",
                    "LD-20260811-001",
                    today.isoformat(),
                    "Call",
                    "Discovery call completed.",
                    "Send pricing sheet",
                    in_three_days.isoformat(),
                    "Vishal",
                    "Positive",
                ),
                (
                    "AC-20260811-002",
                    "LD-20260811-002",
                    today.isoformat(),
                    "Email",
                    "Shared revised proposal.",
                    "Follow up on budget approval",
                    in_ten_days.isoformat(),
                    "Vishal",
                    "Pending",
                ),
            ),
        ),
        TableSpec(
            name="Quotations",
            filename="quotations.csv",
            headers=QUOTE_HEADERS,
            sample_rows=(
                (
                    "QT-20260811-001",
                    "LD-20260811-002",
                    today.isoformat(),
                    "12000",
                    "Sent",
                    in_ten_days.isoformat(),
                    "Annual support package.",
                ),
            ),
        ),
    )


def write_csv(spec: TableSpec) -> Path:
    GENERATED_DIR.mkdir(exist_ok=True)
    path = GENERATED_DIR / spec.filename
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(spec.headers)
        writer.writerows(spec.sample_rows)
    return path


def command_seed() -> None:
    for spec in build_specs():
        path = write_csv(spec)
        print(f"Wrote {path}")


def command_guide() -> None:
    guide = [
        "Excel VBA CRM setup",
        "",
        "1. Create a new macro-enabled workbook (.xlsm).",
        "2. Open the VBA editor with Alt+F11.",
        "3. Import every .bas file from vba_crm/modules.",
        "4. Run SetupCRMApp once.",
        "5. Optional: import generated/*.csv into the matching sheets.",
        "6. Save the workbook and reopen it to trigger Auto_Open reminders.",
        "",
        "Main macros:",
        "- SetupCRMApp",
        "- AddLead",
        "- SearchLead",
        "- AddActivityForSelectedLead",
        "- AddQuotationForSelectedLead",
        "- AdvanceSelectedLeadStage",
        "- RefreshDashboard",
        "- ShowTodayFollowUps",
    ]
    print("\n".join(guide))


def command_manifest() -> None:
    print("VBA CRM source files:")
    for path in sorted((ROOT / "vba_crm").rglob("*")):
        if path.is_file():
            print(path.relative_to(ROOT))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Helper commands for the Excel VBA CRM app.")
    parser.add_argument(
        "command",
        choices=("seed", "guide", "manifest"),
        help="Action to run.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "seed":
        command_seed()
    elif args.command == "guide":
        command_guide()
    else:
        command_manifest()


if __name__ == "__main__":
    main()
