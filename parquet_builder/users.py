"""User identity (GAL Scraper / Entra) CSV processing."""

from __future__ import annotations

import csv
import json
from pathlib import Path

from .constants import SERVICE_ACCOUNT_PATTERNS
from .helpers import _now_iso, _rename_record

# Column renames for user CSVs — handles both GAL Scraper and Entra export field names
USER_RENAMES = {
    # Core identity
    "DisplayName": "display_name",
    "displayName": "display_name",
    "UserPrincipalName": "user_upn",
    "userPrincipalName": "user_upn",
    "Mail": "mail",
    "mail": "mail",
    "AccountEnabled": "account_enabled",
    "accountEnabled": "account_enabled",
    # Org structure
    "Department": "department",
    "department": "department",
    "JobTitle": "job_title",
    "jobTitle": "job_title",
    "CompanyName": "company_name",
    "companyName": "company_name",
    # Location
    "City": "city",
    "city": "city",
    "State": "state",
    "state": "state",
    "OfficeLocation": "office_location",
    "officeLocation": "office_location",
    "OfficeCity": "office_city",
    "UsageLocation": "usage_location",
    "usageLocation": "usage_location",
    # Manager
    "ManagerDisplayName": "manager_display_name",
    "ManagerUPN": "manager_upn",
    "ManagerMail": "manager_mail",
    # Metadata
    "CreatedDateTime": "created_at",
    "createdDateTime": "created_at",
    "AccountType": "account_type",
    "userType": "account_type",
    # On-premises
    "OnPremisesDN": "on_premises_dn",
    "onPremisesDistinguishedName": "on_premises_dn",
    "OnPremisesSamAccount": "on_premises_sam_account",
    "onPremisesSamAccountName": "on_premises_sam_account",
    "OnPremisesDomain": "on_premises_domain",
    "onPremisesDomainName": "on_premises_domain",
    "ADOrgUnit": "ad_org_unit",
    # GAL-specific enrichment
    "AgencyCode": "agency_code",
    "Branch": "branch",
    "RegionOrBU": "region_or_bu",
    "Division": "division",
    "SubBranch": "sub_branch",
    "HistoricalDept": "historical_dept",
    "LastLogonInfo": "last_logon_info",
    # Entra ID
    "id": "entra_id",
    "Id": "entra_id",
}

# Extension attributes (GAL Clean export)
for _i in range(1, 16):
    USER_RENAMES[f"ExtAttr{_i}"] = f"ext_attr_{_i}"
    USER_RENAMES[f"extension_{_i}"] = f"ext_attr_{_i}"


def process_users_csv(csv_path: Path, drift_tracker=None) -> list[dict]:
    """Process a GAL Scraper or Entra user export CSV into unified user records."""
    if not csv_path.exists():
        print(f"  WARNING: Users CSV not found: {csv_path}")
        return []

    ingested_at = _now_iso()
    users = []

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            renamed, extra = _rename_record(raw, USER_RENAMES)
            if drift_tracker is not None:
                drift_tracker.record("users", extra)
            renamed["_source_tool"] = "cmdletexport"
            renamed["_ingested_at"] = ingested_at

            upn = renamed.get("user_upn", "") or ""
            renamed["is_service_account"] = bool(SERVICE_ACCOUNT_PATTERNS.search(upn))
            renamed["extra_fields"] = json.dumps(extra, default=str) if extra else None

            users.append(renamed)

    print(f"  Users: {len(users)} records from {csv_path.name}")
    return users
