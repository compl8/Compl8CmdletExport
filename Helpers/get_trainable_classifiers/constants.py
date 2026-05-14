"""Constants and module-level paths for the trainable-classifier helper."""

from __future__ import annotations

import logging
from pathlib import Path

PURVIEW_URL = "https://purview.microsoft.com"
TC_PAGE_URL = f"{PURVIEW_URL}/informationprotection/dataclassification/trainableclassifiers"
TC_API_AGGREGATES = f"{PURVIEW_URL}/apiproxy/dgs/aggregate/GetTypeAggregates"
TC_API_METADATA = f"{PURVIEW_URL}/apiproxy/gws/ipmlservice/CategoryTrainingModel//ModelMetadata"
TC_API_GETALL = f"{PURVIEW_URL}/apiproxy/gws/ipmlservice/CategoryTrainingModel/getAll"

PROJECT_ROOT = Path(__file__).resolve().parents[2]
STATE_FILE = PROJECT_ROOT / "ConfigFiles" / "PurviewPortalAuth.local.json"
DEFAULT_COMPL8_OUT = PROJECT_ROOT / "ConfigFiles" / "CurrentTenantTCs.local.json"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"
)

TYPE_TC = "TrainableClassifier"
TYPE_SIT = "SensitiveInformationType"

logging.basicConfig(
    format="%(asctime)s %(levelname)-5s  %(message)s",
    datefmt="%H:%M:%S",
    level=logging.INFO,
)
log = logging.getLogger("compl8.tc")
