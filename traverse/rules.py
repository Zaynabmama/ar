import re
from typing import Any


STATUS_BY_CUSTOMER_CODE = {
    "PCHSCTC": "DOUB_INS_CLA",
    "TRDARA": "DOUBTFUL",
    "TRMIX": "DOUBTFUL",
}

REGION_BY_COUNTRY = {
    "JOR": "Jordan",
    "ARE": "UAE",
    "LBN": "Lebanon",
    "PSE": "Palestine",
    "BDI": "Jordan",
}

MAIN_ACCOUNT_BY_CUSTOMER_CODE = {
    "MINDWAREFZ": "12305",
    "PCABM": "12305",
    "PCBLUETECH": "12305",
    "PCBTA": "12305",
    "PCIBS": "12305",
    "PCJBS": "12305",
    "PCJBU": "12305",
    "PCJDS": "12305",
    "PCMAGNOOS": "12305",
    "PCOMNJ": "12305",
    "PCOMS": "12305",
}

COFACE_LIMIT_BY_CUSTNAME = {}


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u00A0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text.upper()


def lookup_with_default(mapping: dict[str, Any], key: Any, default: Any) -> Any:
    normalized = normalize_key(key)
    if not normalized:
        return default
    return mapping.get(normalized, default)

