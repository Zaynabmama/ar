import pandas as pd

from common.identifier_utils import normalize_excel_identifier_series

_TRAVERSE_KEY_COL = "Customer reference"
_TRAVERSE_AMOUNT_COL = "Amount agreed"

_BUDG_REQUIRED_COLS = {
    "Customer Code",
    "Main Account",
    "Insurance Limit",
}

_COLUMN_ALIASES = {
    "Customer Code": ["Customer Code", "Cust Code", "CustCode", "Customercode", _TRAVERSE_KEY_COL],
    "Main Account": ["Main Account", "Main Ac", "MainAc", "Main Account No", "Main Acct"],
    "Insurance Limit": ["Insurance Limit", "Insurance", "Limit", "Ins Limit", _TRAVERSE_AMOUNT_COL],
}


def _sanitize_cols(cols):
    """Trim, replace NBSP with space; keep original case."""
    fixed = []
    for c in cols:
        if c is None:
            fixed.append("")
            continue
        s = str(c).replace("\u00A0", " ").strip()
        fixed.append(s)
    return fixed


def _normalize_col_name(name: str) -> str:
    return "".join(str(name).split()).lower()


def _resolve_column(columns, target: str) -> str | None:
    candidates = _COLUMN_ALIASES.get(target, [target])
    normalized_columns = {_normalize_col_name(col): col for col in columns}
    for candidate in candidates:
        resolved = normalized_columns.get(_normalize_col_name(candidate))
        if resolved is not None:
            return resolved
    return None


def _score_budg_master(df: pd.DataFrame) -> int:
    return sum(int(_resolve_column(df.columns, col) is not None) for col in _BUDG_REQUIRED_COLS) + int(len(df) > 0)


def _score_traverse_master(df: pd.DataFrame) -> int:
    key = _resolve_column(df.columns, _TRAVERSE_KEY_COL)
    amount = _resolve_column(df.columns, _TRAVERSE_AMOUNT_COL)
    return int(key is not None) + int(amount is not None) + int(len(df) > 0)


def _read_master_candidates(xlsx_or_filelike) -> list[pd.DataFrame]:
    xl = pd.ExcelFile(xlsx_or_filelike, engine="openpyxl")
    candidates: list[pd.DataFrame] = []
    for sheet_name in xl.sheet_names:
        for header_row in range(0, 11):
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row, dtype=str)
            except Exception:
                continue
            df.columns = _sanitize_cols(df.columns)
            candidates.append(df)
    return candidates


def _normalize_budg_master(df: pd.DataFrame) -> pd.DataFrame:
    source_df = df.copy()
    resolved = {}
    for target in _BUDG_REQUIRED_COLS:
        actual = _resolve_column(source_df.columns, target)
        if actual is not None:
            resolved[target] = actual

    missing = [c for c in sorted(_BUDG_REQUIRED_COLS) if c not in resolved]
    if missing:
        raise ValueError(
            "Insurance Master is missing required column(s): "
            + ", ".join(missing)
            + ". Found columns: "
            + ", ".join(map(str, df.columns))
        )

    out = source_df[[resolved[c] for c in resolved]].copy()
    out = out.rename(columns={actual: target for target, actual in resolved.items()})
    out["Customer Code"] = out["Customer Code"].fillna("").astype(str).str.strip()
    out["Main Account"] = normalize_excel_identifier_series(out["Main Account"])
    out["Insurance Limit"] = pd.to_numeric(out["Insurance Limit"], errors="coerce")

    for dc in ["Effective From", "Effective To", "Created Date"]:
        actual = _resolve_column(source_df.columns, dc)
        if actual is not None:
            out[dc] = pd.to_datetime(source_df[actual], errors="coerce", dayfirst=True)

    sort_cols = [col for col in ["Effective From", "Created Date"] if col in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, ascending=[False] * len(sort_cols))

    out = out.drop_duplicates(subset=["Customer Code", "Main Account"], keep="first")
    out.attrs["master_format"] = "budg"
    return out.reset_index(drop=True)


def _normalize_traverse_master(df: pd.DataFrame) -> pd.DataFrame:
    key_col = _resolve_column(df.columns, _TRAVERSE_KEY_COL)
    amount_col = _resolve_column(df.columns, _TRAVERSE_AMOUNT_COL)
    if key_col is None or amount_col is None:
        raise ValueError(
            f"Traverse-style master must contain '{_TRAVERSE_KEY_COL}' and '{_TRAVERSE_AMOUNT_COL}' columns."
        )

    out = pd.DataFrame(
        {
            "Customer Code": df[key_col].fillna("").astype(str).str.strip(),
            "Main Account": "",
            "Insurance Limit": pd.to_numeric(
                df[amount_col].fillna("").astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
                errors="coerce",
            ),
        }
    )
    out = out.dropna(subset=["Customer Code"])
    out = out[out["Customer Code"] != ""]
    out = out.drop_duplicates(subset=["Customer Code"], keep="first")
    out.attrs["master_format"] = "traverse"
    return out.reset_index(drop=True)


def load_insurance_master(xlsx_or_filelike) -> pd.DataFrame:
    """
    Read either insurance master format and normalize it for the BUD2026 mapper.

    Supported formats:
    - Budg-style: Customer Code + Main Account + Insurance Limit
    - Traverse-style: Customer reference + Amount agreed
    """
    candidates = _read_master_candidates(xlsx_or_filelike)
    if not candidates:
        raise ValueError("Could not read any sheet from the insurance master workbook.")

    best_df = None
    best_kind = None
    best_score = -1

    for df in candidates:
        budg_score = _score_budg_master(df)
        traverse_score = _score_traverse_master(df)
        has_main_account = _resolve_column(df.columns, "Main Account") is not None
        has_traverse_pair = (
            _resolve_column(df.columns, _TRAVERSE_KEY_COL) is not None
            and _resolve_column(df.columns, _TRAVERSE_AMOUNT_COL) is not None
        )

        if has_traverse_pair and not has_main_account:
            kind = "traverse"
            score = traverse_score
        elif budg_score >= traverse_score:
            kind = "budg"
            score = budg_score
        else:
            kind = "traverse"
            score = traverse_score

        if score > best_score:
            best_df = df
            best_kind = kind
            best_score = score

    if best_df is None:
        raise ValueError("Could not detect a valid insurance master layout.")

    if best_kind == "traverse":
        return _normalize_traverse_master(best_df)
    return _normalize_budg_master(best_df)
