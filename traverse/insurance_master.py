import pandas as pd

_KEY_COL = "Customer reference"
_AMOUNT_COL = "Amount agreed"


def _sanitize_cols(cols):
    cleaned = []
    for col in cols:
        if col is None:
            cleaned.append("")
            continue
        cleaned.append(str(col).replace("\u00A0", " ").strip())
    return cleaned


def load_traverse_insurance_master(xlsx_or_filelike) -> pd.DataFrame:
    """
    Read the Traverse insurance master file.

    Expected lookup key: Customer reference
    Expected value: Amount agreed
    """
    xl = pd.ExcelFile(xlsx_or_filelike, engine="openpyxl")
    best_df = None
    best_score = -1

    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
        except Exception:
            continue

        df.columns = _sanitize_cols(df.columns)
        score = int(_KEY_COL in df.columns) + int(_AMOUNT_COL in df.columns) + (1 if len(df) else 0)
        if score > best_score:
            best_df = df
            best_score = score

    if best_df is None:
        raise ValueError("Could not read any sheet from the insurance master workbook.")

    if _KEY_COL not in best_df.columns or _AMOUNT_COL not in best_df.columns:
        raise ValueError(
            f"Insurance master must contain '{_KEY_COL}' and '{_AMOUNT_COL}' columns."
        )

    out = best_df[[c for c in best_df.columns if c in {_KEY_COL, _AMOUNT_COL}]].copy()
    out[_KEY_COL] = out[_KEY_COL].fillna("").astype(str).str.strip()
    out[_AMOUNT_COL] = (
        out[_AMOUNT_COL]
        .fillna("")
        .astype(str)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    out[_AMOUNT_COL] = pd.to_numeric(out[_AMOUNT_COL], errors="coerce")

    for date_col in ["Decision date", "Effect date", "Request date", "Created Date", "End date"]:
        if date_col in best_df.columns:
            out[date_col] = pd.to_datetime(best_df.loc[out.index, date_col], errors="coerce", dayfirst=True)

    sort_cols = [col for col in ["Decision date", "Effect date", "Request date", "Created Date", "End date"] if col in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, ascending=[False] * len(sort_cols))

    out = out.dropna(subset=[_KEY_COL])
    out = out[out[_KEY_COL] != ""]
    out = out.drop_duplicates(subset=[_KEY_COL], keep="first")
    return out.reset_index(drop=True)
