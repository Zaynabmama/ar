import io
from collections import defaultdict
from datetime import date, datetime

import pandas as pd
import xlsxwriter

from common.quarter_utils import build_customer_output_config
from orion.export import build_col_map, coerce_export_dates, normalize_all_date_strings
from orion.processor import customer_summary
from traverse.rules import (
    MAIN_ACCOUNT_BY_CUSTOMER_CODE,
    REGION_BY_COUNTRY,
    STATUS_BY_CUSTOMER_CODE,
    lookup_with_default,
    normalize_key,
)


TRAVERSE_INPUT_HEADERS = [
    "CustId",
    "CurrencyId",
    "DoubtfulYN",
    "RetentionYN",
    "InvcNum",
    "DistCode",
    "SalesRepId",
    "Region",
    "CustGroup",
    "CustGroupName",
    "CustName",
    "Status",
    "GrpDate",
    "GrpDue",
    "Feb Weeks",
    "OrigAmount",
    "GrossAmount",
    "GrossAmountBase",
    "CurrentPd",
    "Period1",
    "Period2",
    "Period3",
    "Period4",
    "PrdUnassigned",
    "CustType",
    "ClassId",
    "Class Description",
    "CustLevel",
    "CustStatus",
    "SalesRepId2",
    "Phone",
    "PmtMethod",
    "TransId",
    "CredNum",
    "Country",
    "TermsCode",
    "PONum",
    "SalesRepName",
    "SalesRep2Name",
    "ICPYN",
    "ICPID",
    "CustCateg",
    "CustCategDesc",
    "CustCateg2",
    "CustCategDesc2",
    "CustomerStatus",
    "TaxPayerNo",
    "Collector",
    "Addr1",
    "Addr2",
    "GroupCode",
    "CreditLimit",
    "StdCustId",
    "stdCustDesc",
    "ProjectName",
    "CustDistCode",
    "BankGarantee",
    "ShipToID",
    "ShipToName",
    "SplitVATYN",
    "ProductLine",
    "ProductLineDescr",
    "CountProdLine",
    "AgingDays",
    "CustId",
]

TRAVERSE_APPENDED_HEADERS = [
    "Days from due date",
    "Status",
    "Customer Region",
    "Main Acccount",
    "invoice value - USD",
    "Aging Term",
    "Customer Type",
    "% of Depreciation",
    "Advance from customer",
    "Amount - JD",
    "Provision",
    "Credit Limit Coface",
    "Provision - With Coface",
    "Balance",
    "Balance V1",
    "Count",
]

TRAVERSE_OUTPUT_HEADERS = TRAVERSE_INPUT_HEADERS + TRAVERSE_APPENDED_HEADERS


def _normalize_header(value: str) -> str:
    return "".join(str(value).split()).lower()


def _find_nth_occurrence(cols, target, n=1):
    count = 0
    for i, col in enumerate(cols):
        if _normalize_header(col) == _normalize_header(target):
            count += 1
            if count == n:
                return i
    raise ValueError(f"{target!r} occurrence {n} not found")


def _col_letter(idx: int) -> str:
    idx += 1
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def _col_ref(headers, name, occ=1):
    return _col_letter(_find_nth_occurrence(headers, name, occ))


def _date_formula(as_of_date: date | datetime) -> str:
    return f"DATE({as_of_date.year},{as_of_date.month},{as_of_date.day})"


def _source_value(df: pd.DataFrame, row_idx: int, header: str, occurrence: int = 1):
    source_cols = list(df.columns)
    row = df.iloc[row_idx].tolist()
    count = 0
    target = _normalize_header(header)
    for idx, col in enumerate(source_cols):
        if _normalize_header(col) == target:
            count += 1
            if count == occurrence:
                return row[idx]
    return ""


def _write_value(ws, row_idx, col_idx, header, value, text_fmt, num_fmt, date_fmt):
    if header in {"GrossAmount", "GrossAmountBase", "AgingDays", "CreditLimit"}:
        try:
            num_value = float(value or 0)
            if pd.isna(num_value) or num_value in (float("inf"), float("-inf")):
                num_value = 0.0
            ws.write_number(row_idx, col_idx, num_value, num_fmt)
            return
        except Exception:
            pass

    if header in {"GrpDate", "GrpDue"}:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.notna(parsed):
            ws.write_datetime(row_idx, col_idx, parsed.to_pydatetime(), date_fmt)
            return

    if value is None or (isinstance(value, float) and pd.isna(value)):
        ws.write_blank(row_idx, col_idx, None, text_fmt)
    else:
        ws.write(row_idx, col_idx, value, text_fmt)


def _lookup_status(cust_id):
    return lookup_with_default(STATUS_BY_CUSTOMER_CODE, cust_id, "Regular")


def _lookup_region(country):
    return lookup_with_default(REGION_BY_COUNTRY, country, "")


def _lookup_main_account(cust_id):
    return lookup_with_default(MAIN_ACCOUNT_BY_CUSTOMER_CODE, cust_id, "12301")


def _lookup_coface_limit(cust_name):
    return 0.0


def _to_float(value) -> float:
    try:
        if value is None or value == "":
            return 0.0
        out = float(value)
        if pd.isna(out) or out in (float("inf"), float("-inf")):
            return 0.0
        return out
    except Exception:
        return 0.0


def _source_currency_factor(source_currency: str) -> float:
    return 1.0 if str(source_currency).strip().upper() == "USD" else 1 / 0.71


def _gross_amount_to_usd(gross_amount: float, source_currency: str) -> float:
    if not gross_amount:
        return 0.0
    return gross_amount * _source_currency_factor(source_currency)


def _to_date(value):
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime().date()


def _depreciation_bucket(days: float) -> tuple[str, float]:
    if days <= 0:
        return "3%", 0.03
    if 1 <= days <= 60:
        return "3%", 0.03
    if 61 <= days <= 90:
        return "25%", 0.25
    if 91 <= days <= 120:
        return "50%", 0.50
    if 121 <= days <= 150:
        return "75%", 0.75
    if days >= 151:
        return "100%", 1.0
    return "", 0.0


def _aging_term(days: float) -> str:
    if days < 0:
        return "On Account"
    if days <= 0:
        return "Not past due"
    if days <= 30:
        return "1 - 30"
    if days <= 60:
        return "31 - 60"
    if days <= 90:
        return "61 - 90"
    if days <= 120:
        return "91 - 120"
    if days <= 150:
        return "121 - 150"
    if days >= 151:
        return "More than 151"
    return ""


def _build_orion_customer_source(
    df_rows: pd.DataFrame,
    as_of_date: date,
    source_currency: str = "JD",
) -> pd.DataFrame:
    rows = []
    for r_idx in range(len(df_rows)):
        cust_code = _source_value(df_rows, r_idx, "CustId", 1)
        cust_name = _source_value(df_rows, r_idx, "CustName", 1)
        country = _source_value(df_rows, r_idx, "Country", 1)
        status_value = _lookup_status(cust_code)
        main_ac = _lookup_main_account(cust_code)
        region_value = _lookup_region(country)
        grp_date_value = _source_value(df_rows, r_idx, "GrpDate", 1)
        grp_due_value = _source_value(df_rows, r_idx, "GrpDue", 1)
        gross_amount = _to_float(_source_value(df_rows, r_idx, "GrossAmount", 1))
        gross_amount_base = _to_float(_source_value(df_rows, r_idx, "GrossAmountBase", 1))
        due_dt = _to_date(grp_due_value)
        grp_dt = _to_date(grp_date_value)
        days_from_due = (as_of_date - due_dt).days if due_dt else 0
        invoice_value_usd = _gross_amount_to_usd(gross_amount, source_currency)
        positive_invoice_value_usd = invoice_value_usd if invoice_value_usd > 0 else 0.0

        on_account_value = invoice_value_usd if invoice_value_usd < 0 else 0.0

        not_due_value = positive_invoice_value_usd if days_from_due <= 0 else 0.0
        age1_30 = positive_invoice_value_usd if 0 <= days_from_due <= 30 else 0.0
        age31_60 = positive_invoice_value_usd if 31 <= days_from_due <= 60 else 0.0
        age61_90 = positive_invoice_value_usd if 61 <= days_from_due <= 90 else 0.0
        age91_120 = positive_invoice_value_usd if 91 <= days_from_due <= 120 else 0.0
        age121_150 = positive_invoice_value_usd if 121 <= days_from_due <= 150 else 0.0
        age_ge151 = positive_invoice_value_usd if days_from_due >= 151 else 0.0
        age_over_365 = gross_amount if grp_dt and (as_of_date - grp_dt).days > 365 else 0.0

        # For 'Advance from customer' and 'Amount - JD', the logic is in the main export, not here (By_Customer uses only these fields)

        rows.append(
            {
                "Cust Code": cust_code,
                "Cust Name": cust_name,
                "Main Ac": main_ac,
                "Cust Region": region_value,
                "Region": "Qnal",
                "Updated Status": status_value,
                "Document Due Date": grp_due_value,
                "On Account (Derived)": on_account_value,
                "Not Due Amount": not_due_value,
                "Ar Balance (Copy)": invoice_value_usd,
                "Overdue days (Days)": 0,
                "Aging 1 to 30 (Amount)": age1_30,
                "Aging 31 to 60 (Amount)": age31_60,
                "Aging 61 to 90 (Amount)": age61_90,
                "Aging 91 to 120 (Amount)": age91_120,
                "Aging 121 to 150 (Amount)": age121_150,
                "Aging >=151 (Amount)": age_ge151,
                "Ageing > 365 (Amt)": age_over_365,
            }
        )

    return pd.DataFrame(rows)


def _write_by_customer_sheet(ws, wb, df_customer: pd.DataFrame, selected_quarter: str):
    cfg = build_customer_output_config(selected_quarter)
    cust_df = normalize_all_date_strings(df_customer.copy()).fillna("")
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})
    for c_idx, name in enumerate(cust_df.columns):
        ws.write(0, c_idx, str(name), header_fmt)

    col_map = build_col_map(cust_df)

    def idx(name):
        return list(cust_df.columns).index(name)

    for r_idx, row in enumerate(cust_df.to_numpy(), start=1):
        excel_row = r_idx + 1
        ws.write_row(r_idx, 0, row.tolist())

        current_period = col_map.get(cfg["current_period_label"])
        percent = col_map.get(cfg["percent_label"])
        remaining = col_map.get(cfg["remaining_label"])
        to_add = col_map.get(cfg["to_add_label"])
        next_period = col_map.get(cfg["next_period_label"])

        if current_period and percent:
            ws.write_formula(
                r_idx,
                idx(cfg["actual_label"]),
                f"=IFERROR(${current_period}{excel_row}*${percent}{excel_row},0)",
            )

        if percent:
            ws.write_formula(
                r_idx,
                idx(cfg["remaining_label"]),
                f"=IFERROR(1-${percent}{excel_row},0)",
            )

        if remaining and current_period:
            ws.write_formula(
                r_idx,
                idx(cfg["to_add_label"]),
                f"=IFERROR(${remaining}{excel_row}*${current_period}{excel_row},0)",
            )

        if to_add and next_period:
            ws.write_formula(
                r_idx,
                idx(cfg["forecast_label"]),
                f"=IFERROR(${next_period}{excel_row}+${to_add}{excel_row},0)",
            )

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(1, len(cust_df)), len(cust_df.columns) - 1)


def export_traverse_ar(
    df_rows: pd.DataFrame,
    as_of_date: date,
    selected_quarter: str = "Q1",
    source_currency: str = "JD",
) -> io.BytesIO:
    headers = TRAVERSE_OUTPUT_HEADERS

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    ws = wb.add_worksheet("Traverse_AR")

    header_fmt = wb.add_format(
        {"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1}
    )
    text_fmt = wb.add_format({"border": 1})
    num_fmt = wb.add_format({"border": 1, "num_format": "#,##0.00"})
    date_fmt = wb.add_format({"border": 1, "num_format": "dd/mm/yyyy"})

    for c_idx, name in enumerate(headers):
        ws.write(0, c_idx, str(name), header_fmt)

    raw_cols = TRAVERSE_INPUT_HEADERS
    raw_lookup = {}
    for name in {"CustId", "GrossAmount", "GrossAmountBase", "CustCategDesc", "Country", "GrpDate"}:
        try:
            raw_lookup[name] = _col_ref(raw_cols, name, occ=2 if name == "CustId" else 1)
        except ValueError:
            raw_lookup[name] = None

    appended_start = len(raw_cols)
    appended_positions = {name: appended_start + idx for idx, name in enumerate(TRAVERSE_APPENDED_HEADERS)}
    appended_letters = {name: _col_letter(pos) for name, pos in appended_positions.items()}

    as_of_formula = _date_formula(as_of_date)
    balances_by_cust = defaultdict(float)
    positive_counts_by_cust = defaultdict(int)

    for r_idx in range(len(df_rows)):
        excel_row = r_idx + 2
        occ_counts = defaultdict(int)

        for c_idx, header in enumerate(raw_cols):
            occ_counts[_normalize_header(header)] += 1
            value = _source_value(df_rows, r_idx, header, occ_counts[_normalize_header(header)])
            _write_value(ws, r_idx + 1, c_idx, header, value, text_fmt, num_fmt, date_fmt)

        cust_id = _source_value(df_rows, r_idx, "CustId", 1)
        cust_name = _source_value(df_rows, r_idx, "CustName", 1)
        country = _source_value(df_rows, r_idx, "Country", 1)
        grp_date_value = _source_value(df_rows, r_idx, "GrpDate", 1)
        grp_due_value = _source_value(df_rows, r_idx, "GrpDue", 1)
        gross_amount = _to_float(_source_value(df_rows, r_idx, "GrossAmount", 1))
        gross_amount_base = _to_float(_source_value(df_rows, r_idx, "GrossAmountBase", 1))
        cust_categ_desc = _source_value(df_rows, r_idx, "CustCategDesc", 1)
        credit_limit_source = _source_value(df_rows, r_idx, "Credit Limit Coface", 1)

        grp_due = _to_date(grp_due_value)
        days_from_due = (as_of_date - grp_due).days if grp_due else None
        invoice_usd = _gross_amount_to_usd(gross_amount, source_currency)
        status_value = _lookup_status(cust_id)
        region_value = _lookup_region(country)
        main_account_value = _lookup_main_account(cust_id)
        customer_type_value = (
            "affiliated"
            if normalize_key(cust_categ_desc)
            == normalize_key("Non Gov - Inter-Affiliated Company - level 1(MIDIS)")
            else "Non-affiliated"
        )
        dep_label, dep_rate = _depreciation_bucket(days_from_due or 0)
        advance_value = gross_amount_base if gross_amount_base < 0 else 0.0
        amount_jd_value = gross_amount_base if gross_amount_base >= 0 else 0.0
        provision_value = (
            0.0
            if amount_jd_value <= 0 or customer_type_value != "Non-affiliated"
            else dep_rate * gross_amount_base
        )
        credit_limit_value = _to_float(credit_limit_source) if credit_limit_source != "" else _lookup_coface_limit(cust_name)
        balance_value = balances_by_cust[cust_id] + amount_jd_value
        balance_v1_value = balance_value - credit_limit_value
        count_value = positive_counts_by_cust[cust_id] + (1 if balance_v1_value > 0 else 0)

        if customer_type_value != "Non-affiliated":
            provision_with_coface_value = provision_value
        elif balance_value <= credit_limit_value:
            provision_with_coface_value = (amount_jd_value * 0.05) * dep_rate
        elif count_value <= 1:
            if balance_value > credit_limit_value:
                provision_with_coface_value = (
                    (balance_value - credit_limit_value) * dep_rate
                    + ((amount_jd_value - (balance_value - credit_limit_value)) * 0.05) * dep_rate
                )
            else:
                provision_with_coface_value = provision_value
        else:
            provision_with_coface_value = provision_value

        balances_by_cust[cust_id] = balance_value
        positive_counts_by_cust[cust_id] = count_value

        ws.write_formula(
            r_idx + 1,
            appended_positions["Days from due date"],
            f"={as_of_formula}-{_col_ref(raw_cols, 'GrpDue')}{excel_row}",
            num_fmt if days_from_due is not None else text_fmt,
            days_from_due if days_from_due is not None else "",
        )
        ws.write(r_idx + 1, appended_positions["Status"], status_value, text_fmt)
        ws.write(r_idx + 1, appended_positions["Customer Region"], region_value, text_fmt)
        ws.write(r_idx + 1, appended_positions["Main Acccount"], main_account_value, text_fmt)
        ws.write_formula(
            r_idx + 1,
            appended_positions["invoice value - USD"],
            (
                f"=IFERROR({raw_lookup['GrossAmount']}{excel_row}"
                f"{'' if str(source_currency).strip().upper() == 'USD' else '/0.71'},0)"
            ),
            num_fmt,
            invoice_usd,
        )
        days_col = appended_letters["Days from due date"]
        ws.write_formula(
            r_idx + 1,
            appended_positions["Aging Term"],
            (
                f'=IF({appended_letters["invoice value - USD"]}{excel_row}<0,"On Account",'
                f'IF({days_col}{excel_row}<=0,"Not past due",'
                f'IF(AND({days_col}{excel_row}>=1,{days_col}{excel_row}<=30),"1 - 30",'
                f'IF(AND({days_col}{excel_row}>=31,{days_col}{excel_row}<=60),"31 - 60",'
                f'IF(AND({days_col}{excel_row}>=61,{days_col}{excel_row}<=90),"61 - 90",'
                f'IF(AND({days_col}{excel_row}>=91,{days_col}{excel_row}<=120),"91 - 120",'
                f'IF(AND({days_col}{excel_row}>=121,{days_col}{excel_row}<=150),"121 - 150",'
                f'IF({days_col}{excel_row}>=151,"More than 151",""))))))))'
            ),
            text_fmt,
            "On Account" if invoice_usd < 0 else _aging_term(days_from_due or 0),
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Customer Type"],
            f'=IF({raw_lookup["CustCategDesc"]}{excel_row}="Non Gov - Inter-Affiliated Company - level 1(MIDIS)","affiliated","Non-affiliated")',
            text_fmt,
            customer_type_value,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["% of Depreciation"],
            (
                f'=IF({days_col}{excel_row}<=0,"3%",'
                f'IF(AND({days_col}{excel_row}>=1,{days_col}{excel_row}<=60),"3%",'
                f'IF(AND({days_col}{excel_row}>=61,{days_col}{excel_row}<=90),"25%",'
                f'IF(AND({days_col}{excel_row}>=91,{days_col}{excel_row}<=120),"50%",'
                f'IF(AND({days_col}{excel_row}>=121,{days_col}{excel_row}<=150),"75%",'
                f'IF({days_col}{excel_row}>=151,"100%",""))))))'
            ),
            text_fmt,
            dep_label,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Advance from customer"],
            f"=IF({raw_lookup['GrossAmountBase']}{excel_row}<0,{raw_lookup['GrossAmountBase']}{excel_row},0)",
            num_fmt,
            advance_value,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Amount - JD"],
            f"=IF({raw_lookup['GrossAmountBase']}{excel_row}>=0,{raw_lookup['GrossAmountBase']}{excel_row},0)",
            num_fmt,
            amount_jd_value,
        )

        amount_jd_letter = appended_letters["Amount - JD"]
        customer_type_letter = appended_letters["Customer Type"]
        dep_letter = appended_letters["% of Depreciation"]
        provision_letter = appended_letters["Provision"]
        balance_letter = appended_letters["Balance"]
        balance_v1_letter = appended_letters["Balance V1"]
        credit_limit_letter = appended_letters["Credit Limit Coface"]

        ws.write_formula(
            r_idx + 1,
            appended_positions["Provision"],
            (
                f'=IF({amount_jd_letter}{excel_row}<=0,0,'
                f'IF({customer_type_letter}{excel_row}="Non-affiliated",'
                f'{dep_letter}{excel_row}*{raw_lookup["GrossAmountBase"]}{excel_row},0))'
            ),
            num_fmt,
            provision_value,
        )
        ws.write(
            r_idx + 1,
            appended_positions["Credit Limit Coface"],
            credit_limit_value,
            num_fmt,
        )

        cust_id_balance_ref = _col_ref(raw_cols, "CustId", occ=2)
        ws.write_formula(
            r_idx + 1,
            appended_positions["Provision - With Coface"],
            (
                f'=IF({customer_type_letter}{excel_row}="Non-affiliated",'
                f'IF(SUMIF(${cust_id_balance_ref}$2:{cust_id_balance_ref}{excel_row},{cust_id_balance_ref}{excel_row},'
                f'${amount_jd_letter}$2:{amount_jd_letter}{excel_row})<={credit_limit_letter}{excel_row},'
                f'({amount_jd_letter}{excel_row}*5%)*{dep_letter}{excel_row},'
                f'IF({balance_v1_letter}{excel_row}<=1,'
                f'IF({balance_letter}{excel_row}>{credit_limit_letter}{excel_row},'
                f'({balance_letter}{excel_row}-{credit_limit_letter}{excel_row})*{dep_letter}{excel_row}'
                f'+(({amount_jd_letter}{excel_row}-({balance_letter}{excel_row}-{credit_limit_letter}{excel_row}))*5%)*{dep_letter}{excel_row},'
                f'{provision_letter}{excel_row}),'
                f'{provision_letter}{excel_row})),'
                f'{provision_letter}{excel_row})'
            ),
            num_fmt,
            provision_with_coface_value,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Balance"],
            f"=SUMIF(${cust_id_balance_ref}$2:{cust_id_balance_ref}{excel_row},{cust_id_balance_ref}{excel_row},${amount_jd_letter}$2:{amount_jd_letter}{excel_row})",
            num_fmt,
            balance_value,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Balance V1"],
            f"={balance_letter}{excel_row}-{credit_limit_letter}{excel_row}",
            num_fmt,
            balance_v1_value,
        )
        ws.write_formula(
            r_idx + 1,
            appended_positions["Count"],
            f'=COUNTIFS(${cust_id_balance_ref}$2:{cust_id_balance_ref}{excel_row},{cust_id_balance_ref}{excel_row},${balance_v1_letter}$2:{balance_v1_letter}{excel_row},">0")',
            num_fmt,
            count_value,
        )

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(1, len(df_rows)), len(headers) - 1)

    for idx, header in enumerate(headers):
        width = 14
        if header in {"CustName", "CustGroupName", "CustCategDesc", "CustomerStatus"}:
            width = 20
        elif header in {"GrpDate", "GrpDue"}:
            width = 12
        elif header in {"Days from due date", "invoice value - USD", "Provision", "Credit Limit Coface"}:
            width = 16
        ws.set_column(idx, idx, width)

    ws_cust = wb.add_worksheet("By_Customer")
    orion_source = _build_orion_customer_source(df_rows, as_of_date, source_currency=source_currency)
    df_customer = customer_summary(orion_source, selected_quarter=selected_quarter)
    _write_by_customer_sheet(ws_cust, wb, df_customer, selected_quarter)

    wb.close()
    output.seek(0)
    return output
