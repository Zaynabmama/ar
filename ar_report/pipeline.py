"""End-to-end: four raw exports in -> finished workbook out.

This is the one function both the Streamlit app and the test script call, so
whenever file delivery is automated later (Orion pushing files directly, a
mailbox watcher for the backlog report, etc.) that trigger only needs to call
`run_pipeline` with file paths -- nothing else changes.
"""

from __future__ import annotations

from pathlib import Path

from ar_report import aggregate, build_output, ingest, transform
from ar_report.lookups import apply_region_overrides, load_lookups

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "template" / "lookup_data.xlsx"


def run_pipeline(
    ar_path, pdc_path, insurance_path, backlog_path, as_on_date=None,
    template_path=TEMPLATE_PATH, region_override_paths=None,
):
    lookups = load_lookups(template_path)

    region_overrides_applied = 0
    if region_override_paths:
        override_dfs = [ingest.load_region_override_file(p) for p in region_override_paths]
        region_overrides_applied = sum(len(df) for df in override_dfs)
        apply_region_overrides(lookups, override_dfs)

    ar_raw = ingest.load_ar_provision(ar_path)
    pdc_df = ingest.load_pdc(pdc_path)
    insurance_df = ingest.load_insurance(insurance_path)
    backlog_raw = ingest.load_backlog(backlog_path)

    ar_clean = transform.clean_ar_provision(ar_raw, lookups)
    resolved_date = transform.resolve_as_on_date(ar_clean, override=as_on_date)
    ar_full = transform.add_helper_columns(ar_clean, resolved_date, lookups)
    backlog_clean = transform.clean_backlog(backlog_raw)

    customer_pivot = aggregate.pivot_customer_aging(ar_full)
    additional_dues = aggregate.pivot_additional_dues(ar_full)
    backlog_detail_pivot = aggregate.pivot_backlog_detail(backlog_clean)
    backlog_by_customer = aggregate.pivot_backlog_by_customer(backlog_clean)
    by_customer = aggregate.build_by_customer(
        customer_pivot, additional_dues, backlog_by_customer, insurance_df, lookups
    )
    by_invoice = aggregate.build_by_invoice(ar_full)

    workbook_buffer = build_output.build_workbook(
        ar_df=ar_full,
        pdc_df=pdc_df,
        insurance_df=insurance_df,
        backlog_df=backlog_clean,
        as_on_date=resolved_date,
        backlog_detail_pivot=backlog_detail_pivot,
        backlog_by_customer=backlog_by_customer,
        by_customer=by_customer,
        by_invoice=by_invoice,
        lookups=lookups,
    )

    stats = {
        "as_on_date": resolved_date,
        "ar_rows_raw": len(ar_raw),
        "ar_rows_clean": len(ar_clean),
        "backlog_rows_raw": len(backlog_raw),
        "backlog_rows_clean": len(backlog_clean),
        "customers": by_customer["Cust Code"].nunique(),
        "invoices": len(by_invoice),
        "region_overrides_applied": region_overrides_applied,
    }
    return workbook_buffer, stats
