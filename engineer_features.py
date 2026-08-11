#!/usr/bin/env python3
"""
Invoice Feature Engineering
============================

Companion to reconcile.py (KPI reconciliation) and classify_fields.py (field-
usage audit). Where those two tools check whether the data supports what's
already being shown, this one derives NEW columns from the export that
aren't on the dashboard at all - ratios, flags, and buckets useful for
further analysis (a PM building a pivot table, a BI report, or a model)
rather than for reproducing an existing chart.

Every feature declares the export column(s) it needs up front, in
FEATURE_REGISTRY below. At runtime the tool checks the *actual* loaded
export's columns against that list and computes only what's possible - a
feature blocked by a missing column is reported by name, with the exact
column that would unlock it, rather than silently skipped or guessed at.
This is the same "desired/required fields" idea reconcile.py's
ENHANCEMENT_FIELDS registry uses for KPI cards, applied here to engineered
features instead: the "My Team's Open Invoices" view is one export among
several possible ones, and a PM can add columns to it at any time.

Usage:
    python3 engineer_features.py export.xlsx --out enriched.csv
    python3 engineer_features.py export.xlsx --sheet "My Team's Open Invoices"
    python3 engineer_features.py export.xlsx --list-only   # just show what's computable, don't write a file

Works on .xlsx/.xls/.xlsm or .csv, same loader pattern as classify_fields.py.
"""

import argparse
import sys

import numpy as np
import pandas as pd

VAT_RATE_PCT = 15.0
HIGH_VALUE_PERCENTILE = 0.95


# ---------------------------------------------------------------------------
# Feature registry - the extension point. Add a new engineered feature by
# adding an entry here; the runner below handles availability checking,
# computing, and reporting uniformly for every entry.
#
#   requires          - export columns that MUST all be present
#   requires_any_of   - if set, at least one of these columns must be present
#                        (used for fields with a documented fallback, like the
#                        payment-date pair reconcile.py already flags as
#                        missing from the default export view)
#   compute(df)        - returns a pd.Series aligned to df.index; may contain
#                        NaN for individual rows that don't have the needed
#                        values even though the column exists (e.g. a blank
#                        Closed Date on an unpaid invoice)
# ---------------------------------------------------------------------------

def _dt(df, col):
    return pd.to_datetime(df[col], errors="coerce")


def _num(df, col):
    return pd.to_numeric(df[col], errors="coerce")


FEATURE_REGISTRY = {
    "days_to_close": {
        "requires": ["Created On", "Closed Date"],
        "description": "Days between Created On and Closed Date - a proxy for how long an "
                        "invoice took to close out (not the same as payment duration; see "
                        "payment_duration_days below for why that one needs a different field).",
        "compute": lambda df: (_dt(df, "Closed Date") - _dt(df, "Created On")).dt.days,
    },
    "invoice_age_days": {
        "requires": ["Created On"],
        "description": "Days between Created On and now - how long the invoice has existed, "
                        "regardless of status.",
        "compute": lambda df, now: (now - _dt(df, "Created On")).dt.days,
    },
    "created_month": {
        "requires": ["Created On"],
        "description": "Calendar month (YYYY-MM) the invoice was created in - a standard "
                        "bucketing feature for monthly rollups or as a categorical model input.",
        "compute": lambda df: _dt(df, "Created On").dt.strftime("%Y-%m"),
    },
    "created_quarter": {
        "requires": ["Created On"],
        "description": "Calendar quarter (YYYY-Q#) the invoice was created in.",
        "compute": lambda df: _dt(df, "Created On").dt.year.astype("Int64").astype(str) + "-Q" + _dt(df, "Created On").dt.quarter.astype("Int64").astype(str),
    },
    "effective_vat_rate_pct": {
        "requires": ["Subtotal", "Total VAT"],
        "description": f"Total VAT / Subtotal * 100 - should sit close to {VAT_RATE_PCT:.0f}%; "
                        "the same ratio reconcile.py's arithmetic-integrity check flags when it "
                        "drifts, exposed here as a per-row feature rather than a pass/fail.",
        "compute": lambda df: (_num(df, "Total VAT") / _num(df, "Subtotal").replace(0, np.nan)) * 100,
    },
    "vat_rate_deviation_flag": {
        "requires": ["Subtotal", "Total VAT"],
        "description": f"True if effective_vat_rate_pct is more than 2 points from {VAT_RATE_PCT:.0f}%.",
        "compute": lambda df: (
            (_num(df, "Total VAT") / _num(df, "Subtotal").replace(0, np.nan) * 100 - VAT_RATE_PCT).abs() > 2.0
        ),
    },
    "discount_rate_pct": {
        "requires": ["Discount Amount", "Subtotal"],
        "description": "Discount Amount / Subtotal * 100 - how much of the pre-VAT value was discounted.",
        "compute": lambda df: (_num(df, "Discount Amount") / _num(df, "Subtotal").replace(0, np.nan)) * 100,
    },
    "payment_completion_pct": {
        "requires": ["Payment Received", "Total incl. VAT"],
        "description": "Payment Received / Total incl. VAT * 100, capped at [0, 100] - how much "
                        "of the invoice has actually been settled.",
        "compute": lambda df: (
            (_num(df, "Payment Received") / _num(df, "Total incl. VAT").replace(0, np.nan)) * 100
        ).clip(lower=0, upper=100),
    },
    "is_overdue": {
        "requires": ["Due date", "Status Reason"],
        "description": "True if Due date is in the past and the invoice isn't Paid/Cancelled/Written Off.",
        "compute": lambda df, now: (
            (_dt(df, "Due date") < now) & (~df["Status Reason"].isin(["Paid", "Cancelled", "Written Off"]))
        ),
    },
    "days_overdue": {
        "requires": ["Due date", "Status Reason"],
        "description": "Days past Due date for invoices flagged is_overdue, else 0.",
        "compute": lambda df, now: np.where(
            (_dt(df, "Due date") < now) & (~df["Status Reason"].isin(["Paid", "Cancelled", "Written Off"])),
            (now - _dt(df, "Due date")).dt.days,
            0,
        ),
    },
    "is_high_value": {
        "requires": ["Total incl. VAT", "Status Reason"],
        "description": f"True if Total incl. VAT is at/above the {HIGH_VALUE_PERCENTILE*100:.0f}th "
                        "percentile among non-cancelled invoices - flags the invoices that matter "
                        "most to totals, useful for prioritising manual review.",
        "compute": lambda df: _num(df, "Total incl. VAT") >= pd.to_numeric(
            df.loc[df["Status Reason"] != "Cancelled", "Total incl. VAT"], errors="coerce"
        ).quantile(HIGH_VALUE_PERCENTILE),
    },
    "amount_per_lineitem_ratio": {
        "requires": ["Total incl. VAT", "Line Items Total"],
        "description": "Total incl. VAT / Line Items Total - should sit close to 1 unless admin "
                        "fees/interest were added; a ratio far from 1 is worth a look the same way "
                        "reconcile.py's arithmetic checks are.",
        "compute": lambda df: _num(df, "Total incl. VAT") / _num(df, "Line Items Total").replace(0, np.nan),
    },
    "consultant_monthly_invoice_count": {
        "requires": ["Consultant", "Created On"],
        "description": "How many invoices this row's consultant raised in this row's calendar "
                        "month - a contextual/aggregate feature (not a pure per-row ratio) useful "
                        "for consultant productivity or workload analysis.",
        "compute": lambda df: df.groupby([df["Consultant"], _dt(df, "Created On").dt.to_period("M")])[
            "Consultant"
        ].transform("count"),
    },
    "commission_consistency_flag": {
        "requires": ["Commission Amount", "Subtotal", "Commission (%)"],
        "description": "True where a filled-in Commission Amount doesn't match Subtotal * "
                        "Commission (%) within a small tolerance - only evaluated on rows where "
                        "Commission Amount is actually filled (it's rarely-used per the field-usage "
                        "classifier's report, so most rows will be NaN here, not False).",
        "compute": lambda df: pd.Series(
            np.where(
                df["Commission Amount"].notna(),
                (_num(df, "Commission Amount") - _num(df, "Subtotal") * _num(df, "Commission (%)") / 100).abs() > 1.0,
                np.nan,
            ),
            index=df.index,
        ),
    },
    "payment_duration_days": {
        "requires": [],
        "requires_any_of": ["Payment Date", "Stamped Payment Date"],
        "description": "Days between Created On and the invoice's actual payment date - the "
                        "engineered-feature twin of reconcile.py's 'Avg. Payment Duration' KPI "
                        "card, and blocked by the exact same export gap: neither 'Payment Date' "
                        "nor 'Stamped Payment Date' (icon_paymentdatenew / icon_stampedpaymentdate) "
                        "is in the default 'My Team's Open Invoices' view.",
        "compute": lambda df: (
            _dt(df, "Payment Date" if "Payment Date" in df.columns else "Stamped Payment Date")
            - _dt(df, "Created On")
        ).dt.days,
    },
}


def check_availability(df, spec):
    missing_required = [c for c in spec["requires"] if c not in df.columns]
    any_of = spec.get("requires_any_of")
    missing_any_of = []
    if any_of:
        present_any = [c for c in any_of if c in df.columns]
        if not present_any:
            missing_any_of = any_of
    available = not missing_required and not missing_any_of
    return available, missing_required, missing_any_of


def engineer_all(df, now=None):
    """Computes every feature whose required columns are present. Returns
    (enriched_df, computed_names, blocked) where blocked is a list of
    (name, missing_required, missing_any_of) for features that couldn't run."""
    import inspect

    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now("UTC").tz_localize(None)
    enriched = df.copy()
    computed = []
    blocked = []

    for name, spec in FEATURE_REGISTRY.items():
        available, missing_required, missing_any_of = check_availability(df, spec)
        if not available:
            blocked.append((name, missing_required, missing_any_of, spec["description"]))
            continue
        fn = spec["compute"]
        try:
            if "now" in inspect.signature(fn).parameters:
                result = fn(df, now)
            else:
                result = fn(df)
        except Exception as e:  # pragma: no cover - defensive, so one bad feature doesn't kill the run
            blocked.append((name, [], [], f"{spec['description']} [FAILED: {e}]"))
            continue
        enriched[name] = result
        computed.append(name)

    return enriched, computed, blocked


def summarize_feature(df, name):
    """One-line evidence string per computed feature, mirroring the style
    of classify_fields.py's per-field reasons - so the report is legible
    without opening the CSV."""
    s = df[name]
    non_null = s.notna().sum()
    n = len(s)
    if pd.api.types.is_bool_dtype(s) or s.dropna().isin([True, False]).all():
        true_count = (s == True).sum()  # noqa: E712
        return f"{non_null}/{n} rows evaluable, {true_count} flagged True"
    if pd.api.types.is_numeric_dtype(s):
        clean = pd.to_numeric(s, errors="coerce").dropna()
        if clean.empty:
            return f"{non_null}/{n} rows filled, no numeric values"
        return f"{non_null}/{n} rows filled, mean={clean.mean():,.2f}, min={clean.min():,.2f}, max={clean.max():,.2f}"
    return f"{non_null}/{n} rows filled, {s.nunique(dropna=True)} distinct values"


def load_table(path, sheet=None):
    if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        xl = pd.ExcelFile(path)
        sheet_name = sheet or xl.sheet_names[0]
        return xl.parse(sheet_name), sheet_name
    return pd.read_csv(path), path


def print_report(df, enriched, computed, blocked, source):
    print("=" * 78)
    print(f"Feature Engineering - {source} ({len(df)} rows, {len(df.columns)} source columns)")
    print("=" * 78)

    print(f"\n-- Computed ({len(computed)} of {len(FEATURE_REGISTRY)} registered features) --")
    for name in computed:
        print(f"  {name:<34} {summarize_feature(enriched, name)}")
        print(f"      {FEATURE_REGISTRY[name]['description']}")

    print(f"\n-- Blocked ({len(blocked)} feature(s) - add the named column(s) to unlock) --")
    if not blocked:
        print("  none - every registered feature was computable from this export")
    for name, missing_required, missing_any_of, description in blocked:
        need = []
        if missing_required:
            need.append(f"requires: {missing_required}")
        if missing_any_of:
            need.append(f"requires one of: {missing_any_of}")
        print(f"  {name:<34} {'; '.join(need)}")
        print(f"      {description}")

    print(f"\nSummary: {len(computed)} feature(s) computed, {len(blocked)} blocked by a missing "
          f"export column. This is additive - it doesn't change or remove any source column, "
          f"see the tool's README for how to add your own feature to FEATURE_REGISTRY.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="Path to an .xlsx or .csv invoice export")
    ap.add_argument("--sheet", default=None, help="Sheet name for .xlsx (default: first sheet)")
    ap.add_argument("--out", default=None, help="Write source columns + engineered features to this CSV path")
    ap.add_argument("--list-only", action="store_true", help="Print the report; don't compute or write anything")
    args = ap.parse_args()

    df, source = load_table(args.path, args.sheet)

    if args.list_only:
        print(f"Registered features ({len(FEATURE_REGISTRY)}):")
        for name, spec in FEATURE_REGISTRY.items():
            available, missing_required, missing_any_of = check_availability(df, spec)
            status = "AVAILABLE" if available else "BLOCKED"
            print(f"  [{status:<9}] {name}")
        return

    enriched, computed, blocked = engineer_all(df)
    print_report(df, enriched, computed, blocked, source)

    if args.out:
        enriched.to_csv(args.out, index=False)
        print(f"\nWrote {len(df.columns)} source + {len(computed)} engineered columns to {args.out}")


if __name__ == "__main__":
    main()
