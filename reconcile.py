#!/usr/bin/env python3
"""
Invoice KPI Reconciliation Tool
================================

Recomputes the riivo_invoice_ui_dashboard.html KPI cards and three of its
charts (Invoices by Month, Invoice Breakdown by type, Invoices by Status)
directly from an invoice Excel export ("My Team's Open Invoices" view, or
any export with the same 56 columns), so a PM can check the dashboard's
numbers against the underlying data without a manual Excel walkthrough.

This is a *deliberate* line-by-line port of the dashboard's own JS
aggregation functions (computeKpis, computeMonthly, computeDonut,
computeStatusBreakdown in riivo_invoice_ui_dashboard.html, InvoicingAgg
module) - not a re-derivation of the logic from scratch. Every exclusion
rule, VAT fallback, and rounding choice below has a comment pointing back
to the JS line it mirrors, so if the dashboard's logic changes, this file
is easy to re-sync by diffing against the same functions.

Known field-name mapping (Dataverse logical field -> export column):
    riivo_totalinclvat   -> "Total incl. VAT"
    riivo_subtotal        -> "Subtotal"
    ttt_paymentreceived    -> "Payment Received"
    ttt_discountamount     -> "Discount Amount"
    ttt_sarsreimbursement  -> "SARS Reimbursement"
    createdon              -> "Created On"
    riivo_invoicetype      -> "Invoice Type"   (export already carries the
                                                 FormattedValue label, same
                                                 as typeLabelOf() resolves to)
    statuscode              -> "Status Reason"  (export already carries the
                                                  FormattedValue label, same
                                                  as statusOf() resolves to)

Known limitation: Avg. Payment Duration (one of the 12 KPI cards) needs
icon_paymentdatenew / icon_stampedpaymentdate, neither of which is present
in this export view. That KPI is reported as "not checkable from this
export" rather than guessed at - see NOT_CHECKABLE below. This is also the
first entry in ENHANCEMENT_FIELDS: the same "field not in this export"
gap, generalized into a registry so any KPI (or, in engineer_features.py,
any engineered feature) blocked by a missing export column says so
explicitly and names the column that would unlock it - rather than the
tool silently skipping something a PM didn't know was possible.

This version adds a second report beyond the original KPI reconciliation:
an anomaly-detection pass (see run_anomaly_detection) that goes past the
single "25x the 95th percentile" money-field scan to also check arithmetic
integrity (does Subtotal + VAT actually equal Total incl. VAT?), duplicate/
missing Invoice IDs, date logic (due date before created-on, etc.), generic
statistical outliers on every numeric column (not just the 5 money fields),
a month-over-month volume/value anomaly check, and rare Status Reason
values. Every detector is independent and labelled, so a false-positive in
one doesn't hide a real finding in another.

Usage:
    python3 reconcile.py <path-to-export.xlsx> [--years 2026] [--sheet NAME]
    python3 reconcile.py <path-to-export.xlsx> --compare dashboard_values.json

`--compare` takes a JSON file of {kpi_key: value} as currently shown on the
dashboard (see compare_template.json) and reports any KPI whose computed
value differs from the dashboard's by more than a small rounding tolerance.
"""

import argparse
import json
import math
import sys
from datetime import datetime

import pandas as pd

COL = {
    "totalinclvat": "Total incl. VAT",
    "subtotal": "Subtotal",
    "paymentreceived": "Payment Received",
    "discountamount": "Discount Amount",
    "sarsreimbursement": "SARS Reimbursement",
    "createdon": "Created On",
    "invoicetype": "Invoice Type",
    "statusreason": "Status Reason",
}

NOT_CHECKABLE = {
    "avgPaymentDuration": (
        "needs icon_paymentdatenew / icon_stampedpaymentdate - not present "
        "in this export view"
    ),
}

REQUIRED_COLUMNS = list(COL.values())


# ---------------------------------------------------------------------------
# Desired/required-fields registry.
#
# The "My Team's Open Invoices" view is one Dataverse export among many
# possible ones - a PM can add columns to that view at any time. This
# registry names the export columns that are known to be missing from the
# *default* view but would unlock something if added, so the tool can say
# "add column X to unlock Y" instead of silently doing less than it could.
# Add an entry here (and the matching one in engineer_features.py's
# FEATURE_REGISTRY, if it's feature-engineering-facing) whenever a new gap
# like this is found - this is the intended extension point.
# ---------------------------------------------------------------------------
ENHANCEMENT_FIELDS = {
    "Payment Date": {
        "aliases": ["Payment Date", "icon_paymentdatenew"],
        "unlocks": [
            "Avg. Payment Duration KPI card",
            "payment_duration_days feature (engineer_features.py)",
        ],
    },
    "Stamped Payment Date": {
        "aliases": ["Stamped Payment Date", "icon_stampedpaymentdate"],
        "unlocks": [
            "Avg. Payment Duration KPI card (fallback source)",
            "payment_duration_days feature (engineer_features.py)",
        ],
    },
}


def print_field_coverage(df):
    """Reports export-column coverage against both what the tool already
    needs (REQUIRED_COLUMNS) and what it could use if added
    (ENHANCEMENT_FIELDS) - a single place to see 'what would one more
    export column buy me' rather than discovering it KPI by KPI."""
    cols = set(df.columns)
    missing_required = [c for c in REQUIRED_COLUMNS if c not in cols]
    print("\n-- Field coverage --")
    if missing_required:
        print(f"  MISSING required columns (KPIs using these will be wrong or zero): {missing_required}")
    else:
        print(f"  All {len(REQUIRED_COLUMNS)} required columns present.")

    for canonical, spec in ENHANCEMENT_FIELDS.items():
        present = next((a for a in spec["aliases"] if a in cols), None)
        if present:
            print(f"  [available] {canonical!r} present as {present!r} - unlocks: {', '.join(spec['unlocks'])}")
        else:
            print(f"  [missing]   {canonical!r} not in this export - add it to unlock: {', '.join(spec['unlocks'])}")


# ---------------------------------------------------------------------------
# Row-level helpers - each one mirrors a specific JS function by name/line.
# ---------------------------------------------------------------------------

def num(row, field):
    """Mirrors num(row, field) at riivo_invoice_ui_dashboard.html:2880."""
    v = row.get(field)
    try:
        v = float(v)
    except (TypeError, ValueError):
        return 0.0
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return 0.0
    return v


def status_of(row):
    """Mirrors statusOf(row) at :2886. Export already carries the
    FormattedValue label in 'Status Reason', same as the JS annotation path."""
    v = row.get(COL["statusreason"])
    if v is None or (isinstance(v, float) and math.isnan(v)) or v == "":
        return ""
    return str(v)


def is_cancelled(row):
    """Mirrors isCancelled(row) at :2913."""
    return status_of(row) == "Cancelled"


def is_paid(status):
    """Mirrors isPaid(s) at :2959."""
    return status == "Paid"


def type_label_of(row):
    """Mirrors typeLabelOf(row) at :2963."""
    v = row.get(COL["invoicetype"])
    if v is None or (isinstance(v, float) and math.isnan(v)) or v == "":
        return "Unspecified"
    return str(v)


def ym_of(value):
    """Mirrors ymOf(value) at :2918. Returns (year, month0-11) or None.
    NOTE: the dashboard uses getUTCFullYear/getUTCMonth on a Dataverse UTC
    timestamp; this export's 'Created On' column is read as-is (naive
    datetime as exported by the view) - see the README's dated-boundary
    caveat if a created-on invoice sits within a few hours of midnight."""
    if pd.isna(value):
        return None
    d = pd.to_datetime(value)
    if pd.isna(d):
        return None
    return (d.year, d.month - 1)


def in_years(year, years):
    """Mirrors inYears(year, years) at :2953."""
    if year is None:
        return False
    if not years:
        return True
    return year in years


# ---------------------------------------------------------------------------
# Aggregations - each mirrors a compute* function in InvoicingAgg.
# ---------------------------------------------------------------------------

def months_in_period(range_from, range_to, now=None):
    """Mirrors monthsInPeriod(startDate, endDate) at :2971."""
    if range_from is None or range_to is None:
        return 1
    now = now or datetime.utcnow()
    end = min(range_to, now)
    months = (end.year - range_from.year) * 12 + (end.month - range_from.month) + 1
    return max(1, min(12, months))


def compute_kpis(rows, years, range_from=None, range_to=None):
    """Mirrors computeKpis(rows, years, rangeFrom, rangeTo) at :2995-3043."""
    total_invoices = count_paid = accounting_count = tax_count = 0
    total_billed = total_billed_exvat = total_received = discount = refunds = 0.0
    included_ids = []
    excluded_cancelled = excluded_out_of_scope = 0

    for row in rows:
        if is_cancelled(row):
            excluded_cancelled += 1
            continue
        c = ym_of(row.get(COL["createdon"]))
        if not c or not in_years(c[0], years):
            excluded_out_of_scope += 1
            continue

        total_invoices += 1
        included_ids.append(row.get("Invoice ID"))
        if is_paid(status_of(row)):
            count_paid += 1
        if type_label_of(row) == "Accounting":
            accounting_count += 1
        elif type_label_of(row) == "Tax":
            tax_count += 1
        total_billed += num(row, COL["totalinclvat"])
        total_billed_exvat += num(row, COL["subtotal"])
        total_received += num(row, COL["paymentreceived"])
        discount += num(row, COL["discountamount"])
        refunds += num(row, COL["sarsreimbursement"])

    months = months_in_period(range_from, range_to)
    have_subtotal = total_billed > 0 and total_billed_exvat > 0
    vat_ratio = (total_billed_exvat / total_billed) if have_subtotal else 1.0

    return {
        "totalInvoices": total_invoices,
        "paidPct": (count_paid / total_invoices * 100) if total_invoices else 0.0,
        "totalBilled": total_billed,
        "totalBilledExVat": total_billed_exvat if have_subtotal else total_billed,
        "totalReceived": total_received,
        "totalReceivedExVat": total_received * vat_ratio,
        "avgInvoice": (total_billed / total_invoices) if total_invoices else 0.0,
        "avgInvoiceExVat": (
            (total_billed_exvat if have_subtotal else total_billed) / total_invoices
        ) if total_invoices else 0.0,
        "discount": discount,
        "discountExVat": discount * vat_ratio,
        "refunds": refunds,
        "refundsExVat": refunds * vat_ratio,
        "avgAccInvoice": accounting_count / months,
        "avgTaxInvoice": tax_count / months,
        "monthsInPeriod": months,
        "_excludedCancelled": excluded_cancelled,
        "_excludedOutOfScope": excluded_out_of_scope,
        "_includedInvoiceIds": included_ids,
    }


MONTH_NAMES = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


def compute_monthly(rows, years):
    """Mirrors computeMonthly(rows, years) at :3080-3112."""
    billed = [0.0] * 12
    billed_exvat = [0.0] * 12
    received = [0.0] * 12
    invoice_count = [0] * 12
    paid_count = [0] * 12
    total_billed = total_billed_exvat = 0.0

    for row in rows:
        if is_cancelled(row):
            continue
        c = ym_of(row.get(COL["createdon"]))
        if not c or not in_years(c[0], years):
            continue
        m = c[1]
        b = num(row, COL["totalinclvat"])
        bx = num(row, COL["subtotal"])
        billed[m] += b
        billed_exvat[m] += bx
        total_billed += b
        total_billed_exvat += bx
        received[m] += num(row, COL["paymentreceived"])
        invoice_count[m] += 1
        if is_paid(status_of(row)):
            paid_count[m] += 1

    have_exvat = total_billed > 0 and total_billed_exvat > 0
    vat_ratio = (total_billed_exvat / total_billed) if have_exvat else 1.0

    out = []
    for m in range(12):
        out.append({
            "month": MONTH_NAMES[m],
            "billed": billed[m],
            "billedExVat": billed_exvat[m] if have_exvat else billed[m],
            "received": received[m],
            "receivedExVat": received[m] * vat_ratio,
            "invoiceCount": invoice_count[m],
            "paidCount": paid_count[m],
        })
    return out


def compute_donut(rows, years):
    """Mirrors computeDonut(rows, years) at :3120-3159."""
    by_type, by_type_exvat, by_type_count = {}, {}, {}
    for row in rows:
        if is_cancelled(row):
            continue
        c = ym_of(row.get(COL["createdon"]))
        if not c or not in_years(c[0], years):
            continue
        t = type_label_of(row)
        by_type[t] = by_type.get(t, 0.0) + num(row, COL["totalinclvat"])
        by_type_exvat[t] = by_type_exvat.get(t, 0.0) + num(row, COL["subtotal"])
        by_type_count[t] = by_type_count.get(t, 0) + 1

    def segments(by):
        return sorted(
            ({"label": k, "value": v} for k, v in by.items() if v > 0),
            key=lambda s: -s["value"],
        )

    return {
        "segments": segments(by_type),
        "segmentsExVat": segments(by_type_exvat) if sum(by_type_exvat.values()) > 0 else segments(by_type),
        "countSegments": segments(by_type_count),
        "total": sum(by_type.values()),
        "totalExVat": sum(by_type_exvat.values()) if sum(by_type_exvat.values()) > 0 else sum(by_type.values()),
        "totalCount": sum(by_type_count.values()),
    }


def compute_status_breakdown(rows, years):
    """Mirrors computeStatusBreakdown(rows, years) at :3168-3200.
    Deliberately does NOT exclude cancelled invoices - matches the dashboard's
    explicit 2026-08-05 client decision that this one chart shows the full
    status mix, cancellations included."""
    by_status, by_status_exvat, by_status_count = {}, {}, {}
    for row in rows:
        c = ym_of(row.get(COL["createdon"]))
        if not c or not in_years(c[0], years):
            continue
        s = status_of(row) or "Unspecified"
        by_status[s] = by_status.get(s, 0.0) + num(row, COL["totalinclvat"])
        by_status_exvat[s] = by_status_exvat.get(s, 0.0) + num(row, COL["subtotal"])
        by_status_count[s] = by_status_count.get(s, 0) + 1

    def segments(by):
        return sorted(
            ({"label": k, "value": v} for k, v in by.items() if v > 0),
            key=lambda s: -s["value"],
        )

    return {
        "segments": segments(by_status),
        "countSegments": segments(by_status_count),
        "total": sum(by_status.values()),
        "totalCount": sum(by_status_count.values()),
    }


# ---------------------------------------------------------------------------
# Report rendering + compare mode
# ---------------------------------------------------------------------------

def fmt_currency(v):
    return f"R{v:,.2f}"


OUTLIER_FIELDS = [
    COL["totalinclvat"],
    COL["subtotal"],
    COL["sarsreimbursement"],
    COL["paymentreceived"],
    "Outstanding",
]


def scan_outliers(rows, multiple=25):
    """Not a KPI recompute - a sanity pass. Flags any row where a money field
    is more than `multiple`x the 95th percentile of that field's non-zero
    values. This is deliberately loose (catches gross data-entry/mapping
    errors, e.g. one field's value leaking into another), not a statistical
    fraud check - a flagged row is 'worth a manual look', not a proven bug.
    Every compute* aggregation above sums these fields directly, so one
    corrupted row can silently dominate a total (see the tool's README)."""
    import numpy as np

    df = pd.DataFrame(rows)
    flags = []
    for field in OUTLIER_FIELDS:
        if field not in df.columns:
            continue
        vals = pd.to_numeric(df[field], errors="coerce").fillna(0)
        nonzero = vals[vals > 0]
        if len(nonzero) < 10:
            continue
        p95 = np.percentile(nonzero, 95)
        threshold = p95 * multiple
        if threshold <= 0:
            continue
        hits = df[vals > threshold]
        for _, row in hits.iterrows():
            flags.append({
                "field": field,
                "value": row.get(field),
                "p95_of_field": p95,
                "invoice_id": row.get("Invoice ID"),
                "name": row.get("Name"),
                "status": row.get(COL["statusreason"]),
                "created_on": row.get(COL["createdon"]),
            })
    return flags


# ---------------------------------------------------------------------------
# Anomaly detection - six independent detectors beyond the extreme-value
# scan above. Each returns a list of finding dicts; none of them mutate the
# data or the KPI numbers - this is entirely a "worth a manual look" layer
# on top of the reconciliation, same spirit as scan_outliers.
# ---------------------------------------------------------------------------

def check_arithmetic_integrity(df, vat_rate_pct=15.0, vat_tolerance_pts=2.0, money_tolerance=1.0):
    """Row-level bookkeeping checks: does Subtotal + VAT actually equal
    Total incl. VAT? Is the effective VAT rate near the expected 15%? Has
    Payment Received exceeded what was billed? Does Outstanding reconcile
    with Total incl. VAT minus Payment Received? These are things a single
    KPI total can never surface, because KPIs sum a field rather than
    checking it against another field on the same row."""
    findings = []
    need = {"subtotal": COL["subtotal"], "vat": "Total VAT", "total": COL["totalinclvat"],
            "received": COL["paymentreceived"], "outstanding": "Outstanding"}
    if not all(c in df.columns for c in need.values()):
        return findings

    sub = pd.to_numeric(df[need["subtotal"]], errors="coerce")
    vat = pd.to_numeric(df[need["vat"]], errors="coerce")
    total = pd.to_numeric(df[need["total"]], errors="coerce")
    received = pd.to_numeric(df[need["received"]], errors="coerce")
    outstanding = pd.to_numeric(df[need["outstanding"]], errors="coerce")

    calc_total = sub + vat
    mismatch = (calc_total - total).abs() > money_tolerance
    for idx in df.index[mismatch.fillna(False)]:
        findings.append({
            "category": "Arithmetic: Subtotal + VAT != Total incl. VAT",
            "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
            "detail": f"Subtotal {sub[idx]:,.2f} + VAT {vat[idx]:,.2f} = {calc_total[idx]:,.2f}, "
                      f"but Total incl. VAT = {total[idx]:,.2f}",
        })

    rate = (vat / sub.replace(0, pd.NA)) * 100
    rate_off = (rate - vat_rate_pct).abs() > vat_tolerance_pts
    for idx in df.index[rate_off.fillna(False)]:
        findings.append({
            "category": f"Arithmetic: effective VAT rate far from {vat_rate_pct:.0f}%",
            "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
            "detail": f"VAT/Subtotal = {rate[idx]:.1f}% (Subtotal {sub[idx]:,.2f}, VAT {vat[idx]:,.2f})",
        })

    overpaid = (received - total) > money_tolerance
    for idx in df.index[overpaid.fillna(False)]:
        findings.append({
            "category": "Arithmetic: Payment Received exceeds Total incl. VAT",
            "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
            "detail": f"Received {received[idx]:,.2f} > Total incl. VAT {total[idx]:,.2f}",
        })

    calc_outstanding = (total - received).clip(lower=0)
    outstanding_off = (calc_outstanding - outstanding).abs() > money_tolerance
    for idx in df.index[outstanding_off.fillna(False)]:
        findings.append({
            "category": "Arithmetic: Outstanding != Total incl. VAT - Payment Received",
            "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
            "detail": f"Expected {calc_outstanding[idx]:,.2f}, export has {outstanding[idx]:,.2f}",
        })
    return findings


def check_referential(df):
    """Duplicate or missing Invoice IDs. A duplicate ID means two rows are
    either the same invoice exported twice (double-counts every KPI total)
    or two different invoices sharing an ID (a Dataverse numbering issue) -
    either way, worth a manual look before trusting sums that key off it."""
    findings = []
    if "Invoice ID" not in df.columns:
        return findings
    ids = df["Invoice ID"]
    missing = ids.isna() | (ids.astype(str).str.strip() == "")
    for idx in df.index[missing]:
        findings.append({
            "category": "Referential: missing Invoice ID",
            "invoice_id": None,
            "detail": f"row {idx} (Name: {df.at[idx, 'Name'] if 'Name' in df.columns else '?'}) has no Invoice ID",
        })
    dup_mask = ids.duplicated(keep=False) & ~missing
    for dup_id, group in df[dup_mask].groupby("Invoice ID"):
        findings.append({
            "category": "Referential: duplicate Invoice ID",
            "invoice_id": dup_id,
            "detail": f"{len(group)} rows share Invoice ID {dup_id!r}",
        })
    return findings


def check_date_logic(df, now=None):
    """Created-on/due-date/closed-date should obey basic ordering. A
    created-on date in the future or a due date before the invoice was even
    created usually means a date field was mapped wrong in the export, not
    that the business genuinely invoiced ahead of time."""
    findings = []
    now = now or datetime.utcnow()
    if COL["createdon"] in df.columns:
        created = pd.to_datetime(df[COL["createdon"]], errors="coerce")
        future = created > pd.Timestamp(now)
        for idx in df.index[future.fillna(False)]:
            findings.append({
                "category": "Date logic: Created On is in the future",
                "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
                "detail": f"Created On = {created[idx]}",
            })
        if "Due date" in df.columns:
            due = pd.to_datetime(df["Due date"], errors="coerce")
            before = due < created
            for idx in df.index[before.fillna(False)]:
                findings.append({
                    "category": "Date logic: Due date before Created On",
                    "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
                    "detail": f"Created On = {created[idx]}, Due date = {due[idx]}",
                })
        if "Closed Date" in df.columns:
            closed = pd.to_datetime(df["Closed Date"], errors="coerce")
            before_close = closed < created
            for idx in df.index[before_close.fillna(False)]:
                findings.append({
                    "category": "Date logic: Closed Date before Created On",
                    "invoice_id": df.at[idx, "Invoice ID"] if "Invoice ID" in df.columns else None,
                    "detail": f"Created On = {created[idx]}, Closed Date = {closed[idx]}",
                })
    return findings


def check_statistical_outliers_iqr(df, fields=None, k=3.0, max_examples=5):
    """Generic Tukey IQR outlier check across every numeric field passed in
    (default: all numeric columns), not just the 5 money fields the extreme-
    value scan covers. Deliberately a wider net at a looser multiplier (k=3
    on the IQR, the standard 'far outlier' fence) than scan_outliers's 25x
    p95 - this catches subtler things (e.g. one invoice's admin fee an order
    of magnitude off, not just a field mapped into the wrong unit)."""
    findings = []
    if fields is None:
        fields = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    for field in fields:
        if field not in df.columns:
            continue
        vals = pd.to_numeric(df[field], errors="coerce")
        clean = vals.dropna()
        if len(clean) < 20 or clean.nunique() <= 1:
            continue
        q1, q3 = clean.quantile(0.25), clean.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            continue
        lo, hi = q1 - k * iqr, q3 + k * iqr
        hits = df.index[(vals < lo) | (vals > hi)]
        if len(hits) == 0:
            continue
        examples = [
            f"Invoice {df.at[i, 'Invoice ID'] if 'Invoice ID' in df.columns else i}={vals[i]:,.2f}"
            for i in list(hits)[:max_examples]
        ]
        findings.append({
            "category": f"Statistical (IQR) outliers on '{field}'",
            "invoice_id": None,
            "detail": f"{len(hits)} row(s) outside [{lo:,.2f}, {hi:,.2f}] "
                      f"(Q1={q1:,.2f}, Q3={q3:,.2f}); examples: {', '.join(examples)}"
                      + (" ..." if len(hits) > max_examples else ""),
        })
    return findings


def check_monthly_time_series_anomaly(monthly, z_threshold=3.5):
    """Modified z-score (median + MAD, Iglewicz & Hoya's convention: 0.6745
    * (x - median) / MAD, flagged past the standard 3.5 threshold) across
    the months that actually have invoices - zero-activity months (e.g.
    future months with no data yet) are excluded so they don't get treated
    as a 'drop'. Median/MAD instead of mean/std deliberately: a plain
    mean-and-stddev z-score lets the spike and the matching drop it came
    from (invoices moved from one month into another, not lost or gained)
    inflate the std they're both being measured against, which can mask
    exactly the swap this check exists to catch. Median/MAD isn't pulled
    around by the outliers it's trying to measure."""
    findings = []
    active = [m for m in monthly if m["invoiceCount"] > 0]
    if len(active) < 4:
        return findings

    for metric, label in [("billed", "billed amount"), ("invoiceCount", "invoice count")]:
        values = [float(m[metric]) for m in active]
        sorted_vals = sorted(values)
        n = len(sorted_vals)
        median = sorted_vals[n // 2] if n % 2 else (sorted_vals[n // 2 - 1] + sorted_vals[n // 2]) / 2
        abs_devs = sorted([abs(v - median) for v in values])
        mad = abs_devs[n // 2] if n % 2 else (abs_devs[n // 2 - 1] + abs_devs[n // 2]) / 2
        if mad == 0:
            continue
        for m, v in zip(active, values):
            mod_z = 0.6745 * (v - median) / mad
            if abs(mod_z) >= z_threshold:
                direction = "spike" if mod_z > 0 else "drop"
                findings.append({
                    "category": f"Time series: monthly {label} {direction}",
                    "invoice_id": None,
                    "detail": f"{m['month']}: {v:,.2f} vs median {median:,.2f} across active months (modified z={mod_z:+.2f})",
                })
    return findings


def check_rare_categories(df, field=COL["statusreason"], threshold_pct=1.0):
    """Flags values of a categorical field that show up in less than
    threshold_pct of rows. Distinct from the field-usage classifier's
    column-level 'rarely used' bucket - this is about a rare *value* inside
    an otherwise fully-used column (e.g. one invoice with a Status Reason
    that's a typo or a one-off state nobody else uses), which a per-column
    fill-rate check can't see."""
    findings = []
    if field not in df.columns:
        return findings
    counts = df[field].value_counts(dropna=True)
    total = counts.sum()
    if total == 0:
        return findings
    for value, count in counts.items():
        pct = count / total * 100
        if pct < threshold_pct:
            findings.append({
                "category": f"Rare category value in '{field}'",
                "invoice_id": None,
                "detail": f"{value!r}: {count} row(s), {pct:.2f}% of non-null values",
            })
    return findings


def run_anomaly_detection(rows, monthly, vat_rate_pct=15.0, vat_tolerance_pts=2.0,
                           iqr_k=3.0, rare_category_pct=1.0):
    """Runs every detector and prints a grouped report. Returns the total
    finding count so callers (or a future CI check) can decide whether to
    treat 'zero findings' as a gate."""
    df = pd.DataFrame(rows)

    sections = [
        ("Extreme value scan (money fields, 25x 95th pct)",
         [{"category": "Extreme value", "invoice_id": f["invoice_id"],
           "detail": f"field={f['field']!r} value={f['value']:,.2f} vs 95th pct={f['p95_of_field']:,.2f} "
                     f"status={f['status']} created={f['created_on']}"}
          for f in scan_outliers(rows)]),
        ("Arithmetic integrity", check_arithmetic_integrity(df, vat_rate_pct, vat_tolerance_pts)),
        ("Referential (duplicate/missing IDs)", check_referential(df)),
        ("Date logic", check_date_logic(df)),
        ("Statistical (IQR) outliers, all numeric fields", check_statistical_outliers_iqr(df, k=iqr_k)),
        ("Time-series (month-over-month)", check_monthly_time_series_anomaly(monthly)),
        (f"Rare category values (<{rare_category_pct:.1f}%)", check_rare_categories(df, threshold_pct=rare_category_pct)),
    ]

    print("\n" + "=" * 72)
    print("Anomaly detection")
    print("=" * 72)
    total = 0
    for title, findings in sections:
        total += len(findings)
        if not findings:
            print(f"\n-- {title}: none --")
            continue
        print(f"\n-- {title}: {len(findings)} finding(s) --")
        for f in findings:
            inv = f"Invoice {f['invoice_id']!r}  " if f.get("invoice_id") not in (None, "") else ""
            print(f"  [{f['category']}] {inv}{f['detail']}")

    print(f"\nTotal anomaly findings across all detectors: {total}")
    if total:
        print("These are candidates for a manual look, not proven errors - see the tool's README "
              "for what each detector does and doesn't check.")
    return total


def load_rows(path, sheet=None):
    xl = pd.ExcelFile(path)
    sheet_name = sheet or xl.sheet_names[0]
    df = xl.parse(sheet_name)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print(f"WARNING: export is missing expected columns: {missing}", file=sys.stderr)
    return df.to_dict("records"), sheet_name, len(df)


def print_report(rows, years, total_rows, range_from=None, range_to=None):
    if range_from is None or range_to is None:
        created_dates = [ym_of(r.get(COL["createdon"])) for r in rows]
        created_actual = [pd.to_datetime(r.get(COL["createdon"])) for r in rows
                           if not pd.isna(r.get(COL["createdon"]))]
        if created_actual:
            in_scope = [d for d in created_actual if in_years(d.year, years)]
            if in_scope:
                range_from = range_from or min(in_scope).to_pydatetime()
                range_to = range_to or max(in_scope).to_pydatetime()
    print(f"NOTE: 'Avg .../month' KPIs need the dashboard's active date-range filter to match - "
          f"using inferred range [{range_from} .. {range_to}] from the data itself. Pass "
          f"--range-from/--range-to to match whatever window the dashboard was actually showing "
          f"when you compare its numbers.")
    kpis = compute_kpis(rows, years, range_from, range_to)
    print("=" * 72)
    print(f"Invoice KPI Reconciliation - years={years or 'ALL'}")
    print(f"Source rows: {total_rows}  |  Included: {kpis['totalInvoices']}  |  "
          f"Excluded (cancelled): {kpis['_excludedCancelled']}  |  "
          f"Excluded (out of year scope): {kpis['_excludedOutOfScope']}")
    print("=" * 72)

    print("\n-- KPI cards (10 of 12 - see note on Avg. Payment Duration) --")
    rows_out = [
        ("Total Invoices", kpis["totalInvoices"], None),
        ("Paid %", f"{kpis['paidPct']:.1f}%", None),
        ("Total Billed (incl VAT)", fmt_currency(kpis["totalBilled"]), None),
        ("Total Billed (ex VAT)", fmt_currency(kpis["totalBilledExVat"]), None),
        ("Total Received (incl VAT)", fmt_currency(kpis["totalReceived"]), None),
        ("Total Received (ex VAT, est.)", fmt_currency(kpis["totalReceivedExVat"]), None),
        ("Avg Invoice (incl VAT)", fmt_currency(kpis["avgInvoice"]), None),
        ("Avg Invoice (ex VAT)", fmt_currency(kpis["avgInvoiceExVat"]), None),
        ("Discount (incl VAT)", fmt_currency(kpis["discount"]), None),
        ("Refunds / SARS reimbursement (incl VAT)", fmt_currency(kpis["refunds"]), None),
        ("Avg Accounting invoices/month", f"{kpis['avgAccInvoice']:.2f}", None),
        ("Avg Tax invoices/month", f"{kpis['avgTaxInvoice']:.2f}", None),
    ]
    for label, value, _ in rows_out:
        print(f"  {label:<42} {value}")
    for key, reason in NOT_CHECKABLE.items():
        print(f"  {'Avg. Payment Duration':<42} NOT CHECKABLE - {reason}")

    monthly = compute_monthly(rows, years)
    print("\n-- Invoices by Month (billed incl VAT / received / count) --")
    for m in monthly:
        print(f"  {m['month']:<4} billed={fmt_currency(m['billed']):>16}  "
              f"received={fmt_currency(m['received']):>16}  "
              f"invoices={m['invoiceCount']:>4}  paid={m['paidCount']:>4}")

    donut = compute_donut(rows, years)
    print(f"\n-- Invoice Breakdown by type (total incl VAT: {fmt_currency(donut['total'])}) --")
    for seg in donut["segments"]:
        pct = (seg["value"] / donut["total"] * 100) if donut["total"] else 0
        print(f"  {seg['label']:<20} {fmt_currency(seg['value']):>16}  ({pct:.1f}%)")

    status = compute_status_breakdown(rows, years)
    print(f"\n-- Invoices by Status (total incl VAT, cancelled INCLUDED: "
          f"{fmt_currency(status['total'])}) --")
    for seg in status["segments"]:
        pct = (seg["value"] / status["total"] * 100) if status["total"] else 0
        print(f"  {seg['label']:<28} {fmt_currency(seg['value']):>16}  ({pct:.1f}%)")

    return kpis, monthly, donut, status


def run_compare(kpis, compare_path, tolerance_pct=0.5):
    with open(compare_path) as f:
        dashboard_values = json.load(f)

    print("\n" + "=" * 72)
    print(f"COMPARE vs dashboard-reported values ({compare_path})")
    print("=" * 72)
    any_mismatch = False
    for key, dash_val in dashboard_values.items():
        if key not in kpis:
            print(f"  {key:<30} SKIPPED - not a recognised KPI key")
            continue
        computed = kpis[key]
        if isinstance(computed, (int, float)) and isinstance(dash_val, (int, float)):
            if computed == 0 and dash_val == 0:
                diff_pct = 0.0
            elif computed == 0:
                diff_pct = 100.0
            else:
                diff_pct = abs(computed - dash_val) / abs(computed) * 100
            flag = "MISMATCH" if diff_pct > tolerance_pct else "match"
            if diff_pct > tolerance_pct:
                any_mismatch = True
            print(f"  {key:<30} dashboard={dash_val!r:<18} computed={computed!r:<18} "
                  f"diff={diff_pct:.2f}%  [{flag}]")
        else:
            flag = "match" if computed == dash_val else "MISMATCH"
            if flag == "MISMATCH":
                any_mismatch = True
            print(f"  {key:<30} dashboard={dash_val!r:<18} computed={computed!r:<18} [{flag}]")

    print("=" * 72)
    if any_mismatch:
        print("RESULT: at least one KPI does not match within tolerance - see MISMATCH rows above.")
    else:
        print("RESULT: every compared KPI matches within tolerance.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("export_path", help="Path to the invoice Excel export")
    ap.add_argument("--sheet", default=None, help="Sheet name (default: first sheet)")
    ap.add_argument("--years", nargs="*", type=int, default=None,
                     help="Filter to these created-on years, e.g. --years 2026 (default: all years present)")
    ap.add_argument("--compare", default=None,
                     help="Path to a JSON file of {kpiKey: dashboardValue} to diff against")
    ap.add_argument("--tolerance-pct", type=float, default=0.5,
                     help="Allowed %% difference before flagging a mismatch (default 0.5)")
    ap.add_argument("--range-from", default=None, help="YYYY-MM-DD - match the dashboard's active range filter")
    ap.add_argument("--range-to", default=None, help="YYYY-MM-DD - match the dashboard's active range filter")
    ap.add_argument("--vat-rate", type=float, default=15.0,
                     help="Expected VAT rate %% for the arithmetic-integrity check (default 15.0)")
    ap.add_argument("--vat-tolerance-pts", type=float, default=2.0,
                     help="Allowed +/- percentage points before an effective VAT rate is flagged (default 2.0)")
    ap.add_argument("--iqr-k", type=float, default=3.0,
                     help="IQR multiplier for the generic statistical outlier check (default 3.0, Tukey's 'far outlier' fence)")
    ap.add_argument("--rare-category-pct", type=float, default=1.0,
                     help="Flag a Status Reason value used by fewer than this %% of rows (default 1.0)")
    ap.add_argument("--skip-anomalies", action="store_true", help="Skip the anomaly-detection pass entirely")
    args = ap.parse_args()

    rows, sheet_name, total_rows = load_rows(args.export_path, args.sheet)
    print(f"Loaded '{sheet_name}' - {total_rows} rows from {args.export_path}")
    print_field_coverage(pd.DataFrame(rows))

    range_from = datetime.fromisoformat(args.range_from) if args.range_from else None
    range_to = datetime.fromisoformat(args.range_to) if args.range_to else None
    kpis, monthly, donut, status = print_report(rows, args.years, total_rows, range_from, range_to)

    if not args.skip_anomalies:
        run_anomaly_detection(
            rows, monthly,
            vat_rate_pct=args.vat_rate,
            vat_tolerance_pts=args.vat_tolerance_pts,
            iqr_k=args.iqr_k,
            rare_category_pct=args.rare_category_pct,
        )

    if args.compare:
        run_compare(kpis, args.compare, args.tolerance_pct)


if __name__ == "__main__":
    main()
