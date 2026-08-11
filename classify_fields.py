#!/usr/bin/env python3
"""
Generalized Field-Usage Classifier
===================================

Generalizes the by-hand invoice field audit (the one that found 230 of the
underlying Dataverse invoice fields could be trimmed to 143 actually-used
ones) into a reusable tool: point it at any CRM/Dataverse table export
(invoices, cases, accounts, whatever), and it classifies every column as
Actively Used, Sparsely Filled, Effectively Constant, or Unused - each with
the actual numbers behind the call, not just a label - plus a second pass
that flags column PAIRS that look like they might be duplicating each other.

This does not delete or recommend deleting anything on its own. It produces
a classification with evidence; a human still decides what actually gets
removed from a form or entity (same "recommends, doesn't authorize" pattern
as the offering-selector skill).

Usage:
    python3 classify_fields.py <path.xlsx|path.csv> [--sheet NAME] [--out report.csv]
    python3 classify_fields.py <path.xlsx> --pairs-only     # skip per-field, just check duplicates
"""

import argparse
import sys

import pandas as pd

RARE_THRESHOLD = 0.05       # below this fill rate -> "Rarely used"
SPARSE_THRESHOLD = 0.20     # below this fill rate (but above RARE) -> "Sparsely filled"
CONSTANT_SHARE = 0.98       # one value taking >= this share of filled rows -> "Effectively constant"
DUPLICATE_MATCH_THRESHOLD = 0.95  # pairwise agreement rate to flag as a possible duplicate
NUMERIC_CORR_THRESHOLD = 0.98


def is_blank(v):
    if pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def classify_column(series, col_name, n_rows):
    non_blank_mask = ~series.apply(is_blank)
    non_blank = series[non_blank_mask]
    fill_rate = len(non_blank) / n_rows if n_rows else 0.0
    distinct = non_blank.nunique(dropna=True)

    top_value = None
    top_share = 0.0
    if len(non_blank) > 0:
        vc = non_blank.value_counts(normalize=True)
        top_value = vc.index[0]
        top_share = float(vc.iloc[0])

    is_system_field = col_name.strip().startswith("(Do Not Modify)")

    if fill_rate == 0.0:
        label = "Unused (always empty)"
        reason = "every value is null/blank"
    elif distinct <= 1:
        label = "Constant (single value)"
        reason = f"only ever takes one value ({top_value!r}) whenever filled"
    elif top_share >= CONSTANT_SHARE:
        label = "Effectively constant"
        reason = f"{top_share*100:.1f}% of filled rows are the same value ({top_value!r})"
    elif fill_rate < RARE_THRESHOLD:
        label = "Rarely used"
        reason = f"filled in only {fill_rate*100:.1f}% of rows"
    elif fill_rate < SPARSE_THRESHOLD:
        label = "Sparsely filled"
        reason = f"filled in {fill_rate*100:.1f}% of rows - verify still needed before assuming dead"
    else:
        label = "Actively used"
        reason = f"filled in {fill_rate*100:.1f}% of rows with {distinct} distinct values"

    if is_system_field:
        reason = "(Do Not Modify) system field - " + reason

    return {
        "field": col_name,
        "classification": label,
        "fill_rate_pct": round(fill_rate * 100, 2),
        "distinct_values": int(distinct),
        "top_value": top_value,
        "top_value_share_pct": round(top_share * 100, 2),
        "system_field": is_system_field,
        "reason": reason,
    }


CLASSIFICATION_ORDER = [
    "Unused (always empty)",
    "Constant (single value)",
    "Effectively constant",
    "Rarely used",
    "Sparsely filled",
    "Actively used",
]


def classify_all(df):
    n_rows = len(df)
    rows = [classify_column(df[c], c, n_rows) for c in df.columns]
    result = pd.DataFrame(rows)
    order_map = {label: i for i, label in enumerate(CLASSIFICATION_ORDER)}
    result["_sort"] = result["classification"].map(order_map)
    result = result.sort_values(["_sort", "field"]).drop(columns="_sort").reset_index(drop=True)
    return result


def find_possible_duplicates(df):
    """Flags column pairs that appear to carry the same information twice.
    Two heuristics, kept separate so the reason is always legible:
      - categorical/text: exact-value agreement rate on rows where both are filled
      - numeric: Pearson correlation
    A flagged pair is 'worth checking', not a proven duplicate - e.g. Subtotal
    and Total incl. VAT will correlate highly by construction (one is a
    multiple of the other) without being the same field."""
    findings = []
    cols = list(df.columns)
    numeric_cols = [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
    non_numeric_cols = [c for c in cols if c not in numeric_cols]

    for i in range(len(numeric_cols)):
        for j in range(i + 1, len(numeric_cols)):
            a, b = numeric_cols[i], numeric_cols[j]
            sub = df[[a, b]].dropna()
            if len(sub) < 20:
                continue
            if sub[a].nunique() <= 1 or sub[b].nunique() <= 1:
                continue
            corr = sub[a].corr(sub[b])
            if pd.notna(corr) and abs(corr) >= NUMERIC_CORR_THRESHOLD:
                findings.append({
                    "field_a": a, "field_b": b, "kind": "numeric correlation",
                    "score_pct": round(abs(corr) * 100, 2),
                    "rows_compared": len(sub),
                })

    for i in range(len(non_numeric_cols)):
        for j in range(i + 1, len(non_numeric_cols)):
            a, b = non_numeric_cols[i], non_numeric_cols[j]
            sub = df[[a, b]].dropna()
            sub = sub[(sub[a].astype(str).str.strip() != "") & (sub[b].astype(str).str.strip() != "")]
            if len(sub) < 20:
                continue
            # Skip pairs where either side is already ~constant - a 100% "agreement"
            # between two fields that are each always "No" is not a duplicate finding,
            # it's two unrelated constant fields that happen to share a default.
            if sub[a].nunique() <= 1 or sub[b].nunique() <= 1:
                continue
            match_rate = (sub[a].astype(str) == sub[b].astype(str)).mean()
            if match_rate >= DUPLICATE_MATCH_THRESHOLD:
                findings.append({
                    "field_a": a, "field_b": b, "kind": "exact-value agreement",
                    "score_pct": round(match_rate * 100, 2),
                    "rows_compared": len(sub),
                })

    findings.sort(key=lambda f: -f["score_pct"])
    return findings


def load_table(path, sheet=None):
    if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        xl = pd.ExcelFile(path)
        sheet_name = sheet or xl.sheet_names[0]
        return xl.parse(sheet_name), sheet_name
    return pd.read_csv(path), path


def print_report(result, findings, total_rows, source):
    print("=" * 78)
    print(f"Field-Usage Classification - {source} ({total_rows} rows, {len(result)} fields)")
    print("=" * 78)

    for label in CLASSIFICATION_ORDER:
        subset = result[result["classification"] == label]
        if subset.empty:
            continue
        print(f"\n-- {label} ({len(subset)} field(s)) --")
        for _, row in subset.iterrows():
            sys_tag = " [system]" if row["system_field"] else ""
            print(f"  {row['field']:<45}{sys_tag}")
            print(f"      {row['reason']}")

    print("\n" + "=" * 78)
    print(f"Possible duplicate/overlapping field pairs ({len(findings)} found)")
    print("=" * 78)
    if not findings:
        print("  none above threshold")
    for f in findings:
        print(f"  {f['field_a']:<35} <-> {f['field_b']:<35} "
              f"{f['kind']} = {f['score_pct']}% (n={f['rows_compared']})")

    n_actionable = len(result[result["classification"].isin(
        ["Unused (always empty)", "Constant (single value)", "Effectively constant", "Rarely used"]
    )])
    print(f"\nSummary: {n_actionable} of {len(result)} fields are candidates for a closer look "
          f"(unused, constant, or rarely used). This is a starting list for a human review, "
          f"not a deletion order - see the tool's README.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="Path to an .xlsx or .csv export")
    ap.add_argument("--sheet", default=None, help="Sheet name for .xlsx (default: first sheet)")
    ap.add_argument("--out", default=None, help="Write the per-field classification to this CSV path")
    ap.add_argument("--pairs-only", action="store_true", help="Skip the per-field pass, only check for duplicate pairs")
    args = ap.parse_args()

    df, source = load_table(args.path, args.sheet)
    total_rows = len(df)

    findings = find_possible_duplicates(df)

    if args.pairs_only:
        print(f"Possible duplicate/overlapping field pairs in {source} ({len(findings)} found):")
        for f in findings:
            print(f"  {f['field_a']:<35} <-> {f['field_b']:<35} {f['kind']} = {f['score_pct']}%")
        return

    result = classify_all(df)
    print_report(result, findings, total_rows, source)

    if args.out:
        result.to_csv(args.out, index=False)
        print(f"\nWrote per-field classification to {args.out}")


if __name__ == "__main__":
    main()
