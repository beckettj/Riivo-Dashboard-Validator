"""
Invoice Tools - Interactive UI
==============================

A single Streamlit app wrapping the three existing CLI tools (reconcile.py,
classify_fields.py, engineer_features.py) so a PM can drag in an export and
see everything those scripts produce, without a terminal.

This app does NOT reimplement any of the reconciliation, anomaly-detection,
classification, or feature-engineering logic - it imports the three
scripts as modules and calls their functions directly, so results here are
guaranteed to match what the CLI tools would print for the same export.
The three .py files in this folder are synced copies of the canonical CLI
tools; if you edit the logic, edit the CLI tool and re-copy it here (see
README.md).

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import difflib
import re

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import classify_fields
import engineer_features
import reconcile

st.set_page_config(page_title="Invoice Tools", layout="wide")

# The (i) info-popover buttons on KPI cards should read as a plain icon, not
# a bordered button - this strips the default Streamlit button chrome
# (border/background/shadow) from just those buttons, leaving every other
# button (Add to dashboard, Remove, Add KPI, etc.) untouched.
st.markdown("""
<style>
button[data-testid="stPopoverButton"] {
    border: none !important;
    background: transparent !important;
    box-shadow: none !important;
    padding: 0 !important;
}
button[data-testid="stPopoverButton"]:hover {
    background: rgba(10, 42, 74, 0.08) !important;
    border: none !important;
}
button[data-testid="stPopoverButton"]:focus:not(:active) {
    border: none !important;
    box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

NAVY = "#0A2A4A"
YELLOW = "#F2E234"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7"]
CRITICAL = "#d03b3b"
GOOD = "#0ca30c"

# ---------------------------------------------------------------------------
# Safe custom-KPI expression evaluation
# ---------------------------------------------------------------------------

BLOCKED_TOKENS = [
    "__", "import", "exec", "eval(", "os.", "sys.", "subprocess", "open(",
    "lambda", ";", "getattr", "setattr", "globals", "locals", "compile(",
    "\\", "%", "input(",
]


def is_safe_expr(expr):
    lowered = expr.lower()
    return not any(tok in lowered for tok in BLOCKED_TOKENS)


def auto_backtick(expr, columns):
    """Wraps column names containing spaces/punctuation in backticks so
    pandas' eval/query can resolve them, without the user having to type
    backticks themselves."""
    result = expr
    for col in sorted(columns, key=len, reverse=True):
        if col in result and re.search(r"[^A-Za-z0-9_]", col):
            already = re.search(r"`\s*" + re.escape(col) + r"\s*`", result)
            if not already:
                result = result.replace(col, f"`{col}`")
    return result


AGG_FUNCS = {
    "Sum": lambda s: s.sum(),
    "Mean": lambda s: s.mean(),
    "Median": lambda s: s.median(),
    "Min": lambda s: s.min(),
    "Max": lambda s: s.max(),
    "Count": lambda s: s.count(),
    "Count Distinct": lambda s: s.nunique(),
    "% True": lambda s: (s.astype(bool).mean() * 100) if len(s) else 0.0,
}


def compute_custom_kpi(df, expr, agg, filter_expr=None):
    if not is_safe_expr(expr) or (filter_expr and not is_safe_expr(filter_expr)):
        raise ValueError("Expression contains a disallowed token.")
    work = df
    if filter_expr:
        work = work.query(auto_backtick(filter_expr, df.columns), engine="python")
    series = work.eval(auto_backtick(expr, df.columns), engine="python")
    if agg != "% True":
        series = pd.to_numeric(series, errors="coerce")
    return AGG_FUNCS[agg](series)


_KPI_KEYWORDS = {"and", "or", "not", "in", "true", "false", "none"}


def resolve_column_refs(expr, columns, cutoff=0.72):
    """Finds column references inside a formula/filter and reconciles them
    against the file's real column names - correcting typos, wrong casing,
    or missing punctuation (e.g. 'Total Incl VAT' -> 'Total incl. VAT')
    instead of just letting them fail. Returns (resolved_expr, corrections,
    unresolved): resolved_expr has every confidently-matched reference
    backticked; corrections is [(typed, matched), ...] for anything that
    needed fixing; unresolved is any phrase that looked like a column
    reference but had no good match, so the caller can report it instead of
    creating a KPI that just says 'error'."""
    result = expr
    columns = list(columns)

    # Pass 1: exact matches, longest name first so substrings don't clash.
    for col in sorted(columns, key=len, reverse=True):
        if col in result:
            already = re.search(r"`\s*" + re.escape(col) + r"\s*`", result)
            if not already:
                result = result.replace(col, f"`{col}`", 1)

    # Pass 2: whatever's left outside backticks/quotes might be a mistyped
    # column reference. Mask those spans (same length, so offsets still
    # line up with `result`) before hunting for candidate phrases.
    masked = re.sub(r"`[^`]*`", lambda m: "\0" * len(m.group()), result)
    masked = re.sub(r"'[^']*'", lambda m: "\0" * len(m.group()), masked)
    masked = re.sub(r'"[^"]*"', lambda m: "\0" * len(m.group()), masked)

    corrections = []
    unresolved = []
    matches = list(re.finditer(r"[A-Za-z][A-Za-z0-9 ./%()]{1,}[A-Za-z0-9)]", masked))
    # Process right-to-left so earlier match offsets stay valid as we splice.
    for m in reversed(matches):
        phrase = m.group().strip()
        if not phrase or phrase.lower() in _KPI_KEYWORDS or not re.search(r"[A-Za-z]{2,}", phrase):
            continue
        ci_match = next((c for c in columns if c.lower() == phrase.lower()), None)
        if ci_match:
            result = result[: m.start()] + f"`{ci_match}`" + result[m.end():]
            corrections.append((phrase, ci_match))
            continue
        best = difflib.get_close_matches(phrase, columns, n=1, cutoff=cutoff)
        if best:
            result = result[: m.start()] + f"`{best[0]}`" + result[m.end():]
            corrections.append((phrase, best[0]))
        else:
            unresolved.append(phrase)

    corrections.reverse()
    unresolved.reverse()
    return result, corrections, unresolved


def validate_custom_kpi(df, expr, agg, filter_expr=None, cutoff=0.72):
    """End-to-end check used before a custom KPI is ever added: resolves
    (and auto-corrects) column references, then actually computes it
    against this file. Only if this passes does the KPI get created -
    otherwise the caller shows an inline error and nothing is added, so a
    bad field name never turns into a permanent 'error' card."""
    if not is_safe_expr(expr) or (filter_expr and not is_safe_expr(filter_expr)):
        return False, expr, filter_expr, [], "Expression contains a disallowed token."

    resolved_expr, corr_expr, unresolved_expr = resolve_column_refs(expr, df.columns, cutoff)
    resolved_filter, corr_filter, unresolved_filter = (
        resolve_column_refs(filter_expr, df.columns, cutoff) if filter_expr else (None, [], [])
    )
    corrections = corr_expr + corr_filter
    unresolved = unresolved_expr + unresolved_filter
    if unresolved:
        names = ", ".join(f"'{u}'" for u in unresolved)
        return False, resolved_expr, resolved_filter, corrections, (
            f"Couldn't match {names} to any column in this file - check spelling."
        )

    try:
        compute_custom_kpi(df, resolved_expr, agg, resolved_filter)
    except Exception as e:
        return False, resolved_expr, resolved_filter, corrections, str(e)

    return True, resolved_expr, resolved_filter, corrections, None


def fmt_value(value, kind):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "-"
    if kind == "currency":
        return f"R{value:,.2f}"
    if kind == "pct":
        return f"{value:.1f}%"
    if kind == "int":
        return f"{int(value):,}"
    if kind == "float2":
        return f"{value:,.2f}"
    return f"{value:,.2f}" if isinstance(value, (int, float)) else str(value)


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def read_sheet_names(file_bytes, filename):
    if filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
        import io
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        return xl.sheet_names
    return None


@st.cache_data(show_spinner=True)
def load_dataframe(file_bytes, filename, sheet):
    import io
    if filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        return xl.parse(sheet or xl.sheet_names[0])
    return pd.read_csv(io.BytesIO(file_bytes))


DEFAULT_KPI_DEFS = [
    ("totalInvoices", "Total Invoices", "int"),
    ("paidPct", "Paid %", "pct"),
    ("totalBilled", "Total Billed (incl VAT)", "currency"),
    ("totalBilledExVat", "Total Billed (ex VAT)", "currency"),
    ("totalReceived", "Total Received (incl VAT)", "currency"),
    ("totalReceivedExVat", "Total Received (ex VAT, est.)", "currency"),
    ("avgInvoice", "Avg Invoice (incl VAT)", "currency"),
    ("avgInvoiceExVat", "Avg Invoice (ex VAT)", "currency"),
    ("discount", "Discount (incl VAT)", "currency"),
    ("refunds", "Refunds / SARS reimbursement", "currency"),
    ("avgAccInvoice", "Avg Accounting invoices/month", "float2"),
    ("avgTaxInvoice", "Avg Tax invoices/month", "float2"),
]

# Human-readable formula behind each default KPI, shown in the info popover
# next to its card. These describe compute_kpis() in reconcile.py in plain
# language rather than restating the code - keep in sync if that function's
# logic changes.
DEFAULT_KPI_FORMULAS = {
    "totalInvoices": (
        "Count of rows where Status Reason != 'Cancelled' AND Created On "
        "falls within the selected year(s)."
    ),
    "paidPct": (
        "count(Status Reason == 'Paid') / Total Invoices x 100, over the "
        "same included rows as Total Invoices."
    ),
    "totalBilled": "Sum of 'Total incl. VAT' across included rows.",
    "totalBilledExVat": (
        "Sum of 'Subtotal' across included rows (falls back to Total "
        "Billed (incl VAT) if Subtotal isn't usable for this export)."
    ),
    "totalReceived": "Sum of 'Payment Received' across included rows.",
    "totalReceivedExVat": (
        "Total Received (incl VAT) x (Total Billed ex VAT / Total Billed "
        "incl VAT) - an estimate using the overall billed VAT ratio, since "
        "Payment Received isn't itself split into VAT/ex-VAT."
    ),
    "avgInvoice": "Total Billed (incl VAT) / Total Invoices.",
    "avgInvoiceExVat": "Total Billed (ex VAT) / Total Invoices.",
    "discount": "Sum of 'Discount Amount' across included rows.",
    "refunds": "Sum of 'SARS Reimbursement' across included rows.",
    "avgAccInvoice": (
        "count(Invoice Type == 'Accounting') / months in the selected "
        "date range (inferred from Created On, capped 1-12)."
    ),
    "avgTaxInvoice": (
        "count(Invoice Type == 'Tax') / months in the selected date range "
        "(inferred from Created On, capped 1-12)."
    ),
}


def kpi_metric(label, value_str, formula_text, key):
    """Renders a metric card with a small (i) info button in the top-right
    corner. Clicking it opens a popover showing exactly what's behind the
    number - the plain-language formula for a default KPI, the engineered
    feature + aggregation for a recommended one, or the literal expression
    for a custom one - so nobody has to guess or go spelunking in the code."""
    head_l, head_r = st.columns([6, 1])
    with head_r:
        with st.popover("ℹ️", use_container_width=True):
            st.markdown(f"**{label}**")
            st.caption("Formula")
            st.code(formula_text, language=None)
    st.metric(label, value_str)


def recommend_kpis(enriched, computed_names, top_n=5):
    """Scores every successfully-computed engineered feature for how
    'interesting' it looks on THIS export, and returns the top N as
    recommended KPI candidates with a plain-language rationale. This is a
    heuristic, not a model: booleans score highest when they split the data
    into a real minority/majority rather than being almost-always-true or
    almost-always-false; numerics score highest when they have real spread
    (coefficient of variation) rather than being nearly constant."""
    scored = []
    for name in computed_names:
        s = enriched[name]
        non_null = s.dropna()
        if non_null.empty:
            continue
        is_boolish = non_null.isin([True, False]).all()
        if is_boolish:
            rate = float(non_null.mean())
            if rate <= 0.0 or rate >= 1.0:
                continue
            score = rate * (1 - rate) * 4
            rationale = (f"{rate*100:.1f}% of rows flag True - a real minority worth "
                         f"tracking as a headline %, not an edge case.")
            scored.append({"name": name, "score": score, "rationale": rationale,
                            "kind": "bool", "agg": "% True", "value": rate * 100})
        elif pd.api.types.is_numeric_dtype(non_null):
            mean, std = non_null.mean(), non_null.std()
            if mean == 0 or pd.isna(std) or std == 0:
                continue
            cv = abs(std / mean)
            score = min(cv, 3) / 3
            rationale = (f"Coefficient of variation {cv:.2f} - real spread across invoices "
                         f"rather than a near-constant value, worth a summary KPI.")
            scored.append({"name": name, "score": score, "rationale": rationale,
                            "kind": "numeric", "agg": "Mean", "value": mean})
    scored.sort(key=lambda x: -x["score"])
    return scored[:top_n]


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

st.sidebar.title("Invoice Tools")
st.sidebar.caption("Wraps reconcile.py, classify_fields.py and engineer_features.py in one interface.")

uploaded = st.sidebar.file_uploader("Invoice export", type=["xlsx", "xls", "xlsm", "csv"])

if not uploaded:
    st.title("Invoice Tools")
    st.write(
        "Upload an invoice export (the same 'My Team's Open Invoices' shape "
        "reconcile.py expects, or any similarly-shaped CRM export) in the "
        "sidebar to get started."
    )
    st.write(
        "This runs the same reconciliation, anomaly detection, field-usage "
        "classification, and feature engineering as the three CLI tools - "
        "just rendered here instead of printed to a terminal."
    )
    st.stop()

file_bytes = uploaded.getvalue()
sheet_names = read_sheet_names(file_bytes, uploaded.name)
sheet = None
if sheet_names and len(sheet_names) > 1:
    sheet = st.sidebar.selectbox("Sheet", sheet_names)
elif sheet_names:
    sheet = sheet_names[0]

df = load_dataframe(file_bytes, uploaded.name, sheet)
total_rows = len(df)

years_present = sorted({
    d.year for d in pd.to_datetime(df.get(reconcile.COL["createdon"]), errors="coerce").dropna()
}) if reconcile.COL["createdon"] in df.columns else []

years_selected = st.sidebar.multiselect(
    "Created-on years (empty = all)", years_present, default=years_present
)
years_filter = years_selected or None

with st.sidebar.expander("Anomaly detection thresholds"):
    vat_rate = st.slider("Expected VAT rate %", 0.0, 30.0, 15.0, 0.5)
    vat_tol = st.slider("VAT rate tolerance (+/- points)", 0.5, 10.0, 2.0, 0.5)
    iqr_k = st.slider("IQR outlier multiplier (k)", 1.0, 6.0, 3.0, 0.5)
    rare_pct = st.slider("Rare category threshold %", 0.1, 5.0, 1.0, 0.1)
    extreme_multiple = st.slider("Extreme-value multiple (x 95th pct)", 5, 100, 25, 5)

rows = df.to_dict("records")

# ---------------------------------------------------------------------------
# Shared computations (used across tabs)
# ---------------------------------------------------------------------------

def infer_range(df, years_filter):
    """Mirrors reconcile.print_report's own inferred-range logic, so
    avgAccInvoice/avgTaxInvoice divide by the actual number of active
    months in scope rather than defaulting to 1 (compute_kpis returns 1
    month when range_from/range_to are None - fine for the CLI, which
    always infers a range first, but the UI needs to do the same)."""
    if reconcile.COL["createdon"] not in df.columns:
        return None, None
    created = pd.to_datetime(df[reconcile.COL["createdon"]], errors="coerce").dropna()
    if years_filter:
        created = created[created.dt.year.isin(years_filter)]
    if created.empty:
        return None, None
    return created.min().to_pydatetime(), created.max().to_pydatetime()


range_from, range_to = infer_range(df, years_filter)
kpis = reconcile.compute_kpis(rows, years_filter, range_from, range_to)
monthly = reconcile.compute_monthly(rows, years_filter)
donut = reconcile.compute_donut(rows, years_filter)
status = reconcile.compute_status_breakdown(rows, years_filter)

if "custom_kpis" not in st.session_state:
    st.session_state.custom_kpis = []
if "kpi_feedback" not in st.session_state:
    st.session_state.kpi_feedback = None

st.title("Invoice Tools")
st.caption(f"{uploaded.name}" + (f" - sheet '{sheet}'" if sheet else "") + f" - {total_rows} rows, {len(df.columns)} columns")

tab_kpi, tab_anomaly, tab_fields, tab_features = st.tabs(
    ["KPI Dashboard", "Anomaly Detection", "Field Usage", "Feature Engineering"]
)

# ---------------------------------------------------------------------------
# Tab 1: KPI Dashboard
# ---------------------------------------------------------------------------

with tab_kpi:
    st.subheader("Field coverage")
    missing_required = [c for c in reconcile.REQUIRED_COLUMNS if c not in df.columns]
    cols = st.columns(3)
    with cols[0]:
        if missing_required:
            st.error(f"Missing required columns: {missing_required}")
        else:
            st.success(f"All {len(reconcile.REQUIRED_COLUMNS)} required columns present.")
    enh_items = list(reconcile.ENHANCEMENT_FIELDS.items())
    for i, (canonical, spec) in enumerate(enh_items):
        present = next((a for a in spec["aliases"] if a in df.columns), None)
        with cols[(i + 1) % 3]:
            if present:
                st.success(f"**{canonical}** available (as '{present}') - unlocks: {', '.join(spec['unlocks'])}")
            else:
                st.warning(f"**{canonical}** missing - add it to unlock: {', '.join(spec['unlocks'])}")

    st.divider()
    st.subheader("Default KPIs")
    st.caption(f"Source rows: {total_rows}  |  Included: {kpis['totalInvoices']}  |  "
               f"Excluded (cancelled): {kpis['_excludedCancelled']}  |  "
               f"Excluded (out of scope): {kpis['_excludedOutOfScope']}")

    kpi_labels = {key: label for key, label, _ in DEFAULT_KPI_DEFS}
    selected_kpis = st.multiselect(
        "Show these default KPI cards", list(kpi_labels.keys()),
        default=list(kpi_labels.keys()), format_func=lambda k: kpi_labels[k],
    )

    kpi_cols = st.columns(4)
    for i, (key, label, kind) in enumerate(DEFAULT_KPI_DEFS):
        if key not in selected_kpis:
            continue
        value = kpis[key]
        with kpi_cols[i % 4]:
            kpi_metric(label, fmt_value(value, kind), DEFAULT_KPI_FORMULAS.get(key, "-"), key=f"default_{key}")
    st.info(f"Avg. Payment Duration - NOT CHECKABLE: {reconcile.NOT_CHECKABLE['avgPaymentDuration']}")

    st.divider()
    st.subheader("Recommended KPIs")
    st.caption("Data-driven suggestions from the feature-engineering pass below - scored for real spread/split "
               "on THIS export, not a fixed list. Heuristic, not a model - use judgement.")
    st.caption("Note: the spread score isn't anomaly-filtered - a single extreme row (see Anomaly Detection) "
               "can inflate a numeric feature's coefficient of variation and push it up this list for the "
               "wrong reason. Check the Anomaly Detection tab before trusting a numeric recommendation's scale.")
    enriched_all, computed_all, blocked_all = engineer_features.engineer_all(df)
    recs = recommend_kpis(enriched_all, computed_all)
    if not recs:
        st.write("No standout candidates on this export.")
    else:
        rec_cols = st.columns(len(recs))
        for i, rec in enumerate(recs):
            with rec_cols[i]:
                display_val = f"{rec['value']:.1f}%" if rec["kind"] == "bool" else f"{rec['value']:,.2f}"
                feature_desc = engineer_features.FEATURE_REGISTRY.get(rec["name"], {}).get(
                    "description", "No description available for this engineered feature."
                )
                formula_text = f"{rec['agg']} of engineered feature '{rec['name']}':\n{feature_desc}"
                kpi_metric(rec["name"], display_val, formula_text, key=f"rec_{rec['name']}")
                st.caption(rec["rationale"])
                if st.button("Add to dashboard", key=f"add_rec_{rec['name']}"):
                    st.session_state.custom_kpis.append({
                        "name": rec["name"], "expr": rec["name"], "agg": rec["agg"],
                        "filter": None, "format": "pct" if rec["kind"] == "bool" else "float2",
                        "source": "recommended", "_enriched_col": rec["name"],
                    })
                    st.rerun()

    st.divider()
    st.subheader("Custom KPIs")
    if st.session_state.kpi_feedback:
        kind, msg = st.session_state.kpi_feedback
        getattr(st, kind)(msg)
        st.session_state.kpi_feedback = None
    if st.session_state.custom_kpis:
        custom_cols = st.columns(4)
        for i, cfg in enumerate(list(st.session_state.custom_kpis)):
            with custom_cols[i % 4]:
                formula_text = f"{cfg['agg']} of `{cfg['expr']}`"
                if cfg.get("filter"):
                    formula_text += f"\nwhere {cfg['filter']}"
                if cfg.get("source") == "recommended":
                    formula_text += "\n\n(added from a Recommended KPI - see the Feature Engineering tab for how this feature is computed)"
                try:
                    source_df = enriched_all if cfg.get("_enriched_col") else df
                    value = compute_custom_kpi(source_df, cfg["expr"], cfg["agg"], cfg.get("filter"))
                    kpi_metric(cfg["name"], fmt_value(value, cfg["format"]), formula_text, key=f"custom_{i}_{cfg['name']}")
                except Exception as e:
                    # Only reachable if a *previously valid* KPI stops working
                    # because a new file was uploaded without that column -
                    # new KPIs are validated before they're ever added, so
                    # this shouldn't happen for anything created just now.
                    st.metric(cfg["name"], "N/A")
                    st.caption(f"Not available in this file: {e}")
                if st.button("Remove", key=f"remove_{i}_{cfg['name']}"):
                    st.session_state.custom_kpis.pop(i)
                    st.rerun()
    else:
        st.write("No custom KPIs yet - add one below, or add a recommendation above.")

    with st.expander("Add a new KPI"):
        st.caption("Pick a column, or write a formula referencing column names (spaces/punctuation are "
                    "handled automatically - no need to type backticks). Example formula: "
                    "`Total incl. VAT - Payment Received`. Optional filter, e.g. `Status Reason == 'Paid'`. "
                    "A misspelled or slightly-off column name is matched to the closest real column "
                    "automatically where possible; if nothing added shows up, check the error message above "
                    "the cards - that's the actual reason it wasn't created.")
        new_name = st.text_input("Name", key="new_kpi_name")
        new_expr = st.text_input("Column or formula", key="new_kpi_expr")
        new_agg = st.selectbox("Aggregation", list(AGG_FUNCS.keys()), key="new_kpi_agg")
        new_filter = st.text_input("Filter (optional)", key="new_kpi_filter")
        new_format = st.selectbox("Display format", ["currency", "float2", "int", "pct"], key="new_kpi_format")
        if st.button("Add KPI"):
            if not new_name or not new_expr:
                st.error("Name and column/formula are required.")
            else:
                ok, resolved_expr, resolved_filter, corrections, err = validate_custom_kpi(
                    df, new_expr, new_agg, new_filter or None
                )
                if not ok:
                    st.error(f"Couldn't add '{new_name}': {err}")
                else:
                    if corrections:
                        note = "; ".join(f"'{typed}' matched to '{fixed}'" for typed, fixed in corrections)
                        st.session_state.kpi_feedback = ("info", f"'{new_name}' added - {note}.")
                    st.session_state.custom_kpis.append({
                        "name": new_name, "expr": resolved_expr, "agg": new_agg,
                        "filter": resolved_filter, "format": new_format, "source": "manual",
                    })
                    st.rerun()

    st.divider()
    st.subheader("Invoices by month")
    months = [m["month"] for m in monthly]
    fig = go.Figure()
    fig.add_bar(x=months, y=[m["billed"] for m in monthly], name="Billed (incl VAT)", marker_color=SERIES[0])
    fig.add_bar(x=months, y=[m["received"] for m in monthly], name="Received", marker_color=SERIES[2])
    fig.update_layout(barmode="group", height=360, margin=dict(t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Invoice breakdown by type")
        fig2 = go.Figure(go.Pie(labels=[s["label"] for s in donut["segments"]],
                                 values=[s["value"] for s in donut["segments"]],
                                 hole=0.6, marker_colors=SERIES))
        fig2.update_layout(height=320, margin=dict(t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)
    with c2:
        st.subheader("Invoices by status (cancelled included)")
        labels = [s["label"] for s in status["segments"]]
        values = [s["value"] for s in status["segments"]]
        colors = [CRITICAL if l == "Cancelled" else SERIES[0] for l in labels]
        fig3 = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker_color=colors))
        fig3.update_layout(xaxis_type="log", height=320, margin=dict(t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True)

# ---------------------------------------------------------------------------
# Tab 2: Anomaly Detection
# ---------------------------------------------------------------------------

with tab_anomaly:
    st.subheader("Anomaly detection")
    st.caption("Seven independent detectors - a false positive in one doesn't hide a real finding in another.")

    extreme_flags = reconcile.scan_outliers(rows, multiple=extreme_multiple)
    arithmetic_findings = reconcile.check_arithmetic_integrity(df, vat_rate, vat_tol)
    referential_findings = reconcile.check_referential(df)
    date_findings = reconcile.check_date_logic(df)
    iqr_findings = reconcile.check_statistical_outliers_iqr(df, k=iqr_k)
    ts_findings = reconcile.check_monthly_time_series_anomaly(monthly)
    rare_findings = reconcile.check_rare_categories(df, threshold_pct=rare_pct)

    sections = [
        ("Extreme value scan", extreme_flags),
        ("Arithmetic integrity", arithmetic_findings),
        ("Referential", referential_findings),
        ("Date logic", date_findings),
        ("Statistical (IQR) outliers", iqr_findings),
        ("Time-series (month-over-month)", ts_findings),
        ("Rare category values", rare_findings),
    ]

    total_findings = sum(len(f) for _, f in sections)
    m1, m2 = st.columns(2)
    m1.metric("Total findings", total_findings)
    m2.metric("Detectors with findings", sum(1 for _, f in sections if f))

    fig = go.Figure(go.Bar(
        x=[title for title, _ in sections],
        y=[len(f) for _, f in sections],
        marker_color=CRITICAL,
    ))
    fig.update_layout(height=320, margin=dict(t=10, b=10), yaxis_title="Findings")
    st.plotly_chart(fig, use_container_width=True)

    for title, findings in sections:
        with st.expander(f"{title} ({len(findings)})", expanded=bool(findings) and len(findings) <= 10):
            if not findings:
                st.write("None.")
                continue
            if title == "Extreme value scan":
                st.dataframe(pd.DataFrame(findings), use_container_width=True, hide_index=True)
            else:
                st.dataframe(pd.DataFrame(findings), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Tab 3: Field Usage
# ---------------------------------------------------------------------------

with tab_fields:
    st.subheader("Field-usage classification")
    classification = classify_fields.classify_all(df)
    duplicates = classify_fields.find_possible_duplicates(df)

    counts = classification["classification"].value_counts().reindex(classify_fields.CLASSIFICATION_ORDER).fillna(0)
    fig = go.Figure(go.Bar(x=counts.index.tolist(), y=counts.values.tolist(),
                            marker_color=[CRITICAL, "#ec835a", "#fab219", "#eda100", "#9a988f", GOOD]))
    fig.update_layout(height=320, margin=dict(t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

    label_filter = st.multiselect("Filter by classification", classify_fields.CLASSIFICATION_ORDER,
                                   default=classify_fields.CLASSIFICATION_ORDER)
    st.dataframe(classification[classification["classification"].isin(label_filter)],
                 use_container_width=True, hide_index=True)

    st.subheader(f"Possible duplicate/overlapping field pairs ({len(duplicates)})")
    if duplicates:
        st.dataframe(pd.DataFrame(duplicates), use_container_width=True, hide_index=True)
    else:
        st.write("None above threshold.")

# ---------------------------------------------------------------------------
# Tab 4: Feature Engineering
# ---------------------------------------------------------------------------

with tab_features:
    st.subheader("Engineered features")
    enriched, computed, blocked = engineer_features.engineer_all(df)

    m1, m2 = st.columns(2)
    m1.metric("Computed", f"{len(computed)} / {len(engineer_features.FEATURE_REGISTRY)}")
    m2.metric("Blocked (missing column)", len(blocked))

    rows_out = []
    for name in computed:
        rows_out.append({"feature": name, "status": "Computed",
                          "evidence": engineer_features.summarize_feature(enriched, name),
                          "description": engineer_features.FEATURE_REGISTRY[name]["description"]})
    for name, missing_required, missing_any_of, description in blocked:
        need = []
        if missing_required:
            need.append(f"requires: {missing_required}")
        if missing_any_of:
            need.append(f"requires one of: {missing_any_of}")
        rows_out.append({"feature": name, "status": "Blocked", "evidence": "; ".join(need),
                          "description": description})
    st.dataframe(pd.DataFrame(rows_out), use_container_width=True, hide_index=True)

    st.download_button(
        "Download enriched CSV (source columns + engineered features)",
        enriched.to_csv(index=False).encode("utf-8"),
        file_name="enriched_export.csv",
        mime="text/csv",
    )
