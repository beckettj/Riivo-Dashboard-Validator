# Invoice Tools - Interactive UI

A single browser interface wrapping the three CLI tools (`reconcile.py`,
`classify_fields.py`, `engineer_features.py`) so a PM can drag in an export
and see everything those scripts produce, without a terminal.

**This does not reimplement any logic.** `app.py` imports the three scripts
as Python modules and calls their functions directly - every number here is
guaranteed to match what the CLI tools would print for the same export,
because it's the same code running underneath. The three `.py` files in
this folder are synced copies of the canonical CLI tools; if you change the
underlying logic, change it in the CLI tool's own folder and re-copy the
file here (there's no packaging step - `import reconcile` just needs
`reconcile.py` sitting next to `app.py`).

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens in your browser at `http://localhost:8501`. Drag in an invoice
export (the same "My Team's Open Invoices" shape `reconcile.py` expects, or
any similarly-shaped CRM export) in the sidebar.

## What's in each tab

- **KPI Dashboard** - the field-coverage report, the 12 default KPI cards
  (pick which ones to show), a Recommended KPIs row, a Custom KPIs builder,
  and the three reconciliation charts (monthly, type donut, status bar).
- **Anomaly Detection** - all seven detectors from `reconcile.py`, with live
  sliders in the sidebar for VAT tolerance, IQR sensitivity, rare-category
  threshold, and the extreme-value multiple - move a slider and every
  detector re-runs against the same export.
- **Field Usage** - the field-usage classification table and duplicate-pair
  pass from `classify_fields.py`, filterable by classification.
- **Feature Engineering** - the computed/blocked feature table from
  `engineer_features.py`, plus a button to download the enriched export
  (source columns + every computed feature) as a CSV.

## Default KPIs, Recommended KPIs, Custom KPIs - what's the difference

This was the point of building the UI in the first place: a fixed set of
defaults, a place to add your own, and data-driven suggestions in between.

- **Default KPIs** are the same 12 cards `reconcile.py` always computes
  (mirroring the dashboard). The multiselect just controls which are shown
  here - it doesn't change what's computed.
- **Recommended KPIs** are generated fresh from whatever the feature-
  engineering pass computed on *this* export. Each computed feature is
  scored for how "interesting" it looks - a boolean feature scores highest
  when it splits the data into a real minority/majority (not
  almost-always-true or almost-always-false); a numeric feature scores
  highest when it has real spread (coefficient of variation) rather than
  being nearly constant. This is a heuristic, not a model - it's meant to
  surface a short list worth a human look, not a verdict. **It's also not
  anomaly-filtered**: a single extreme row can inflate a numeric feature's
  spread score and push it up the list for the wrong reason - check the
  Anomaly Detection tab before trusting a numeric recommendation's scale
  (this is exactly what happens with `amount_per_lineitem_ratio` on the
  bundled test export - see Known limitations).
- **Custom KPIs** are yours: pick a column, or write a formula referencing
  column names (e.g. `Total incl. VAT - Payment Received`), choose an
  aggregation (Sum/Mean/Median/Min/Max/Count/Count Distinct/% True), and
  optionally a filter (e.g. `Status Reason == 'Paid'`). Column names with
  spaces or punctuation don't need manual backtick-quoting - the app adds
  them for you. A recommendation's "Add to dashboard" button just appends
  one of these under the hood, so it shows up in the same section and can
  be removed the same way.

## Formula safety

Custom KPI/filter expressions go through pandas' own `eval`/`query`
(restricted to referencing columns and arithmetic/comparison/boolean
operations - not a general Python `eval`), plus a denylist on top for
defense in depth (blocks `import`, `exec`, `__`, `os.`, `subprocess`,
`lambda`, and similar tokens before the expression is ever parsed). A
blocked or invalid expression shows an inline error on that KPI's card
rather than crashing the app - tested against `__import__("os").system(...)`
during development, which the denylist catches before it reaches pandas.

## Known limitations

- Every tab's computation runs on every interaction (Streamlit re-executes
  the whole script top-to-bottom on any widget change), regardless of which
  tab is open. Fine up to at least tens of thousands of rows on the export
  sizes this was built for; if you're loading something much larger,
  cache the per-tab computations with `st.cache_data` keyed on the file
  bytes + settings.
- The Recommended KPIs' "coefficient of variation" scoring is naive about
  outliers - see the caveat in the UI itself and above. `days_overdue` and
  `amount_per_lineitem_ratio` both surfaced as recommendations on the
  bundled 1,200-row test export partly *because* of the seeded anomaly row,
  not because they're organically high-spread - worth cross-checking any
  numeric recommendation against the Anomaly Detection tab's findings for
  the same columns.
- The custom-KPI formula box is a thin layer over `df.eval`/`df.query` -
  it inherits pandas' own quirks (e.g. `and`/`or` don't work in filters,
  use `&`/`|` with parenthesized comparisons instead: `` (`Status Reason` == 'Paid') & (`Total incl. VAT` > 1000) ``).
- No authentication, no persistence - custom KPIs live in
  `st.session_state` for the current browser session only. Refreshing the
  page or restarting the app clears them.
