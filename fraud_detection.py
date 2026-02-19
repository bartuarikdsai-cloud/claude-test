"""
Auto Insurance Fraud / Anomaly Detection Script
================================================
Analyzes auto insurance claims data using a score-based detection system.
Outputs flagged claims to CSV and generates an HTML summary report.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import html as html_module

# ── Configuration ──────────────────────────────────────────────────────────
DATA_FILE = Path(__file__).parent / "auto_insurance_data.csv"
OUTPUT_CSV = Path(__file__).parent / "flagged_claims.csv"
OUTPUT_HTML = Path(__file__).parent / "fraud_detection_summary.html"

RULE_DEFINITIONS = {
    "extreme_loss_ratio": {
        "label": "Extreme Loss Ratio (>15x)",
        "score": 3,
    },
    "statistical_outlier": {
        "label": "Statistical Outlier (>mean+3*std by age group)",
        "score": 2,
    },
    "new_car_high_loss": {
        "label": "New Car (>=2022), High Loss (>$10k)",
        "score": 2,
    },
    "young_driver_extreme": {
        "label": "Young Driver (<25), Extreme Claim (>$15k)",
        "score": 2,
    },
    "premium_loss_mismatch": {
        "label": "Premium-Loss Mismatch (top 5% loss, bottom 25% premium)",
        "score": 1,
    },
}


# ── Data Loading ───────────────────────────────────────────────────────────
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    print(f"Loaded {len(df):,} records from {path.name}")
    return df


# ── Detection Rules ────────────────────────────────────────────────────────
def apply_rules(df: pd.DataFrame) -> pd.DataFrame:
    """Apply all fraud-detection rules to customers with claims (total_loss > 0)."""

    claims = df[df["total_loss"] > 0].copy()
    print(f"Customers with claims (total_loss > 0): {len(claims):,}\n")

    # Initialise tracking columns
    claims["risk_score"] = 0
    claims["flags"] = ""

    flag_counts = {}

    # ── Rule 1: Extreme Loss Ratio (score +3) ─────────────────────────────
    rule = "extreme_loss_ratio"
    mask = claims["loss_ratio"] > 15
    claims.loc[mask, "risk_score"] += RULE_DEFINITIONS[rule]["score"]
    claims.loc[mask, "flags"] += rule + ","
    flag_counts[rule] = int(mask.sum())

    # ── Rule 2: Statistical Outlier by age group (score +2) ────────────────
    rule = "statistical_outlier"
    # Define age groups
    bins = [0, 25, 35, 45, 55, 65, 200]
    labels = ["<25", "25-34", "35-44", "45-54", "55-64", "65+"]
    claims["age_group"] = pd.cut(claims["age"], bins=bins, labels=labels, right=False)

    # Compute mean + 3*std per age group from claiming population
    group_stats = claims.groupby("age_group", observed=True)["total_loss"].agg(["mean", "std"])
    group_stats["threshold"] = group_stats["mean"] + 3 * group_stats["std"]

    # Map threshold back (convert to float to avoid Categorical comparison issues)
    claims["outlier_threshold"] = claims["age_group"].map(group_stats["threshold"]).astype(float)
    mask = claims["total_loss"] > claims["outlier_threshold"]
    claims.loc[mask, "risk_score"] += RULE_DEFINITIONS[rule]["score"]
    claims.loc[mask, "flags"] += rule + ","
    flag_counts[rule] = int(mask.sum())

    # ── Rule 3: New Car, High Loss (score +2) ─────────────────────────────
    rule = "new_car_high_loss"
    mask = (claims["car_model_year"] >= 2022) & (claims["total_loss"] > 10_000)
    claims.loc[mask, "risk_score"] += RULE_DEFINITIONS[rule]["score"]
    claims.loc[mask, "flags"] += rule + ","
    flag_counts[rule] = int(mask.sum())

    # ── Rule 4: Young Driver, Extreme Claim (score +2) ────────────────────
    rule = "young_driver_extreme"
    mask = (claims["age"] < 25) & (claims["total_loss"] > 15_000)
    claims.loc[mask, "risk_score"] += RULE_DEFINITIONS[rule]["score"]
    claims.loc[mask, "flags"] += rule + ","
    flag_counts[rule] = int(mask.sum())

    # ── Rule 5: Premium-Loss Mismatch (score +1) ──────────────────────────
    rule = "premium_loss_mismatch"
    loss_95 = claims["total_loss"].quantile(0.95)
    premium_25 = claims["annual_premium"].quantile(0.25)
    mask = (claims["total_loss"] >= loss_95) & (claims["annual_premium"] <= premium_25)
    claims.loc[mask, "risk_score"] += RULE_DEFINITIONS[rule]["score"]
    claims.loc[mask, "flags"] += rule + ","
    flag_counts[rule] = int(mask.sum())

    # Clean trailing commas from flags
    claims["flags"] = claims["flags"].str.rstrip(",")

    # Drop helper columns
    claims.drop(columns=["age_group", "outlier_threshold"], inplace=True)

    return claims, flag_counts, {
        "loss_95_threshold": loss_95,
        "premium_25_threshold": premium_25,
        "age_group_thresholds": group_stats["threshold"].to_dict(),
    }


# ── Console Summary ───────────────────────────────────────────────────────
def print_summary(total_records: int, claims: pd.DataFrame, flag_counts: dict, thresholds: dict):
    flagged = claims[claims["risk_score"] > 0]
    score_dist = flagged["risk_score"].value_counts().sort_index()

    print("=" * 62)
    print("  AUTO INSURANCE FRAUD / ANOMALY DETECTION — SUMMARY")
    print("=" * 62)
    print(f"  Total records in dataset      : {total_records:,}")
    print(f"  Total claims analysed          : {len(claims):,}")
    print(f"  Unique customers flagged       : {len(flagged):,}")
    print(f"  Flag rate                      : {len(flagged)/len(claims)*100:.1f}%")
    print("-" * 62)
    print("  FLAGS BY RULE")
    print("-" * 62)
    for key, count in flag_counts.items():
        label = RULE_DEFINITIONS[key]["label"]
        print(f"    {label:<52} {count:>5}")
    print("-" * 62)
    print("  RISK SCORE DISTRIBUTION (flagged customers only)")
    print("-" * 62)
    for score, cnt in score_dist.items():
        bar = "#" * min(cnt, 60)
        print(f"    Score {score:>2}: {cnt:>5}  {bar}")
    print("-" * 62)
    print(f"  Computed thresholds:")
    print(f"    Loss top-5% cutoff           : ${thresholds['loss_95_threshold']:,.2f}")
    print(f"    Premium bottom-25% cutoff    : ${thresholds['premium_25_threshold']:,.2f}")
    for grp, thr in thresholds["age_group_thresholds"].items():
        print(f"    Age group {str(grp):<8} outlier   : ${thr:,.2f}")
    print("=" * 62)


# ── CSV Output ─────────────────────────────────────────────────────────────
def save_flagged_csv(claims: pd.DataFrame, path: Path):
    flagged = claims[claims["risk_score"] > 0].sort_values("risk_score", ascending=False)
    cols = [
        "customer_id", "gender", "age", "car_model_year",
        "annual_premium", "total_loss", "loss_ratio", "risk_score", "flags",
    ]
    flagged[cols].to_csv(path, index=False)
    print(f"\nSaved {len(flagged):,} flagged claims to {path.name}")


# ── HTML Report ────────────────────────────────────────────────────────────
def generate_html_report(
    total_records: int,
    claims: pd.DataFrame,
    flag_counts: dict,
    thresholds: dict,
    path: Path,
):
    flagged = claims[claims["risk_score"] > 0].sort_values("risk_score", ascending=False)
    top30 = flagged.head(30)
    score_dist = flagged["risk_score"].value_counts().sort_index()

    # Prepare chart data
    rule_labels = [RULE_DEFINITIONS[k]["label"] for k in flag_counts]
    rule_values = list(flag_counts.values())

    score_labels = [f"Score {s}" for s in score_dist.index]
    score_values = score_dist.values.tolist()

    # Build table rows
    def risk_color(score):
        if score >= 7:
            return "#EF4444"   # red
        elif score >= 5:
            return "#F59E0B"   # amber
        else:
            return "#EAB308"   # yellow

    table_rows = ""
    for _, row in top30.iterrows():
        color = risk_color(row["risk_score"])
        flags_display = ", ".join(
            RULE_DEFINITIONS[f]["label"] for f in row["flags"].split(",") if f in RULE_DEFINITIONS
        )
        table_rows += f"""
        <tr>
          <td>{int(row['customer_id'])}</td>
          <td>{html_module.escape(str(row['gender']))}</td>
          <td>{int(row['age'])}</td>
          <td>{int(row['car_model_year'])}</td>
          <td>${row['annual_premium']:,.2f}</td>
          <td>${row['total_loss']:,.2f}</td>
          <td>{row['loss_ratio']:.2f}</td>
          <td style="color:{color}; font-weight:700;">{int(row['risk_score'])}</td>
          <td class="flags-cell">{html_module.escape(flags_display)}</td>
        </tr>"""

    # Age group thresholds for display
    age_thresh_rows = ""
    for grp, thr in thresholds["age_group_thresholds"].items():
        age_thresh_rows += f"<tr><td>{grp}</td><td>${thr:,.2f}</td></tr>"

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fraud Detection Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg: #0F1117;
    --surface: #1A1D27;
    --surface2: #252833;
    --border: #2D3140;
    --text: #E2E8F0;
    --text-dim: #94A3B8;
    --accent: #6366F1;
    --red: #EF4444;
    --amber: #F59E0B;
    --yellow: #EAB308;
    --green: #22C55E;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: var(--bg);
    color: var(--text);
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    line-height: 1.6;
    padding: 2rem;
  }}
  .container {{ max-width: 1400px; margin: 0 auto; }}

  /* Header */
  .header {{
    text-align: center;
    margin-bottom: 2.5rem;
    padding-bottom: 1.5rem;
    border-bottom: 1px solid var(--border);
  }}
  .header h1 {{
    font-size: 2rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    margin-bottom: 0.3rem;
  }}
  .header h1 span {{ color: var(--red); }}
  .header p {{ color: var(--text-dim); font-size: 0.95rem; }}

  /* KPI cards */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1rem;
    margin-bottom: 2rem;
  }}
  .kpi {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    text-align: center;
  }}
  .kpi .value {{
    font-size: 1.75rem;
    font-weight: 700;
    color: var(--accent);
  }}
  .kpi .value.red {{ color: var(--red); }}
  .kpi .value.amber {{ color: var(--amber); }}
  .kpi .label {{
    font-size: 0.82rem;
    color: var(--text-dim);
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-top: 0.2rem;
  }}

  /* Charts */
  .charts-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1.5rem;
    margin-bottom: 2rem;
  }}
  @media (max-width: 900px) {{
    .charts-grid {{ grid-template-columns: 1fr; }}
  }}
  .chart-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
  }}
  .chart-card h3 {{
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 1rem;
    color: var(--text-dim);
  }}

  /* Thresholds section */
  .thresholds {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 2rem;
  }}
  .thresholds h3 {{
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 1rem;
    color: var(--text-dim);
  }}
  .thresh-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 1rem;
  }}
  .thresh-item {{
    background: var(--surface2);
    border-radius: 8px;
    padding: 0.75rem 1rem;
  }}
  .thresh-item .tl {{ font-size: 0.8rem; color: var(--text-dim); }}
  .thresh-item .tv {{ font-size: 1.1rem; font-weight: 600; }}

  /* Table */
  .table-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    overflow-x: auto;
  }}
  .table-card h3 {{
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 1rem;
    color: var(--text-dim);
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
  }}
  th {{
    text-align: left;
    padding: 0.6rem 0.75rem;
    border-bottom: 2px solid var(--border);
    color: var(--text-dim);
    font-weight: 600;
    text-transform: uppercase;
    font-size: 0.75rem;
    letter-spacing: 0.04em;
    white-space: nowrap;
  }}
  td {{
    padding: 0.55rem 0.75rem;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }}
  tr:hover td {{ background: var(--surface2); }}
  .flags-cell {{
    white-space: normal;
    max-width: 350px;
    font-size: 0.78rem;
    color: var(--text-dim);
  }}

  /* Footer */
  .footer {{
    text-align: center;
    margin-top: 2.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--border);
    color: var(--text-dim);
    font-size: 0.82rem;
  }}
</style>
</head>
<body>
<div class="container">

  <!-- Header -->
  <div class="header">
    <h1><span>Fraud</span> / Anomaly Detection Report</h1>
    <p>Auto Insurance Claims Analysis &mdash; {len(claims):,} claims evaluated from {total_records:,} records</p>
  </div>

  <!-- KPI Cards -->
  <div class="kpi-grid">
    <div class="kpi">
      <div class="value">{total_records:,}</div>
      <div class="label">Total Records</div>
    </div>
    <div class="kpi">
      <div class="value">{len(claims):,}</div>
      <div class="label">Claims Analysed</div>
    </div>
    <div class="kpi">
      <div class="value red">{len(flagged):,}</div>
      <div class="label">Customers Flagged</div>
    </div>
    <div class="kpi">
      <div class="value amber">{len(flagged)/len(claims)*100:.1f}%</div>
      <div class="label">Flag Rate</div>
    </div>
    <div class="kpi">
      <div class="value">{flagged['risk_score'].max() if len(flagged) > 0 else 0}</div>
      <div class="label">Highest Risk Score</div>
    </div>
  </div>

  <!-- Charts -->
  <div class="charts-grid">
    <div class="chart-card">
      <h3>Flags by Detection Rule</h3>
      <canvas id="ruleChart"></canvas>
    </div>
    <div class="chart-card">
      <h3>Risk Score Distribution</h3>
      <canvas id="scoreChart"></canvas>
    </div>
  </div>

  <!-- Computed Thresholds -->
  <div class="thresholds">
    <h3>Computed Thresholds (derived from data)</h3>
    <div class="thresh-grid">
      <div class="thresh-item">
        <div class="tl">Loss Ratio Extreme (&gt;15x)</div>
        <div class="tv">15.0x</div>
      </div>
      <div class="thresh-item">
        <div class="tl">Top 5% Loss Cutoff</div>
        <div class="tv">${thresholds['loss_95_threshold']:,.2f}</div>
      </div>
      <div class="thresh-item">
        <div class="tl">Bottom 25% Premium Cutoff</div>
        <div class="tv">${thresholds['premium_25_threshold']:,.2f}</div>
      </div>
      {age_thresh_rows and f'''
      </div>
      <h3 style="margin-top:1rem; font-size:0.9rem; color:var(--text-dim);">
        Statistical Outlier Thresholds by Age Group (mean + 3&sigma;)
      </h3>
      <div class="thresh-grid">
      ''' + ''.join(
          f'<div class="thresh-item"><div class="tl">Age {grp}</div><div class="tv">${thr:,.2f}</div></div>'
          for grp, thr in thresholds["age_group_thresholds"].items()
      ) + '''
      '''}
    </div>
  </div>

  <!-- Top 30 Flagged Claims Table -->
  <div class="table-card">
    <h3>Top 30 Flagged Claims (sorted by risk score)</h3>
    <table>
      <thead>
        <tr>
          <th>Customer ID</th>
          <th>Gender</th>
          <th>Age</th>
          <th>Car Year</th>
          <th>Premium</th>
          <th>Total Loss</th>
          <th>Loss Ratio</th>
          <th>Risk Score</th>
          <th>Triggered Rules</th>
        </tr>
      </thead>
      <tbody>
        {table_rows}
      </tbody>
    </table>
  </div>

  <div class="footer">
    Generated by fraud_detection.py &mdash; Score-based anomaly detection system
  </div>

</div>

<script>
  Chart.defaults.color = '#94A3B8';
  Chart.defaults.borderColor = '#2D3140';

  // Rule chart
  new Chart(document.getElementById('ruleChart'), {{
    type: 'bar',
    data: {{
      labels: {rule_labels},
      datasets: [{{
        data: {rule_values},
        backgroundColor: ['#EF4444','#F59E0B','#6366F1','#22C55E','#8B5CF6'],
        borderRadius: 6,
        maxBarThickness: 48,
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: ctx => ctx.raw + ' claims'
          }}
        }}
      }},
      scales: {{
        x: {{ grid: {{ color: '#2D314044' }} }},
        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 11 }} }} }}
      }}
    }}
  }});

  // Score distribution chart
  new Chart(document.getElementById('scoreChart'), {{
    type: 'bar',
    data: {{
      labels: {score_labels},
      datasets: [{{
        data: {score_values},
        backgroundColor: {score_labels}.map((_, i) => {{
          const scores = {list(score_dist.index)};
          const s = scores[i];
          if (s >= 7) return '#EF4444';
          if (s >= 5) return '#F59E0B';
          return '#EAB308';
        }}),
        borderRadius: 6,
        maxBarThickness: 48,
      }}]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: ctx => ctx.raw + ' customers'
          }}
        }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }} }},
        y: {{ grid: {{ color: '#2D314044' }}, beginAtZero: true }}
      }}
    }}
  }});
</script>
</body>
</html>"""

    path.write_text(html_content, encoding="utf-8")
    print(f"Saved HTML report to {path.name}")


# ── Main ───────────────────────────────────────────────────────────────────
def main():
    df = load_data(DATA_FILE)
    claims, flag_counts, thresholds = apply_rules(df)
    print_summary(len(df), claims, flag_counts, thresholds)
    save_flagged_csv(claims, OUTPUT_CSV)
    generate_html_report(len(df), claims, flag_counts, thresholds, OUTPUT_HTML)
    print("\nDone.")


if __name__ == "__main__":
    main()
