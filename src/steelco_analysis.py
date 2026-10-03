"""
SteelCo Business Analysis
=========================
Data cleaning, KPI reconciliation and visualisation for the SteelCo Ltd
Q3-Q4 2022 sales data (``SteelCo_Data_Analysis.xlsx``, sheet ``Taul1``).

The Excel dashboard was built with pivot tables; this script reproduces the
same KPIs in Python, documents every cleaning decision, and adds extra
analysis (margin *rate*, quarter-on-quarter change, price per kg, Pareto).

Usage
-----
    python src/steelco_analysis.py --data data/SteelCo_Data_Analysis.xlsx --out outputs

Outputs (written to ``--out``)
    steelco_clean.csv      cleaned transaction-level data
    cleaning_audit.csv     what each cleaning step changed
    figures/*.png          charts used in the notebook / report
"""
from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
RAW_SHEET = "Taul1"

# The raw headers are long and contain typos ("Quater"); rename by position.
COLUMN_NAMES = [
    "product_group",
    "steel_grade",
    "region",
    "country",
    "quarter",
    "volume_kg",
    "revenue_eur",
    "gross_margin_eur",
]

# Shorten SAP-style abbreviations so country labels are readable in charts.
COUNTRY_FIXES = {
    "Russian Fed.": "Russia",
    "Utd.Arab Emir.": "United Arab Emirates",
    "Amer.Virgin Is.": "US Virgin Islands",
    "Bosnia-Herz.": "Bosnia and Herzegovina",
    "Dem. Rep. Congo": "DR Congo",
    "Rep. Congo": "Congo",
    "Dominican Rep.": "Dominican Republic",
    "Trinidad,Tobago": "Trinidad and Tobago",
    "St. Lucia": "Saint Lucia",
    "Macedonia": "North Macedonia",
    "USA": "United States",
}

# Headline numbers shown on the Excel dashboard (for reconciliation).
DASHBOARD = {
    "revenue_eur": 1105.4e6,
    "volume_kg": 4345.3e6,
    "gross_margin_eur": 403.9e6,
}

PRODUCT_ORDER = ["CORE", "THERMA", "MODA", "DURA", "SUPRA"]
sns.set_theme(style="whitegrid", context="notebook")
PALETTE = sns.color_palette("viridis", 5)


# --------------------------------------------------------------------------
# Load
# --------------------------------------------------------------------------
def load_raw(path: str | Path, sheet: str = RAW_SHEET) -> pd.DataFrame:
    """Read the raw transaction sheet and give columns clean snake_case names."""
    df = pd.read_excel(path, sheet_name=sheet)
    if df.shape[1] != len(COLUMN_NAMES):
        raise ValueError(f"Expected {len(COLUMN_NAMES)} columns, found {df.shape[1]}")
    df.columns = COLUMN_NAMES
    return df


# --------------------------------------------------------------------------
# Clean
# --------------------------------------------------------------------------
def clean_data(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Clean the data and return ``(clean_df, audit_table)``.

    Decisions (each is logged in the audit table):
      1. Standardise text (trim spaces, upper-case product group, fix country labels).
      2. Drop the single 'Not assigned' placeholder row (zero revenue, no market).
      3. Drop rows with negative volume (physically impossible).
      4. KEEP negative revenue / margin rows: they are credit notes, returns or
         genuine losses and belong in the P&L. They are flagged, not removed.
      5. Flag (not drop) exact duplicate rows: there is no transaction ID, so
         identical rows can still be separate shipments.
    """
    audit = []

    def log(step: str, before: pd.DataFrame, after: pd.DataFrame, note: str = "") -> None:
        audit.append(
            {
                "step": step,
                "rows_affected": len(before) - len(after),
                "rows_remaining": len(after),
                "revenue_removed_eur": before["revenue_eur"].sum() - after["revenue_eur"].sum(),
                "note": note,
            }
        )

    df = raw.copy()

    # 1. Standardise text ---------------------------------------------------
    for col in ["product_group", "steel_grade", "region", "country", "quarter"]:
        df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df["product_group"] = df["product_group"].str.upper()
    df["country"] = df["country"].replace(COUNTRY_FIXES)
    df["quarter"] = pd.Categorical(df["quarter"], categories=["Q3 2022", "Q4 2022"], ordered=True)
    audit.append(
        {"step": "Standardise text / country labels", "rows_affected": 0,
         "rows_remaining": len(df), "revenue_removed_eur": 0.0,
         "note": f"{len(COUNTRY_FIXES)} country abbreviations mapped to full names"}
    )

    # 2. Placeholder row ----------------------------------------------------
    nxt = df[(df["region"] != "Not assigned") & (df["country"] != "Not assigned")]
    log("Drop 'Not assigned' placeholder row", df, nxt,
        "No region/country; revenue 0 but a €47.5k margin entry that cannot be allocated")
    df = nxt

    # 3. Negative volume ----------------------------------------------------
    nxt = df[df["volume_kg"] >= 0]
    log("Drop negative-volume rows", df, nxt, "Volume cannot be negative (returns booked as -kg)")
    df = nxt

    # 4/5. Flags (rows kept) ------------------------------------------------
    df = df.assign(
        is_possible_duplicate=df.duplicated(keep="first"),
        is_loss_making=df["gross_margin_eur"] < 0,
        has_negative_revenue=df["revenue_eur"] < 0,
    )
    audit.append(
        {"step": "Flag exact duplicate rows (kept)", "rows_affected": 0,
         "rows_remaining": len(df), "revenue_removed_eur": 0.0,
         "note": f"{int(df['is_possible_duplicate'].sum())} repeated rows flagged; no transaction ID to confirm"}
    )
    audit.append(
        {"step": "Flag negative revenue / margin rows (kept)", "rows_affected": 0,
         "rows_remaining": len(df), "revenue_removed_eur": 0.0,
         "note": f"{int(df['has_negative_revenue'].sum())} negative-revenue rows, "
                 f"{int(df['is_loss_making'].sum())} loss-making rows retained"}
    )

    # Feature engineering ---------------------------------------------------
    positive_rev = df["revenue_eur"] > 0
    positive_vol = df["volume_kg"] > 0
    df["margin_pct"] = np.where(positive_rev, df["gross_margin_eur"] / df["revenue_eur"], np.nan)
    df["price_per_kg"] = np.where(positive_vol & positive_rev, df["revenue_eur"] / df["volume_kg"], np.nan)

    return df.reset_index(drop=True), pd.DataFrame(audit)


# --------------------------------------------------------------------------
# Analysis helpers
# --------------------------------------------------------------------------
def totals(df: pd.DataFrame) -> pd.Series:
    """Headline KPIs for any slice of the data."""
    rev, gm, vol = df["revenue_eur"].sum(), df["gross_margin_eur"].sum(), df["volume_kg"].sum()
    return pd.Series(
        {
            "revenue_eur": rev,
            "gross_margin_eur": gm,
            "volume_kg": vol,
            "margin_pct": gm / rev,
            "avg_price_eur_per_kg": rev / vol,
        }
    )


def reconcile(raw: pd.DataFrame, clean: pd.DataFrame) -> pd.DataFrame:
    """Compare raw and cleaned totals with the numbers on the Excel dashboard."""
    out = pd.DataFrame(
        {
            "excel_dashboard": pd.Series(DASHBOARD),
            "python_raw": totals(raw)[list(DASHBOARD)],
            "python_clean": totals(clean)[list(DASHBOARD)],
        }
    )
    out["clean_vs_raw_pct"] = (out["python_clean"] / out["python_raw"] - 1) * 100
    return out


def summarise(df: pd.DataFrame, by: str | list[str]) -> pd.DataFrame:
    """Revenue, volume, margin, margin rate and shares grouped by a column."""
    g = (
        df.groupby(by, observed=True)[["volume_kg", "revenue_eur", "gross_margin_eur"]]
        .sum()
        .assign(
            margin_pct=lambda d: d["gross_margin_eur"] / d["revenue_eur"],
            revenue_share=lambda d: d["revenue_eur"] / d["revenue_eur"].sum(),
            volume_share=lambda d: d["volume_kg"] / d["volume_kg"].sum(),
            margin_share=lambda d: d["gross_margin_eur"] / d["gross_margin_eur"].sum(),
        )
    )
    return g.sort_values("revenue_eur", ascending=False)


def quarter_on_quarter(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Q3 -> Q4 change in revenue by a dimension."""
    pivot = df.pivot_table(index=by, columns="quarter", values="revenue_eur", aggfunc="sum", observed=True)
    pivot["change_eur"] = pivot["Q4 2022"] - pivot["Q3 2022"]
    pivot["change_pct"] = pivot["change_eur"] / pivot["Q3 2022"]
    return pivot.sort_values("Q4 2022", ascending=False)


# --------------------------------------------------------------------------
# Charts (each returns a Figure so the notebook can display it)
# --------------------------------------------------------------------------
def _millions(x, _pos=None):
    return f"€{x / 1e6:,.0f}M"


def _label_bars(ax, fmt, horizontal=False, pad=3):
    for c in ax.containers:
        ax.bar_label(c, labels=[fmt(v) for v in c.datavalues], padding=pad, fontsize=9)


def plot_region_revenue_volume(df: pd.DataFrame) -> plt.Figure:
    s = summarise(df, "region")
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
    sns.barplot(x=s.index, y=s["revenue_eur"], hue=s.index, palette="viridis", legend=False, ax=axes[0])
    axes[0].set(title="Revenue by region", xlabel="", ylabel="")
    axes[0].yaxis.set_major_formatter(_millions)
    _label_bars(axes[0], lambda v: f"€{v / 1e6:,.0f}M")
    sns.barplot(x=s.index, y=s["volume_kg"] / 1e6, hue=s.index, palette="viridis", legend=False, ax=axes[1])
    axes[1].set(title="Volume by region (million kg)", xlabel="", ylabel="")
    _label_bars(axes[1], lambda v: f"{v:,.0f}")
    for ax in axes:
        ax.margins(y=0.12)
    fig.tight_layout()
    return fig


def plot_top_countries(df: pd.DataFrame, n: int = 10) -> plt.Figure:
    s = summarise(df, "country")
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    rev = s.nlargest(n, "revenue_eur")["revenue_eur"].iloc[::-1]
    vol = s.nlargest(n, "volume_kg")["volume_kg"].iloc[::-1] / 1e6
    axes[0].barh(rev.index, rev.values / 1e6, color=PALETTE[1])
    axes[0].set(title=f"Top {n} countries by revenue (€M)", xlabel="€ million")
    axes[1].barh(vol.index, vol.values, color=PALETTE[3])
    axes[1].set(title=f"Top {n} countries by volume (million kg)", xlabel="million kg")
    for ax, vals in zip(axes, [rev / 1e6, vol]):
        ax.bar_label(ax.containers[0], labels=[f"{v:,.0f}" for v in vals.values], padding=3, fontsize=9)
        ax.margins(x=0.12)
    fig.tight_layout()
    return fig


def plot_product_mix(df: pd.DataFrame) -> plt.Figure:
    """Revenue share, volume share and margin RATE per product group."""
    s = summarise(df, "product_group").reindex(PRODUCT_ORDER)
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.8))
    x = np.arange(len(s))
    w = 0.38
    axes[0].bar(x - w / 2, s["revenue_share"] * 100, w, label="Revenue share", color=PALETTE[1])
    axes[0].bar(x + w / 2, s["volume_share"] * 100, w, label="Volume share", color=PALETTE[3])
    axes[0].set_xticks(x, s.index)
    axes[0].set(title="Share of revenue vs share of volume", ylabel="% of total")
    axes[0].legend()
    bars = axes[1].bar(s.index, s["margin_pct"] * 100, color=PALETTE[2])
    axes[1].bar_label(bars, labels=[f"{v:.1f}%" for v in s["margin_pct"] * 100], padding=3)
    axes[1].axhline(totals(df)["margin_pct"] * 100, color="crimson", ls="--", lw=1.2, label="Company average")
    axes[1].set(title="Gross margin rate (margin / revenue)", ylabel="%")
    axes[1].legend()
    axes[1].margins(y=0.12)
    fig.tight_layout()
    return fig


def plot_region_product_heatmap(df: pd.DataFrame) -> plt.Figure:
    pivot = df.pivot_table(index="region", columns="product_group", values="revenue_eur", aggfunc="sum")
    pivot = pivot[PRODUCT_ORDER] / 1e6
    fig, ax = plt.subplots(figsize=(8, 4))
    sns.heatmap(pivot, annot=True, fmt=",.0f", cmap="YlGnBu", cbar_kws={"label": "Revenue (€M)"}, ax=ax)
    ax.set(title="Revenue (€M): region x product group", xlabel="", ylabel="")
    fig.tight_layout()
    return fig


def plot_quarter_comparison(df: pd.DataFrame) -> plt.Figure:
    q = quarter_on_quarter(df, "region")
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.5))
    q[["Q3 2022", "Q4 2022"]].div(1e6).plot.bar(ax=axes[0], color=[PALETTE[0], PALETTE[3]], rot=0)
    axes[0].set(title="Revenue by region and quarter (€M)", xlabel="", ylabel="€ million")
    pq = quarter_on_quarter(df, "product_group").reindex(PRODUCT_ORDER)
    colors = ["seagreen" if v >= 0 else "firebrick" for v in pq["change_pct"]]
    bars = axes[1].bar(pq.index, pq["change_pct"] * 100, color=colors)
    axes[1].bar_label(bars, labels=[f"{v:+.1f}%" for v in pq["change_pct"] * 100], padding=3)
    axes[1].axhline(0, color="black", lw=0.8)
    axes[1].set(title="Q3 to Q4 revenue change by product group", ylabel="%")
    axes[1].margins(y=0.15)
    fig.tight_layout()
    return fig


def plot_pareto_countries(df: pd.DataFrame, top: int = 20) -> plt.Figure:
    s = summarise(df, "country")
    cum = s["revenue_eur"].cumsum() / s["revenue_eur"].sum() * 100
    head = s.head(top)
    fig, ax = plt.subplots(figsize=(13, 4.8))
    ax.bar(head.index, head["revenue_eur"] / 1e6, color=PALETTE[1])
    ax.set(ylabel="Revenue (€M)", title=f"Pareto: top {top} countries and cumulative share of revenue")
    ax.tick_params(axis="x", rotation=60)
    for lbl in ax.get_xticklabels():
        lbl.set_ha("right")
    ax2 = ax.twinx()
    ax2.plot(head.index, cum.head(top), color="crimson", marker="o")
    ax2.set(ylabel="Cumulative % of revenue", ylim=(0, 105))
    ax2.grid(False)
    n80 = int((cum < 80).sum() + 1)
    ax2.axhline(80, color="grey", ls=":")
    ax2.annotate(f"{n80} countries = 80% of revenue", xy=(n80 - 1, 80), xytext=(n80 + 1, 60),
                 arrowprops=dict(arrowstyle="->"), fontsize=10)
    fig.tight_layout()
    return fig


def plot_price_per_kg(df: pd.DataFrame) -> plt.Figure:
    d = df.dropna(subset=["price_per_kg"])
    d = d[d["price_per_kg"] > 0]
    fig, ax = plt.subplots(figsize=(9, 4.8))
    sns.boxplot(data=d, x="product_group", y="price_per_kg", order=PRODUCT_ORDER, hue="product_group",
                palette="viridis", legend=False, fliersize=2, ax=ax)
    ax.set_yscale("log")
    ax.set(title="Price per kg by product group (log scale)", xlabel="", ylabel="EUR per kg")
    fig.tight_layout()
    return fig


def plot_margin_distribution(df: pd.DataFrame) -> plt.Figure:
    d = df["margin_pct"].dropna().clip(-1, 1) * 100
    fig, ax = plt.subplots(figsize=(9, 4.2))
    sns.histplot(d, bins=50, color=PALETTE[2], ax=ax)
    ax.axvline(0, color="black", lw=0.8)
    ax.set(title="Distribution of row-level gross margin % (clipped to +/-100%)", xlabel="margin %")
    fig.tight_layout()
    return fig


def build_figures(df: pd.DataFrame) -> dict[str, plt.Figure]:
    return {
        "01_region_revenue_volume": plot_region_revenue_volume(df),
        "02_top_countries": plot_top_countries(df),
        "03_product_mix_and_margin": plot_product_mix(df),
        "04_region_product_heatmap": plot_region_product_heatmap(df),
        "05_quarter_comparison": plot_quarter_comparison(df),
        "06_pareto_countries": plot_pareto_countries(df),
        "07_price_per_kg": plot_price_per_kg(df),
        "08_margin_distribution": plot_margin_distribution(df),
    }


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="SteelCo sales data cleaning and analysis")
    parser.add_argument("--data", default="data/SteelCo_Data_Analysis.xlsx", help="path to the Excel file")
    parser.add_argument("--out", default="outputs", help="output folder")
    args = parser.parse_args()
    matplotlib.use("Agg")  # headless-safe when run as a script

    out = Path(args.out)
    (out / "figures").mkdir(parents=True, exist_ok=True)

    raw = load_raw(args.data)
    clean, audit = clean_data(raw)

    print(f"Raw rows: {len(raw):,} | Clean rows: {len(clean):,}\n")
    print("Cleaning audit:\n", audit.to_string(index=False), "\n")
    print("Reconciliation with Excel dashboard:\n", reconcile(raw, clean).round(4).to_string(), "\n")
    print("Product group summary:\n", summarise(clean, "product_group").round(3).to_string(), "\n")

    clean.to_csv(out / "steelco_clean.csv", index=False)
    audit.to_csv(out / "cleaning_audit.csv", index=False)
    for name, fig in build_figures(clean).items():
        fig.savefig(out / "figures" / f"{name}.png", dpi=150, bbox_inches="tight")
        plt.close(fig)
    print(f"Saved cleaned data, audit table and charts to: {out.resolve()}")


if __name__ == "__main__":
    main()
