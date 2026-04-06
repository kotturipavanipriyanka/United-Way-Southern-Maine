"""
UWSM Community Needs Assessment — Analysis Pipeline
CS5100 Final Project | Kotturi & Rachita Sharma
=====================================================
What this does:
    Loads three survey Excel files, cleans and merges them,
    runs statistical comparisons across demographic groups,
    generates publication-ready charts, and exports summary tables.

Data:
    data/UWSM_Data.xlsx               → raw survey responses
    data/UWSM_Analysis_Datasets.xlsx  → cleaned & coded dataset (primary)
    data/Respondents_dataset.xlsx     → respondent demographics

Outputs:
    figures/  → all charts as .png
    outputs/  → summary Excel workbook + stats CSV
"""

import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import seaborn as sns
from scipy.stats import chi2_contingency

warnings.filterwarnings("ignore")

# ── Output directories ────────────────────────────────────────────────────────
os.makedirs("figures", exist_ok=True)
os.makedirs("outputs", exist_ok=True)

# ── File paths ────────────────────────────────────────────────────────────────
PATH_RAW         = "data/UWSM_Data.xlsx"
PATH_ANALYSIS    = "data/UWSM_Analysis_Datasets.xlsx"
PATH_RESPONDENTS = "data/Respondents_dataset.xlsx"

# ── Brand colors & style ──────────────────────────────────────────────────────
C = {
    "navy"   : "#1B3F6E",
    "orange" : "#E87722",
    "green"  : "#4CAF50",
    "purple" : "#9B59B6",
    "red"    : "#E74C3C",
    "teal"   : "#1ABC9C",
    "gold"   : "#F39C12",
    "gray"   : "#7F8C8D",
}
PALETTE = list(C.values())
sns.set_theme(style="whitegrid", font="DejaVu Sans")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DATA LOADING
# Load all three Excel files, normalize headers, report what we find
# ══════════════════════════════════════════════════════════════════════════════

def normalize_columns(df):
    """Lowercase, strip, and sanitize column names for consistent access."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[\s/]+", "_", regex=True)
        .str.replace(r"[^\w]", "", regex=True)
    )
    return df


def load_excel(path, label):
    """Load all sheets from an Excel file and return as a dict of DataFrames."""
    xl     = pd.ExcelFile(path)
    frames = {}
    print(f"\n  {label}  |  sheets: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        df = normalize_columns(xl.parse(sheet))
        print(f"    [{sheet}]  {df.shape[0]:,} rows × {df.shape[1]} cols")
        frames[sheet] = df
    return frames


print("=" * 60)
print("SECTION 1 — Loading data")
print("=" * 60)

raw_sheets  = load_excel(PATH_RAW,         "UWSM_Data.xlsx")
anal_sheets = load_excel(PATH_ANALYSIS,    "UWSM_Analysis_Datasets.xlsx")
resp_sheets = load_excel(PATH_RESPONDENTS, "Respondents_dataset.xlsx")

# Primary working frame is the analysis dataset (fully coded)
df = list(anal_sheets.values())[0].copy()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — COLUMN DETECTION
# Instead of hardcoding headers (which change), we keyword-match column names.
# This makes the script resilient to minor header variations across file versions.
# ══════════════════════════════════════════════════════════════════════════════

def find_col(df, *keywords):
    """Return first column whose name contains any of the given keywords."""
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None


print("\n" + "=" * 60)
print("SECTION 2 — Detecting columns")
print("=" * 60)

COL = {
    "id"          : find_col(df, "id", "unique"),
    "survey_type" : find_col(df, "survey_type", "survey"),
    "challenges"  : find_col(df, "challenge", "biggest"),
    "county"      : find_col(df, "county"),
    "alice"       : find_col(df, "alice_status", "alice"),
    "unheard"     : find_col(df, "unheard", "heard"),
    "community"   : find_col(df, "community", "group"),
    "age"         : find_col(df, "age_group", "age"),
    "race"        : find_col(df, "race", "ethnicity"),
    "connection"  : find_col(df, "connection", "uwsm", "connect"),
    "engagement"  : find_col(df, "engagement", "involved"),
    "hardest_bill": find_col(df, "hardest", "bill", "expense"),
    "barriers"    : find_col(df, "barrier", "getting_in"),
    "supports"    : find_col(df, "support", "helpful"),
    "five_yr"     : find_col(df, "five", "5_year", "goal"),
    "zip"         : find_col(df, "zip"),
}

for k, v in COL.items():
    status = v if v else "NOT FOUND"
    print(f"    {k:15s} → {status}")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — CLEANING & FILTERING
# Fix duplicate columns, standardize strings, filter to Southern Maine,
# and create boolean helper flags for ALICE and unheard respondents.
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 3 — Cleaning & filtering")
print("=" * 60)

# Drop duplicate column names before any string ops — pandas chokes on them
df = df.loc[:, ~df.columns.duplicated()].copy()

# Standardize all text fields: strip whitespace, replace "nan" strings with NaN
for col in df.select_dtypes("object").columns:
    df[col] = df[col].astype(str).str.strip().replace({"nan": np.nan, "": np.nan})

# Keep only Cumberland and York — the two Southern Maine counties in scope
county_col = COL["county"]
if county_col:
    before = len(df)
    df = df[df[county_col].isin(["Cumberland", "York", "Cumberland & York"])].copy()
    print(f"  County filter : {before:,} → {len(df):,} rows")

# ALICE flag — households that are asset-limited, income-constrained
alice_col = COL["alice"]
if alice_col:
    df["is_alice"] = df[alice_col].str.contains("Below", na=False)
    print(f"  ALICE (Below) : {df['is_alice'].sum():,} respondents")

# Unheard flag — respondents who feel their community voice isn't heard
# Try multiple possible value formats since coding varies across survey versions
unheard_col = COL["unheard"]
if unheard_col:
    raw_vals = df[unheard_col].dropna().unique()
    print(f"  Unheard values found: {raw_vals[:10]}")
    df["is_unheard"] = df[unheard_col].str.strip().str.lower().isin(
        ["unheard", "yes", "maybe", "not heard"]
    )
    print(f"  Unheard       : {df['is_unheard'].sum():,} respondents")

print(f"  Final dataset : {len(df):,} rows × {df.shape[1]} cols")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — EXPLODING MULTI-SELECT COLUMNS
# "Check all that apply" answers are stored as semicolon-separated strings.
# We split them into one row per answer so they can be counted and compared.
# Example: "Housing Cost; Food Access" → two separate rows
# ══════════════════════════════════════════════════════════════════════════════

def explode_multi(df, col_key, sep=";"):
    """
    Convert a semicolon-separated multi-select column into long-form rows.
    Each answer gets its own row, with all demographic context columns attached.
    Returns an empty DataFrame if the column doesn't exist.
    """
    col = COL.get(col_key)
    if not col or col not in df.columns:
        return pd.DataFrame()

    id_col = COL["id"] or df.columns[0]

    # Gather demographic context columns — skip any that weren't detected
    context_cols = [
        alice_col, county_col, COL["age"], COL["race"],
        COL["connection"], unheard_col
    ]
    keep = [id_col, col]
    for c in context_cols:
        if c and c in df.columns and c not in keep:
            keep.append(c)

    tmp = df[keep].copy()
    tmp = tmp.loc[:, ~tmp.columns.duplicated()].reset_index(drop=True)

    # Split the semicolon string and explode into separate rows
    tmp[col] = tmp[col].astype(str).str.split(sep)
    tmp      = tmp.explode(col).reset_index(drop=True)
    tmp[col] = tmp[col].astype(str).str.strip().replace({"nan": np.nan, "": np.nan})

    return tmp.dropna(subset=[col])


print("\n" + "=" * 60)
print("SECTION 4 — Exploding multi-select columns")
print("=" * 60)

EX = {k: explode_multi(df, k)
      for k in ["challenges", "engagement", "hardest_bill", "barriers", "supports"]}

for k, v in EX.items():
    print(f"  [{k}] → {len(v):,} response rows")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — STATISTICAL ANALYSIS
# Chi-square tests measure whether two categorical variables are independent.
# Cramér's V gives us effect size — how strong the association actually is.
# We test every meaningful demographic × outcome pairing.
# ══════════════════════════════════════════════════════════════════════════════

def cramers_v(ct):
    """Effect size for chi-square. Range 0 (none) to 1 (perfect association)."""
    chi2, _, _, _ = chi2_contingency(ct)
    n             = ct.sum().sum()
    r, k          = ct.shape
    denom         = n * (min(r, k) - 1)
    return np.sqrt(chi2 / denom) if denom > 0 else 0


def chi_test(df, col_a, col_b, min_n=5):
    """
    Run chi-square test between two categorical columns.
    Drops cells with fewer than min_n responses to avoid unreliable results.
    Returns None if there isn't enough data to run the test.
    """
    sub = df[[col_a, col_b]].dropna().reset_index(drop=True)
    ct  = pd.crosstab(sub[col_a], sub[col_b])
    ct  = ct.loc[ct.sum(axis=1) >= min_n, ct.sum(axis=0) >= min_n]

    if ct.shape[0] < 2 or ct.shape[1] < 2:
        return None

    chi2, p, dof, _ = chi2_contingency(ct)
    return {
        "comparison" : f"{col_a} × {col_b}",
        "chi2"       : round(chi2, 3),
        "p_value"    : round(p, 4),
        "dof"        : dof,
        "cramers_v"  : round(cramers_v(ct), 3),
        "significant": p < 0.05,
    }


print("\n" + "=" * 60)
print("SECTION 5 — Statistical analysis")
print("=" * 60)

results  = []
grp_cols = [c for c in [alice_col, COL["age"], COL["race"],
                         county_col, COL["connection"], unheard_col]
            if c and c in df.columns]

# Test each demographic group variable against each multi-select outcome
for gv in grp_cols:
    for ex_key, ex_df in EX.items():
        val_col = COL.get(ex_key)
        if val_col and val_col in ex_df.columns and gv in ex_df.columns:
            r = chi_test(ex_df, gv, val_col)
            if r:
                results.append(r)

    # Also test against single-answer outcome columns
    for col in [unheard_col, COL["five_yr"]]:
        if col and col in df.columns and col != gv:
            r = chi_test(df, gv, col)
            if r:
                results.append(r)

stats_df = pd.DataFrame(results).drop_duplicates("comparison")
stats_df.to_csv("outputs/statistical_tests.csv", index=False)
print(f"  {len(stats_df)} tests run | "
      f"{stats_df['significant'].sum()} significant (p < 0.05)")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — VISUALIZATION HELPERS
# Three reusable chart builders that handle layout, labels, and styling.
# Every chart saves directly to figures/ via save().
# ══════════════════════════════════════════════════════════════════════════════

def save(name):
    """Save current figure to figures/ and close it cleanly."""
    path = f"figures/{name}.png"
    plt.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"  Saved → {path}")


def hbar(series, title, xlabel="Count", top_n=12, color=C["navy"],
         pct=False, figsize=(10, 6)):
    """Horizontal bar chart — best for ranked categorical data."""
    data = series.value_counts().head(top_n)
    if pct:
        data = (data / data.sum() * 100).round(1)

    fig, ax = plt.subplots(figsize=figsize)
    bars = ax.barh(data.index[::-1], data.values[::-1],
                   color=color, edgecolor="white", linewidth=0.5)

    # Inline value labels on each bar
    for bar, val in zip(bars, data.values[::-1]):
        lbl = f"{val:.1f}%" if pct else f"{int(val):,}"
        ax.text(bar.get_width() + 0.2,
                bar.get_y() + bar.get_height() / 2,
                lbl, va="center", ha="left", fontsize=8.5, color=C["gray"])

    ax.set_title(title, fontsize=13, fontweight="bold", color=C["navy"], pad=10)
    ax.set_xlabel(xlabel, fontsize=10)
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    return fig


def grouped_pct_bar(ct, title, figsize=(12, 6)):
    """
    Stacked 100% bar chart — shows composition within each group.
    Ideal for comparing how response distributions shift across demographics.
    """
    pct = ct.div(ct.sum(axis=1), axis=0) * 100
    fig, ax = plt.subplots(figsize=figsize)
    pct.plot(kind="bar", stacked=True, ax=ax,
             color=PALETTE[:pct.shape[1]], edgecolor="white", linewidth=0.4)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter())
    ax.set_title(title, fontsize=13, fontweight="bold", color=C["navy"], pad=10)
    ax.set_xlabel("")
    ax.legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    plt.xticks(rotation=30, ha="right", fontsize=9)
    plt.tight_layout()
    return fig


def heatmap_chart(ct, title, figsize=(13, 7)):
    """
    Percentage heatmap — useful when there are many categories on both axes.
    Normalizes by row so each group's distribution is directly comparable.
    """
    pct = ct.div(ct.sum(axis=1), axis=0) * 100
    fig, ax = plt.subplots(figsize=figsize)
    sns.heatmap(pct, annot=True, fmt=".1f", cmap="Blues",
                linewidths=0.4, ax=ax, cbar_kws={"label": "% within row"})
    ax.set_title(title, fontsize=13, fontweight="bold", color=C["navy"], pad=10)
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — GENERATING ALL CHARTS
# Organized into four themes: Demographics, Challenges, Unheard Communities,
# Engagement & ALICE-specific findings.
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 6 — Generating charts → figures/")
print("=" * 60)

# Shorthand references for frequently used column names
ch_col   = COL["challenges"]
eng_col  = COL["engagement"]
bl_col   = COL["hardest_bill"]
fy_col   = COL["five_yr"]
age_col  = COL["age"]
race_col = COL["race"]
con_col  = COL["connection"]
id_col   = COL["id"] or df.columns[0]

# ── Theme 1: Who responded? ───────────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle("Survey Respondent Demographics — Southern Maine",
             fontsize=15, fontweight="bold", color=C["navy"])

for (col, label), ax in zip(
    [(county_col, "County"), (alice_col, "ALICE Status"),
     (age_col,   "Age Group"), (race_col, "Race / Ethnicity")],
    axes.flatten()
):
    if col and col in df.columns:
        counts = df[col].value_counts().head(8)
        ax.barh(counts.index[::-1], counts.values[::-1],
                color=C["navy"], edgecolor="white")
        ax.set_title(label, fontsize=11, fontweight="bold")
        ax.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
save("00_demographics_overview")

# ── Theme 2: Community Challenges ────────────────────────────────────────────

if ch_col and len(EX["challenges"]) > 0:

    # Overall top challenges
    hbar(EX["challenges"][ch_col],
         "Top Community Challenges — Southern Maine",
         color=C["navy"])
    save("01_top_challenges_overall")

    # Challenges by ALICE status — heatmap works well here (2 groups, many challenges)
    if alice_col:
        top10 = EX["challenges"][ch_col].value_counts().head(10).index
        sub   = EX["challenges"][EX["challenges"][ch_col].isin(top10)]
        ct    = pd.crosstab(sub[alice_col], sub[ch_col])
        heatmap_chart(ct, "Challenges by ALICE Status (% within group)")
        save("02_challenges_by_alice")

    # Challenges by age — stacked bar shows how priorities shift with age
    if age_col:
        top8 = EX["challenges"][ch_col].value_counts().head(8).index
        sub  = EX["challenges"][EX["challenges"][ch_col].isin(top8)]
        ct   = pd.crosstab(sub[age_col], sub[ch_col])
        grouped_pct_bar(ct, "Challenges by Age Group")
        save("03_challenges_by_age")

    # Challenges by county — Cumberland vs York
    if county_col:
        top8 = EX["challenges"][ch_col].value_counts().head(8).index
        sub  = EX["challenges"][EX["challenges"][ch_col].isin(top8)]
        ct   = pd.crosstab(sub[county_col], sub[ch_col])
        grouped_pct_bar(ct, "Challenges by County — Cumberland vs York")
        save("04_challenges_by_county")

    # Challenges by race/ethnicity — heatmap handles many groups cleanly
    if race_col:
        top8 = EX["challenges"][ch_col].value_counts().head(8).index
        sub  = EX["challenges"][EX["challenges"][ch_col].isin(top8)]
        ct   = pd.crosstab(sub[race_col], sub[ch_col])
        heatmap_chart(ct, "Challenges by Race / Ethnicity (% within group)",
                      figsize=(14, 8))
        save("05_challenges_by_race")

    # Challenges by UWSM connection level
    if con_col:
        top8 = EX["challenges"][ch_col].value_counts().head(8).index
        sub  = EX["challenges"][EX["challenges"][ch_col].isin(top8)]
        ct   = pd.crosstab(sub[con_col], sub[ch_col])
        grouped_pct_bar(ct, "Challenges by UWSM Connection Level")
        save("06_challenges_by_connection")

# ── Theme 3: Who Feels Unheard? ───────────────────────────────────────────────

if unheard_col:

    if alice_col:
        ct = pd.crosstab(df[alice_col], df[unheard_col])
        grouped_pct_bar(ct, "Feeling Unheard by ALICE Status", figsize=(8, 5))
        save("07_unheard_by_alice")

    if age_col:
        ct = pd.crosstab(df[age_col], df[unheard_col])
        grouped_pct_bar(ct, "Feeling Unheard by Age Group", figsize=(9, 5))
        save("08_unheard_by_age")

    if race_col:
        ct = pd.crosstab(df[race_col], df[unheard_col])
        grouped_pct_bar(ct, "Feeling Unheard by Race / Ethnicity", figsize=(10, 6))
        save("09_unheard_by_race")

    # Do unheard communities face systematically different challenges?
    if ch_col and len(EX["challenges"]) > 0:
        top8 = EX["challenges"][ch_col].value_counts().head(8).index
        sub  = EX["challenges"][EX["challenges"][ch_col].isin(top8)]
        ct   = pd.crosstab(sub[unheard_col], sub[ch_col])
        grouped_pct_bar(ct, "Challenges: Unheard vs Heard Communities")
        save("10_challenges_unheard_vs_heard")

# ── Theme 4: Engagement & ALICE-Specific Findings ────────────────────────────

if eng_col and len(EX["engagement"]) > 0:

    hbar(EX["engagement"][eng_col],
         "How Residents Want to Get Involved", color=C["orange"])
    save("11_engagement_overall")

    if alice_col:
        ct = pd.crosstab(EX["engagement"][alice_col], EX["engagement"][eng_col])
        grouped_pct_bar(ct, "Engagement Preferences by ALICE Status")
        save("12_engagement_by_alice")

    if age_col:
        ct = pd.crosstab(EX["engagement"][age_col], EX["engagement"][eng_col])
        grouped_pct_bar(ct, "Engagement Preferences by Age Group")
        save("13_engagement_by_age")

# Hardest bills — ALICE households only
if bl_col and "is_alice" in df.columns and len(EX["hardest_bill"]) > 0:
    alice_ids  = df[df["is_alice"]][id_col]
    alice_bills = EX["hardest_bill"][EX["hardest_bill"][id_col].isin(alice_ids)]
    if len(alice_bills) > 0:
        hbar(alice_bills[bl_col],
             "Hardest Bills to Afford — ALICE Households", color=C["green"])
        save("14_hardest_bills_alice")

# 5-year goals — ALICE households only
if fy_col and "is_alice" in df.columns:
    alice_goals = df[df["is_alice"]][fy_col].dropna()
    if len(alice_goals) > 0:
        hbar(alice_goals, "5-Year Goals — ALICE Households", color=C["purple"])
        save("15_five_yr_goals_alice")

# Statistical significance summary — which associations are actually meaningful?
if len(stats_df) > 0:
    sig = (stats_df[stats_df["significant"]]
           .sort_values("cramers_v", ascending=False)
           .head(15))
    if len(sig) > 0:
        fig, ax = plt.subplots(figsize=(10, max(5, len(sig) * 0.55 + 1)))
        bar_colors = [C["orange"] if v > 0.3 else C["navy"]
                      for v in sig["cramers_v"]]
        ax.barh(sig["comparison"].str.replace("_", " ")[::-1],
                sig["cramers_v"][::-1],
                color=bar_colors[::-1], edgecolor="white")
        ax.axvline(0.1, ls="--", color=C["gray"],   lw=1, label="Small effect (0.1)")
        ax.axvline(0.3, ls="--", color=C["orange"], lw=1, label="Medium effect (0.3)")
        ax.set_title("Strongest Associations — Cramér's V Effect Size",
                     fontsize=13, fontweight="bold", color=C["navy"])
        ax.set_xlabel("Cramér's V")
        ax.legend(fontsize=8)
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        save("16_cramers_v_summary")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — EXPORT SUMMARY TABLES
# Everything a report writer needs: counts, percentages, crosstabs, and the
# full cleaned dataset — all in one workbook with clearly named sheets.
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 7 — Exporting summary tables → outputs/")
print("=" * 60)

with pd.ExcelWriter("outputs/uwsm_summary_tables.xlsx", engine="openpyxl") as writer:

    # Overall challenge frequency
    if ch_col and len(EX["challenges"]) > 0:
        c = EX["challenges"][ch_col].value_counts().reset_index()
        c.columns = ["Challenge", "Count"]
        c["Pct_of_Respondents"] = (c["Count"] / len(df) * 100).round(1)
        c.to_excel(writer, sheet_name="Challenges_Overall", index=False)

    # Challenge breakdown by ALICE status
    if ch_col and alice_col and len(EX["challenges"]) > 0:
        ct = pd.crosstab(EX["challenges"][ch_col], EX["challenges"][alice_col],
                         normalize="columns").mul(100).round(1)
        ct.to_excel(writer, sheet_name="Challenges_by_ALICE")

    # Challenge breakdown by county
    if ch_col and county_col and len(EX["challenges"]) > 0:
        ct = pd.crosstab(EX["challenges"][ch_col], EX["challenges"][county_col],
                         normalize="columns").mul(100).round(1)
        ct.to_excel(writer, sheet_name="Challenges_by_County")

    # Engagement preferences
    if eng_col and len(EX["engagement"]) > 0:
        e = EX["engagement"][eng_col].value_counts().reset_index()
        e.columns = ["Engagement_Type", "Count"]
        e["Pct"] = (e["Count"] / len(df) * 100).round(1)
        e.to_excel(writer, sheet_name="Engagement_Overall", index=False)

    # Who feels unheard — broken down by each demographic
    if unheard_col:
        for gv in [alice_col, age_col, race_col, county_col]:
            if gv and gv in df.columns:
                ct = pd.crosstab(df[gv], df[unheard_col],
                                 normalize="index").mul(100).round(1)
                ct.to_excel(writer, sheet_name=f"Unheard_by_{gv[:12]}")

    # ALICE-specific: 5-year goals
    if fy_col and "is_alice" in df.columns:
        goals = df[df["is_alice"]][fy_col].value_counts().reset_index()
        goals.columns = ["Goal", "Count"]
        goals.to_excel(writer, sheet_name="ALICE_5yr_Goals", index=False)

    # ALICE-specific: hardest bills
    if bl_col and "is_alice" in df.columns and len(EX["hardest_bill"]) > 0:
        alice_ids   = df[df["is_alice"]][id_col]
        alice_bills = EX["hardest_bill"][EX["hardest_bill"][id_col].isin(alice_ids)]
        b = alice_bills[bl_col].value_counts().reset_index()
        b.columns = ["Bill_Type", "Count"]
        b.to_excel(writer, sheet_name="ALICE_Hardest_Bills", index=False)

    # All chi-square test results
    if len(stats_df) > 0:
        stats_df.to_excel(writer, sheet_name="Chi_Square_Tests", index=False)

    # Full cleaned & filtered dataset for reference
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)

print("  Saved → outputs/uwsm_summary_tables.xlsx")
print("  Saved → outputs/statistical_tests.csv")

# ── Final summary ─────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("PIPELINE COMPLETE")
print(f"  Figures  : figures/   ({len(os.listdir('figures'))} files)")
print(f"  Outputs  : outputs/   ({len(os.listdir('outputs'))} files)")
print("=" * 60)
