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
# Load all sheets from the analysis Excel file.
# The workbook already ships pre-exploded long-form sheets for each
# multi-select question, so we read them directly instead of splitting strings.
# ══════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("SECTION 1 — Loading data")
print("=" * 60)

xl = pd.ExcelFile(PATH_ANALYSIS)
print(f"  Sheets found: {xl.sheet_names}")

# Primary respondent-level frame (one row per person)
df = xl.parse("Respondents")
print(f"  Respondents  : {df.shape[0]:,} rows × {df.shape[1]} cols")

# Pre-exploded long-form sheets (one row per response option selected)
challenge_long = xl.parse("Challenge_Long")
engagement_long = xl.parse("Engagement_Long")
expense_barriers_long = xl.parse("Expense_Barriers_Long")
supports_long = xl.parse("Supports_Long")

print(f"  Challenge_Long       : {len(challenge_long):,} rows")
print(f"  Engagement_Long      : {len(engagement_long):,} rows")
print(f"  Expense_Barriers_Long: {len(expense_barriers_long):,} rows")
print(f"  Supports_Long        : {len(supports_long):,} rows")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — COLUMN NAMES
# Exact column names as they appear in the dataset (confirmed from inspection).
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 2 — Column references")
print("=" * 60)

# Respondents sheet columns
COL_ID          = "id"
COL_SURVEY_TYPE = "survey_type"
COL_COUNTY      = "County"
COL_ALICE       = "ALICE_status"          # values: "Below ALICE", "Above ALICE"
COL_UNHEARD     = "unheard"               # values: "Yes", "No", "Maybe"
COL_AGE         = "age_group"
COL_RACE        = "race"
COL_CONNECTION  = "UW_connection"
COL_HARDEST_EXP = "hardest_expenses"
COL_FIVE_YR     = "Community_Voice"       # closest proxy for 5-year / voice goal
COL_ZIP         = "zip"

# Long-form sheet value columns (the exploded answer per row)
COL_CHALLENGE_CODE  = "challenge_code"    # Challenge_Long
COL_ENGAGEMENT      = "Engagement"        # Engagement_Long
COL_EXP_BARRIER     = "Expense_barriers"  # Expense_Barriers_Long
COL_SUPPORT         = "Supports"          # Supports_Long

for name, val in {
    "county": COL_COUNTY, "alice": COL_ALICE, "unheard": COL_UNHEARD,
    "age": COL_AGE, "race": COL_RACE, "connection": COL_CONNECTION,
    "hardest_expenses": COL_HARDEST_EXP, "community_voice": COL_FIVE_YR,
    "challenge_code (long)": COL_CHALLENGE_CODE,
    "engagement (long)": COL_ENGAGEMENT,
}.items():
    print(f"    {name:25s} → {val}")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — CLEANING & FILTERING
# Standardize strings, filter to Southern Maine counties,
# and create boolean helper flags for ALICE and unheard respondents.
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 3 — Cleaning & filtering")
print("=" * 60)

def clean_strings(frame):
    """Strip whitespace and replace empty / 'nan' strings with NaN."""
    for col in frame.select_dtypes("object").columns:
        frame[col] = frame[col].astype(str).str.strip().replace(
            {"nan": np.nan, "": np.nan, "None": np.nan}
        )
    return frame

df                    = clean_strings(df)
challenge_long        = clean_strings(challenge_long)
engagement_long       = clean_strings(engagement_long)
expense_barriers_long = clean_strings(expense_barriers_long)
supports_long         = clean_strings(supports_long)

# Keep only Cumberland and York — the two Southern Maine counties in scope
before = len(df)
df = df[df[COL_COUNTY].isin(["Cumberland", "York", "Cumberland & York"])].copy()
print(f"  County filter : {before:,} → {len(df):,} rows")

# Apply same county filter to all long-form sheets
valid_ids = set(df[COL_ID])
challenge_long        = challenge_long[challenge_long[COL_ID].isin(valid_ids)].copy()
engagement_long       = engagement_long[engagement_long[COL_ID].isin(valid_ids)].copy()
expense_barriers_long = expense_barriers_long[expense_barriers_long[COL_ID].isin(valid_ids)].copy()
supports_long         = supports_long[supports_long[COL_ID].isin(valid_ids)].copy()

# ALICE flag — households that are asset-limited, income-constrained
df["is_alice"] = df[COL_ALICE].str.contains("Below", na=False)
print(f"  ALICE (Below) : {df['is_alice'].sum():,} respondents")

# Unheard flag — "Yes" or "Maybe" both indicate some degree of feeling unheard
df["is_unheard"] = df[COL_UNHEARD].isin(["Yes", "Maybe"])
print(f"  Unheard (Yes/Maybe): {df['is_unheard'].sum():,} respondents")

print(f"  Final dataset : {len(df):,} rows × {df.shape[1]} cols")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STATISTICAL ANALYSIS
# Chi-square tests measure whether two categorical variables are independent.
# Cramér's V gives effect size — how strong the association actually is.
# ══════════════════════════════════════════════════════════════════════════════

def cramers_v(ct):
    """Effect size for chi-square. Range 0 (none) to 1 (perfect association)."""
    chi2, _, _, _ = chi2_contingency(ct)
    n             = ct.sum().sum()
    r, k          = ct.shape
    denom         = n * (min(r, k) - 1)
    return np.sqrt(chi2 / denom) if denom > 0 else 0


def chi_test(frame, col_a, col_b, min_n=5):
    """
    Run chi-square test between two categorical columns.
    Returns None if there isn't enough data.
    """
    sub = frame[[col_a, col_b]].dropna()
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
print("SECTION 4 — Statistical analysis")
print("=" * 60)

results  = []
grp_cols = [COL_ALICE, COL_AGE, COL_RACE, COL_COUNTY, COL_CONNECTION, COL_UNHEARD]

# Test each demographic group against each long-form outcome
long_frames = {
    "challenge_code"  : (challenge_long,        COL_CHALLENGE_CODE),
    "engagement"      : (engagement_long,        COL_ENGAGEMENT),
    "expense_barriers": (expense_barriers_long,  COL_EXP_BARRIER),
    "supports"        : (supports_long,          COL_SUPPORT),
}

for gv in grp_cols:
    for key, (frame, val_col) in long_frames.items():
        if val_col in frame.columns and gv in frame.columns:
            r = chi_test(frame, gv, val_col)
            if r:
                results.append(r)

    # Single-answer outcome columns on the main respondents frame
    for col in [COL_UNHEARD, COL_FIVE_YR]:
        if col in df.columns and col != gv:
            r = chi_test(df, gv, col)
            if r:
                results.append(r)

stats_df = pd.DataFrame(results).drop_duplicates("comparison")
stats_df.to_csv("outputs/statistical_tests.csv", index=False)
print(f"  {len(stats_df)} tests run | "
      f"{stats_df['significant'].sum()} significant (p < 0.05)")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — VISUALIZATION HELPERS
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
    data = series.dropna().value_counts().head(top_n)
    if pct:
        data = (data / data.sum() * 100).round(1)

    fig, ax = plt.subplots(figsize=figsize)
    bars = ax.barh(data.index[::-1], data.values[::-1],
                   color=color, edgecolor="white", linewidth=0.5)

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
    """Stacked 100% bar chart — shows composition within each group."""
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
    """Percentage heatmap — normalizes by row for direct group comparison."""
    pct = ct.div(ct.sum(axis=1), axis=0) * 100
    fig, ax = plt.subplots(figsize=figsize)
    sns.heatmap(pct, annot=True, fmt=".1f", cmap="Blues",
                linewidths=0.4, ax=ax, cbar_kws={"label": "% within row"})
    ax.set_title(title, fontsize=13, fontweight="bold", color=C["navy"], pad=10)
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — GENERATING ALL CHARTS
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 6 — Generating charts → figures/")
print("=" * 60)

# ── Theme 1: Who responded? ───────────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle("Survey Respondent Demographics — Southern Maine",
             fontsize=15, fontweight="bold", color=C["navy"])

for (col, label), ax in zip(
    [(COL_COUNTY, "County"), (COL_ALICE, "ALICE Status"),
     (COL_AGE,   "Age Group"), (COL_RACE, "Race / Ethnicity")],
    axes.flatten()
):
    if col in df.columns:
        counts = df[col].dropna().value_counts().head(8)
        ax.barh(counts.index[::-1], counts.values[::-1],
                color=C["navy"], edgecolor="white")
        ax.set_title(label, fontsize=11, fontweight="bold")
        ax.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
save("00_demographics_overview")

# ── Theme 2: Community Challenges ────────────────────────────────────────────

if COL_CHALLENGE_CODE in challenge_long.columns and len(challenge_long) > 0:

    # Overall top challenges
    hbar(challenge_long[COL_CHALLENGE_CODE],
         "Top Community Challenges — Southern Maine",
         color=C["navy"])
    save("01_top_challenges_overall")

    # Challenges by ALICE status
    if COL_ALICE in challenge_long.columns:
        top10 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(10).index
        sub   = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top10)]
        ct    = pd.crosstab(sub[COL_ALICE], sub[COL_CHALLENGE_CODE])
        heatmap_chart(ct, "Challenges by ALICE Status (% within group)")
        save("02_challenges_by_alice")

    # Challenges by age group
    if COL_AGE in challenge_long.columns:
        top8 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(8).index
        sub  = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top8)]
        ct   = pd.crosstab(sub[COL_AGE], sub[COL_CHALLENGE_CODE])
        grouped_pct_bar(ct, "Challenges by Age Group")
        save("03_challenges_by_age")

    # Challenges by county
    if COL_COUNTY in challenge_long.columns:
        top8 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(8).index
        sub  = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top8)]
        ct   = pd.crosstab(sub[COL_COUNTY], sub[COL_CHALLENGE_CODE])
        grouped_pct_bar(ct, "Challenges by County — Cumberland vs York")
        save("04_challenges_by_county")

    # Challenges by race/ethnicity
    if COL_RACE in challenge_long.columns:
        top8 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(8).index
        sub  = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top8)]
        ct   = pd.crosstab(sub[COL_RACE], sub[COL_CHALLENGE_CODE])
        heatmap_chart(ct, "Challenges by Race / Ethnicity (% within group)",
                      figsize=(14, 8))
        save("05_challenges_by_race")

    # Challenges by UWSM connection level
    if COL_CONNECTION in challenge_long.columns:
        top8 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(8).index
        sub  = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top8)]
        ct   = pd.crosstab(sub[COL_CONNECTION], sub[COL_CHALLENGE_CODE])
        grouped_pct_bar(ct, "Challenges by UWSM Connection Level")
        save("06_challenges_by_connection")

# ── Theme 3: Who Feels Unheard? ───────────────────────────────────────────────

if COL_UNHEARD in df.columns:

    if COL_ALICE in df.columns:
        ct = pd.crosstab(df[COL_ALICE], df[COL_UNHEARD])
        grouped_pct_bar(ct, "Feeling Unheard by ALICE Status", figsize=(8, 5))
        save("07_unheard_by_alice")

    if COL_AGE in df.columns:
        ct = pd.crosstab(df[COL_AGE], df[COL_UNHEARD])
        grouped_pct_bar(ct, "Feeling Unheard by Age Group", figsize=(9, 5))
        save("08_unheard_by_age")

    if COL_RACE in df.columns:
        ct = pd.crosstab(df[COL_RACE], df[COL_UNHEARD])
        grouped_pct_bar(ct, "Feeling Unheard by Race / Ethnicity", figsize=(10, 6))
        save("09_unheard_by_race")

    # Do unheard communities face systematically different challenges?
    if COL_UNHEARD in challenge_long.columns and len(challenge_long) > 0:
        top8 = challenge_long[COL_CHALLENGE_CODE].value_counts().head(8).index
        sub  = challenge_long[challenge_long[COL_CHALLENGE_CODE].isin(top8)]
        ct   = pd.crosstab(sub[COL_UNHEARD], sub[COL_CHALLENGE_CODE])
        grouped_pct_bar(ct, "Challenges: Unheard vs Heard Communities")
        save("10_challenges_unheard_vs_heard")

# ── Theme 4: Engagement & ALICE-Specific Findings ────────────────────────────

if COL_ENGAGEMENT in engagement_long.columns and len(engagement_long) > 0:

    hbar(engagement_long[COL_ENGAGEMENT],
         "How Residents Want to Get Involved", color=C["orange"])
    save("11_engagement_overall")

    if COL_ALICE in engagement_long.columns:
        ct = pd.crosstab(engagement_long[COL_ALICE], engagement_long[COL_ENGAGEMENT])
        grouped_pct_bar(ct, "Engagement Preferences by ALICE Status")
        save("12_engagement_by_alice")

    if COL_AGE in engagement_long.columns:
        ct = pd.crosstab(engagement_long[COL_AGE], engagement_long[COL_ENGAGEMENT])
        grouped_pct_bar(ct, "Engagement Preferences by Age Group")
        save("13_engagement_by_age")

# Hardest expenses — ALICE households only
if COL_HARDEST_EXP in df.columns and "is_alice" in df.columns:
    alice_exp = df[df["is_alice"]][COL_HARDEST_EXP].dropna()
    if len(alice_exp) > 0:
        hbar(alice_exp,
             "Hardest Expenses to Afford — ALICE Households", color=C["green"])
        save("14_hardest_expenses_alice")

# Community voice / 5-year goal — ALICE households only
if COL_FIVE_YR in df.columns and "is_alice" in df.columns:
    alice_voice = df[df["is_alice"]][COL_FIVE_YR].dropna()
    if len(alice_voice) > 0:
        hbar(alice_voice, "Community Voice Responses — ALICE Households",
             color=C["purple"])
        save("15_community_voice_alice")

# Statistical significance summary
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
# SECTION 7 — EXPORT SUMMARY TABLES
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("SECTION 7 — Exporting summary tables → outputs/")
print("=" * 60)

with pd.ExcelWriter("outputs/uwsm_summary_tables.xlsx", engine="openpyxl") as writer:

    # Overall challenge frequency
    if COL_CHALLENGE_CODE in challenge_long.columns:
        c = challenge_long[COL_CHALLENGE_CODE].value_counts().reset_index()
        c.columns = ["Challenge", "Count"]
        c["Pct_of_Respondents"] = (c["Count"] / len(df) * 100).round(1)
        c.to_excel(writer, sheet_name="Challenges_Overall", index=False)

    # Challenge breakdown by ALICE status
    if COL_CHALLENGE_CODE in challenge_long.columns and COL_ALICE in challenge_long.columns:
        ct = pd.crosstab(challenge_long[COL_CHALLENGE_CODE],
                         challenge_long[COL_ALICE],
                         normalize="columns").mul(100).round(1)
        ct.to_excel(writer, sheet_name="Challenges_by_ALICE")

    # Challenge breakdown by county
    if COL_CHALLENGE_CODE in challenge_long.columns and COL_COUNTY in challenge_long.columns:
        ct = pd.crosstab(challenge_long[COL_CHALLENGE_CODE],
                         challenge_long[COL_COUNTY],
                         normalize="columns").mul(100).round(1)
        ct.to_excel(writer, sheet_name="Challenges_by_County")

    # Engagement preferences
    if COL_ENGAGEMENT in engagement_long.columns:
        e = engagement_long[COL_ENGAGEMENT].value_counts().reset_index()
        e.columns = ["Engagement_Type", "Count"]
        e["Pct"] = (e["Count"] / len(df) * 100).round(1)
        e.to_excel(writer, sheet_name="Engagement_Overall", index=False)

    # Who feels unheard — broken down by each demographic
    if COL_UNHEARD in df.columns:
        for gv in [COL_ALICE, COL_AGE, COL_RACE, COL_COUNTY]:
            if gv in df.columns:
                ct = pd.crosstab(df[gv], df[COL_UNHEARD],
                                 normalize="index").mul(100).round(1)
                ct.to_excel(writer, sheet_name=f"Unheard_by_{gv[:12]}")

    # ALICE-specific: hardest expenses
    if COL_HARDEST_EXP in df.columns and "is_alice" in df.columns:
        b = df[df["is_alice"]][COL_HARDEST_EXP].value_counts().reset_index()
        b.columns = ["Expense_Type", "Count"]
        b.to_excel(writer, sheet_name="ALICE_Hardest_Expenses", index=False)

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
