# UWSM Community Needs Assessment — Analysis Pipeline
### CS5100 Final Project | Kotturi & Sharma | Northeastern University, Roux Institute

---

## Overview

This project analyzes community survey data collected by the **United Way of Southern Maine (UWSM)** to understand what challenges Southern Maine residents face, which communities feel unheard, and how people want to get involved in creating change.

The pipeline loads three Excel datasets, cleans and merges them, runs statistical analysis across demographic groups, generates 17 publication-ready charts, and exports summary tables — all in a single Python script.

---

## Project Structure

```
CS5100 Final Project Kotturi and Sharma/
│
├── data/
│   ├── UWSM_Data.xlsx                  # Raw survey responses (1,235 rows)
│   ├── UWSM_Analysis_Datasets.xlsx     # Cleaned & pre-exploded long-form sheets
│   │     Sheets: Respondents,
│   │             Challenge_Long,
│   │             Race_Long,
│   │             Engagement_Long,
│   │             Expense_Barriers_Long,
│   │             Supports_Long
│   └── Respondents_dataset.xlsx        # Full respondent demographics (8,743 rows)
│
├── figures/                            # Auto-created — all 17 charts saved here
├── outputs/                            # Auto-created — summary tables + stats CSV
│
├── project.py                          # Main analysis pipeline
├── requirements.txt                    # Python dependencies
└── README.md                           # This file
```

---

## Setup & Installation

**1. Clone or download the project folder**

**2. Create and activate a virtual environment (recommended)**
```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# Mac/Linux
source .venv/bin/activate
```

**3. Install dependencies**
```bash
pip install -r requirements.txt
```

**4. Run the pipeline**
```bash
python project.py
```

---

## Dependencies

```
pandas
numpy
matplotlib
seaborn
scipy
openpyxl
```

All installable via `pip install -r requirements.txt`.

---

## What the Pipeline Does

### Section 1 — Data Loading
Loads all three Excel files and all 6 sheets from the analysis dataset. Normalizes column headers (lowercase, strip whitespace, remove special characters) and reports shape for every sheet.

### Section 2 — Column Detection
Uses keyword matching to detect relevant columns automatically — no hardcoded header names. This makes the script resilient to minor variations across file versions.

### Section 3 — Cleaning & Filtering
- Strips whitespace and replaces `"nan"` strings with proper `NaN`
- Drops duplicate column names (a known issue in the source files)
- Filters to **Cumberland and York counties only** (Southern Maine scope)
- Creates boolean flags: `is_alice` (Below ALICE threshold) and `is_unheard` (Yes or Maybe)

### Section 4 — Attaching Demographics
Each long-form sheet gets ALICE status, county, age group, and UWSM connection level merged in from the respondent frame using the shared ID column.

### Section 5 — Statistical Analysis
Runs **chi-square tests** across every meaningful demographic × outcome pairing and computes **Cramér's V** as effect size. Results exported to `outputs/statistical_tests.csv`.

| Effect Size | Cramér's V |
|---|---|
| Small | 0.1 |
| Medium | 0.3 |
| Large | 0.5+ |

### Section 6 — Visualizations
17 charts saved to `figures/` across four themes (see chart list below).

### Section 7 — Export
Summary tables exported to `outputs/uwsm_summary_tables.xlsx` with 10+ named sheets.

---

## Output Charts (`figures/`)

### Theme 1 — Who Responded?
| File | Description |
|---|---|
| `00_demographics_overview.png` | 2×2 grid: county, ALICE status, age group, survey type |

### Theme 2 — Community Challenges
| File | Description |
|---|---|
| `01_top_challenges_overall.png` | Top 12 challenges ranked across all respondents |
| `02_challenges_by_alice.png` | Heatmap — challenges by ALICE vs non-ALICE |
| `03_challenges_by_age.png` | Stacked bar — challenges by age group |
| `04_challenges_by_county.png` | Stacked bar — Cumberland vs York |
| `05_challenges_by_race.png` | Heatmap — challenges by race/ethnicity |
| `06_challenges_by_connection.png` | Stacked bar — challenges by UWSM connection level |

### Theme 3 — Who Feels Unheard?
| File | Description |
|---|---|
| `07_unheard_by_alice.png` | Unheard feeling by ALICE status |
| `08_unheard_by_age.png` | Unheard feeling by age group |
| `09_unheard_by_race.png` | Unheard feeling by race/ethnicity |
| `10_challenges_unheard_vs_heard.png` | Do unheard communities face different challenges? |

### Theme 4 — Engagement & ALICE-Specific Findings
| File | Description |
|---|---|
| `11_engagement_overall.png` | Top ways residents want to get involved |
| `12_engagement_by_alice.png` | Engagement preferences by ALICE status |
| `13_engagement_by_age.png` | Engagement preferences by age group |
| `14_hardest_bills_alice.png` | Hardest bills to afford — ALICE households only |
| `15_five_yr_goals_alice.png` | 5-year goals — ALICE households only |
| `16_cramers_v_summary.png` | Top statistically significant associations by effect size |

---

## Output Tables (`outputs/`)

### `uwsm_summary_tables.xlsx`

| Sheet | Contents |
|---|---|
| `Challenges_Overall` | Count + % of each challenge across all respondents |
| `Challenges_by_ALICE` | % breakdown by ALICE vs non-ALICE |
| `Challenges_by_County` | % breakdown by Cumberland vs York |
| `Engagement_Overall` | Count + % of each engagement type |
| `Unheard_by_ALICE` | % feeling unheard by ALICE status |
| `Unheard_by_Age` | % feeling unheard by age group |
| `Unheard_by_Race` | % feeling unheard by race/ethnicity |
| `Unheard_by_County` | % feeling unheard by county |
| `ALICE_5yr_Goals` | 5-year goal categories for ALICE households |
| `ALICE_Hardest_Bills` | Bill categories hardest for ALICE households |
| `Chi_Square_Tests` | All chi-square results with p-values and Cramér's V |
| `Cleaned_Data` | Full filtered, cleaned respondent dataset |

### `statistical_tests.csv`
Raw chi-square results for all tested pairings — comparison label, chi², p-value, degrees of freedom, Cramér's V, and significance flag.

---

## Dataset Notes

| Dataset | Rows | Description |
|---|---|---|
| `UWSM_Data.xlsx` | 1,235 | Raw survey — all regions |
| `Respondents` sheet | 1,053 | Filtered to Cumberland & York |
| ALICE households | 259 | Below cost-of-living threshold |
| Unheard respondents | 242 | Answered Yes or Maybe to feeling unheard |

### Known Anomalies
- **Duplicate column names** in source files — handled with `loc[:, ~df.columns.duplicated()]`
- **Mixed types after explode** — handled with `.astype(str)` before string operations
- **Duplicate index after explode** — handled with `reset_index(drop=True)`
- **Multi-select answers** stored as semicolon-separated strings in raw data — the analysis file provides pre-exploded long-form sheets for each question type

---

## Research Questions Addressed

1. What are the most pressing challenges in Southern Maine communities?
2. How do challenges differ by ALICE status, age, race, and county?
3. Which communities feel unheard — and what challenges do they face?
4. How do residents want to engage to drive community change?
5. What are the specific needs and goals of ALICE households?

---

## Authors

**Pavani Kotturi & Rachita Sharma**
MS in Artificial Intelligence — Northeastern University, Roux Institute
CS5100 Final Project