# SEC 10-K Ten-Year Financial & Sentiment Analysis: Apple vs. Intel

An end-to-end automated financial intelligence pipeline that downloads, parses, standardizes, and analyzes ten consecutive fiscal years (**FY2016 through FY2025**) of SEC Form 10-K annual reports for **Apple Inc. (AAPL)** and **Intel Corporation (INTC)**.

The pipeline extracts multi-period financial statements, computes essential liquidity, solvency, efficiency, and profitability ratios, performs natural language processing (NLP) sentiment scoring on Item 7 Management's Discussion and Analysis (MD&A) narratives using the **Loughran-McDonald Financial Lexicon**, correlates MD&A tone with historical closing stock prices, produces audited eight-tab Excel workbooks, generates comparative figures, and compiles a self-contained interactive web dashboard.

---

## Live Interactive Dashboard

The dashboard is self-contained and ready for GitHub Pages:
- **Local Viewing**: Open [`index.html`](file:///Users/amirtha/Documents/Tenk_project/index.html) in any modern web browser.
- **GitHub Pages**: Automatically deployed from repository root (`/index.html`).

---

## Key Pipeline Features

1. **Automated SEC EDGAR 10-K Retrieval**:
   - Queries SEC Submissions API for verified CIK identifiers (`0000320193` for Apple, `0000050863` for Intel).
   - Downloads official 10-K HTML filings and parses XBRL company concept endpoints.
2. **Statement Normalization & Multi-Period Stacking**:
   - Parses Consolidated Statements of Operations, Balance Sheets, and Cash Flows.
   - Reconciles comparative prior-year reporting periods to isolate exactly the 10 target fiscal years.
   - Guarantees the fundamental accounting equation holds for all years: $\text{Assets} = \text{Liabilities} + \text{Stockholders' Equity}$.
3. **Loughran-McDonald MD&A Sentiment Analysis**:
   - Robustly isolates Item 7 (MD&A) while stripping exhibits, signatures, and Item 8 financial statements.
   - Evaluates text across 7 Loughran-McDonald sentiment categories: *Positive*, *Negative*, *Uncertainty*, *Litigious*, *Strong Modal*, *Weak Modal*, and *Constraining*.
   - Calculates the definitive Positive / Negative (P/N) sentiment ratio ($\text{Positive} \div \text{Negative}$).
4. **Yahoo Finance Stock Price Integration**:
   - Fetches historical adjusted split-adjusted closing stock prices matching each fiscal year-end period date.
5. **Eight-Tab Excel Workbook Generation**:
   - Outputs publication-ready `.xlsx` workbooks formatted with custom headers, borders, auto-fitted columns, and embedded dual-axis Chart tabs.
6. **Automated Auditing & Verification**:
   - Includes a standalone verification harness (`verify_outputs.py`) testing tab existence, 10-year coverage, balance sheet balancing, revenue/income alignment, and exact sentiment ratio calculations.
7. **Comparative Visualization Suite & Interactive Dashboard**:
   - Six high-resolution matplotlib charts (`build_comparison.py`) and a single self-contained responsive Chart.js dashboard (`make_dashboard.py`).

---

## Project Structure

```text
Tenk_project/
├── main.py                     # Main end-to-end pipeline CLI
├── verify_outputs.py           # Automated 5-point verification harness
├── build_comparison.py         # Matplotlib 6-chart comparative visualization generator
├── make_dashboard.py           # Interactive self-contained HTML dashboard builder
├── audit_mda.py                # Standalone MD&A extraction validator & word frequency auditor
├── requirements.txt            # Python dependencies
├── .gitignore                  # Git exclusion rules (data/, filings/, __pycache__/)
├── README.md                   # Project documentation
├── index.html                  # Self-contained web dashboard (project root)
├── dashboard/
│   └── index.html              # Dashboard source file
└── output/
    ├── AAPL_10K_Analysis.xlsx  # Apple 10-year 8-tab workbook
    ├── INTC_10K_Analysis.xlsx  # Intel 10-year 8-tab workbook
    └── figures/
        ├── aapl_sentiment_vs_stock_price.png  # Apple dual-axis sentiment vs price
        ├── intc_sentiment_vs_stock_price.png  # Intel dual-axis sentiment vs price
        ├── ratios_comparison.png              # 2x2 grid of 4 financial ratios
        ├── rd_intensity.png                   # Overlaid R&D expense / revenue %
        ├── gross_margin.png                   # Overlaid gross margin %
        └── capital_allocation.png             # Capex, buybacks, dividends vs revenue
```

*Note: Large raw filing files (`filings/`) and the 9 MB master dictionary (`data/`) are excluded from Git tracking via `.gitignore`.*

---

## Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/<your-username>/<repo-name>.git
   cd <repo-name>
   ```

2. **Create and activate a virtual environment (optional but recommended)**:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Loughran-McDonald Master Dictionary**:
   - Place `Loughran-McDonald_MasterDictionary_1993-2025.csv` into the `data/` directory.

---

## How to Run

### 1. Run Complete 10-K Pipeline for a Company
Pass the ticker symbol and the desired count of 10-K filings as command-line arguments:

```bash
# Run Apple for 10 fiscal years (FY2016 – FY2025)
python3 main.py AAPL 10

# Run Intel for 10 fiscal years (FY2016 – FY2025)
python3 main.py INTC 10
```

The pipeline will:
1. Resolve CIK and retrieve filing metadata from SEC EDGAR.
2. Download official Form 10-K annual reports into `filings/<TICKER>/`.
3. Extract Item 7 MD&A text into `filings/<TICKER>/<TICKER>_MD&A_*.txt`.
4. Parse Income Statement, Balance Sheet, and Cash Flows.
5. Score sentiment using Loughran-McDonald word lists.
6. Fetch fiscal period-end closing stock prices from Yahoo Finance.
7. Compute financial ratios and verify balance sheet equations.
8. Build `output/<TICKER>_10K_Analysis.xlsx` containing all 8 tabs.

---

### 2. Verify Output Workbooks
Run `verify_outputs.py` to perform an automated audit of the generated Excel workbooks:

```bash
# Verify Apple workbook
python3 verify_outputs.py output/AAPL_10K_Analysis.xlsx

# Verify Intel workbook
python3 verify_outputs.py output/INTC_10K_Analysis.xlsx
```

**Verification Checklist (5/5 PASS per workbook):**
- **Check 1**: All 8 required tabs exist (`README`, `Income Statement`, `Balance Sheet`, `Cash Flows`, `Sentiment Analysis`, `Stock Prices`, `Ratios`, `Chart`).
- **Check 2**: All tabs cover exactly 10 fiscal years (FY2016 through FY2025).
- **Check 3**: $\text{Assets} = \text{Liabilities} + \text{Equity}$ holds to the dollar (\$0 variance) across all 10 periods.
- **Check 4**: Income Statement Revenue and Net Income match Ratios tab inputs to the dollar.
- **Check 5**: Sentiment P/N ratio precisely equals positive count divided by negative count ($\text{Positive} \div \text{Negative}$).

---

### 3. Generate Comparative Matplotlib Charts
Generate the six high-resolution visual comparison figures:

```bash
python3 build_comparison.py
```
Output charts saved to `output/figures/`:
1. `aapl_sentiment_vs_stock_price.png`: Dual-axis plot of Apple stock price vs. MD&A sentiment P/N ratio.
2. `intc_sentiment_vs_stock_price.png`: Dual-axis plot of Intel stock price vs. MD&A sentiment P/N ratio.
3. `ratios_comparison.png`: 2×2 grid comparing Current Ratio, Debt/Equity, Asset Turnover, and Net Margin.
4. `rd_intensity.png`: Overlaid R&D expense as a percentage of revenue.
5. `gross_margin.png`: Overlaid gross profit margin illustrating the historical crossover.
6. `capital_allocation.png`: Side-by-side capital allocation (Capex, Buybacks, Dividends) vs. Revenue (\$B).

---

### 4. Build Interactive HTML Dashboard
Compile the self-contained single-page dashboard:

```bash
python3 make_dashboard.py
```
Outputs `dashboard/index.html` and updates root `index.html`.

---

## Accounting & Financial Methodology Notes

1. **Apple Cash Flow Dividends**:
   - In accordance with comprehensive cash distributions, Apple's dividend line item is strictly extracted from **Payments for dividends and dividend equivalents** in the Consolidated Statements of Cash Flows (\$15,421M in FY2025 down to \$12,150M in FY2016).
2. **Intel Balance Sheet Reconciliation & Non-Controlling Interests**:
   - In Intel's consolidated financial statements (e.g. FY2025), total equity comprises parent stockholders' equity (\$114,281M) and non-controlling interests (\$12,079M), totaling \$126,360M.
   - Total Liabilities are properly imputed as $\text{Total Assets} - \text{Total Consolidated Equity}$ (\$211,429M - \$126,360M = **\$85,069M**), ensuring non-controlling interests are strictly excluded from liabilities and that $\text{Assets} = \text{Liabilities} + \text{Equity}$ balances to the exact dollar.
3. **Intel Period Date Mapping**:
   - Intel's fiscal year ends on the last Saturday of December, resulting in January/February filings in the subsequent calendar year. Filing metadata labels are derived strictly from the period-of-report date (`reportDate`) rather than filing date (`filingDate`), ensuring FY2022 corresponds to report date 2022-12-31 (filed 2023-01-27).

---

## Ten-Year Comparative Highlights (FY2016 – FY2025)

| Metric | Apple Inc. (AAPL) | Intel Corporation (INTC) | Comparative Takeaway |
|:---|:---:|:---:|:---|
| **10-Year Revenue Growth** | **+93.0%** (\$215.6B $\to$ \$416.2B) | **-11.0%** (\$59.4B $\to$ \$52.9B) | Apple doubled top-line scale; Intel contracted top-line by 11%. |
| **10-Year Stock Return** | **+806.5%** (\$28.18 $\to$ \$255.46) | **-0.2%** (\$36.27 $\to$ \$36.20) | Apple delivered a 9.1x equity return; Intel remained flat over a decade. |
| **FY2025 Net Margin** | **26.9%** (\$112.0B profit) | **-0.5%** (-\$267M loss) | Apple sustained elite profitability; Intel stabilized after FY2024 (-35.3%). |
| **10-Year Avg R&D Intensity** | **6.5%** of Revenue | **22.9%** of Revenue | Intel spends ~3.5x more on R&D relative to revenue for manufacturing nodes. |
| **Gross Margin Trend** | Rose from **38.5%** $\to$ **46.9%** | Fell from **62.3%** $\to$ **34.8%** | Historical margin crossover driven by Apple Services & Intel fab costs. |
| **Capital Allocation Focus** | **Share Buybacks** (20–27% of rev/yr) | **Capex Fab Buildout** (reached 47.5% of rev) | Intel paused buybacks ($0) and suspended dividends ($0 in FY2025). |
| **MD&A Sentiment P/N Trend** | Stable: **0.27 – 0.59** | Collapsed: **1.67 $\to$ 0.23** | Intel negative terms climbed from 72 to 888 words per MD&A. |
