"""make_dashboard.py.

Generates dashboard/index.html — a single self-contained HTML file with:
- All financial & sentiment data inlined (no external fetch/XHR required)
- Chart.js loaded from CDN (works locally via file:// and on GitHub Pages)
- KPI cards (10-yr revenue growth, 10-yr stock growth, FY2025 net margin)
- MD&A Sentiment P/N ratio vs. Stock Price charts for AAPL and INTC
- Overlaid ratio comparison charts (Current, Debt/Equity, Turnover, Margin)
- Overlaid R&D Intensity and Gross Margin charts
- Capital allocation charts (Capex, Buybacks, Dividends vs Revenue)
- Interactive Sentiment detail table with filtering and search
"""

import json
import os
import openpyxl

OUTPUT_DIR = "dashboard"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "index.html")
YEARS = [f"FY{y}" for y in range(2016, 2026)]


def extract_data(ticker):
    """Extract complete ten-year data from ticker's 10-K workbook."""
    wb_path = f"output/{ticker}_10K_Analysis.xlsx"
    if not os.path.exists(wb_path):
        raise FileNotFoundError(f"Workbook not found: {wb_path}")

    wb = openpyxl.load_workbook(wb_path, data_only=True)

    # 1. Chart tab: Stock Price and MD&A P/N Ratio
    ws_chart = wb["Chart"]
    stock_prices = {}
    pn_ratios = {}
    for r in list(ws_chart.iter_rows(values_only=True))[3:13]:
        if r[0]:
            fy = str(r[0]).strip()
            stock_prices[fy] = round(float(r[1]), 2) if r[1] else 0.0
            pn_ratios[fy] = round(float(r[2]), 4) if r[2] else 0.0

    # 2. Ratios tab
    ws_ratios = wb["Ratios"]
    ratios_dict = {}
    for r in list(ws_ratios.iter_rows(values_only=True))[3:13]:
        if r[0]:
            fy = str(r[0]).strip()
            ratios_dict[fy] = {
                "period_end": str(r[1]).strip() if r[1] else "",
                "current_ratio": round(float(r[2]), 3) if r[2] else 0.0,
                "liab_to_equity": round(float(r[3]), 3) if r[3] else 0.0,
                "asset_turnover": round(float(r[4]), 3) if r[4] else 0.0,
                "net_profit_margin": (
                    round(float(r[5]) * 100.0, 2) if r[5] else 0.0
                ),
                "revenue": float(r[11]) if r[11] else 0.0,
                "net_income": float(r[12]) if r[12] else 0.0,
            }

    # 3. Income Statement tab
    col_to_fy = {i + 1: f"FY{2025 - i}" for i in range(10)}
    ws_is = wb["Income Statement"]
    rd_expense = {}
    gross_profit = {}
    for r in list(ws_is.iter_rows(values_only=True))[3:]:
        label = str(r[0]).strip() if r[0] else ""
        if label == "Research and development":
            for col_idx in range(1, 11):
                fy = col_to_fy[col_idx]
                if r[col_idx] is not None and fy not in rd_expense:
                    rd_expense[fy] = float(r[col_idx])
        elif label in ["Gross margin", "Gross profit"]:
            for col_idx in range(1, 11):
                fy = col_to_fy[col_idx]
                if r[col_idx] is not None and fy not in gross_profit:
                    gross_profit[fy] = float(r[col_idx])

    # 4. Cash Flows tab
    ws_cf = wb["Cash Flows"]
    capex = {}
    dividends = {}
    buybacks = {}
    for r in list(ws_cf.iter_rows(values_only=True))[3:]:
        label = str(r[0]).strip() if r[0] else ""
        for col_idx in range(1, 11):
            fy = col_to_fy[col_idx]
            val = r[col_idx]
            if val is not None:
                try:
                    amt = abs(float(val))
                except ValueError:
                    continue
                if ticker == "AAPL":
                    if (
                        label
                        == "Payments for acquisition of property, "
                        "plant and equipment"
                    ):
                        capex[fy] = amt
                    elif (
                        label
                        == "Payments for dividends and dividend equivalents"
                    ):
                        dividends[fy] = amt
                    elif label == "Repurchases of common stock":
                        buybacks[fy] = amt
                else:
                    if label in [
                        "Additions to property, plant and equipment",
                        "Additions to property, plant, and equipment",
                    ]:
                        if fy not in capex:
                            capex[fy] = amt
                    elif label == "Payment of dividends to stockholders":
                        dividends[fy] = amt
                    elif label == "Repurchase of common stock":
                        buybacks[fy] = amt

    # 5. Sentiment Analysis tab
    ws_sent = wb["Sentiment Analysis"]
    sentiment_rows = []
    for r in list(ws_sent.iter_rows(values_only=True))[3:13]:
        if r[0]:
            fy = str(r[0]).strip()
            sentiment_rows.append(
                {
                    "company": ticker,
                    "fiscal_year": fy,
                    "filing_date": str(r[1])[:10] if r[1] else "",
                    "report_date": str(r[2])[:10] if r[2] else "",
                    "total_words": int(r[3]) if r[3] else 0,
                    "positive": int(r[4]) if r[4] else 0,
                    "pos_pct": (
                        round(float(r[5]) * 100.0, 3) if r[5] else 0.0
                    ),
                    "negative": int(r[6]) if r[6] else 0,
                    "neg_pct": (
                        round(float(r[7]) * 100.0, 3) if r[7] else 0.0
                    ),
                    "pn_ratio": round(float(r[8]), 4) if r[8] else 0.0,
                    "uncertainty": int(r[9]) if r[9] else 0,
                    "litigious": int(r[11]) if r[11] else 0,
                    "stock_price": stock_prices.get(fy, 0.0),
                }
            )

    sentiment_rows = sorted(sentiment_rows, key=lambda x: x["fiscal_year"])

    # Build metric series
    rev_series = [ratios_dict[fy]["revenue"] for fy in YEARS]
    ni_series = [ratios_dict[fy]["net_income"] for fy in YEARS]
    cr_series = [ratios_dict[fy]["current_ratio"] for fy in YEARS]
    tle_series = [ratios_dict[fy]["liab_to_equity"] for fy in YEARS]
    at_series = [ratios_dict[fy]["asset_turnover"] for fy in YEARS]
    npm_series = [ratios_dict[fy]["net_profit_margin"] for fy in YEARS]
    price_series = [stock_prices.get(fy, 0.0) for fy in YEARS]
    pn_series = [pn_ratios.get(fy, 0.0) for fy in YEARS]

    rd_intensity = [
        round(
            rd_expense.get(fy, 0.0) / ratios_dict[fy]["revenue"] * 100.0, 2
        )
        if ratios_dict[fy]["revenue"]
        else 0.0
        for fy in YEARS
    ]
    gross_margin = [
        round(
            gross_profit.get(fy, 0.0) / ratios_dict[fy]["revenue"] * 100.0, 2
        )
        if ratios_dict[fy]["revenue"]
        else 0.0
        for fy in YEARS
    ]

    capex_pct = [
        round(capex.get(fy, 0.0) / ratios_dict[fy]["revenue"] * 100.0, 2)
        if ratios_dict[fy]["revenue"]
        else 0.0
        for fy in YEARS
    ]
    buybacks_pct = [
        round(buybacks.get(fy, 0.0) / ratios_dict[fy]["revenue"] * 100.0, 2)
        if ratios_dict[fy]["revenue"]
        else 0.0
        for fy in YEARS
    ]
    dividends_pct = [
        round(dividends.get(fy, 0.0) / ratios_dict[fy]["revenue"] * 100.0, 2)
        if ratios_dict[fy]["revenue"]
        else 0.0
        for fy in YEARS
    ]
    revenue_b = [
        round(ratios_dict[fy]["revenue"] / 1000.0, 2) for fy in YEARS
    ]

    # KPIs
    rev_start, rev_end = rev_series[0], rev_series[-1]
    rev_growth_pct = round((rev_end - rev_start) / rev_start * 100.0, 1)
    rev_cagr = round(((rev_end / rev_start) ** (1.0 / 9.0) - 1.0) * 100.0, 2)

    price_start, price_end = price_series[0], price_series[-1]
    price_growth_pct = round(
        (price_end - price_start) / price_start * 100.0, 1
    )
    price_cagr = round(
        ((price_end / price_start) ** (1.0 / 9.0) - 1.0) * 100.0, 2
    )

    latest_npm = npm_series[-1]
    avg_rd_int = round(sum(rd_intensity) / len(rd_intensity), 2)

    return {
        "ticker": ticker,
        "years": YEARS,
        "revenue": rev_series,
        "revenue_b": revenue_b,
        "net_income": ni_series,
        "current_ratio": cr_series,
        "liab_to_equity": tle_series,
        "asset_turnover": at_series,
        "net_profit_margin": npm_series,
        "stock_prices": price_series,
        "pn_ratios": pn_series,
        "rd_intensity": rd_intensity,
        "gross_margin": gross_margin,
        "capex_pct": capex_pct,
        "buybacks_pct": buybacks_pct,
        "dividends_pct": dividends_pct,
        "sentiment_rows": sentiment_rows,
        "kpi": {
            "rev_growth_pct": rev_growth_pct,
            "rev_cagr": rev_cagr,
            "price_growth_pct": price_growth_pct,
            "price_cagr": price_cagr,
            "latest_npm": latest_npm,
            "avg_rd_int": avg_rd_int,
            "rev_start_b": round(rev_start / 1000.0, 1),
            "rev_end_b": round(rev_end / 1000.0, 1),
            "price_start": price_start,
            "price_end": price_end,
        },
    }


def build_html(aapl, intc):
    """Generate self-contained HTML dashboard with embedded JSON data."""
    dashboard_data = {
        "years": YEARS,
        "aapl": aapl,
        "intc": intc,
    }
    json_str = json.dumps(dashboard_data, indent=2)

    all_sentiment_rows = aapl["sentiment_rows"] + intc["sentiment_rows"]
    all_sentiment_rows.sort(
        key=lambda r: (r["fiscal_year"], r["company"]), reverse=True
    )

    table_rows_html = []
    for r in all_sentiment_rows:
        badge_cls = "badge-aapl" if r["company"] == "AAPL" else "badge-intc"
        pos_str = f"{r['positive']:,} ({r['pos_pct']:.2f}%)"
        neg_str = f"{r['negative']:,} ({r['neg_pct']:.2f}%)"
        price_str = f"${r['stock_price']:,.2f}"
        table_rows_html.append(
            f"""<tr data-company="{r['company']}">
  <td><span class="badge {badge_cls}">{r['company']}</span></td>
  <td class="font-bold">{r['fiscal_year']}</td>
  <td>{r['report_date']}</td>
  <td>{r['filing_date']}</td>
  <td class="text-right">{price_str}</td>
  <td class="text-right">{r['total_words']:,}</td>
  <td class="text-right text-green">{pos_str}</td>
  <td class="text-right text-red">{neg_str}</td>
  <td class="text-right font-bold">{r['pn_ratio']:.4f}</td>
  <td class="text-right">{r['uncertainty']:,}</td>
  <td class="text-right">{r['litigious']:,}</td>
</tr>"""
        )
    table_body = "\n".join(table_rows_html)

    # Build KPI subtitle strings
    aapl_rev_sub = (
        f"${aapl['kpi']['rev_start_b']}B &rarr; "
        f"${aapl['kpi']['rev_end_b']}B "
        f"({aapl['kpi']['rev_cagr']}%/yr)"
    )
    intc_rev_sub = (
        f"${intc['kpi']['rev_start_b']}B &rarr; "
        f"${intc['kpi']['rev_end_b']}B "
        f"({intc['kpi']['rev_cagr']}%/yr)"
    )
    aapl_price_sub = (
        f"${aapl['kpi']['price_start']:.2f} &rarr; "
        f"${aapl['kpi']['price_end']:.2f} "
        f"({aapl['kpi']['price_cagr']}%/yr)"
    )
    intc_price_sub = (
        f"${intc['kpi']['price_start']:.2f} &rarr; "
        f"${intc['kpi']['price_end']:.2f} "
        f"({intc['kpi']['price_cagr']}%/yr)"
    )

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AAPL vs. INTC — 10-Year 10-K Dashboard</title>
  <!-- Chart.js from CDN for local & web compatibility -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root {{
      --primary: #0f172a;
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --border: #e2e8f0;
      --text: #1e293b;
      --text-muted: #64748b;
      --aapl: #0071e3;
      --aapl-light: rgba(0, 113, 227, 0.12);
      --intc: #ff7f0e;
      --intc-light: rgba(255, 127, 14, 0.12);
      --green: #10b981;
      --red: #ef4444;
      --radius: 12px;
      --shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05),
        0 2px 4px -2px rgba(0, 0, 0, 0.05);
    }}
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
        Roboto, Helvetica, Arial, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding: 24px;
    }}
    .container {{
      max-width: 1400px;
      margin: 0 auto;
    }}
    header {{
      margin-bottom: 24px;
      padding-bottom: 16px;
      border-bottom: 1px solid var(--border);
    }}
    h1 {{
      font-size: 1.75rem;
      font-weight: 700;
      color: var(--primary);
      margin-bottom: 6px;
    }}
    .subtitle {{
      color: var(--text-muted);
      font-size: 0.95rem;
    }}
    /* KPI Cards */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
      margin-bottom: 28px;
    }}
    .kpi-card {{
      background: var(--card-bg);
      border-radius: var(--radius);
      padding: 20px;
      box-shadow: var(--shadow);
      border: 1px solid var(--border);
      transition: transform 0.15s ease;
    }}
    .kpi-card:hover {{
      transform: translateY(-2px);
    }}
    .kpi-title {{
      font-size: 0.85rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-bottom: 12px;
    }}
    .kpi-values {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
    }}
    .kpi-company {{
      flex: 1;
    }}
    .kpi-company-label {{
      font-size: 0.8rem;
      font-weight: 700;
      margin-bottom: 2px;
    }}
    .aapl-label {{ color: var(--aapl); }}
    .intc-label {{ color: var(--intc); }}
    .kpi-num {{
      font-size: 1.35rem;
      font-weight: 800;
      line-height: 1.2;
    }}
    .kpi-sub {{
      font-size: 0.75rem;
      color: var(--text-muted);
      margin-top: 2px;
    }}
    /* Charts Grid */
    .section-title {{
      font-size: 1.25rem;
      font-weight: 700;
      margin: 28px 0 16px;
      color: var(--primary);
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .charts-grid-2 {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(600px, 1fr));
      gap: 20px;
      margin-bottom: 24px;
    }}
    .chart-card {{
      background: var(--card-bg);
      border-radius: var(--radius);
      padding: 20px;
      box-shadow: var(--shadow);
      border: 1px solid var(--border);
    }}
    .chart-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 14px;
    }}
    .chart-title {{
      font-size: 1rem;
      font-weight: 700;
      color: var(--primary);
    }}
    .chart-container {{
      position: relative;
      height: 320px;
      width: 100%;
    }}
    /* Table Section */
    .table-card {{
      background: var(--card-bg);
      border-radius: var(--radius);
      padding: 20px;
      box-shadow: var(--shadow);
      border: 1px solid var(--border);
      margin-top: 28px;
      overflow-x: auto;
    }}
    .table-controls {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 16px;
    }}
    .filter-pills {{
      display: flex;
      gap: 8px;
    }}
    .filter-btn {{
      background: var(--bg);
      border: 1px solid var(--border);
      padding: 6px 14px;
      border-radius: 20px;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
      color: var(--text-muted);
      transition: all 0.15s ease;
    }}
    .filter-btn.active {{
      background: var(--primary);
      color: #fff;
      border-color: var(--primary);
    }}
    .search-input {{
      padding: 6px 14px;
      border-radius: 20px;
      border: 1px solid var(--border);
      font-size: 0.85rem;
      outline: none;
      min-width: 220px;
    }}
    .search-input:focus {{
      border-color: var(--aapl);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.85rem;
    }}
    th {{
      text-align: left;
      padding: 10px 12px;
      background: var(--bg);
      color: var(--text-muted);
      font-weight: 600;
      border-bottom: 2px solid var(--border);
      white-space: nowrap;
    }}
    td {{
      padding: 10px 12px;
      border-bottom: 1px solid var(--border);
      white-space: nowrap;
    }}
    tr:hover td {{
      background: rgba(241, 245, 249, 0.6);
    }}
    .text-right {{ text-align: right; }}
    .font-bold {{ font-weight: 700; }}
    .text-green {{ color: var(--green); }}
    .text-red {{ color: var(--red); }}
    .badge {{
      display: inline-block;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 0.75rem;
      font-weight: 700;
    }}
    .badge-aapl {{
      background: var(--aapl-light);
      color: var(--aapl);
    }}
    .badge-intc {{
      background: var(--intc-light);
      color: var(--intc);
    }}
    footer {{
      margin-top: 40px;
      padding-top: 20px;
      border-top: 1px solid var(--border);
      font-size: 0.8rem;
      color: var(--text-muted);
      text-align: center;
    }}
    @media (max-width: 768px) {{
      body {{ padding: 12px; }}
      .charts-grid-2 {{ grid-template-columns: 1fr; }}
      .chart-container {{ height: 260px; }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>Apple Inc. (AAPL) vs. Intel Corporation (INTC)</h1>
      <p class="subtitle">
        Ten-Year Comparative Financial &amp; MD&amp;A Sentiment Dashboard
        (FY2016 – FY2025) &bull; Verified SEC 10-K Filings
      </p>
    </header>

    <!-- KPI Cards -->
    <div class="kpi-grid">
      <!-- KPI 1 -->
      <div class="kpi-card">
        <div class="kpi-title">10-Year Revenue Growth</div>
        <div class="kpi-values">
          <div class="kpi-company">
            <div class="kpi-company-label aapl-label">AAPL</div>
            <div class="kpi-num text-green">
              +{aapl['kpi']['rev_growth_pct']}%
            </div>
            <div class="kpi-sub">{aapl_rev_sub}</div>
          </div>
          <div class="kpi-company">
            <div class="kpi-company-label intc-label">INTC</div>
            <div class="kpi-num text-red">
              {intc['kpi']['rev_growth_pct']}%
            </div>
            <div class="kpi-sub">{intc_rev_sub}</div>
          </div>
        </div>
      </div>

      <!-- KPI 2 -->
      <div class="kpi-card">
        <div class="kpi-title">10-Year Stock Return</div>
        <div class="kpi-values">
          <div class="kpi-company">
            <div class="kpi-company-label aapl-label">AAPL</div>
            <div class="kpi-num text-green">
              +{aapl['kpi']['price_growth_pct']}%
            </div>
            <div class="kpi-sub">{aapl_price_sub}</div>
          </div>
          <div class="kpi-company">
            <div class="kpi-company-label intc-label">INTC</div>
            <div class="kpi-num">{intc['kpi']['price_growth_pct']}%</div>
            <div class="kpi-sub">{intc_price_sub}</div>
          </div>
        </div>
      </div>

      <!-- KPI 3 -->
      <div class="kpi-card">
        <div class="kpi-title">FY2025 Net Profit Margin</div>
        <div class="kpi-values">
          <div class="kpi-company">
            <div class="kpi-company-label aapl-label">AAPL</div>
            <div class="kpi-num text-green">{aapl['kpi']['latest_npm']}%</div>
            <div class="kpi-sub">$112.0B Net Income</div>
          </div>
          <div class="kpi-company">
            <div class="kpi-company-label intc-label">INTC</div>
            <div class="kpi-num text-red">{intc['kpi']['latest_npm']}%</div>
            <div class="kpi-sub">-$267M Net Loss</div>
          </div>
        </div>
      </div>

      <!-- KPI 4 -->
      <div class="kpi-card">
        <div class="kpi-title">10-Year Avg R&amp;D Intensity</div>
        <div class="kpi-values">
          <div class="kpi-company">
            <div class="kpi-company-label aapl-label">AAPL</div>
            <div class="kpi-num">{aapl['kpi']['avg_rd_int']}%</div>
            <div class="kpi-sub">$194.8B Total Invested</div>
          </div>
          <div class="kpi-company">
            <div class="kpi-company-label intc-label">INTC</div>
            <div class="kpi-num">{intc['kpi']['avg_rd_int']}%</div>
            <div class="kpi-sub">$148.8B Total Invested</div>
          </div>
        </div>
      </div>
    </div>

    <!-- Section 1: Sentiment vs Stock Price -->
    <h2 class="section-title">
      1. MD&amp;A Sentiment P/N Ratio vs. Stock Price
    </h2>
    <div class="charts-grid-2">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">
            Apple (AAPL) — Sentiment P/N vs. Stock Price
          </div>
        </div>
        <div class="chart-container">
          <canvas id="aaplSentimentChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">
            Intel (INTC) — Sentiment P/N vs. Stock Price
          </div>
        </div>
        <div class="chart-container">
          <canvas id="intcSentimentChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Section 2: Overlaid Financial Ratios -->
    <h2 class="section-title">2. Ten-Year Financial Ratios Comparison</h2>
    <div class="charts-grid-2">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Current Ratio (Liquidity)</div>
        </div>
        <div class="chart-container">
          <canvas id="crChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Total-Liabilities-to-Equity (Solvency)</div>
        </div>
        <div class="chart-container">
          <canvas id="tleChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Asset Turnover (Asset Efficiency)</div>
        </div>
        <div class="chart-container">
          <canvas id="atChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Net Profit Margin (%) (Profitability)</div>
        </div>
        <div class="chart-container">
          <canvas id="npmChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Section 3: Operating Performance -->
    <h2 class="section-title">
      3. Operating Performance &amp; Margin Evolution
    </h2>
    <div class="charts-grid-2">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">R&amp;D Intensity (% of Revenue)</div>
        </div>
        <div class="chart-container">
          <canvas id="rdChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Gross Margin (% of Revenue)</div>
        </div>
        <div class="chart-container">
          <canvas id="gmChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Section 4: Capital Allocation -->
    <h2 class="section-title">4. Capital Allocation vs. Revenue</h2>
    <div class="charts-grid-2">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">
            Apple (AAPL) — Capex, Buybacks &amp; Dividends vs. Revenue
          </div>
        </div>
        <div class="chart-container">
          <canvas id="aaplCapAllocChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">
            Intel (INTC) — Capex, Buybacks &amp; Dividends vs. Revenue
          </div>
        </div>
        <div class="chart-container">
          <canvas id="intcCapAllocChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Section 5: Sentiment Detail Table -->
    <div class="table-card">
      <div class="table-controls">
        <div>
          <h2 class="chart-title" style="margin-bottom: 4px;">
            MD&amp;A Sentiment Analysis Detail (FY2016 – FY2025)
          </h2>
          <p class="subtitle">
            Loughran-McDonald Financial Sentiment Lexicon Word Counts
          </p>
        </div>
        <div style="display: flex; gap: 12px; align-items: center;">
          <div class="filter-pills">
            <button class="filter-btn active" onclick="filterCompany('ALL')">
              All (20)
            </button>
            <button class="filter-btn" onclick="filterCompany('AAPL')">
              AAPL
            </button>
            <button class="filter-btn" onclick="filterCompany('INTC')">
              INTC
            </button>
          </div>
          <input
            type="text"
            id="tableSearch"
            class="search-input"
            placeholder="Search year or keyword..."
            onkeyup="filterTable()"
          >
        </div>
      </div>
      <table id="sentimentTable">
        <thead>
          <tr>
            <th>Company</th>
            <th>Fiscal Year</th>
            <th>Report Date</th>
            <th>Filing Date</th>
            <th class="text-right">Stock Price ($)</th>
            <th class="text-right">MD&amp;A Words</th>
            <th class="text-right">Positive</th>
            <th class="text-right">Negative</th>
            <th class="text-right">P/N Ratio</th>
            <th class="text-right">Uncertainty</th>
            <th class="text-right">Litigious</th>
          </tr>
        </thead>
        <tbody>
          {table_body}
        </tbody>
      </table>
    </div>

    <footer>
      <p>
        Data derived directly from audited SEC Form 10-K Annual Reports
        (FY2016 – FY2025) for Apple Inc. and Intel Corporation.
      </p>
      <p style="margin-top: 4px;">
        Notes: Apple dividends strictly reflect <em>Payments for dividends
        and dividend equivalents</em> from Consolidated Cash Flows.
        Intel Total Liabilities properly excludes non-controlling interests.
      </p>
    </footer>
  </div>

  <!-- Inlined Financial & Sentiment Data -->
  <script>
    const DATA = {json_str};

    // Shared Chart.js defaults
    Chart.defaults.font.family = (
      '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    );
    Chart.defaults.color = '#64748b';
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 8;

    const years = DATA.years;
    const aapl = DATA.aapl;
    const intc = DATA.intc;

    // 1. AAPL Sentiment vs Stock Price
    new Chart(document.getElementById('aaplSentimentChart'), {{
      type: 'line',
      data: {{
        labels: years,
        datasets: [
          {{
            label: 'AAPL Stock Price ($)',
            data: aapl.stock_prices,
            borderColor: '#0071e3',
            backgroundColor: '#0071e3',
            yAxisID: 'yPrice',
            borderWidth: 2.5,
            tension: 0.2,
            pointRadius: 4,
          }},
          {{
            label: 'MD&A P/N Ratio',
            data: aapl.pn_ratios,
            borderColor: '#ef4444',
            backgroundColor: '#ef4444',
            yAxisID: 'yRatio',
            borderWidth: 2,
            borderDash: [5, 5],
            tension: 0.2,
            pointRadius: 4,
          }}
        ]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: 'index', intersect: false }},
        scales: {{
          yPrice: {{
            type: 'linear',
            position: 'left',
            title: {{ display: true, text: 'Stock Price ($)' }},
            ticks: {{ callback: v => '$' + v }}
          }},
          yRatio: {{
            type: 'linear',
            position: 'right',
            title: {{ display: true, text: 'MD&A P/N Ratio' }},
            grid: {{ drawOnChartArea: false }}
          }}
        }}
      }}
    }});

    // 2. INTC Sentiment vs Stock Price
    new Chart(document.getElementById('intcSentimentChart'), {{
      type: 'line',
      data: {{
        labels: years,
        datasets: [
          {{
            label: 'INTC Stock Price ($)',
            data: intc.stock_prices,
            borderColor: '#ff7f0e',
            backgroundColor: '#ff7f0e',
            yAxisID: 'yPrice',
            borderWidth: 2.5,
            tension: 0.2,
            pointRadius: 4,
          }},
          {{
            label: 'MD&A P/N Ratio',
            data: intc.pn_ratios,
            borderColor: '#ef4444',
            backgroundColor: '#ef4444',
            yAxisID: 'yRatio',
            borderWidth: 2,
            borderDash: [5, 5],
            tension: 0.2,
            pointRadius: 4,
          }}
        ]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: 'index', intersect: false }},
        scales: {{
          yPrice: {{
            type: 'linear',
            position: 'left',
            title: {{ display: true, text: 'Stock Price ($)' }},
            ticks: {{ callback: v => '$' + v }}
          }},
          yRatio: {{
            type: 'linear',
            position: 'right',
            title: {{ display: true, text: 'MD&A P/N Ratio' }},
            grid: {{ drawOnChartArea: false }}
          }}
        }}
      }}
    }});

    // Helper for Overlaid Ratio Charts
    function createRatioChart(
      canvasId, labelAapl, dataAapl, labelIntc, dataIntc, yTitle, isPct
    ) {{
      new Chart(document.getElementById(canvasId), {{
        type: 'line',
        data: {{
          labels: years,
          datasets: [
            {{
              label: labelAapl,
              data: dataAapl,
              borderColor: '#0071e3',
              backgroundColor: '#0071e3',
              borderWidth: 2.5,
              tension: 0.2,
              pointRadius: 4,
            }},
            {{
              label: labelIntc,
              data: dataIntc,
              borderColor: '#ff7f0e',
              backgroundColor: '#ff7f0e',
              borderWidth: 2.5,
              borderDash: [4, 4],
              tension: 0.2,
              pointRadius: 4,
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          scales: {{
            y: {{
              title: {{ display: true, text: yTitle }},
              ticks: {{
                callback: v => isPct ? v + '%' : v
              }}
            }}
          }}
        }}
      }});
    }}

    createRatioChart(
      'crChart', 'AAPL Current Ratio', aapl.current_ratio,
      'INTC Current Ratio', intc.current_ratio,
      'Current Ratio (x)', false
    );
    createRatioChart(
      'tleChart', 'AAPL Liab/Equity', aapl.liab_to_equity,
      'INTC Liab/Equity', intc.liab_to_equity,
      'Liabilities-to-Equity (x)', false
    );
    createRatioChart(
      'atChart', 'AAPL Asset Turnover', aapl.asset_turnover,
      'INTC Asset Turnover', intc.asset_turnover,
      'Asset Turnover (x)', false
    );
    createRatioChart(
      'npmChart', 'AAPL Net Margin', aapl.net_profit_margin,
      'INTC Net Margin', intc.net_profit_margin,
      'Net Profit Margin (%)', true
    );

    createRatioChart(
      'rdChart', 'AAPL R&D Intensity', aapl.rd_intensity,
      'INTC R&D Intensity', intc.rd_intensity,
      'R&D Expense (% of Rev)', true
    );
    createRatioChart(
      'gmChart', 'AAPL Gross Margin', aapl.gross_margin,
      'INTC Gross Margin', intc.gross_margin,
      'Gross Margin (% of Rev)', true
    );

    // Capital Allocation Charts (Capex, Buybacks, Dividends vs Revenue)
    function createCapAllocChart(canvasId, compData, isAapl) {{
      new Chart(document.getElementById(canvasId), {{
        data: {{
          labels: years,
          datasets: [
            {{
              type: 'bar',
              label: 'Capex (% of Rev)',
              data: compData.capex_pct,
              backgroundColor: '#2b5c8f',
              yAxisID: 'yPct',
              borderRadius: 4,
            }},
            {{
              type: 'bar',
              label: 'Buybacks (% of Rev)',
              data: compData.buybacks_pct,
              backgroundColor: '#10b981',
              yAxisID: 'yPct',
              borderRadius: 4,
            }},
            {{
              type: 'bar',
              label: 'Dividends (% of Rev)',
              data: compData.dividends_pct,
              backgroundColor: '#ef4444',
              yAxisID: 'yPct',
              borderRadius: 4,
            }},
            {{
              type: 'line',
              label: 'Revenue ($B)',
              data: compData.revenue_b,
              borderColor: '#64748b',
              backgroundColor: '#64748b',
              yAxisID: 'yRev',
              borderWidth: 2,
              borderDash: [4, 4],
              pointStyle: 'rectRot',
              pointRadius: 5,
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          scales: {{
            yPct: {{
              type: 'linear',
              position: 'left',
              title: {{ display: true, text: 'Allocation (% of Revenue)' }},
              ticks: {{ callback: v => v + '%' }},
              suggestedMax: isAapl ? 35 : 55,
            }},
            yRev: {{
              type: 'linear',
              position: 'right',
              title: {{ display: true, text: 'Revenue (USD $B)' }},
              ticks: {{ callback: v => '$' + v + 'B' }},
              grid: {{ drawOnChartArea: false }},
              suggestedMax: isAapl ? 500 : 100,
            }}
          }}
        }}
      }});
    }}

    createCapAllocChart('aaplCapAllocChart', aapl, true);
    createCapAllocChart('intcCapAllocChart', intc, false);

    // Table Filtering Functions
    let currentFilter = 'ALL';

    function filterCompany(comp) {{
      currentFilter = comp;
      document.querySelectorAll('.filter-btn').forEach(btn => {{
        btn.classList.toggle('active', btn.textContent.includes(comp));
      }});
      filterTable();
    }}

    function filterTable() {{
      const query = document.getElementById('tableSearch').value.toLowerCase();
      const rows = document.querySelectorAll('#sentimentTable tbody tr');

      rows.forEach(row => {{
        const comp = row.getAttribute('data-company');
        const text = row.textContent.toLowerCase();
        const matchesComp = (
          currentFilter === 'ALL' || comp === currentFilter
        );
        const matchesQuery = text.includes(query);

        row.style.display = (matchesComp && matchesQuery) ? '' : 'none';
      }});
    }}
  </script>
</body>
</html>
"""
    return html_content


def main():
    """Extract data, build dashboard HTML, and save to dashboard/index.html."""
    print("Extracting data for AAPL and INTC...")
    aapl_data = extract_data("AAPL")
    intc_data = extract_data("INTC")

    print("Generating self-contained HTML dashboard...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    html_output = build_html(aapl_data, intc_data)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html_output)

    file_size_kb = os.path.getsize(OUTPUT_FILE) / 1024.0
    print(f"Successfully generated: {OUTPUT_FILE} ({file_size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
