"""build_comparison.py.

Generates six comparison charts comparing Apple Inc. (AAPL) and Intel
Corporation (INTC) using data extracted from their respective 10-K analysis
workbooks:
- 'output/AAPL_10K_Analysis.xlsx'
- 'output/INTC_10K_Analysis.xlsx'

Charts generated in 'output/figures/':
1. aapl_sentiment_vs_stock_price.png: AAPL sentiment P/N vs. stock price
2. intc_sentiment_vs_stock_price.png: INTC sentiment P/N vs. stock price
3. ratios_comparison.png: Four financial ratios overlaid (2x2 grid)
4. rd_intensity.png: R&D intensity (R&D / Revenue %)
5. gross_margin.png: Gross margin (Gross Profit / Revenue %)
6. capital_allocation.png: Capex, Buybacks, Dividends vs Revenue
"""

import os

# Set writable cache dir for matplotlib in sandbox environments
os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib")

import matplotlib.pyplot as plt
import numpy as np
import openpyxl

# Configuration constants
OUTPUT_DIR = "output/figures"
AAPL_WORKBOOK = "output/AAPL_10K_Analysis.xlsx"
INTC_WORKBOOK = "output/INTC_10K_Analysis.xlsx"
YEARS = [f"FY{y}" for y in range(2016, 2026)]

# Consistent color palette
AAPL_COLOR = "#1f77b4"  # Steel blue
INTC_COLOR = "#ff7f0e"  # Vivid orange
CAPEX_COLOR = "#2b5c8f"  # Deep blue
BUYBACK_COLOR = "#2ca02c"  # Emerald green
DIVIDEND_COLOR = "#d62728"  # Crimson
REVENUE_COLOR = "#555555"  # Slate grey


def extract_company_data(ticker):
    """Extract financial metrics, ratios, sentiment, and capital allocation.

    Reads the company workbook and returns a comprehensive dictionary with
    chronologically ordered data for FY2016 through FY2025.
    """
    wb_path = f"output/{ticker}_10K_Analysis.xlsx"
    if not os.path.exists(wb_path):
        raise FileNotFoundError(f"Workbook not found: {wb_path}")

    wb = openpyxl.load_workbook(wb_path, data_only=True)

    # 1. Chart tab: Stock Price and MD&A P/N Ratio
    ws_chart = wb["Chart"]
    chart_rows = list(ws_chart.iter_rows(values_only=True))
    stock_prices = {}
    pn_ratios = {}
    for r in chart_rows[3:13]:
        if r[0]:
            fy = str(r[0]).strip()
            stock_prices[fy] = float(r[1]) if r[1] is not None else 0.0
            pn_ratios[fy] = float(r[2]) if r[2] is not None else 0.0

    # 2. Ratios tab: 4 Ratios + Revenue and Net Income
    ws_ratios = wb["Ratios"]
    ratios_rows = list(ws_ratios.iter_rows(values_only=True))
    current_ratio = {}
    liab_to_equity = {}
    asset_turnover = {}
    net_profit_margin = {}
    revenue = {}
    net_income = {}

    for r in ratios_rows[3:13]:
        if r[0]:
            fy = str(r[0]).strip()
            current_ratio[fy] = float(r[2]) if r[2] is not None else 0.0
            liab_to_equity[fy] = float(r[3]) if r[3] is not None else 0.0
            asset_turnover[fy] = float(r[4]) if r[4] is not None else 0.0
            net_profit_margin[fy] = float(r[5]) if r[5] is not None else 0.0
            revenue[fy] = float(r[11]) if r[11] is not None else 0.0
            net_income[fy] = float(r[12]) if r[12] is not None else 0.0

    # 3. Income Statement tab: R&D and Gross Profit / Margin
    # Columns 1 to 10 are in reverse chronological order (FY2025 -> FY2016)
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

    # 4. Cash Flows tab: Capex, Buybacks, and Dividends
    # Apple dividends strictly uses 'Payments for dividends and dividend
    # equivalents' as instructed by user specifications.
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

    # Ensure all years FY2016-FY2025 are populated with default 0.0
    for fy in YEARS:
        stock_prices.setdefault(fy, 0.0)
        pn_ratios.setdefault(fy, 0.0)
        current_ratio.setdefault(fy, 0.0)
        liab_to_equity.setdefault(fy, 0.0)
        asset_turnover.setdefault(fy, 0.0)
        net_profit_margin.setdefault(fy, 0.0)
        revenue.setdefault(fy, 0.0)
        net_income.setdefault(fy, 0.0)
        rd_expense.setdefault(fy, 0.0)
        gross_profit.setdefault(fy, 0.0)
        capex.setdefault(fy, 0.0)
        dividends.setdefault(fy, 0.0)
        buybacks.setdefault(fy, 0.0)

    # Compute intensity and margin percentages
    rd_intensity = {
        fy: (rd_expense[fy] / revenue[fy] * 100.0 if revenue[fy] else 0.0)
        for fy in YEARS
    }
    gross_margin_pct = {
        fy: (gross_profit[fy] / revenue[fy] * 100.0 if revenue[fy] else 0.0)
        for fy in YEARS
    }
    capex_pct = {
        fy: (capex[fy] / revenue[fy] * 100.0 if revenue[fy] else 0.0)
        for fy in YEARS
    }
    buybacks_pct = {
        fy: (buybacks[fy] / revenue[fy] * 100.0 if revenue[fy] else 0.0)
        for fy in YEARS
    }
    dividends_pct = {
        fy: (dividends[fy] / revenue[fy] * 100.0 if revenue[fy] else 0.0)
        for fy in YEARS
    }

    return {
        "ticker": ticker,
        "years": YEARS,
        "stock_prices": stock_prices,
        "pn_ratios": pn_ratios,
        "current_ratio": current_ratio,
        "liab_to_equity": liab_to_equity,
        "asset_turnover": asset_turnover,
        "net_profit_margin": net_profit_margin,
        "revenue": revenue,
        "net_income": net_income,
        "rd_expense": rd_expense,
        "rd_intensity": rd_intensity,
        "gross_profit": gross_profit,
        "gross_margin_pct": gross_margin_pct,
        "capex": capex,
        "capex_pct": capex_pct,
        "buybacks": buybacks,
        "buybacks_pct": buybacks_pct,
        "dividends": dividends,
        "dividends_pct": dividends_pct,
    }


def plot_sentiment_vs_stock_price(data, output_dir):
    """Plot dual-axis MD&A Sentiment P/N Ratio vs. Stock Price."""
    ticker = data["ticker"]
    years = data["years"]
    prices = [data["stock_prices"][fy] for fy in years]
    pn = [data["pn_ratios"][fy] for fy in years]

    fig, ax1 = plt.subplots(figsize=(10, 5.5))

    color_price = "#1f77b4" if ticker == "AAPL" else "#0071c5"
    color_pn = "#d62728" if ticker == "AAPL" else "#e65100"

    # Left axis: Stock Price ($)
    line1 = ax1.plot(
        years,
        prices,
        color=color_price,
        marker="o",
        linewidth=2.2,
        markersize=6,
        label=f"{ticker} Stock Price ($)",
    )
    ax1.set_xlabel("Fiscal Year", fontsize=11, fontweight="bold")
    ax1.set_ylabel(
        "Stock Price (USD $)",
        color=color_price,
        fontsize=11,
        fontweight="bold",
    )
    ax1.tick_params(axis="y", labelcolor=color_price)
    ax1.yaxis.set_major_formatter("${x:,.0f}")
    ax1.grid(True, linestyle="--", alpha=0.4)

    # Right axis: Sentiment P/N Ratio
    ax2 = ax1.twinx()
    line2 = ax2.plot(
        years,
        pn,
        color=color_pn,
        marker="s",
        linewidth=2.2,
        linestyle="--",
        markersize=6,
        label=f"{ticker} MD&A P/N Ratio",
    )
    ax2.set_ylabel(
        "MD&A Positive / Negative Ratio",
        color=color_pn,
        fontsize=11,
        fontweight="bold",
    )
    ax2.tick_params(axis="y", labelcolor=color_pn)
    ax2.grid(False)

    # Combined legend
    lines = line1 + line2
    labels = [l.get_label() for l in lines]
    ax1.legend(lines, labels, loc="upper left", framealpha=0.9)

    company_name = (
        "Apple Inc. (AAPL)" if ticker == "AAPL" else "Intel Corp. (INTC)"
    )
    plt.title(
        f"{company_name}: MD&A Sentiment P/N Ratio vs. Stock Price\n"
        f"(FY2016 – FY2025)",
        fontsize=13,
        fontweight="bold",
        pad=12,
    )
    fig.tight_layout()

    filename = (
        "aapl_sentiment_vs_stock_price.png"
        if ticker == "AAPL"
        else "intc_sentiment_vs_stock_price.png"
    )
    out_path = os.path.join(output_dir, filename)
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"Saved: {out_path}")


def plot_ratios_comparison(aapl, intc, output_dir):
    """Plot four financial ratios in 2x2 grid with both companies overlaid."""
    years = aapl["years"]
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))

    # Metric configs: (title, y_label, key, is_percent, ax)
    metrics = [
        (
            "Current Ratio (Liquidity)",
            "Current Ratio (x)",
            "current_ratio",
            False,
            axes[0, 0],
        ),
        (
            "Total-Liabilities-to-Equity (Solvency)",
            "Liabilities-to-Equity (x)",
            "liab_to_equity",
            False,
            axes[0, 1],
        ),
        (
            "Asset Turnover (Efficiency)",
            "Asset Turnover (x)",
            "asset_turnover",
            False,
            axes[1, 0],
        ),
        (
            "Net Profit Margin (Profitability)",
            "Net Profit Margin (%)",
            "net_profit_margin",
            True,
            axes[1, 1],
        ),
    ]

    for title, y_label, key, is_pct, ax in metrics:
        if is_pct:
            y_aapl = [aapl[key][fy] * 100.0 for fy in years]
            y_intc = [intc[key][fy] * 100.0 for fy in years]
        else:
            y_aapl = [aapl[key][fy] for fy in years]
            y_intc = [intc[key][fy] for fy in years]

        ax.plot(
            years,
            y_aapl,
            color=AAPL_COLOR,
            marker="o",
            linewidth=2.2,
            markersize=6,
            label="Apple (AAPL)",
        )
        ax.plot(
            years,
            y_intc,
            color=INTC_COLOR,
            marker="s",
            linewidth=2.2,
            linestyle="--",
            markersize=6,
            label="Intel (INTC)",
        )

        ax.set_title(title, fontsize=12, fontweight="bold", pad=8)
        ax.set_xlabel("Fiscal Year", fontsize=10)
        ax.set_ylabel(y_label, fontsize=10, fontweight="bold")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend(loc="best", framealpha=0.9)

        if is_pct:
            ax.yaxis.set_major_formatter("{x:.1f}%")
            ax.axhline(0, color="grey", linestyle=":", linewidth=0.8)

    fig.suptitle(
        "Financial Ratios Comparison: Apple vs. Intel (FY2016 – FY2025)",
        fontsize=15,
        fontweight="bold",
        y=0.99,
    )
    fig.tight_layout()

    out_path = os.path.join(output_dir, "ratios_comparison.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"Saved: {out_path}")


def plot_rd_intensity(aapl, intc, output_dir):
    """Plot R&D intensity (R&D Expense / Revenue %) for both companies."""
    years = aapl["years"]
    aapl_rd = [aapl["rd_intensity"][fy] for fy in years]
    intc_rd = [intc["rd_intensity"][fy] for fy in years]

    fig, ax = plt.subplots(figsize=(10, 5.5))

    ax.plot(
        years,
        aapl_rd,
        color=AAPL_COLOR,
        marker="o",
        linewidth=2.4,
        markersize=7,
        label="Apple (AAPL)",
    )
    ax.plot(
        years,
        intc_rd,
        color=INTC_COLOR,
        marker="s",
        linewidth=2.4,
        linestyle="--",
        markersize=7,
        label="Intel (INTC)",
    )

    # Annotate end points
    ax.annotate(
        f"{aapl_rd[-1]:.1f}%",
        xy=(years[-1], aapl_rd[-1]),
        xytext=(8, -2),
        textcoords="offset points",
        fontweight="bold",
        color=AAPL_COLOR,
    )
    ax.annotate(
        f"{intc_rd[-1]:.1f}%",
        xy=(years[-1], intc_rd[-1]),
        xytext=(8, -2),
        textcoords="offset points",
        fontweight="bold",
        color=INTC_COLOR,
    )

    ax.set_title(
        "R&D Intensity Comparison: Apple vs. Intel\n"
        "(Research & Development Expense ÷ Revenue, FY2016 – FY2025)",
        fontsize=13,
        fontweight="bold",
        pad=12,
    )
    ax.set_xlabel("Fiscal Year", fontsize=11, fontweight="bold")
    ax.set_ylabel(
        "R&D Intensity (% of Revenue)", fontsize=11, fontweight="bold"
    )
    ax.yaxis.set_major_formatter("{x:.1f}%")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend(loc="upper left", framealpha=0.9, fontsize=11)

    fig.tight_layout()
    out_path = os.path.join(output_dir, "rd_intensity.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"Saved: {out_path}")


def plot_gross_margin(aapl, intc, output_dir):
    """Plot Gross Margin (Gross Profit / Revenue %) for both companies."""
    years = aapl["years"]
    aapl_gm = [aapl["gross_margin_pct"][fy] for fy in years]
    intc_gm = [intc["gross_margin_pct"][fy] for fy in years]

    fig, ax = plt.subplots(figsize=(10, 5.5))

    ax.plot(
        years,
        aapl_gm,
        color=AAPL_COLOR,
        marker="o",
        linewidth=2.4,
        markersize=7,
        label="Apple (AAPL)",
    )
    ax.plot(
        years,
        intc_gm,
        color=INTC_COLOR,
        marker="s",
        linewidth=2.4,
        linestyle="--",
        markersize=7,
        label="Intel (INTC)",
    )

    # Annotate end points
    ax.annotate(
        f"{aapl_gm[-1]:.1f}%",
        xy=(years[-1], aapl_gm[-1]),
        xytext=(8, -2),
        textcoords="offset points",
        fontweight="bold",
        color=AAPL_COLOR,
    )
    ax.annotate(
        f"{intc_gm[-1]:.1f}%",
        xy=(years[-1], intc_gm[-1]),
        xytext=(8, -2),
        textcoords="offset points",
        fontweight="bold",
        color=INTC_COLOR,
    )

    ax.set_title(
        "Gross Margin Comparison: Apple vs. Intel\n"
        "(Gross Profit ÷ Revenue, FY2016 – FY2025)",
        fontsize=13,
        fontweight="bold",
        pad=12,
    )
    ax.set_xlabel("Fiscal Year", fontsize=11, fontweight="bold")
    ax.set_ylabel(
        "Gross Margin (% of Revenue)", fontsize=11, fontweight="bold"
    )
    ax.yaxis.set_major_formatter("{x:.1f}%")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend(loc="center left", framealpha=0.9, fontsize=11)

    fig.tight_layout()
    out_path = os.path.join(output_dir, "gross_margin.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"Saved: {out_path}")


def plot_capital_allocation(aapl, intc, output_dir):
    """Plot Capex, Buybacks, and Dividends vs. Revenue for both companies.

    Features a 2-panel figure (Left: Apple, Right: Intel) showing:
    - Grouped bars for Capex, Buybacks, and Dividends (% of Revenue)
    - Secondary axis line for Total Revenue ($ in Billions)
    - Strict adherence: Apple dividends uses 'Payments for dividends and
      dividend equivalents'.
    """
    years = aapl["years"]
    x = np.arange(len(years))
    width = 0.25

    fig, (ax_aapl, ax_intc) = plt.subplots(1, 2, figsize=(16, 6))

    panels = [
        (
            ax_aapl,
            aapl,
            "Apple (AAPL) — Capital Allocation vs. Revenue",
            [aapl["capex_pct"][fy] for fy in years],
            [aapl["buybacks_pct"][fy] for fy in years],
            [aapl["dividends_pct"][fy] for fy in years],
            [aapl["revenue"][fy] / 1000.0 for fy in years],  # in $B
            35.0,  # y-max left
            500.0,  # y-max right ($B)
        ),
        (
            ax_intc,
            intc,
            "Intel (INTC) — Capital Allocation vs. Revenue",
            [intc["capex_pct"][fy] for fy in years],
            [intc["buybacks_pct"][fy] for fy in years],
            [intc["dividends_pct"][fy] for fy in years],
            [intc["revenue"][fy] / 1000.0 for fy in years],  # in $B
            55.0,  # y-max left
            100.0,  # y-max right ($B)
        ),
    ]

    for (
        ax,
        data,
        title,
        capex_p,
        buyback_p,
        div_p,
        rev_b,
        y_max_left,
        y_max_right,
    ) in panels:
        # Primary axis: Bars for Capital Allocation (% of Revenue)
        b1 = ax.bar(
            x - width,
            capex_p,
            width,
            label="Capex (% of Rev)",
            color=CAPEX_COLOR,
            alpha=0.9,
        )
        b2 = ax.bar(
            x,
            buyback_p,
            width,
            label="Buybacks (% of Rev)",
            color=BUYBACK_COLOR,
            alpha=0.9,
        )
        b3 = ax.bar(
            x + width,
            div_p,
            width,
            label="Dividends (% of Rev)",
            color=DIVIDEND_COLOR,
            alpha=0.9,
        )

        ax.set_title(title, fontsize=12, fontweight="bold", pad=10)
        ax.set_xlabel("Fiscal Year", fontsize=10, fontweight="bold")
        ax.set_ylabel(
            "Capital Allocation (% of Revenue)",
            fontsize=10,
            fontweight="bold",
        )
        ax.set_xticks(x)
        ax.set_xticklabels(years, rotation=45, ha="right", fontsize=9)
        ax.yaxis.set_major_formatter("{x:.0f}%")
        ax.set_ylim(0, y_max_left)
        ax.grid(True, linestyle="--", alpha=0.3, axis="y")

        # Secondary axis: Revenue Line ($ Billions)
        ax_rev = ax.twinx()
        l_rev = ax_rev.plot(
            x,
            rev_b,
            color=REVENUE_COLOR,
            marker="D",
            linewidth=2.0,
            linestyle="-.",
            markersize=5,
            label="Revenue ($B)",
        )
        ax_rev.set_ylabel(
            "Revenue (USD $ Billions)",
            color=REVENUE_COLOR,
            fontsize=10,
            fontweight="bold",
        )
        ax_rev.tick_params(axis="y", labelcolor=REVENUE_COLOR)
        ax_rev.yaxis.set_major_formatter("${x:,.0f}B")
        ax_rev.set_ylim(0, y_max_right)
        ax_rev.grid(False)

        # Combined Legend
        bars = [b1, b2, b3] + l_rev
        labels = [b.get_label() for b in bars]
        ax.legend(
            bars, labels, loc="upper right", framealpha=0.88, fontsize=8.5
        )

    fig.suptitle(
        "Capital Allocation: Capex, Buybacks, & Dividends vs. Revenue "
        "(FY2016 – FY2025)",
        fontsize=14,
        fontweight="bold",
        y=0.98,
    )
    fig.tight_layout()

    out_path = os.path.join(output_dir, "capital_allocation.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"Saved: {out_path}")


def main():
    """Execute complete chart generation pipeline."""
    print("Generating comparison charts for AAPL and INTC...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    aapl_data = extract_company_data("AAPL")
    intc_data = extract_company_data("INTC")

    # Chart 1: AAPL Sentiment vs Stock Price
    plot_sentiment_vs_stock_price(aapl_data, OUTPUT_DIR)

    # Chart 2: INTC Sentiment vs Stock Price
    plot_sentiment_vs_stock_price(intc_data, OUTPUT_DIR)

    # Chart 3: Financial Ratios Comparison (2x2 grid)
    plot_ratios_comparison(aapl_data, intc_data, OUTPUT_DIR)

    # Chart 4: R&D Intensity
    plot_rd_intensity(aapl_data, intc_data, OUTPUT_DIR)

    # Chart 5: Gross Margin
    plot_gross_margin(aapl_data, intc_data, OUTPUT_DIR)

    # Chart 6: Capital Allocation vs Revenue
    plot_capital_allocation(aapl_data, intc_data, OUTPUT_DIR)

    print("\nAll 6 comparison charts generated in output/figures/!")


if __name__ == "__main__":
    main()
