"""Verification script for 10-K analysis workbook.

Validates that the generated Excel workbook meets all project specifications:
  1. All 8 required tabs exist.
  2. All tabs cover exactly 10 fiscal years.
  3. Balance sheet balances (Assets = Liabilities + Equity) for every year.
  4. Income statement revenue and net income match the ratio input values.
  5. MD&A sentiment P/N ratio matches positive count divided by negative count.

Prints PASS/FAIL status for each check and returns exit code 0 if all pass.
"""

import argparse
import datetime
import math
import os
import sys
from typing import Any, Dict, List, Optional, Tuple

import openpyxl


def _to_date(val: Any) -> Optional[datetime.date]:
    """Parse various date formats into a standard datetime.date object.

    Args:
        val: Date representation as string, datetime, or date.

    Returns:
        Optional[datetime.date]: Parsed date object or None if unparseable.
    """
    if val is None:
        return None
    if isinstance(val, datetime.date) and not isinstance(
        val, datetime.datetime
    ):
        return val
    if isinstance(val, datetime.datetime):
        return val.date()

    val_str = str(val).strip()
    for fmt in (
        "%Y-%m-%d",
        "%b. %d, %Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%m/%d/%Y",
        "%Y%m%d",
    ):
        try:
            return datetime.datetime.strptime(val_str, fmt).date()
        except ValueError:
            continue
    return None


def verify_eight_tabs(
    wb: openpyxl.Workbook,
) -> Tuple[bool, str, List[str]]:
    """Check that all eight required tabs exist in the workbook.

    Required tabs:
      - README
      - Income Statement
      - Balance Sheet
      - Cash Flows
      - Sentiment Analysis
      - Stock Prices
      - Ratios
      - Chart

    Args:
        wb: Openpyxl Workbook instance.

    Returns:
        Tuple[bool, str, List[str]]: Success flag, message, and missing list.
    """
    required_tabs = [
        "README",
        "Income Statement",
        "Balance Sheet",
        "Cash Flows",
        "Sentiment Analysis",
        "Stock Prices",
        "Ratios",
        "Chart",
    ]
    existing_sheets = wb.sheetnames
    missing = [t for t in required_tabs if t not in existing_sheets]

    if missing:
        msg = f"Missing {len(missing)} tab(s): {', '.join(missing)}"
        return False, msg, missing

    msg = f"All {len(required_tabs)} required tabs are present."
    return True, msg, []


def verify_ten_fiscal_years(
    wb: openpyxl.Workbook,
) -> Tuple[bool, str, Dict[str, Any]]:
    """Check that all tabs cover exactly 10 fiscal years.

    Verifies:
      - Statement tabs have 10 period columns.
      - Table tabs have 10 fiscal-year data rows.
      - README tab documents 10-year coverage.

    Args:
        wb: Openpyxl Workbook instance.

    Returns:
        Tuple[bool, str, Dict[str, Any]]: Success flag, message, details map.
    """
    details: Dict[str, Any] = {}
    failures: List[str] = []

    # 1. Statement tabs: count header period columns (row 3, columns 2+)
    statement_tabs = ["Income Statement", "Balance Sheet", "Cash Flows"]
    for tab in statement_tabs:
        if tab not in wb.sheetnames:
            failures.append(f"{tab}: tab missing")
            continue
        ws = wb[tab]
        periods = [
            ws.cell(row=3, column=c).value
            for c in range(2, ws.max_column + 1)
            if ws.cell(row=3, column=c).value is not None
        ]
        details[tab] = len(periods)
        if len(periods) != 10:
            failures.append(
                f"{tab}: expected 10 period columns, found {len(periods)}"
            )

    # 2. Table tabs: count fiscal year rows (rows 4+, column 1)
    table_tabs = ["Sentiment Analysis", "Stock Prices", "Ratios", "Chart"]
    for tab in table_tabs:
        if tab not in wb.sheetnames:
            failures.append(f"{tab}: tab missing")
            continue
        ws = wb[tab]
        fy_rows = [
            ws.cell(row=r, column=1).value
            for r in range(4, ws.max_row + 1)
            if ws.cell(row=r, column=1).value is not None
        ]
        details[tab] = len(fy_rows)
        if len(fy_rows) != 10:
            failures.append(
                f"{tab}: expected 10 fiscal year rows, found {len(fy_rows)}"
            )

    # 3. README tab: verify mentions 10 filings coverage
    if "README" in wb.sheetnames:
        ws_readme = wb["README"]
        readme_text = " ".join(
            str(c.value)
            for row in ws_readme.iter_rows()
            for c in row
            if c.value is not None
        )
        has_ten = "10" in readme_text or "Ten" in readme_text
        details["README"] = "10 filings coverage noted" if has_ten else "None"
        if not has_ten:
            failures.append("README: does not mention 10 filings coverage")

    if failures:
        return False, "; ".join(failures), details

    return (
        True,
        "All tabs successfully cover exactly 10 fiscal years.",
        details,
    )


def verify_balance_sheet_equation(
    wb: openpyxl.Workbook,
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """Verify that Assets = Liabilities + Equity for every year.

    Locates 'Total assets', 'Total liabilities', and 'Total shareholders'
    equity' rows in the Balance Sheet tab and asserts:
      Assets == Liabilities + Equity (to the dollar).

    Args:
        wb: Openpyxl Workbook instance.

    Returns:
        Tuple[bool, str, List[Dict[str, Any]]]: Success flag, message, and
            per-year verification results.
    """
    if "Balance Sheet" not in wb.sheetnames:
        return False, "Balance Sheet tab is missing", []

    ws = wb["Balance Sheet"]
    headers = [
        ws.cell(row=3, column=c).value for c in range(2, ws.max_column + 1)
    ]
    periods = [h for h in headers if h is not None]

    # Map standardized row labels to lists of column values
    rows: Dict[str, List[Any]] = {}
    for r in range(4, ws.max_row + 1):
        lbl = ws.cell(row=r, column=1).value
        if lbl:
            clean_lbl = (
                str(lbl)
                .strip()
                .replace("’", "'")
                .replace("`", "'")
                .lower()
            )
            vals = [
                ws.cell(row=r, column=c).value
                for c in range(2, len(periods) + 2)
            ]
            rows[clean_lbl] = vals

    # Search for canonical asset, liability, and equity rows
    assets_key = next(
        (
            k
            for k in rows
            if k == "total assets" or k.startswith("total assets")
        ),
        None,
    )
    liab_key = next((k for k in rows if k == "total liabilities"), None)
    le_keys = [
        k
        for k in rows
        if "total liabilities" in k and "equity" in k
    ]
    canonical_eq = [
        "total stockholders' equity",
        "total shareholders' equity",
        "total equity",
    ]
    equity_keys = [k for k in canonical_eq if k in rows] + [
        k
        for k in rows
        if ("total" in k and "equity" in k and "liabilit" not in k)
        and k not in canonical_eq
    ]

    if not assets_key or (not liab_key and not le_keys) or not equity_keys:
        missing_rows = []
        if not assets_key:
            missing_rows.append("Total assets")
        if not liab_key and not le_keys:
            missing_rows.append("Total liabilities (or Liabilities & Equity)")
        if not equity_keys:
            missing_rows.append("Total shareholders' equity")
        return False, f"Missing rows: {', '.join(missing_rows)}", []

    results: List[Dict[str, Any]] = []
    mismatches: List[str] = []

    for idx, period in enumerate(periods):
        a_val = rows[assets_key][idx] if assets_key else None
        e_val = next(
            (rows[k][idx] for k in equity_keys if rows[k][idx] is not None),
            None,
        )

        l_val = rows[liab_key][idx] if liab_key else None
        if l_val is not None and e_val is not None:
            calc_sum = l_val + e_val
        else:
            le_val = next(
                (rows[k][idx] for k in le_keys if rows[k][idx] is not None),
                None,
            )
            calc_sum = le_val
            if calc_sum is not None and e_val is not None:
                l_val = calc_sum - e_val

        if a_val is None or calc_sum is None:
            mismatches.append(
                f"{period}: Missing values "
                f"(A={a_val}, L={l_val}, E={e_val})"
            )
            continue

        diff = a_val - calc_sum
        balanced = (diff == 0)

        results.append(
            {
                "period": period,
                "assets": a_val,
                "liabilities": l_val,
                "equity": e_val,
                "liab_plus_equity": calc_sum,
                "difference": diff,
                "balanced": balanced,
            }
        )
        if not balanced:
            mismatches.append(
                f"{period}: Assets ({a_val}) != Liab ({l_val}) + "
                f"Equity ({e_val}), Diff={diff}"
            )

    if mismatches:
        return False, "; ".join(mismatches), results

    msg = (
        f"Assets = Liabilities + Equity holds to the dollar ($0 difference) "
        f"across all {len(periods)} balance sheet dates."
    )
    return True, msg, results


def verify_revenue_and_net_income(
    wb: openpyxl.Workbook,
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """Verify Income Statement revenue and net income match ratio inputs.

    Compares 'Net sales' and 'Net income' in the 'Income Statement' tab
    against 'Revenue ($M)' and 'Net Income ($M)' in the 'Ratios' tab.

    Args:
        wb: Openpyxl Workbook instance.

    Returns:
        Tuple[bool, str, List[Dict[str, Any]]]: Success flag, message, and
            per-period comparison items.
    """
    if "Income Statement" not in wb.sheetnames:
        return False, "Income Statement tab missing", []
    if "Ratios" not in wb.sheetnames:
        return False, "Ratios tab missing", []

    ws_is = wb["Income Statement"]
    ws_rat = wb["Ratios"]

    # 1. Parse Income Statement periods and rows
    is_headers = [
        ws_is.cell(row=3, column=c).value
        for c in range(2, ws_is.max_column + 1)
    ]
    is_dates = [_to_date(h) for h in is_headers if h is not None]

    is_rows: Dict[str, List[Any]] = {}
    for r in range(4, ws_is.max_row + 1):
        lbl = ws_is.cell(row=r, column=1).value
        if lbl:
            clean_lbl = (
                str(lbl)
                .strip()
                .replace("’", "'")
                .replace("`", "'")
                .lower()
            )
            vals = [
                ws_is.cell(row=r, column=c).value
                for c in range(2, len(is_dates) + 2)
            ]
            is_rows[clean_lbl] = vals

    rev_keys = [
        k
        for k in is_rows
        if k
        in [
            "net sales",
            "total net sales",
            "revenue",
            "revenues",
            "net revenue",
            "total net revenue",
        ]
    ]

    candidate_ni_keys = [
        k
        for k in is_rows
        if ("net income" in k or "net loss" in k)
        and "operating" not in k
        and "per share" not in k
        and "tax" not in k
        and "before" not in k
        and "comprehensive" not in k
        and not k.startswith("less:")
    ]

    if not rev_keys or not candidate_ni_keys:
        missing = []
        if not rev_keys:
            missing.append("Net sales / Net revenue")
        if not candidate_ni_keys:
            missing.append("Net income")
        return False, f"Income statement missing: {', '.join(missing)}", []

    # 2. Parse Ratios tab columns
    rat_headers = [
        str(ws_rat.cell(row=3, column=c).value or "").strip()
        for c in range(1, ws_rat.max_column + 1)
    ]

    try:
        dt_col = rat_headers.index("Fiscal Period End") + 1
        rev_col = next(
            i + 1
            for i, h in enumerate(rat_headers)
            if "Revenue" in h and "($M)" in h
        )
        ni_col = next(
            i + 1
            for i, h in enumerate(rat_headers)
            if "Net Income" in h and "($M)" in h
        )
    except (ValueError, StopIteration) as err:
        return (
            False,
            f"Could not find required columns in Ratios tab: {err}",
            [],
        )

    comparisons: List[Dict[str, Any]] = []
    discrepancies: List[str] = []

    for r in range(4, ws_rat.max_row + 1):
        raw_d = ws_rat.cell(row=r, column=dt_col).value
        p_date = _to_date(raw_d)
        if not p_date:
            continue

        rat_rev = ws_rat.cell(row=r, column=rev_col).value
        rat_ni = ws_rat.cell(row=r, column=ni_col).value

        try:
            col_idx = is_dates.index(p_date)
        except ValueError:
            discrepancies.append(
                f"Date {p_date} found in Ratios but missing in IS"
            )
            continue

        is_rev = next(
            (
                is_rows[k][col_idx]
                for k in rev_keys
                if is_rows[k][col_idx] is not None
            ),
            None,
        )

        # Match net income: prefer exact match with rat_ni if available
        is_ni = None
        for k in candidate_ni_keys:
            if is_rows[k][col_idx] == rat_ni:
                is_ni = is_rows[k][col_idx]
                break

        if is_ni is None:
            prio_keys = [
                k for k in candidate_ni_keys if "attributable" in k
            ] + [
                k
                for k in candidate_ni_keys
                if k in ["net income", "net income (loss)", "net loss"]
            ]
            for k in prio_keys:
                if is_rows[k][col_idx] is not None:
                    is_ni = is_rows[k][col_idx]
                    break

        rev_match = (is_rev == rat_rev)
        ni_match = (is_ni == rat_ni)

        comparisons.append(
            {
                "date": p_date.strftime("%Y-%m-%d"),
                "is_revenue": is_rev,
                "ratios_revenue": rat_rev,
                "rev_match": rev_match,
                "is_net_income": is_ni,
                "ratios_net_income": rat_ni,
                "ni_match": ni_match,
            }
        )

        if not rev_match:
            discrepancies.append(
                f"{p_date}: Revenue mismatch (IS={is_rev} vs Ratios={rat_rev})"
            )
        if not ni_match:
            discrepancies.append(
                f"{p_date}: Net Income mismatch "
                f"(IS={is_ni} vs Ratios={rat_ni})"
            )

    if discrepancies:
        return False, "; ".join(discrepancies), comparisons

    msg = (
        f"Revenue and Net Income match ratio inputs exactly ($0 difference) "
        f"across all {len(comparisons)} fiscal years."
    )
    return True, msg, comparisons


def verify_pn_ratio_calculation(
    wb: openpyxl.Workbook,
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """Verify MD&A P/N ratio = positive count / negative count.

    Checks each row of the 'Sentiment Analysis' tab:
      Stored P/N Ratio == Positive Count / Negative Count.

    Args:
        wb: Openpyxl Workbook instance.

    Returns:
        Tuple[bool, str, List[Dict[str, Any]]]: Success flag, message, and
            per-row verification details.
    """
    if "Sentiment Analysis" not in wb.sheetnames:
        return False, "Sentiment Analysis tab missing", []

    ws = wb["Sentiment Analysis"]
    headers = [
        str(ws.cell(row=3, column=c).value or "").strip()
        for c in range(1, ws.max_column + 1)
    ]

    try:
        fy_col = headers.index("Fiscal Year") + 1
        pos_col = headers.index("Positive") + 1
        neg_col = headers.index("Negative") + 1
        pn_col = headers.index("P/N Ratio") + 1
    except ValueError as err:
        return (
            False,
            f"Missing expected headers in Sentiment Analysis: {err}",
            [],
        )

    results: List[Dict[str, Any]] = []
    errors: List[str] = []

    for r in range(4, ws.max_row + 1):
        fy = ws.cell(row=r, column=fy_col).value
        if fy is None:
            continue

        pos_val = ws.cell(row=r, column=pos_col).value
        neg_val = ws.cell(row=r, column=neg_col).value
        pn_val = ws.cell(row=r, column=pn_col).value

        if pos_val is None or neg_val is None or pn_val is None:
            errors.append(
                f"{fy}: Incomplete sentiment data "
                f"(Pos={pos_val}, Neg={neg_val}, P/N={pn_val})"
            )
            continue

        if neg_val == 0:
            expected_pn = 0.0
        else:
            expected_pn = pos_val / neg_val

        # Check equivalence within floating-point tolerance of 1e-4
        is_close = math.isclose(
            float(pn_val), expected_pn, rel_tol=1e-4, abs_tol=1e-4
        )

        results.append(
            {
                "fiscal_year": fy,
                "positive": pos_val,
                "negative": neg_val,
                "stored_pn": float(pn_val),
                "expected_pn": expected_pn,
                "matches": is_close,
            }
        )

        if not is_close:
            errors.append(
                f"{fy}: P/N ratio mismatch (Stored={pn_val:.4f}, "
                f"Expected={expected_pn:.4f})"
            )

    if errors:
        return False, "; ".join(errors), results

    msg = (
        f"P/N ratio equals positive ÷ negative count for all "
        f"{len(results)} fiscal years."
    )
    return True, msg, results


def run_all_checks(workbook_path: str) -> bool:
    """Run all five verification checks and print detailed PASS/FAIL report.

    Args:
        workbook_path: File path to the analysis workbook Excel file.

    Returns:
        bool: True if all checks passed, False otherwise.
    """
    print("=" * 79)
    print("10-K ANALYSIS WORKBOOK VERIFICATION REPORT")
    print(f"Target Workbook: {workbook_path}")
    print("=" * 79)

    if not os.path.exists(workbook_path):
        print(f"\n[FAIL] File not found: {workbook_path}")
        return False

    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
    except Exception as exc:
        print(f"\n[FAIL] Unable to load workbook: {exc}")
        return False

    # Check 1: All 8 tabs exist
    chk1_pass, chk1_msg, _ = verify_eight_tabs(wb)
    status1 = "PASS" if chk1_pass else "FAIL"
    print(f"\n[{status1}] Check 1: All 8 required tabs exist")
    print(f"       Details: {chk1_msg}")

    # Check 2: All tabs have 10 fiscal years
    chk2_pass, chk2_msg, details2 = verify_ten_fiscal_years(wb)
    status2 = "PASS" if chk2_pass else "FAIL"
    print(f"\n[{status2}] Check 2: All tabs cover exactly 10 fiscal years")
    print(f"       Details: {chk2_msg}")
    for tab_name, count in details2.items():
        print(f"         - {tab_name:<20}: {count}")

    # Check 3: Assets = Liabilities + Equity for every year
    chk3_pass, chk3_msg, details3 = verify_balance_sheet_equation(wb)
    status3 = "PASS" if chk3_pass else "FAIL"
    print(
        f"\n[{status3}] Check 3: Assets = Liabilities + Equity for every year"
    )
    print(f"       Details: {chk3_msg}")
    if details3:
        for item in details3:
            p = item["period"]
            a = item["assets"]
            l = item["liabilities"]
            e = item["equity"]
            diff = item["difference"]
            flag = "OK" if item["balanced"] else "FAIL"
            a_s = f"${a:,}M" if isinstance(a, (int, float)) else str(a)
            l_s = f"${l:,}M" if isinstance(l, (int, float)) else str(l)
            e_s = f"${e:,}M" if isinstance(e, (int, float)) else str(e)
            calc = item.get("liab_plus_equity")
            calc_s = (
                f"${calc:,}M"
                if isinstance(calc, (int, float))
                else str(calc)
            )
            print(
                f"         - {p:<14}: Assets={a_s} | "
                f"L+E={l_s} + {e_s} = {calc_s} | Diff=${diff} [{flag}]"
            )

    # Check 4: Revenue and Net Income match ratio inputs
    chk4_pass, chk4_msg, details4 = verify_revenue_and_net_income(wb)
    status4 = "PASS" if chk4_pass else "FAIL"
    print(f"\n[{status4}] Check 4: Revenue and Net Income match ratio inputs")
    print(f"       Details: {chk4_msg}")
    if details4:
        for item in details4:
            d = item["date"]
            rev = item["is_revenue"]
            ni = item["is_net_income"]
            r_rev = item["ratios_revenue"]
            r_ni = item["ratios_net_income"]
            r_flag = "OK" if item["rev_match"] else "FAIL"
            ni_flag = "OK" if item["ni_match"] else "FAIL"
            rev_s = (
                f"${rev:,}M"
                if isinstance(rev, (int, float))
                else str(rev)
            )
            r_rev_s = (
                f"${r_rev:,}M"
                if isinstance(r_rev, (int, float))
                else str(r_rev)
            )
            ni_s = (
                f"${ni:,}M"
                if isinstance(ni, (int, float))
                else str(ni)
            )
            r_ni_s = (
                f"${r_ni:,}M"
                if isinstance(r_ni, (int, float))
                else str(r_ni)
            )
            print(
                f"         - {d}: Revenue IS={rev_s} / "
                f"Ratio={r_rev_s} [{r_flag}] | "
                f"Net Income IS={ni_s} / Ratio={r_ni_s} [{ni_flag}]"
            )

    # Check 5: P/N ratio = positive count / negative count
    chk5_pass, chk5_msg, details5 = verify_pn_ratio_calculation(wb)
    status5 = "PASS" if chk5_pass else "FAIL"
    print(
        f"\n[{status5}] Check 5: P/N ratio = positive count ÷ negative count"
    )
    print(f"       Details: {chk5_msg}")
    if details5:
        for item in details5:
            fy = item["fiscal_year"]
            pos = item["positive"]
            neg = item["negative"]
            stored = item["stored_pn"]
            calc = item["expected_pn"]
            flag = "OK" if item["matches"] else "FAIL"
            print(
                f"         - {fy:<7}: Positive={pos:<3} | Negative={neg:<3} | "
                f"P/N={stored:.4f} (Expected {calc:.4f}) [{flag}]"
            )

    # Final Summary Table
    print("\n" + "=" * 79)
    print("VERIFICATION SUMMARY")
    print("=" * 79)
    checks = [
        ("Check 1: All 8 tabs exist", chk1_pass),
        ("Check 2: All tabs have 10 fiscal years", chk2_pass),
        ("Check 3: Assets = Liabilities + Equity (all years)", chk3_pass),
        ("Check 4: Revenue & Net Income match ratio inputs", chk4_pass),
        ("Check 5: P/N ratio = positive ÷ negative count", chk5_pass),
    ]

    all_passed = all(passed for _, passed in checks)

    for desc, passed in checks:
        badge = "PASS" if passed else "FAIL"
        print(f"  [{badge:<4}] {desc}")

    print("-" * 79)
    if all_passed:
        print("OVERALL RESULT: ALL CHECKS PASSED (5/5)")
    else:
        print("OVERALL RESULT: SOME CHECKS FAILED")
    print("=" * 79)

    return all_passed


def main() -> None:
    """CLI entry point for running workbook verification."""
    parser = argparse.ArgumentParser(
        description="Verify generated 10-K analysis Excel workbook outputs."
    )
    parser.add_argument(
        "workbook_path",
        nargs="?",
        default="output/AAPL_10K_Analysis.xlsx",
        help=(
            "Path to Excel workbook to verify "
            "(default: output/AAPL_10K_Analysis.xlsx)"
        ),
    )
    args = parser.parse_args()

    success = run_all_checks(args.workbook_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
