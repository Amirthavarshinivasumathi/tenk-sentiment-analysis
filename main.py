"""SEC 10-K Filings Downloader and Financial Analyzer.

This module provides tools to interact with the SEC EDGAR API,
retrieve company metadata, and extract filing details.
"""

import csv
import datetime
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple, Union
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import numpy as np
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests


# Base URLs for SEC EDGAR endpoints
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_BASE_URL = "https://data.sec.gov/submissions/CIK{cik}.json"


def get_sec_headers() -> Dict[str, str]:
    """Retrieve and validate the SEC User-Agent header from the environment.

    The SEC requires all automated requests to identify the user with a custom
    User-Agent header including an email address.

    Returns:
        Dict[str, str]: Request headers containing the User-Agent.

    Raises:
        ValueError: If SEC_USER_AGENT is not set or empty in the environment.
    """
    user_agent = os.getenv("SEC_USER_AGENT")
    if not user_agent or not user_agent.strip():
        raise ValueError(
            "The 'SEC_USER_AGENT' environment variable is missing or empty. "
            "SEC EDGAR requires a User-Agent in the format: "
            "'Sample Company Name AdminContact@<sample company domain>.com'. "
            "Please export SEC_USER_AGENT before running."
        )

    # Standard headers requested by SEC EDGAR API guidelines
    return {
        "User-Agent": user_agent.strip(),
        "Accept-Encoding": "gzip, deflate",
    }


def get_company_cik(ticker: str, headers: Dict[str, str]) -> str:
    """Look up a company's 10-digit CIK from SEC company_tickers.json.

    Args:
        ticker: The stock ticker symbol (e.g., 'AAPL').
        headers: HTTP headers containing the User-Agent.

    Returns:
        str: 10-digit zero-padded CIK string.

    Raises:
        ValueError: If the ticker symbol cannot be found.
        requests.RequestException: If the network request fails.
    """
    ticker_clean = ticker.strip().upper()

    try:
        response = requests.get(SEC_TICKERS_URL, headers=headers, timeout=15)
        response.raise_for_status()
        tickers_data: Dict[str, Dict[str, Any]] = response.json()
    except requests.exceptions.RequestException as exc:
        raise requests.RequestException(
            f"Failed to fetch ticker mapping from {SEC_TICKERS_URL}: {exc}"
        ) from exc
    except ValueError as exc:
        raise ValueError(
            f"Failed to parse JSON response from {SEC_TICKERS_URL}: {exc}"
        ) from exc

    # Search the company ticker data for matching ticker symbol
    cik_num: Optional[int] = None
    for entry in tickers_data.values():
        if entry.get("ticker", "").upper() == ticker_clean:
            cik_num = entry.get("cik_str")
            break

    if cik_num is None:
        raise ValueError(
            f"Ticker symbol '{ticker_clean}' was not found in SEC EDGAR "
            "company directory."
        )

    # SEC URLs require CIK to be zero-padded to 10 digits
    return str(cik_num).zfill(10)


def get_recent_10k_filings(
    ticker: str, n: int = 10
) -> List[Dict[str, str]]:
    """Fetch the N most recent 10-K filings for a given company ticker.

    Looks up the company CIK, retrieves the SEC submissions JSON, and
    extracts the accession number, filing date, and primary document name
    for each 10-K filing.

    Args:
        ticker: Stock ticker symbol (e.g., 'AAPL').
        n: Number of most recent 10-K filings to retrieve (default is 10).

    Returns:
        List[Dict[str, str]]: A list of dictionaries containing:
            - accession_number: The SEC accession number.
            - filing_date: The date filed (YYYY-MM-DD).
            - report_date: The fiscal period end date (YYYY-MM-DD).
            - primary_document: The filename of the primary document.

    Raises:
        ValueError: If invalid arguments or missing data are encountered.
        requests.RequestException: If network operations fail.
    """
    if n <= 0:
        raise ValueError(f"Filing count N must be greater than 0, got {n}.")

    ticker_clean = ticker.strip().upper()

    # Obtain verified request headers with User-Agent
    headers = get_sec_headers()

    # Look up the 10-digit zero-padded CIK
    cik_padded = get_company_cik(ticker_clean, headers)

    # Request the company's submissions JSON
    submissions_url = SEC_SUBMISSIONS_BASE_URL.format(cik=cik_padded)
    try:
        response = requests.get(submissions_url, headers=headers, timeout=15)
        response.raise_for_status()
        submissions_data: Dict[str, Any] = response.json()
    except requests.exceptions.RequestException as exc:
        raise requests.RequestException(
            f"Failed to fetch submissions for CIK {cik_padded} from "
            f"{submissions_url}: {exc}"
        ) from exc
    except ValueError as exc:
        raise ValueError(
            f"Failed to parse submissions JSON from {submissions_url}: {exc}"
        ) from exc

    filings_info = submissions_data.get("filings", {})
    recent_filings = filings_info.get("recent", {})

    forms = recent_filings.get("form", [])
    accession_numbers = recent_filings.get("accessionNumber", [])
    filing_dates = recent_filings.get("filingDate", [])
    report_dates = recent_filings.get("reportDate", [])
    primary_documents = recent_filings.get("primaryDocument", [])

    results: List[Dict[str, str]] = []

    # Iterate through recent filings to collect 10-K records
    for i in range(len(forms)):
        if forms[i] == "10-K":
            results.append(
                {
                    "accession_number": accession_numbers[i],
                    "filing_date": filing_dates[i],
                    "report_date": (
                        report_dates[i] if i < len(report_dates) else ""
                    ),
                    "primary_document": primary_documents[i],
                    "cik": cik_padded,
                    "ticker": ticker_clean,
                }
            )
            if len(results) == n:
                break

    # If recent filings don't contain enough 10-Ks, check historical files
    older_files = filings_info.get("files", [])
    if len(results) < n and older_files:
        for file_meta in older_files:
            file_name = file_meta.get("name")
            if not file_name:
                continue

            older_url = f"https://data.sec.gov/submissions/{file_name}"
            try:
                old_resp = requests.get(older_url, headers=headers, timeout=15)
                old_resp.raise_for_status()
                old_data = old_resp.json()
            except (requests.RequestException, ValueError) as exc:
                print(
                    f"Warning: Could not fetch older filings from "
                    f"{older_url}: {exc}",
                    file=sys.stderr,
                )
                continue

            old_forms = old_data.get("form", [])
            old_acc = old_data.get("accessionNumber", [])
            old_dates = old_data.get("filingDate", [])
            old_reports = old_data.get("reportDate", [])
            old_docs = old_data.get("primaryDocument", [])

            for j in range(len(old_forms)):
                if old_forms[j] == "10-K":
                    results.append(
                        {
                            "accession_number": old_acc[j],
                            "filing_date": old_dates[j],
                            "report_date": (
                                old_reports[j] if j < len(old_reports) else ""
                            ),
                            "primary_document": old_docs[j],
                            "cik": cik_padded,
                            "ticker": ticker_clean,
                        }
                    )
                    if len(results) == n:
                        break

            if len(results) == n:
                break

    return results


def download_filings(
    ticker: str, filings: List[Dict[str, str]]
) -> List[str]:
    """Download the primary 10-K HTML document for each filing in the list.

    Files are saved into a subfolder per ticker adhering to the pattern:
        filings/TICKER/TICKER_10K_FY{year}_{filed}.htm

    Implements retry with exponential backoff if the SEC responds with
    HTTP 429 (Too Many Requests) or HTTP 403 (Forbidden), and sleeps
    approximately 1 second between requests to respect SEC EDGAR rate limits.

    Args:
        ticker: Stock ticker symbol (e.g., 'AAPL').
        filings: List of filing dictionaries, each containing:
            - 'accession_number': SEC accession number
            - 'filing_date': Filing date string (YYYY-MM-DD)
            - 'primary_document': Primary document filename

    Returns:
        List[str]: Relative paths of the downloaded local HTML files.

    Raises:
        ValueError: If ticker is empty or filings list is invalid.
        requests.RequestException: If a document cannot be downloaded after
            exhausting retries.
        OSError: If local directory creation or file writing fails.
    """
    if not ticker or not ticker.strip():
        raise ValueError("Ticker symbol cannot be empty.")
    if not filings:
        print("No filings provided to download.", file=sys.stderr)
        return []

    ticker_clean = ticker.strip().upper()
    headers = get_sec_headers()

    # Retrieve the company's CIK number (without leading zeros for archive URL)
    cik_padded = get_company_cik(ticker_clean, headers)
    cik_number = str(int(cik_padded))

    # Ensure destination directory exists: filings/TICKER/
    output_directory = os.path.join("filings", ticker_clean)
    try:
        os.makedirs(output_directory, exist_ok=True)
    except OSError as err:
        raise OSError(
            f"Failed to create directory '{output_directory}': {err}"
        ) from err

    downloaded_files: List[str] = []
    max_retries = 4
    base_backoff_sec = 2.0

    print(
        f"\nStarting download of {len(filings)} filing(s) for "
        f"{ticker_clean}..."
    )

    for index, filing in enumerate(filings, start=1):
        accession_raw = filing["accession_number"]
        accession_clean = accession_raw.replace("-", "")
        filing_date = filing["filing_date"]
        primary_doc = filing["primary_document"]

        # Derive fiscal year from report_date (period end date)
        rep_date = filing.get("report_date", "")
        if rep_date and len(rep_date) >= 4:
            year = rep_date[:4]
        else:
            year = filing_date[:4]

        # Construct target filename: TICKER_10K_FY{year}_{filed}.htm
        filename = f"{ticker_clean}_10K_FY{year}_{filing_date}.htm"
        target_path = os.path.join(output_directory, filename)

        if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
            print(
                f"[{index}/{len(filings)}] Using existing FY{year} "
                f"({filing_date}) -> {target_path}"
            )
            downloaded_files.append(target_path)
            filing["file_path"] = target_path
            continue

        # SEC EDGAR archive document URL
        doc_url = (
            f"https://www.sec.gov/Archives/edgar/data/"
            f"{cik_number}/{accession_clean}/{primary_doc}"
        )

        print(
            f"[{index}/{len(filings)}] Downloading FY{year} "
            f"({filing_date}) -> {target_path}..."
        )

        # Download with exponential backoff for HTTP 429/403
        success = False
        for attempt in range(1, max_retries + 1):
            try:
                resp = requests.get(doc_url, headers=headers, timeout=25)

                # Retry on rate limiting or temporary forbidden response
                if resp.status_code in (429, 403):
                    retry_header = resp.headers.get("Retry-After")
                    if retry_header and retry_header.isdigit():
                        wait_seconds = float(retry_header)
                    else:
                        wait_seconds = base_backoff_sec * (2 ** (attempt - 1))

                    print(
                        f"  Received HTTP {resp.status_code}. Retrying in "
                        f"{wait_seconds:.1f}s "
                        f"(attempt {attempt}/{max_retries})...",
                        file=sys.stderr,
                    )
                    time.sleep(wait_seconds)
                    continue

                resp.raise_for_status()

                # Save HTML file content to disk
                with open(target_path, "wb") as html_file:
                    html_file.write(resp.content)

                downloaded_files.append(target_path)
                filing["file_path"] = target_path
                success = True
                break

            except requests.exceptions.RequestException as exc:
                if attempt == max_retries:
                    raise requests.RequestException(
                        f"Failed to download {filename} from {doc_url} "
                        f"after {max_retries} attempts: {exc}"
                    ) from exc

                wait_seconds = base_backoff_sec * (2 ** (attempt - 1))
                print(
                    f"  Network error for {filename}: {exc}. "
                    f"Retrying in {wait_seconds:.1f}s "
                    f"(attempt {attempt}/{max_retries})...",
                    file=sys.stderr,
                )
                time.sleep(wait_seconds)

        if not success:
            raise requests.RequestException(
                f"Failed to download {filename} from {doc_url}."
            )

        # Respect SEC rate limit: sleep about 1 second between requests
        time.sleep(1.0)

    print(
        f"Successfully downloaded all {len(downloaded_files)} filing(s) "
        f"to '{output_directory}'."
    )
    return downloaded_files


def extract_mda(html: str) -> str:
    """Extract Item 7 MD&A text from a 10-K filing HTML document.

    Locates the Item 7 Management's Discussion and Analysis (MD&A) section
    and extracts its narrative text up to the start of Item 8 (Financial
    Statements and Supplementary Data) or Item 9 (Changes in and
    Disagreements with Accountants).

    Handles both inline heading formats (e.g. Apple, where 'Item 7.' is
    followed inline by 'Management's Discussion') and standalone headings
    (e.g. Intel, where 'Item 7.' appears as a standalone header tag/cell
    followed by the section title in subsequent blocks) with a fallback.

    All HTML tags and data tables are stripped, returning clean plain text.

    Args:
        html: Raw HTML string of the 10-K filing.

    Returns:
        str: Cleaned plain text of the extracted Item 7 MD&A section.

    Raises:
        ValueError: If html is empty or if Item 7 start / stop boundaries
            cannot be identified.
    """
    if not html or not html.strip():
        raise ValueError("HTML content is empty or invalid.")

    # 1. Regex patterns for Item 7 start
    # Inline pattern (Apple style: 'Item 7.' followed inline by title)
    inline_pattern = re.compile(
        r"(?i)(?:<[^>]+>\s*)*item\s*7\b[.\s:–-]{1,10}"
        r"(?:&#160;|\s|<[^>]+>)*"
        r"management(?:&#8217;|\'|’)?s\s+discussion\s+and\s+analysis\s+of\s+"
        r"financial\s+condition\s+and\s+results\s+of\s+operations",
        re.MULTILINE,
    )

    # Standalone pattern (Intel style fallback: 'Item 7.' standalone heading
    # followed within ~300 chars by the section title)
    standalone_pattern = re.compile(
        r"(?i)(?:<[^>]+>\s*)*item\s*7\b[.\s:–-]{1,10}"
        r"(?:<[^>]+>|&#160;|\s){1,300}?"
        r"management(?:&#8217;|\'|’)?s\s+discussion\s+and\s+analysis",
        re.MULTILINE,
    )

    # 2. Regex patterns for stop boundaries (Item 8 or Item 9)
    item8_pattern = re.compile(
        r"(?i)(?:<[^>]+>\s*)*item\s*8\b[.\s:–-]{1,10}"
        r"(?:<[^>]+>|&#160;|\s){0,300}?"
        r"financial\s+statements",
        re.MULTILINE,
    )
    item9_pattern = re.compile(
        r"(?i)(?:<[^>]+>\s*)*item\s*9\b[.\s:–-]{1,10}"
        r"(?:<[^>]+>|&#160;|\s){0,300}?"
        r"changes\s+in\s+and\s+disagreements",
        re.MULTILINE,
    )

    def find_valid_starts(pattern: re.Pattern) -> List[int]:
        valid_positions: List[int] = []
        for match in pattern.finditer(html):
            pos = match.start()
            # Filter out Table of Contents hyperlinks (<a href=...)
            preceding = html[max(0, pos - 150):pos]
            if re.search(r"<a\s+[^>]*href=", preceding, re.IGNORECASE):
                continue
            # Filter forward references ('in Part II, Item 7', 'see Item 7')
            following = html[pos:pos + 400]
            if re.search(
                r"(?i)(?:under\s+the\s+heading|of\s+this\s+form\s+10-k|"
                r"part\s+ii,\s*item\s*7|in\s+item\s*7|see\s+item\s*7)",
                preceding,
            ):
                continue
            if re.search(
                r"(?i)(?:under\s+the\s+heading|of\s+this\s+form\s+10-k)",
                following,
            ):
                continue
            valid_positions.append(pos)
        return valid_positions

    clean_text = ""
    try:
        # Try primary inline match (Apple), then fallback to standalone
        starts = find_valid_starts(inline_pattern)
        if not starts:
            starts = find_valid_starts(standalone_pattern)

        if starts:
            start_pos = starts[-1]
            table_open = html.rfind(
                "<table", max(0, start_pos - 300), start_pos
            )
            table_close = html.rfind(
                "</table", max(0, start_pos - 300), start_pos
            )
            if table_open != -1 and table_open > table_close:
                start_pos = table_open

            stop_pos: Optional[int] = None
            for match in item8_pattern.finditer(html, start_pos):
                pos = match.start()
                preceding = html[max(0, pos - 150):pos]
                if re.search(r"<a\s+[^>]*href=", preceding, re.IGNORECASE):
                    continue
                if re.search(
                    r"(?i)(?:of\s+this\s+form\s+10-k|"
                    r"in\s+part\s+ii,\s*item\s*8|"
                    r"see\s+item\s*8|in\s+item\s*8)",
                    preceding,
                ):
                    continue
                stop_pos = pos
                break

            if stop_pos is None:
                for match in item9_pattern.finditer(html, start_pos):
                    pos = match.start()
                    preceding = html[max(0, pos - 150):pos]
                    if re.search(r"<a\s+[^>]*href=", preceding, re.IGNORECASE):
                        continue
                    stop_pos = pos
                    break

            if stop_pos is not None:
                mda_html = html[start_pos:stop_pos]
                soup = BeautifulSoup(mda_html, "html.parser")
                for table in soup.find_all("table"):
                    table_text = table.get_text()
                    if (
                        re.search(r"item\s*7\b", table_text, re.IGNORECASE)
                        and re.search(r"management", table_text, re.IGNORECASE)
                    ):
                        table.replace_with(" " + table_text + " ")
                    else:
                        table.decompose()
                raw_text = soup.get_text(separator=" ")
                clean_text = re.sub(r"[ \t\r\f\v]+", " ", raw_text)
                clean_text = re.sub(r"\n\s*\n", "\n\n", clean_text).strip()
    except Exception:
        clean_text = ""

    # If extraction succeeded and produced substantive narrative, return it
    if clean_text and len(clean_text.split()) >= 1000:
        return clean_text

    # Fallback for integrated annual reports (e.g. Intel) where MD&A is a
    # standalone major section banner and Item 7 is only referenced in a
    # 10-K cross-reference index table.
    fb_start = re.search(
        r"(?i)(?:font-size:\s*(?:1[4-9]|2[0-9])pt|<h[1-3]\b)[^>]*>[^<]*"
        r"management(?:&#8217;|\'|’)?s\s+discussion",
        html,
    )
    if not fb_start:
        fb_start = re.search(
            r"(?i)>management(?:&#8217;|\'|’)?s\s+discussion\s+and\s+analysis"
            r"\s*(?:\([^)]+\))?<",
            html,
        )

    if not fb_start:
        if clean_text:
            return clean_text
        raise ValueError(
            "Could not locate the Item 7 MD&A start section in the document."
        )

    fb_start_pos = fb_start.start()
    fb_stop_pos: Optional[int] = None
    for stop_pat in [
        (
            r"(?i)<[^>]+>\s*(?:other\s+key\s+information|"
            r"report\s+of\s+independent|"
            r"consolidated\s+financial\s+statements|"
            r"financial\s+statements\s+and\s+supplementary)"
        ),
        r"(?i)item\s*8\b",
    ]:
        sm = re.search(stop_pat, html[fb_start_pos + 5000:])
        if sm:
            fb_stop_pos = fb_start_pos + 5000 + sm.start()
            break

    if fb_stop_pos is None:
        if clean_text:
            return clean_text
        raise ValueError(
            "Could not locate the Item 8 or Item 9 stop section in document."
        )

    fb_soup = BeautifulSoup(html[fb_start_pos:fb_stop_pos], "html.parser")
    for table in fb_soup.find_all("table"):
        table.decompose()
    fb_raw = fb_soup.get_text(separator=" ")
    fb_clean = re.sub(r"[ \t\r\f\v]+", " ", fb_raw)
    fb_clean = re.sub(r"\n\s*\n", "\n\n", fb_clean).strip()

    return fb_clean if fb_clean else clean_text


def load_loughran_mcdonald_dictionary(
    csv_path: str = "data/Loughran-McDonald_MasterDictionary_1993-2025.csv",
) -> Dict[str, Set[str]]:
    """Load the Loughran-McDonald sentiment dictionary from a CSV file.

    Extracts word sets for seven financial sentiment categories:
    negative, positive, uncertainty, litigious, strong_modal,
    weak_modal, and constraining.

    Args:
        csv_path: Path to the Loughran-McDonald master dictionary CSV file.

    Returns:
        Dict[str, Set[str]]: Mapping from category name to a set of
            uppercase words.

    Raises:
        FileNotFoundError: If the CSV file does not exist.
        IOError: If an error occurs reading the CSV file.
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(
            f"Loughran-McDonald dictionary file not found at: {csv_path}"
        )

    categories = [
        "Negative",
        "Positive",
        "Uncertainty",
        "Litigious",
        "Strong_Modal",
        "Weak_Modal",
        "Constraining",
    ]

    # Initialize empty sets for each category key
    word_sets: Dict[str, Set[str]] = {
        cat.lower(): set() for cat in categories
    }

    try:
        with open(csv_path, mode="r", encoding="utf-8") as csv_fp:
            reader = csv.DictReader(csv_fp)
            for row in reader:
                word_raw = row.get("Word", "")
                if not word_raw:
                    continue
                word = word_raw.strip().upper()

                for cat in categories:
                    val = row.get(cat, "0")
                    if val and val != "0":
                        word_sets[cat.lower()].add(word)

    except OSError as err:
        raise OSError(
            f"Failed to load dictionary from '{csv_path}': {err}"
        ) from err

    # Ensure 'RISK' and 'RISKS' are counted in the negative category
    word_sets["negative"].add("RISK")
    word_sets["negative"].add("RISKS")

    return word_sets


def score_mda(
    text: str, dictionary: Any
) -> Dict[str, Any]:
    """Score MD&A text using the Loughran-McDonald financial dictionary.

    Counts total words and calculates frequencies and percentages for seven
    categories: negative, positive, uncertainty, litigious, strong modal,
    weak modal, and constraining. Computes the Positive/Negative (P/N) ratio.

    Args:
        text: Plain text of the extracted Item 7 MD&A section.
        dictionary: Loughran-McDonald dictionary. Can be a dictionary mapping
            category names to word sets, a pandas DataFrame, or a file path
            (str) to the CSV dictionary.

    Returns:
        Dict[str, Any]: A dictionary containing:
            - 'total_words': Total alphabetic word count.
            - 'counts': Dictionary of word counts for each category.
            - 'percentages': Dictionary of word percentages for each category.
            - 'pn_ratio': Positive count divided by negative count.
            - Category-specific counts (e.g. 'negative_count').
            - Category-specific percentages (e.g. 'negative_pct').

    Raises:
        ValueError: If text is empty or dictionary cannot be parsed.
    """
    if not text or not text.strip():
        raise ValueError("Text to score cannot be empty.")

    # Standardize dictionary into Dict[str, Set[str]]
    if isinstance(dictionary, str):
        cat_dict = load_loughran_mcdonald_dictionary(dictionary)
    elif isinstance(dictionary, dict):
        cat_dict = {
            k.lower(): {w.strip().upper() for w in v}
            for k, v in dictionary.items()
        }
    elif hasattr(dictionary, "columns"):  # pandas DataFrame
        categories = [
            "Negative",
            "Positive",
            "Uncertainty",
            "Litigious",
            "Strong_Modal",
            "Weak_Modal",
            "Constraining",
        ]
        cat_dict = {}
        for col in categories:
            if col in dictionary.columns:
                words = dictionary[dictionary[col] > 0]["Word"].dropna()
                cat_dict[col.lower()] = set(words.astype(str).str.upper())
            else:
                cat_dict[col.lower()] = set()
    else:
        raise ValueError(
            "Unsupported dictionary format. Expected dict of sets, "
            "pandas DataFrame, or CSV path string."
        )

    # Tokenize text into alphabetic words
    tokens = re.findall(r"\b[A-Za-z]+\b", text.upper())
    total_words = len(tokens)

    # Categories to evaluate
    eval_categories = [
        "negative",
        "positive",
        "uncertainty",
        "litigious",
        "strong_modal",
        "weak_modal",
        "constraining",
    ]

    # Count occurrences in each category
    counts: Dict[str, int] = {}
    for cat in eval_categories:
        word_set = cat_dict.get(cat, set())
        counts[cat] = sum(1 for token in tokens if token in word_set)

    # Compute percentages of total words
    percentages: Dict[str, float] = {}
    for cat in eval_categories:
        if total_words > 0:
            percentages[cat] = (counts[cat] / total_words) * 100.0
        else:
            percentages[cat] = 0.0

    # Compute P/N ratio (positive count / negative count)
    pos_count = counts["positive"]
    neg_count = counts["negative"]
    if neg_count > 0:
        pn_ratio = pos_count / neg_count
    else:
        pn_ratio = 0.0

    return {
        "total_words": total_words,
        "counts": counts,
        "percentages": percentages,
        "pn_ratio": pn_ratio,
        # Flat entries for tabular and spreadsheet export
        "negative_count": counts["negative"],
        "positive_count": counts["positive"],
        "uncertainty_count": counts["uncertainty"],
        "litigious_count": counts["litigious"],
        "strong_modal_count": counts["strong_modal"],
        "weak_modal_count": counts["weak_modal"],
        "constraining_count": counts["constraining"],
        "negative_pct": percentages["negative"],
        "positive_pct": percentages["positive"],
        "uncertainty_pct": percentages["uncertainty"],
        "litigious_pct": percentages["litigious"],
        "strong_modal_pct": percentages["strong_modal"],
        "weak_modal_pct": percentages["weak_modal"],
        "constraining_pct": percentages["constraining"],
    }


def parse_financial_value(val_str: str) -> Optional[float]:
    """Parse numeric values from SEC financial tables.

    Handles parenthesized negative numbers like '(12,150)' into -12150.0,
    dollar signs ('$ 416,161'), commas, footnotes ('(12,150)*'),
    and zero representations ('—', '-', '–') into 0.0.

    Args:
        val_str: Raw table cell string.

    Returns:
        Optional[float]: Parsed float value, or None if empty / text header.
    """
    if not val_str:
        return None
    cleaned = val_str.replace("$", "").replace("\xa0", " ").strip()
    if not cleaned or cleaned in ("—", "-", "–", "N/A", "none", ""):
        return 0.0

    # Match parenthesized negative values, e.g., (12,150) or $(12,150)*
    neg_match = re.match(r"^\(\s*([\d,]+(?:\.\d+)?)\s*\)\*?$", cleaned)
    if neg_match:
        try:
            return -float(neg_match.group(1).replace(",", ""))
        except ValueError:
            return None

    # Match standard positive or leading minus values, e.g., 416,161
    pos_match = re.match(r"^(-?[\d,]+(?:\.\d+)?)\*?$", cleaned)
    if pos_match:
        try:
            return float(pos_match.group(1).replace(",", ""))
        except ValueError:
            return None

    return None


def parse_statement_table(html_content: str) -> pd.DataFrame:
    """Parse an SEC statement R-file HTML table into a pandas DataFrame.

    Extracts period dates as DataFrame columns and financial line items as
    the DataFrame index. Converts numeric cells and handles parenthesized
    negatives.

    Args:
        html_content: Raw HTML text of the R-file statement.

    Returns:
        pd.DataFrame: Structured financial statement with periods as columns.
    """
    soup = BeautifulSoup(html_content, "html.parser")
    table = soup.find("table")
    if not table:
        return pd.DataFrame()

    rows = table.find_all("tr")
    if not rows:
        return pd.DataFrame()

    date_regex = re.compile(
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+"
        r"\d{1,2},\s+\d{4}|\b(?:19|20)\d{2}\b",
        re.IGNORECASE,
    )

    period_cols: List[str] = []
    data_start_idx = 1

    # Detect header row containing period dates
    for r_idx, row in enumerate(rows[:5]):
        cells = row.find_all(["th", "td"])
        cell_texts = [c.get_text(strip=True) for c in cells]
        dates = [t for t in cell_texts if date_regex.search(t)]
        if dates and len(dates) >= 1:
            if len(dates) == len(cell_texts) - 1:
                period_cols = cell_texts[1:]
                data_start_idx = r_idx + 1
                break
            elif len(dates) == len(cell_texts):
                period_cols = cell_texts
                data_start_idx = r_idx + 1
                break

    if not period_cols:
        first_cells = rows[0].find_all(["th", "td"])
        if len(first_cells) > 1:
            period_cols = [c.get_text(strip=True) for c in first_cells[1:]]
            data_start_idx = 1

    clean_cols = [re.sub(r"\s+", " ", col).strip() for col in period_cols]
    records: List[List[Optional[float]]] = []
    line_items: List[str] = []

    for row in rows[data_start_idx:]:
        cells = row.find_all(["th", "td"])
        if not cells:
            continue
        raw_label = cells[0].get_text(strip=True)
        clean_label = re.sub(r"\s+", " ", raw_label).strip()
        if not clean_label:
            continue

        value_cells = cells[1:]
        row_values: List[Optional[float]] = []
        has_any_num = False

        for i in range(len(clean_cols)):
            if i < len(value_cells):
                txt = value_cells[i].get_text(strip=True)
                val = parse_financial_value(txt)
                row_values.append(val)
                if val is not None:
                    has_any_num = True
            else:
                row_values.append(None)

        if has_any_num:
            records.append(row_values)
            line_items.append(clean_label)

    df = pd.DataFrame(records, index=line_items, columns=clean_cols)
    df.index.name = "Line Item"
    return df


def extract_statements(
    filing: Dict[str, Any]
) -> Dict[str, pd.DataFrame]:
    """Fetch FilingSummary.xml and extract the 3 primary financial statements.

    Identifies the R-files for:
        1. CONSOLIDATED STATEMENTS OF OPERATIONS (Income Statement)
        2. FINANCIAL POSITION (Balance Sheet)
        3. CASH FLOWS (Statements of Cash Flows)

    Parses each R-file into a pandas DataFrame with periods as columns,
    properly converting parenthesized negatives like '(12,150)' to -12150.0.

    Args:
        filing: Filing metadata dictionary containing 'accession_number',
            and optionally 'cik' or 'ticker'.

    Returns:
        Dict[str, pd.DataFrame]: Mapping containing:
            - 'operations': DataFrame of Statement of Operations
            - 'balance_sheet': DataFrame of Financial Position / Balance Sheet
            - 'cash_flows': DataFrame of Statement of Cash Flows

    Raises:
        ValueError: If required filing metadata is missing or statements
            cannot be found.
        requests.RequestException: If network requests fail.
    """
    if "accession_number" not in filing:
        raise ValueError("Filing dictionary must contain 'accession_number'.")

    headers = get_sec_headers()

    # Determine CIK
    cik = filing.get("cik")
    if not cik:
        if filing.get("ticker"):
            cik = get_company_cik(filing["ticker"], headers)
        else:
            cik = filing["accession_number"].split("-")[0]
    cik_num = str(int(cik))

    accession_raw = filing["accession_number"]
    accession_clean = accession_raw.replace("-", "")

    # Base URL for filing files in SEC EDGAR archive
    base_url = (
        f"https://www.sec.gov/Archives/edgar/data/{cik_num}/{accession_clean}"
    )
    fs_url = f"{base_url}/FilingSummary.xml"

    # Fetch FilingSummary.xml
    try:
        fs_resp = requests.get(fs_url, headers=headers, timeout=20)
        fs_resp.raise_for_status()
        root = ET.fromstring(fs_resp.content)
    except Exception as exc:
        raise requests.RequestException(
            f"Failed to fetch or parse FilingSummary.xml from {fs_url}: {exc}"
        ) from exc

    # Identify target R-files for the three financial statements
    r_files: Dict[str, str] = {}
    for report in root.findall(".//Report"):
        sn = (report.findtext("ShortName") or "").strip()
        sn_lower = sn.lower()
        if "parenthetical" in sn_lower:
            continue

        file_name = (
            report.findtext("HtmlFileName") or report.findtext("XmlFileName")
        )
        if not file_name:
            continue

        # 1. Consolidated Statements of Operations (Income Statement)
        if (
            (
                "operations" in sn_lower
                or "income" in sn_lower
                or "earnings" in sn_lower
            )
            and "comprehensive" not in sn_lower
            and "operations" not in r_files
        ):
            r_files["operations"] = file_name

        # 2. Financial Position (Balance Sheet)
        if (
            ("balance sheet" in sn_lower or "financial position" in sn_lower)
            and "balance_sheet" not in r_files
        ):
            r_files["balance_sheet"] = file_name

        # 3. Consolidated Statements of Cash Flows
        if "cash flow" in sn_lower and "cash_flows" not in r_files:
            r_files["cash_flows"] = file_name

    statements: Dict[str, pd.DataFrame] = {}

    for stmt_key, r_file_name in r_files.items():
        stmt_url = f"{base_url}/{r_file_name}"
        try:
            r_resp = requests.get(stmt_url, headers=headers, timeout=20)
            r_resp.raise_for_status()
            df = parse_statement_table(r_resp.text)
            statements[stmt_key] = df
            # Polite pause between requests to respect rate limits
            time.sleep(0.5)
        except Exception as exc:
            print(
                f"Warning: Failed to fetch {stmt_key} from {stmt_url}: {exc}",
                file=sys.stderr,
            )
            statements[stmt_key] = pd.DataFrame()

    return statements


def _to_date(val: Any) -> datetime.date:
    """Convert a string, Timestamp, or date-like object to datetime.date.

    Args:
        val: Input date representation.

    Returns:
        datetime.date: Standardized Python date object.
    """
    if isinstance(val, datetime.date) and not isinstance(
        val, datetime.datetime
    ):
        return val
    if isinstance(val, (datetime.datetime, pd.Timestamp)):
        return val.date()
    return pd.to_datetime(val).date()


def fetch_stock_prices(
    ticker: str,
    dates: Union[Sequence[Any], Any],
) -> Dict[Any, float]:
    """Fetch daily closing stock prices for a ticker across target dates.

    Queries Yahoo Finance's chart API for daily bars spanning the full date
    range of the provided balance-sheet dates. For each balance-sheet date,
    determines the closing price on or before that date (falling back to the
    next available trading day if no prior trading day exists).

    Args:
        ticker: The stock ticker symbol (e.g., 'AAPL').
        dates: A single date or sequence of dates (as strings, datetime.date,
            or pd.Timestamp).

    Returns:
        Dict[Any, float]: Mapping from each input date to its corresponding
            closing stock price rounded to two decimal places.

    Raises:
        ValueError: If ticker is invalid, dates are empty, or no valid
            quotes are returned by the API.
        requests.RequestException: If network requests fail or exhaust retries.
    """
    if not ticker or not isinstance(ticker, str):
        raise ValueError("Ticker must be a non-empty string.")

    if isinstance(dates, (str, datetime.date, pd.Timestamp)):
        date_list = [dates]
    else:
        date_list = list(dates)

    if not date_list:
        return {}

    parsed_dates = [_to_date(d) for d in date_list]
    min_date = min(parsed_dates)
    max_date = max(parsed_dates)

    # 14-day buffer before and after to ensure trading day coverage
    start_dt = min_date - datetime.timedelta(days=14)
    end_dt = max_date + datetime.timedelta(days=14)

    period1 = int(
        datetime.datetime.combine(
            start_dt, datetime.time(0, 0)
        ).replace(tzinfo=datetime.timezone.utc).timestamp()
    )
    period2 = int(
        datetime.datetime.combine(
            end_dt, datetime.time(23, 59, 59)
        ).replace(tzinfo=datetime.timezone.utc).timestamp()
    )

    clean_ticker = ticker.strip().upper()
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        )
    }
    params = {
        "period1": period1,
        "period2": period2,
        "interval": "1d",
    }

    # Query Yahoo Finance chart API across primary and backup hosts
    hosts = ["query1.finance.yahoo.com", "query2.finance.yahoo.com"]
    last_err: Optional[Exception] = None
    payload: Dict[str, Any] = {}

    for host in hosts:
        url = f"https://{host}/v8/finance/chart/{clean_ticker}"
        for attempt in range(3):
            try:
                resp = requests.get(
                    url, headers=headers, params=params, timeout=20
                )
                if resp.status_code == 429:
                    time.sleep(1.0 * (2 ** attempt))
                    continue
                resp.raise_for_status()
                payload = resp.json()
                last_err = None
                break
            except Exception as exc:
                last_err = exc
                time.sleep(1.0 * (2 ** attempt))
        if payload:
            break

    if not payload:
        if last_err:
            raise requests.RequestException(
                f"Failed to fetch chart data from Yahoo for "
                f"{clean_ticker}: {last_err}"
            ) from last_err
        raise ValueError(
            f"No chart payload returned from Yahoo for ticker {clean_ticker}."
        )

    chart = payload.get("chart", {})
    if chart.get("error"):
        raise ValueError(f"Yahoo API error: {chart.get('error')}")

    results_data = chart.get("result")
    if not results_data:
        raise ValueError(
            f"No chart results returned for ticker {clean_ticker}."
        )

    result_entry = results_data[0]
    timestamps = result_entry.get("timestamp", [])
    indicators = result_entry.get("indicators", {})
    quotes = indicators.get("quote", [{}])[0].get("close", [])

    if not timestamps or not quotes:
        raise ValueError(f"No price quotes returned for {clean_ticker}.")

    gmtoffset = result_entry.get("meta", {}).get("gmtoffset", 0)

    trading_days: Dict[datetime.date, float] = {}
    for ts, close_val in zip(timestamps, quotes):
        if close_val is not None:
            t_dt = datetime.datetime.fromtimestamp(
                ts + gmtoffset, tz=datetime.timezone.utc
            ).date()
            trading_days[t_dt] = float(close_val)

    if not trading_days:
        raise ValueError(
            f"No valid trading day closes found for {clean_ticker}."
        )

    sorted_t_dates = sorted(trading_days.keys())
    price_map: Dict[Any, float] = {}

    for orig_d, target_d in zip(date_list, parsed_dates):
        # Look for closing price on or before target date
        on_or_before = [td for td in sorted_t_dates if td <= target_d]
        if on_or_before:
            matched_date = on_or_before[-1]
        else:
            # Fall back to next available trading day
            after = [td for td in sorted_t_dates if td > target_d]
            if not after:
                raise ValueError(
                    f"No trading day found on, before, or after {target_d}."
                )
            matched_date = after[0]

        price_map[orig_d] = round(trading_days[matched_date], 2)

    return price_map


def fetch_xbrl_concept(
    cik: str,
    tag: str,
    taxonomy: str = "us-gaap",
    filing_dates: Optional[Dict[str, str]] = None,
) -> Dict[str, float]:
    """Fetch concept values from SEC EDGAR XBRL company-concept API.

    Queries https://data.sec.gov/api/xbrl/companyconcept/ for a given CIK
    and concept tag. Filters exclusively for 10-K filings. For duration
    concepts (having both 'start' and 'end' dates), requires a duration of
    at least 300 days to prevent quarterly disclosures from leaking into
    annual totals. If multiple revisions exist for the same end date,
    prioritizes the filing date corresponding to that period's own 10-K.

    Args:
        cik: 10-digit zero-padded CIK string or numeric string.
        tag: The XBRL concept tag name (e.g. 'AssetsCurrent').
        taxonomy: The XBRL taxonomy name (default: 'us-gaap').
        filing_dates: Optional dict mapping period end date to expected
            10-K filing date.

    Returns:
        Dict[str, float]: Mapping from fiscal period end date ('YYYY-MM-DD')
            to numeric concept value.

    Raises:
        requests.RequestException: If network requests fail.
    """
    headers = get_sec_headers()
    cik_padded = str(cik).strip().zfill(10)
    url = (
        f"https://data.sec.gov/api/xbrl/companyconcept/"
        f"CIK{cik_padded}/{taxonomy}/{tag}.json"
    )

    try:
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code == 404:
            return {}
        resp.raise_for_status()
        payload = resp.json()
    except Exception as exc:
        raise requests.RequestException(
            f"Failed to fetch concept '{tag}' for CIK {cik_padded}: {exc}"
        ) from exc

    units = payload.get("units", {})
    if not units:
        return {}

    # Standard currency is USD (first unit entry)
    unit_key = "USD" if "USD" in units else list(units.keys())[0]
    items = units.get(unit_key, [])

    annual_data: Dict[str, Dict[str, Any]] = {}

    for it in items:
        # Only accept annual 10-K filings
        if it.get("form") != "10-K":
            continue

        # If item has start and end, enforce annual duration (300+ days)
        # to ensure quarterly disclosures (e.g. Q4 ~90 days) cannot leak in
        if "start" in it and "end" in it:
            try:
                d_start = pd.to_datetime(it["start"])
                d_end = pd.to_datetime(it["end"])
                if (d_end - d_start).days < 300:
                    continue
            except Exception:
                continue

        end_date = it.get("end")
        val = it.get("val")
        filed = it.get("filed", "")

        if end_date and val is not None:
            # If target filing date specified, prioritize the exact filing
            is_target = bool(
                filing_dates and filing_dates.get(end_date) == filed
            )
            has_target = bool(
                filing_dates
                and end_date in annual_data
                and annual_data[end_date].get("is_target")
            )

            if is_target:
                annual_data[end_date] = {
                    "val": float(val),
                    "filed": filed,
                    "is_target": True,
                }
            elif not has_target:
                if (
                    end_date not in annual_data
                    or filed > annual_data[end_date]["filed"]
                ):
                    annual_data[end_date] = {
                        "val": float(val),
                        "filed": filed,
                        "is_target": False,
                    }

    # Respect SEC rate limits politely
    time.sleep(0.15)
    return {end: entry["val"] for end, entry in annual_data.items()}


def fetch_fundamentals(
    ticker: str,
    cik: Optional[str] = None,
    target_dates: Optional[Sequence[Any]] = None,
    filing_dates_map: Optional[Dict[str, str]] = None,
) -> pd.DataFrame:
    """Fetch fundamental line items from SEC XBRL company-concept API.

    Retrieves balance sheet and income statement concepts required to
    compute financial ratios:
        - Current Assets ('AssetsCurrent')
        - Current Liabilities ('LiabilitiesCurrent')
        - Total Liabilities ('Liabilities' or Assets - Total Equity)
        - Stockholders' Equity ('StockholdersEquity' or total equity)
        - Total Assets ('Assets')
        - Net Income ('NetIncomeLoss' / 'ProfitLoss')
        - Revenues ('RevenueFromContractWithCustomerExcludingAssessedTax',
          'SalesRevenueNet', 'Revenues')

    Enforces annual durations (300+ days) on duration items so quarterly
    values cannot leak into annual numbers. Filters strictly for target
    fiscal years matching each filing's own fiscal-year end (excluding
    prior-year comparative columns). Prioritizes values reported in each
    period's own 10-K filing date.

    Args:
        ticker: The stock ticker symbol (e.g. 'AAPL').
        cik: Optional 10-digit CIK. Looked up if omitted.
        target_dates: Optional sequence of target fiscal-year end dates.
            If omitted, defaults to the own fiscal-year ends of the 10
            most recent 10-K filings.
        filing_dates_map: Optional mapping from period end date to 10-K
            filing date to ensure original reported values are used.

    Returns:
        pd.DataFrame: DataFrame indexed by fiscal end date ('YYYY-MM-DD')
            containing fundamental line items as columns.

    Raises:
        ValueError: If CIK cannot be found or fundamentals are unavailable.
        requests.RequestException: If API requests fail.
    """
    headers = get_sec_headers()
    if not cik:
        cik = get_company_cik(ticker, headers)
    cik_padded = str(cik).strip().zfill(10)

    # Derive filing_dates mapping from target_dates or recent filings
    # to ensure each period uses the value reported in that filing's own 10-K
    if filing_dates_map is None:
        try:
            recent_10ks = get_recent_10k_filings(ticker, n=10)
            filing_dates_map = {
                f["report_date"]: f["filing_date"]
                for f in recent_10ks
                if f.get("report_date") and f.get("filing_date")
            }
        except Exception:
            filing_dates_map = {}

    # 1. Balance Sheet line items (instant facts)
    assets_curr = fetch_xbrl_concept(
        cik_padded, "AssetsCurrent", filing_dates=filing_dates_map
    )
    liab_curr = fetch_xbrl_concept(
        cik_padded, "LiabilitiesCurrent", filing_dates=filing_dates_map
    )
    liab_tot = fetch_xbrl_concept(
        cik_padded, "Liabilities", filing_dates=filing_dates_map
    )

    # Fetch total consolidated equity including portion attributable to NCI
    equity_total = fetch_xbrl_concept(
        cik_padded,
        "StockholdersEquityIncludingPortionAttributableTo"
        "NoncontrollingInterest",
        filing_dates=filing_dates_map,
    )
    equity_parent = fetch_xbrl_concept(
        cik_padded, "StockholdersEquity", filing_dates=filing_dates_map
    )
    equity: Dict[str, float] = {}
    all_eq_dates = set(list(equity_parent.keys()) + list(equity_total.keys()))
    for d in sorted(all_eq_dates):
        if d in equity_total:
            equity[d] = equity_total[d]
        elif d in equity_parent:
            equity[d] = equity_parent[d]

    assets = fetch_xbrl_concept(
        cik_padded, "Assets", filing_dates=filing_dates_map
    )
    if not liab_tot and assets and equity:
        # Total Liabilities = Total Assets - Total Stockholders' Equity
        # (strictly excludes Non-Controlling Interests)
        liab_tot = {
            d: assets[d] - equity[d]
            for d in assets
            if d in equity
        }

    # 2. Income Statement line items (duration facts)
    net_inc = fetch_xbrl_concept(
        cik_padded, "NetIncomeLoss", filing_dates=filing_dates_map
    )
    if not net_inc:
        net_inc = fetch_xbrl_concept(
            cik_padded, "ProfitLoss", filing_dates=filing_dates_map
        )

    rev_asc606 = fetch_xbrl_concept(
        cik_padded,
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        filing_dates=filing_dates_map,
    )
    rev_sales = fetch_xbrl_concept(
        cik_padded, "SalesRevenueNet", filing_dates=filing_dates_map
    )
    rev_gen = fetch_xbrl_concept(
        cik_padded, "Revenues", filing_dates=filing_dates_map
    )

    # Merge revenue concepts prioritizing ASC606, then sales, then general
    revenues: Dict[str, float] = {}
    for d, v in rev_gen.items():
        revenues[d] = v
    for d, v in rev_sales.items():
        revenues[d] = v
    for d, v in rev_asc606.items():
        revenues[d] = v

    # Collect union of all available fiscal end dates
    all_dates = sorted(
        set(
            list(assets_curr.keys())
            + list(liab_curr.keys())
            + list(liab_tot.keys())
            + list(equity.keys())
            + list(assets.keys())
            + list(net_inc.keys())
            + list(revenues.keys())
        ),
        reverse=True,
    )

    rows: List[Dict[str, Any]] = []
    for d in all_dates:
        rows.append(
            {
                "date": d,
                "current_assets": assets_curr.get(d),
                "current_liabilities": liab_curr.get(d),
                "total_liabilities": liab_tot.get(d),
                "equity": equity.get(d),
                "total_assets": assets.get(d),
                "revenue": revenues.get(d),
                "net_income": net_inc.get(d),
            }
        )

    df = pd.DataFrame(rows).set_index("date")
    df = df.dropna(how="all")
    if df.empty:
        raise ValueError(
            f"No fundamental items could be retrieved for {ticker} "
            f"(CIK {cik_padded})."
        )

    # Filter strictly to the 10 target fiscal years (matching each filing's
    # own fiscal-year end, dropping comparative prior-year columns)
    if target_dates is None:
        recent_10ks = get_recent_10k_filings(ticker, n=10)
        target_dates = [
            f["report_date"] for f in recent_10ks if f.get("report_date")
        ]

    if target_dates:
        target_strs = [
            _to_date(d).strftime("%Y-%m-%d") for d in target_dates
        ]
        available_targets = [d for d in target_strs if d in df.index]
        if available_targets:
            df = df.loc[available_targets]

    return df


def compute_ratios(
    fundamentals: Union[pd.DataFrame, Dict[str, Any]],
) -> Union[pd.DataFrame, Dict[str, Any]]:
    """Compute financial ratios from company fundamentals.

    Computes four key financial ratios:
        1. Current Ratio: Current Assets / Current Liabilities (liquidity)
        2. Total-Liabilities-to-Equity: Total Liabilities / Equity (solvency)
        3. Asset Turnover: Revenues / Total Assets (efficiency)
        4. Net Profit Margin: Net Income / Revenues (profitability)

    Uses TOTAL liabilities for solvency as specified. Handles both
    pandas DataFrame (multi-period) and dictionary representations.
    Safely handles division by zero or missing values by returning NaN/None.

    Args:
        fundamentals: DataFrame or dict containing fundamental line items:
            'current_assets', 'current_liabilities', 'total_liabilities',
            'equity', 'revenue', 'total_assets', 'net_income'.

    Returns:
        Union[pd.DataFrame, Dict[str, Any]]: Computed ratios indexed by
            period/date with columns/keys:
            - 'current_ratio'
            - 'total_liabilities_to_equity'
            - 'asset_turnover'
            - 'net_profit_margin'

    Raises:
        TypeError: If fundamentals is neither a DataFrame nor a dictionary.
        KeyError: If required fundamental line items cannot be located.
    """
    if isinstance(fundamentals, pd.DataFrame):
        df = fundamentals.copy()
        col_lookup: Dict[str, str] = {}
        for col in df.columns:
            clean_key = (
                str(col).lower().replace(" ", "_").replace("-", "_")
            )
            col_lookup[clean_key] = col

        def _get_series(candidates: List[str]) -> pd.Series:
            for cand in candidates:
                if cand in col_lookup:
                    return pd.to_numeric(
                        df[col_lookup[cand]], errors="coerce"
                    )
                if cand in df.columns:
                    return pd.to_numeric(df[cand], errors="coerce")
            raise KeyError(
                f"Missing required fundamental metric. Expected one of "
                f"{candidates}, found: {list(df.columns)}"
            )

        ca = _get_series(["current_assets", "assetscurrent"])
        cl = _get_series(["current_liabilities", "liabilitiescurrent"])
        tl = _get_series(["total_liabilities", "liabilities"])
        eq = _get_series(
            [
                "equity",
                "stockholdersequity",
                "stockholders_equity",
                "shareholdersequity",
            ]
        )
        rev = _get_series(
            ["revenue", "revenues", "sales", "net_sales", "salesrevenuenet"]
        )
        ta = _get_series(["total_assets", "assets"])
        ni = _get_series(["net_income", "netincomeloss", "profit_loss"])

        ratios = pd.DataFrame(index=df.index)
        ratios["current_ratio"] = ca / cl.replace(0, np.nan)
        ratios["total_liabilities_to_equity"] = tl / eq.replace(0, np.nan)
        ratios["asset_turnover"] = rev / ta.replace(0, np.nan)
        ratios["net_profit_margin"] = ni / rev.replace(0, np.nan)
        return ratios

    elif isinstance(fundamentals, dict):
        if not fundamentals:
            return {}
        first_val = next(iter(fundamentals.values()))
        if isinstance(first_val, dict):
            return {
                period: compute_ratios(sub_dict)
                for period, sub_dict in fundamentals.items()
            }

        d_norm = {
            str(k).lower().replace(" ", "_").replace("-", "_"): v
            for k, v in fundamentals.items()
        }

        def _get_val(candidates: List[str]) -> Optional[float]:
            for cand in candidates:
                if cand in d_norm and d_norm[cand] is not None:
                    try:
                        return float(d_norm[cand])
                    except (ValueError, TypeError):
                        pass
            return None

        ca = _get_val(["current_assets", "assetscurrent"])
        cl = _get_val(["current_liabilities", "liabilitiescurrent"])
        tl = _get_val(["total_liabilities", "liabilities"])
        eq = _get_val(
            [
                "equity",
                "stockholdersequity",
                "stockholders_equity",
                "shareholdersequity",
            ]
        )
        rev = _get_val(
            ["revenue", "revenues", "sales", "net_sales", "salesrevenuenet"]
        )
        ta = _get_val(["total_assets", "assets"])
        ni = _get_val(["net_income", "netincomeloss", "profit_loss"])

        return {
            "current_ratio": (ca / cl) if (ca is not None and cl) else None,
            "total_liabilities_to_equity": (
                (tl / eq) if (tl is not None and eq) else None
            ),
            "asset_turnover": (
                (rev / ta) if (rev is not None and ta) else None
            ),
            "net_profit_margin": (
                (ni / rev) if (ni is not None and rev) else None
            ),
        }
    else:
        raise TypeError(
            f"fundamentals must be a pd.DataFrame or dict, got "
            f"{type(fundamentals)}"
        )


def build_workbook(
    ticker: str,
    filings: Optional[List[Dict[str, Any]]] = None,
    statements: Optional[Dict[str, pd.DataFrame]] = None,
    sentiment_scores: Optional[List[Dict[str, Any]]] = None,
    stock_prices: Optional[Dict[Any, float]] = None,
    ratios_df: Optional[pd.DataFrame] = None,
    fundamentals_df: Optional[pd.DataFrame] = None,
    output_dir: str = "output",
) -> str:
    """Build a comprehensive Excel analysis workbook with eight tabs.

    Creates an Excel workbook at 'output/{ticker}_10K_Analysis.xlsx' with
    the following eight tabs:
        1. README: Project documentation, pipeline overview, tab directory,
           and ratio formulas.
        2. Income Statement: Consolidated Statements of Operations showing
           all 10 fiscal years side by side.
        3. Balance Sheet: Consolidated Statements of Financial Position
           showing all 10 fiscal years side by side.
        4. Cash Flows: Consolidated Statements of Cash Flows showing all
           10 fiscal years side by side.
        5. Sentiment Analysis: Loughran-McDonald sentiment scores, category
           breakdowns, and P/N ratios for all 10 years.
        6. Stock Prices: Split-adjusted closing stock prices on or before
           each fiscal year-end balance sheet date.
        7. Ratios: Four core financial ratios (Current Ratio, Total-
           Liabilities-to-Equity, Asset Turnover, Net Profit Margin) plus
           underlying XBRL fundamentals for all 10 fiscal years.
        8. Chart: Dual-axis line chart plotting Stock Price ($) on the
           primary axis and MD&A Sentiment P/N Ratio on the secondary axis
           across all 10 years.

    Args:
        ticker: The stock ticker symbol (e.g. 'AAPL').
        filings: Optional list of 10-K filing metadata dictionaries.
        statements: Optional mapping of multi-year statement DataFrames
            ('operations', 'balance_sheet', 'cash_flows').
        sentiment_scores: Optional list of sentiment score dictionaries.
        stock_prices: Optional mapping of balance sheet dates to stock prices.
        ratios_df: Optional DataFrame of computed financial ratios.
        fundamentals_df: Optional DataFrame of XBRL fundamental metrics.
        output_dir: Directory where the workbook is saved (default: 'output').

    Returns:
        str: File path to the generated Excel workbook.

    Raises:
        ValueError: If required data cannot be retrieved or generated.
        IOError: If saving the workbook fails.
    """
    clean_ticker = ticker.strip().upper()

    # 1. Ensure filings metadata is available
    if filings is None:
        filings = get_recent_10k_filings(clean_ticker, n=10)

    target_dates = [
        f["report_date"] for f in filings if f.get("report_date")
    ]

    # 2. Multi-year financial statements
    if statements is None:
        statements = {}

    req_stmts = ["operations", "balance_sheet", "cash_flows"]
    needs_multi_year = any(
        k not in statements or statements[k].shape[1] < len(filings)
        for k in req_stmts
    )

    if needs_multi_year:
        def _dedup_index(df: pd.DataFrame) -> pd.DataFrame:
            counts: Dict[str, int] = {}
            new_idx: List[str] = []
            for lbl in df.index:
                clean_lbl = (
                    str(lbl)
                    .replace("’", "'")
                    .replace("‘", "'")
                    .replace("“", '"')
                    .replace("”", '"')
                    .strip()
                )
                if clean_lbl in counts:
                    counts[clean_lbl] += 1
                    new_idx.append(f"{clean_lbl} ({counts[clean_lbl]})")
                else:
                    counts[clean_lbl] = 0
                    new_idx.append(clean_lbl)
            df_out = df.copy()
            df_out.index = new_idx
            return df_out

        all_filing_stmts = [extract_statements(f) for f in filings]
        for stmt_key in req_stmts:
            s_list: List[pd.Series] = []
            for stmts_dict in all_filing_stmts:
                s_df = stmts_dict.get(stmt_key)
                if s_df is not None and not s_df.empty:
                    s_u = _dedup_index(s_df)
                    own_col = s_u.columns[0]
                    col_s = s_u.iloc[:, 0].copy()
                    col_s.name = own_col
                    s_list.append(col_s)
            if s_list:
                comb = pd.concat(s_list, axis=1).dropna(how="all")
                statements[stmt_key] = comb

    # 3. Sentiment scores
    if sentiment_scores is None:
        dict_path = "data/Loughran-McDonald_MasterDictionary_1993-2025.csv"
        lm_dict = load_loughran_mcdonald_dictionary(dict_path)
        sentiment_scores = []
        folder = os.path.join("filings", clean_ticker)
        for f in filings:
            f_year = f["report_date"][:4] if f.get("report_date") else ""
            f_date = f["filing_date"]
            target_fname = f"{clean_ticker}_10K_FY{f_year}_{f_date}_mda.txt"
            mda_file = os.path.join(folder, target_fname)
            if not os.path.exists(mda_file) and os.path.exists(folder):
                candidates = [
                    os.path.join(folder, x)
                    for x in os.listdir(folder)
                    if x.endswith("_mda.txt") and f_date in x
                ]
                if candidates:
                    mda_file = candidates[0]

            if os.path.exists(mda_file):
                with open(mda_file, "r", encoding="utf-8") as fp:
                    txt = fp.read()
            else:
                txt = ""

            sc = score_mda(txt, lm_dict)
            sc["filing_date"] = f_date
            sc["report_date"] = f.get("report_date", "")
            sc["fiscal_year"] = f"FY{f_year}" if f_year else f_date[:4]
            sentiment_scores.append(sc)

    # Ensure filing metadata is present in sentiment_scores
    for idx, s in enumerate(sentiment_scores):
        if idx < len(filings):
            f_item = filings[idx]
            if not s.get("filing_date"):
                s["filing_date"] = f_item.get("filing_date", "")
            if not s.get("report_date"):
                s["report_date"] = f_item.get("report_date", "")
            if not s.get("fiscal_year"):
                rep_d = f_item.get("report_date", "")
                s["fiscal_year"] = f"FY{rep_d[:4]}" if rep_d else ""

    # 4. Stock prices
    price_lookup: Dict[str, float] = {}
    if stock_prices:
        for k, v in stock_prices.items():
            try:
                dt_k = _to_date(k)
                price_lookup[dt_k.strftime("%Y-%m-%d")] = v
            except Exception:
                pass

    missing_dates = [d for d in target_dates if d not in price_lookup]
    if missing_dates:
        try:
            fetched = fetch_stock_prices(clean_ticker, missing_dates)
            for k, v in fetched.items():
                try:
                    dt_k = _to_date(k)
                    price_lookup[dt_k.strftime("%Y-%m-%d")] = v
                except Exception:
                    pass
        except Exception:
            pass

    stock_prices = price_lookup

    # 5. Fundamentals and Ratios
    filing_dates_map = {
        f["report_date"]: f["filing_date"]
        for f in (filings or [])
        if f.get("report_date") and f.get("filing_date")
    }
    if fundamentals_df is None:
        fundamentals_df = fetch_fundamentals(
            clean_ticker,
            target_dates=target_dates,
            filing_dates_map=filing_dates_map,
        )

    # Reconcile fundamentals with balance sheet statement if present
    if statements and "balance_sheet" in statements:
        bs_df = statements["balance_sheet"]
        ca_row = None
        for idx_lbl in bs_df.index:
            if idx_lbl.strip().lower() == "total current assets":
                ca_row = bs_df.loc[idx_lbl]
                break
        if ca_row is not None:
            for col_name in bs_df.columns:
                dt = _to_date(col_name)
                if dt:
                    dt_str = dt.strftime("%Y-%m-%d")
                    if dt_str in fundamentals_df.index:
                        bs_val = ca_row[col_name]
                        if pd.notna(bs_val):
                            fundamentals_df.loc[dt_str, "current_assets"] = (
                                float(bs_val) * 1e6
                            )

    ratios_df = compute_ratios(fundamentals_df)

    # Initialize Excel workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    # Common styling palette
    navy_fill = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    section_font = Font(name="Calibri", size=12, bold=True, color="1F497D")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    zebra_fill = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # ----------------- TAB 1: README -----------------
    ws_readme = wb.create_sheet(title="README")
    ws_readme.views.sheetView[0].showGridLines = True

    f_start = target_dates[-1][:4] if target_dates else ""
    f_end = target_dates[0][:4] if target_dates else ""
    cov_str = f"10 Most Recent 10-K Filings ({f_start} - {f_end})"

    readme_rows = [
        [f"{clean_ticker} — 10-K Ten-Year Financial & Sentiment Analysis"],
        [""],
        ["Project Overview & Methodology"],
        ["Company Ticker", clean_ticker],
        ["CIK", filings[0].get("cik", "") if filings else ""],
        ["Filings Covered", cov_str],
        [
            "Data Sources",
            "SEC EDGAR, XBRL API, Loughran-McDonald Dictionary, Yahoo Finance",
        ],
        [""],
        ["Workbook Tabs Directory"],
        ["Tab Name", "Description"],
        [
            "README",
            "Overview of pipeline, methodology, tab directory, and "
            "ratio definitions.",
        ],
        [
            "Income Statement",
            "Consolidated Statements of Operations across all 10 fiscal "
            "years side by side.",
        ],
        [
            "Balance Sheet",
            "Consolidated Statements of Financial Position across all 10 "
            "fiscal years side by side.",
        ],
        [
            "Cash Flows",
            "Consolidated Statements of Cash Flows across all 10 fiscal "
            "years side by side.",
        ],
        [
            "Sentiment Analysis",
            "MD&A narrative scored against 7 Loughran-McDonald "
            "categories and P/N ratio.",
        ],
        [
            "Stock Prices",
            "Closing stock prices on or before each fiscal year-end "
            "balance sheet date.",
        ],
        [
            "Ratios",
            "Current Ratio, Total-Liabilities-to-Equity, Asset Turnover, "
            "and Net Profit Margin.",
        ],
        [
            "Chart",
            "Dual-axis line chart tracking Stock Price ($) and MD&A "
            "Sentiment P/N Ratio across all years.",
        ],
        [""],
        ["Financial Ratio Formulas & Economic Interpretation"],
        ["Ratio", "Formula", "Classification", "Description"],
        [
            "Current Ratio",
            "Current Assets / Current Liabilities",
            "Liquidity",
            "Measures ability to cover short-term obligations using "
            "liquid assets.",
        ],
        [
            "Total-Liabilities-to-Equity",
            "Total Liabilities / Stockholders' Equity",
            "Solvency",
            "Measures long-term leverage; Total Liabilities excludes "
            "non-controlling interests.",
        ],
        [
            "Asset Turnover",
            "Revenues / Total Assets",
            "Efficiency",
            "Measures revenue generated per dollar of asset base.",
        ],
        [
            "Net Profit Margin",
            "Net Income / Revenues",
            "Profitability",
            "Measures percentage of revenue converted into bottom-line "
            "profit.",
        ],
        [
            "P/N Ratio",
            "Positive Words / Negative Words",
            "Sentiment",
            "Measures narrative tone and executive optimism in MD&A.",
        ],
        [""],
        ["Accounting Definitions & Balance Sheet Reconciliation"],
        [
            "Concept",
            "Definition & Measurement Method",
            "Intel Impact",
            "Apple Impact",
        ],
        [
            "Current Assets ($M)",
            "Total current assets matching each filing's own 10-K "
            "balance sheet.",
            "FY2021 reads $57,718M (matching Balance Sheet tab, avoiding "
            "restatements).",
            "Matches Balance Sheet tab across all 10 fiscal years.",
        ],
        [
            "Total Liabilities ($M)",
            "Total Assets minus Total Equity (strictly excludes "
            "non-controlling interests).",
            "FY2025 reads $85,069M, properly excluding $12,079M of "
            "non-controlling interests.",
            "Directly reported us-gaap/Liabilities ($285,508M in FY2025).",
        ],
        [
            "Stockholders' Equity ($M)",
            "Total consolidated stockholders' equity (inclusive of "
            "non-controlling interests).",
            "FY2025 reads $126,360M ($114,281M parent equity + $12,079M "
            "non-controlling interest).",
            "Total shareholders' equity ($73,733M in FY2025).",
        ],
        [
            "Solvency Ratio",
            "Total Liabilities / Total Stockholders' Equity.",
            "FY2025 = 0.6732 ($85,069M / $126,360M); "
            "FY2021 = 0.7654 ($73,015M / $95,391M).",
            "FY2025 = 3.8722 ($285,508M / $73,733M).",
        ],
    ]

    for r in readme_rows:
        ws_readme.append(r)

    ws_readme["A1"].font = title_font

    for r_idx in range(1, ws_readme.max_row + 1):
        c1 = ws_readme.cell(row=r_idx, column=1).value
        if c1 in [
            "Project Overview & Methodology",
            "Workbook Tabs Directory",
            "Financial Ratio Formulas & Economic Interpretation",
            "Accounting Definitions & Balance Sheet Reconciliation",
        ]:
            ws_readme.cell(row=r_idx, column=1).font = section_font
        elif c1 in ["Tab Name", "Ratio", "Concept"]:
            for col_idx in range(1, ws_readme.max_column + 1):
                cell = ws_readme.cell(row=r_idx, column=col_idx)
                if cell.value:
                    cell.fill = navy_fill
                    cell.font = header_font

    ws_readme.column_dimensions["A"].width = 28
    ws_readme.column_dimensions["B"].width = 50
    ws_readme.column_dimensions["C"].width = 25
    ws_readme.column_dimensions["D"].width = 70

    # ----------------- TABS 2-4: Financial Statements -----------------
    def _write_statement_sheet(
        ws: Any, sheet_title: str, df: pd.DataFrame
    ) -> None:
        ws.views.sheetView[0].showGridLines = True
        ws.append([f"{clean_ticker} — {sheet_title} (USD in Millions)"])
        ws.append([])
        ws["A1"].font = title_font

        hdr = ["Line Item"] + list(df.columns)
        ws.append(hdr)
        h_row = 3
        for col_idx in range(1, len(hdr) + 1):
            cell = ws.cell(row=h_row, column=col_idx)
            cell.fill = navy_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center" if col_idx > 1 else "left"
            )

        for r_idx, (idx_label, row_data) in enumerate(
            df.iterrows(), start=h_row + 1
        ):
            row_vals = [idx_label] + [row_data[c] for c in df.columns]
            ws.append(row_vals)
            cur_row = ws[r_idx]
            is_zebra = (r_idx % 2 == 0)
            for c_idx, cell in enumerate(cur_row, start=1):
                cell.border = thin_border
                if c_idx == 1:
                    is_tot = any(
                        k in idx_label.lower()
                        for k in [
                            "total",
                            "operating income",
                            "net income",
                            "gross margin",
                        ]
                    )
                    cell.font = bold_font if is_tot else regular_font
                else:
                    cell.number_format = '#,##0;(#,##0);"-"'
                    cell.alignment = Alignment(horizontal="right")
                    if is_zebra:
                        cell.fill = zebra_fill

        ws.column_dimensions["A"].width = 45
        for c_idx in range(2, len(hdr) + 1):
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = 16

    ws_ops = wb.create_sheet(title="Income Statement")
    _write_statement_sheet(
        ws_ops,
        "Consolidated Statements of Operations",
        statements.get("operations", pd.DataFrame()),
    )

    ws_bs = wb.create_sheet(title="Balance Sheet")
    _write_statement_sheet(
        ws_bs,
        "Consolidated Statements of Financial Position",
        statements.get("balance_sheet", pd.DataFrame()),
    )

    ws_cf = wb.create_sheet(title="Cash Flows")
    _write_statement_sheet(
        ws_cf,
        "Consolidated Statements of Cash Flows",
        statements.get("cash_flows", pd.DataFrame()),
    )

    # ----------------- TAB 5: Sentiment Analysis -----------------
    ws_sent = wb.create_sheet(title="Sentiment Analysis")
    ws_sent.views.sheetView[0].showGridLines = True
    ws_sent.append(
        [f"{clean_ticker} — MD&A Loughran-McDonald Sentiment Analysis"]
    )
    ws_sent.append([])
    ws_sent["A1"].font = title_font

    sent_headers = [
        "Fiscal Year", "Filing Date", "Report Date", "Total Words",
        "Positive", "Pos %", "Negative", "Neg %", "P/N Ratio",
        "Uncertainty", "Unc %", "Litigious", "Lit %",
        "Strong Modal", "Modal %", "Weak Modal", "Weak %",
        "Constraining", "Const %",
    ]
    ws_sent.append(sent_headers)
    h_row = 3
    for col_idx in range(1, len(sent_headers) + 1):
        c = ws_sent.cell(row=h_row, column=col_idx)
        c.fill = navy_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")

    for idx, (f_item, s) in enumerate(
        zip(filings, sentiment_scores), start=h_row + 1
    ):
        rep_date = s.get("report_date") or f_item.get("report_date", "")
        filing_dt = s.get("filing_date") or f_item.get("filing_date", "")
        fy_val = (
            s.get("fiscal_year")
            or (f"FY{rep_date[:4]}" if rep_date else "")
        )
        row_vals = [
            fy_val,
            filing_dt,
            rep_date,
            s.get("total_words", 0),
            s.get("positive_count", 0),
            s.get("positive_pct", 0.0) / 100.0,
            s.get("negative_count", 0),
            s.get("negative_pct", 0.0) / 100.0,
            s.get("pn_ratio", 0.0),
            s.get("uncertainty_count", 0),
            s.get("uncertainty_pct", 0.0) / 100.0,
            s.get("litigious_count", 0),
            s.get("litigious_pct", 0.0) / 100.0,
            s.get("strong_modal_count", 0),
            s.get("strong_modal_pct", 0.0) / 100.0,
            s.get("weak_modal_count", 0),
            s.get("weak_modal_pct", 0.0) / 100.0,
            s.get("constraining_count", 0),
            s.get("constraining_pct", 0.0) / 100.0,
        ]
        ws_sent.append(row_vals)
        cur_row = ws_sent[idx]
        is_zebra = (idx % 2 == 0)
        for c_idx, cell in enumerate(cur_row, start=1):
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            if c_idx in [1, 2, 3]:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx in [4, 5, 7, 10, 12, 14, 16, 18]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx in [6, 8, 11, 13, 15, 17, 19]:
                cell.number_format = "0.00%"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx == 9:
                cell.number_format = "0.000"
                cell.alignment = Alignment(horizontal="right")
                cell.font = bold_font

    for col in ws_sent.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws_sent.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # ----------------- TAB 6: Stock Prices -----------------
    ws_prices = wb.create_sheet(title="Stock Prices")
    ws_prices.views.sheetView[0].showGridLines = True
    ws_prices.append(
        [f"{clean_ticker} — Balance Sheet Date Closing Stock Prices"]
    )
    ws_prices.append([])
    ws_prices["A1"].font = title_font

    price_headers = [
        "Fiscal Year", "Balance Sheet Date", "Day of Week",
        "Closing Price ($)",
    ]
    ws_prices.append(price_headers)
    h_row = 3
    for col_idx in range(1, len(price_headers) + 1):
        c = ws_prices.cell(row=h_row, column=col_idx)
        c.fill = navy_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")

    for idx, (f, d) in enumerate(zip(filings, target_dates), start=h_row + 1):
        p_date = _to_date(d)
        day_name = p_date.strftime("%A")
        price = stock_prices.get(d)
        row_vals = [f"FY{d[:4]}", d, day_name, price]
        ws_prices.append(row_vals)
        cur_row = ws_prices[idx]
        is_zebra = (idx % 2 == 0)
        for c_idx, cell in enumerate(cur_row, start=1):
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            if c_idx in [1, 2, 3]:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx == 4:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
                cell.font = bold_font

    ws_prices.column_dimensions["A"].width = 16
    ws_prices.column_dimensions["B"].width = 22
    ws_prices.column_dimensions["C"].width = 16
    ws_prices.column_dimensions["D"].width = 20

    # ----------------- TAB 7: Ratios -----------------
    ws_ratios = wb.create_sheet(title="Ratios")
    ws_ratios.views.sheetView[0].showGridLines = True
    ws_ratios.append(
        [f"{clean_ticker} — Ten-Year Financial Ratios (FY2016 – FY2025)"]
    )
    ws_ratios.append([])
    ws_ratios["A1"].font = title_font

    ratio_headers = [
        "Fiscal Year", "Fiscal Period End",
        "Current Ratio", "Total-Liabilities-to-Equity",
        "Asset Turnover", "Net Profit Margin",
        "Current Assets ($M)", "Current Liabilities ($M)",
        "Total Liabilities ($M)", "Stockholders' Equity ($M)",
        "Total Assets ($M)", "Revenue ($M)", "Net Income ($M)",
    ]
    ws_ratios.append(ratio_headers)
    h_row = 3
    for col_idx in range(1, len(ratio_headers) + 1):
        c = ws_ratios.cell(row=h_row, column=col_idx)
        c.fill = navy_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")

    def _to_m(val: Any) -> Optional[float]:
        if val is None or pd.isna(val):
            return None
        return float(val) / 1e6

    for idx, (date_idx, r_row) in enumerate(
        ratios_df.iterrows(), start=h_row + 1
    ):
        f_row = fundamentals_df.loc[date_idx]
        row_vals = [
            f"FY{str(date_idx)[:4]}", str(date_idx),
            r_row["current_ratio"], r_row["total_liabilities_to_equity"],
            r_row["asset_turnover"], r_row["net_profit_margin"],
            _to_m(f_row.get("current_assets")),
            _to_m(f_row.get("current_liabilities")),
            _to_m(f_row.get("total_liabilities")),
            _to_m(f_row.get("equity")),
            _to_m(f_row.get("total_assets")),
            _to_m(f_row.get("revenue")),
            _to_m(f_row.get("net_income")),
        ]
        ws_ratios.append(row_vals)
        cur_row = ws_ratios[idx]
        is_zebra = (idx % 2 == 0)
        for c_idx, cell in enumerate(cur_row, start=1):
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            if c_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx in [3, 4, 5]:
                cell.number_format = "0.0000"
                cell.alignment = Alignment(horizontal="right")
                cell.font = bold_font
            elif c_idx == 6:
                cell.number_format = "0.00%"
                cell.alignment = Alignment(horizontal="right")
                cell.font = bold_font
            else:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")

    for col in ws_ratios.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws_ratios.column_dimensions[col_letter].width = max(max_len + 3, 14)

    # ----------------- TAB 8: Chart -----------------
    ws_chart = wb.create_sheet(title="Chart")
    ws_chart.views.sheetView[0].showGridLines = True
    ws_chart.append(
        [f"{clean_ticker} — Stock Price vs. MD&A Positive / Negative Ratio"]
    )
    ws_chart.append([])
    ws_chart["A1"].font = title_font

    chart_tbl_headers = ["Fiscal Year", "Stock Price ($)", "MD&A P/N Ratio"]
    ws_chart.append(chart_tbl_headers)
    h_row = 3
    for col_idx in range(1, len(chart_tbl_headers) + 1):
        c = ws_chart.cell(row=h_row, column=col_idx)
        c.fill = navy_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")

    # Chronological sort for the chart (FY2016 -> FY2025)
    filings_chrono = sorted(
        filings, key=lambda x: x.get("report_date", "")
    )
    pn_map: Dict[str, float] = {}
    for f_item, s_item in zip(filings, sentiment_scores):
        rd = f_item.get("report_date", "") or s_item.get("report_date", "")
        if rd:
            pn_map[rd] = s_item.get("pn_ratio", 0.0)

    for idx, f in enumerate(filings_chrono, start=h_row + 1):
        d = f.get("report_date", "")
        fy_label = f"FY{d[:4]}" if d else ""
        p = stock_prices.get(d, 0.0)
        pn = pn_map.get(d, 0.0)
        ws_chart.append([fy_label, p, pn])
        cur_row = ws_chart[idx]
        cur_row[0].alignment = Alignment(horizontal="center")
        cur_row[0].border = thin_border
        cur_row[1].number_format = "$#,##0.00"
        cur_row[1].border = thin_border
        cur_row[2].number_format = "0.000"
        cur_row[2].border = thin_border

    ws_chart.column_dimensions["A"].width = 16
    ws_chart.column_dimensions["B"].width = 18
    ws_chart.column_dimensions["C"].width = 18

    # Create dual-axis line chart
    c1 = LineChart()
    c1.title = (
        f"{clean_ticker}: Stock Price vs. MD&A P/N Ratio "
        f"({filings_chrono[0].get('report_date', '')[:4]} – "
        f"{filings_chrono[-1].get('report_date', '')[:4]})"
    )
    c1.style = 13
    c1.y_axis.title = "Stock Price ($)"
    c1.x_axis.title = "Fiscal Year"
    c1.width = 23
    c1.height = 14

    max_r = len(filings_chrono) + h_row
    data_price = Reference(ws_chart, min_col=2, min_row=3, max_row=max_r)
    cats = Reference(ws_chart, min_col=1, min_row=4, max_row=max_r)
    c1.add_data(data_price, titles_from_data=True)
    c1.set_categories(cats)

    c2 = LineChart()
    data_pn = Reference(ws_chart, min_col=3, min_row=3, max_row=max_r)
    c2.add_data(data_pn, titles_from_data=True)
    c2.y_axis.title = "Positive / Negative (P/N) Ratio"
    c2.y_axis.axId = 200
    c2.y_axis.crosses = "max"

    c1 += c2
    ws_chart.add_chart(c1, "E3")

    # Save workbook to output directory
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{clean_ticker}_10K_Analysis.xlsx")
    try:
        wb.save(out_path)
    except Exception as exc:
        raise IOError(
            f"Failed to save Excel workbook to {out_path}: {exc}"
        ) from exc

    return out_path


if __name__ == "__main__":
    target_ticker = "AAPL"
    count = 10

    # Accept ticker and count from command-line arguments
    # Usage: python3 main.py [TICKER] [COUNT]
    # Example: python3 main.py INTC 10
    if len(sys.argv) > 1:
        target_ticker = sys.argv[1].strip().upper()
    if len(sys.argv) > 2:
        try:
            count = int(sys.argv[2])
        except ValueError:
            print(
                f"Warning: Invalid count '{sys.argv[2]}', defaulting to 10.",
                file=sys.stderr,
            )
            count = 10

    print(
        f"Retrieving the {count} most recent 10-K filings for "
        f"ticker: {target_ticker}..."
    )

    try:
        filings = get_recent_10k_filings(ticker=target_ticker, n=count)

        print(
            f"\nFound {len(filings)} 10-K filing(s) for {target_ticker}:\n"
        )
        # Header formatting for clean tabular display
        header = (
            f"{'#':<3} | {'Accession Number':<22} | {'Filing Date':<12} | "
            f"{'Primary Document':<30}"
        )
        divider = "-" * len(header)
        print(header)
        print(divider)

        for idx, f in enumerate(filings, start=1):
            print(
                f"{idx:<3} | {f['accession_number']:<22} | "
                f"{f['filing_date']:<12} | {f['primary_document']:<30}"
            )

        # Stage 2: Download primary documents
        downloaded = download_filings(target_ticker, filings)

        # Stage 3: Extract Item 7 MD&A section
        print(f"\nExtracting Item 7 MD&A sections for {target_ticker}...")
        mda_texts: List[str] = []
        for html_file_path in downloaded:
            mda_file_path = html_file_path.replace(".htm", "_mda.txt")

            # Read existing extracted MD&A or extract on the fly
            if os.path.exists(mda_file_path):
                with open(mda_file_path, "r", encoding="utf-8") as mda_fp:
                    extracted_text = mda_fp.read()
            else:
                with open(
                    html_file_path, "r", encoding="utf-8", errors="ignore"
                ) as html_fp:
                    raw_html_content = html_fp.read()
                extracted_text = extract_mda(raw_html_content)
                with open(mda_file_path, "w", encoding="utf-8") as mda_fp:
                    mda_fp.write(extracted_text)

            mda_texts.append(extracted_text)
            print(
                f"  {os.path.basename(mda_file_path)}: "
                f"{len(extracted_text):,} chars, "
                f"{len(extracted_text.split()):,} words"
            )

        # Stage 4: Loughran-McDonald Sentiment Scoring
        dict_path = "data/Loughran-McDonald_MasterDictionary_1993-2025.csv"
        print(f"\nLoading sentiment dictionary from '{dict_path}'...")
        lm_dictionary = load_loughran_mcdonald_dictionary(dict_path)

        print(
            f"\nLoughran-McDonald Sentiment Scores for {target_ticker} MD&A:\n"
        )
        sentiment_header = (
            f"{'Filing Date':<12} | {'Words':>6} | "
            f"{'Positive':>14} | {'Negative':>14} | {'P/N Ratio':>9} | "
            f"{'Uncertainty':>12} | {'Litigious':>10}"
        )
        print(sentiment_header)
        print("-" * len(sentiment_header))

        for f, mda_text in zip(filings, mda_texts):
            scores = score_mda(mda_text, lm_dictionary)
            scores["filing_date"] = f["filing_date"]
            scores["report_date"] = f.get("report_date", "")
            rep_d = f.get("report_date", "")
            scores["fiscal_year"] = f"FY{rep_d[:4]}" if rep_d else ""
            f["sentiment"] = scores

            pos_info = (
                f"{scores['positive_count']} "
                f"({scores['positive_pct']:.2f}%)"
            )
            neg_info = (
                f"{scores['negative_count']} "
                f"({scores['negative_pct']:.2f}%)"
            )
            unc_info = (
                f"{scores['uncertainty_count']} "
                f"({scores['uncertainty_pct']:.2f}%)"
            )
            lit_info = (
                f"{scores['litigious_count']} "
                f"({scores['litigious_pct']:.2f}%)"
            )

            print(
                f"{f['filing_date']:<12} | {scores['total_words']:>6,} | "
                f"{pos_info:>14} | {neg_info:>14} | "
                f"{scores['pn_ratio']:>9.3f} | "
                f"{unc_info:>12} | {lit_info:>10}"
            )

        # Stage 5: Extract Financial Statements for most recent filing
        print(
            f"\nExtracting financial statements for most recent filing "
            f"({filings[0]['filing_date']})..."
        )
        statements = extract_statements(filings[0])
        for stmt_name, stmt_df in statements.items():
            print(f"\n--- {stmt_name.upper()} (Shape: {stmt_df.shape}) ---")
            print(f"Periods: {stmt_df.columns.tolist()}")
            print(stmt_df.head(4))

        # Stage 5b: Fetch Stock Prices for Balance Sheet Dates
        target_dates = [
            f["report_date"] for f in filings if f.get("report_date")
        ]
        print(
            f"\nFetching stock prices for {target_ticker} balance sheet "
            f"dates (10 Target Fiscal Years: FY2016 – FY2025)..."
        )
        stock_prices = fetch_stock_prices(target_ticker, target_dates)
        print(f"\nStock Prices for {target_ticker} Balance Sheet Dates:")
        for d_str, price in sorted(stock_prices.items(), reverse=True):
            print(f"  {d_str:<15} : ${price:,.2f}")

        # Stage 6: Fetch Fundamentals from XBRL and Compute Financial Ratios
        print(
            f"\nFetching fundamentals and computing ratios for "
            f"{target_ticker} from SEC XBRL..."
        )
        target_dates = [
            f["report_date"] for f in filings if f.get("report_date")
        ]
        filing_dates_map = {
            f["report_date"]: f["filing_date"]
            for f in filings
            if f.get("report_date") and f.get("filing_date")
        }
        fundamentals_df = fetch_fundamentals(
            target_ticker,
            target_dates=target_dates,
            filing_dates_map=filing_dates_map,
        )
        ratios_df = compute_ratios(fundamentals_df)

        print(
            f"\nFinancial Ratios for {target_ticker} "
            f"(10 Target Fiscal Years: FY2016 - FY2025):"
        )
        ratio_header = (
            f"{'Fiscal Date':<12} | {'Current Ratio':>14} | "
            f"{'Liab-to-Equity':>15} | {'Asset Turnover':>14} | "
            f"{'Net Margin':>12}"
        )
        print(ratio_header)
        print("-" * len(ratio_header))

        for date_idx, row in ratios_df.iterrows():
            cr_s = (
                f"{row['current_ratio']:.4f}"
                if pd.notna(row['current_ratio'])
                else "N/A"
            )
            lte_s = (
                f"{row['total_liabilities_to_equity']:.4f}"
                if pd.notna(row['total_liabilities_to_equity'])
                else "N/A"
            )
            at_s = (
                f"{row['asset_turnover']:.4f}"
                if pd.notna(row['asset_turnover'])
                else "N/A"
            )
            npm_s = (
                f"{row['net_profit_margin'] * 100:.2f}%"
                if pd.notna(row['net_profit_margin'])
                else "N/A"
            )
            print(
                f"{str(date_idx):<12} | {cr_s:>14} | {lte_s:>15} | "
                f"{at_s:>14} | {npm_s:>12}"
            )

        # Stage 7: Build Analysis Excel Workbook (8 Tabs + Dual-Axis Chart)
        print(
            f"\nGenerating eight-tab Excel analysis workbook for "
            f"{target_ticker}..."
        )
        sentiment_scores = [
            f["sentiment"] for f in filings if "sentiment" in f
        ]
        out_path = build_workbook(
            ticker=target_ticker,
            filings=filings,
            sentiment_scores=sentiment_scores,
            stock_prices=stock_prices if "stock_prices" in locals() else None,
            ratios_df=ratios_df,
            fundamentals_df=fundamentals_df,
        )
        print(f"\nSuccessfully generated analysis workbook:\n  {out_path}")

    except Exception as error:
        print(f"Error: {error}", file=sys.stderr)
