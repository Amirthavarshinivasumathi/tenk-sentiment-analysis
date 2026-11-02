"""Audit Item 7 MD&A Extractions.

This module inspects extracted MD&A plain-text files to verify that each
document cleanly begins at the Item 7 section header and terminates
prior to Item 8 (Financial Statements) or Item 9 (Accountant Changes).
"""

import glob
import os
import re
import sys
from typing import Any, Dict, List, Optional


def audit_mda_text(file_path: str) -> Dict[str, Any]:
    """Verify an extracted MD&A file starts and ends at section boundaries.

    Checks:
        1. Start boundary: The document must begin with the 'Item 7' heading
           and 'Management's Discussion' title.
        2. End boundary: The document must terminate cleanly before Item 8
           (Financial Statements) or Item 9 without leaking next section text.
        3. Completeness: The file must contain substantial narrative text
           (character count > 1,000).

    Args:
        file_path: Path to the extracted plain text file.

    Returns:
        Dict[str, Any]: Dictionary containing audit metrics:
            - 'file_path': Path of the audited file.
            - 'file_name': Basename of the file.
            - 'char_count': Total characters in the extracted text.
            - 'word_count': Total words in the extracted text.
            - 'starts_valid': True if text starts with Item 7 header.
            - 'ends_valid': True if text does not contain Item 8/9 headers.
            - 'passed': True if all audit checks pass.
            - 'start_snippet': First 80 characters of the text.
            - 'end_snippet': Last 80 characters of the text.

    Raises:
        FileNotFoundError: If the specified file does not exist.
        IOError: If reading the file fails.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as fp:
            text = fp.read().strip()
    except OSError as err:
        raise OSError(f"Could not read {file_path}: {err}") from err

    char_count = len(text)
    words = text.split()
    word_count = len(words)

    # 1. Start boundary check:
    # Must start with 'Item 7' followed by 'Management's Discussion'
    # within the first 120 characters
    start_snippet = text[:120].replace("\n", " ")
    starts_valid = bool(
        re.match(r"(?i)^item\s*7\b", text)
        and re.search(r"(?i)management", start_snippet)
    )

    # 2. End boundary check:
    # Must NOT contain Item 8 or Item 9 section titles
    has_item8 = bool(
        re.search(
            r"(?i)item\s*8\b[.\s:–-]{1,10}(?:<[^>]+>|\s){0,50}"
            r"financial\s+statements",
            text,
        )
    )
    has_item9 = bool(
        re.search(
            r"(?i)item\s*9\b[.\s:–-]{1,10}(?:<[^>]+>|\s){0,50}changes\s+in",
            text,
        )
    )
    ends_valid = (not has_item8) and (not has_item9) and (char_count > 1000)

    passed = starts_valid and ends_valid

    end_snippet = text[-120:].replace("\n", " ") if text else ""

    return {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "char_count": char_count,
        "word_count": word_count,
        "starts_valid": starts_valid,
        "ends_valid": ends_valid,
        "passed": passed,
        "start_snippet": text[:70].replace("\n", " "),
        "end_snippet": text[-70:].replace("\n", " "),
    }


def audit_all_mda(search_directory: str = "filings") -> List[Dict[str, Any]]:
    """Scan and audit all extracted MD&A text files in the given directory.

    If extracted '_mda.txt' files do not yet exist, attempts to generate
    them from existing '.htm' filings using extract_mda from main.py.

    Args:
        search_directory: Root directory to search for MD&A files.

    Returns:
        List[Dict[str, Any]]: Audit results for each discovered file.
    """
    pattern = os.path.join(search_directory, "**", "*_mda.txt")
    mda_files = sorted(glob.glob(pattern, recursive=True))

    # If no _mda.txt files found, attempt extraction from .htm files
    if not mda_files:
        htm_pattern = os.path.join(search_directory, "**", "*.htm")
        htm_files = sorted(glob.glob(htm_pattern, recursive=True))
        if htm_files:
            print(
                f"No _mda.txt files found. Extracting MD&A from "
                f"{len(htm_files)} HTML filing(s)..."
            )
            try:
                from main import extract_mda

                for htm_path in htm_files:
                    with open(
                        htm_path, "r", encoding="utf-8", errors="ignore"
                    ) as fp:
                        html_content = fp.read()
                    extracted = extract_mda(html_content)
                    out_path = htm_path.replace(".htm", "_mda.txt")
                    with open(out_path, "w", encoding="utf-8") as out_fp:
                        out_fp.write(extracted)
                    mda_files.append(out_path)
            except Exception as err:
                print(
                    f"Error during automatic extraction: {err}",
                    file=sys.stderr,
                )

    if not mda_files:
        print(
            f"No MD&A files found in '{search_directory}'. "
            "Please run main.py first to download and extract filings.",
            file=sys.stderr,
        )
        return []

    results: List[Dict[str, Any]] = []
    for file_path in mda_files:
        try:
            audit_result = audit_mda_text(file_path)
            results.append(audit_result)
        except Exception as err:
            print(
                f"Error auditing {file_path}: {err}",
                file=sys.stderr,
            )

    return results


def print_audit_report(results: List[Dict[str, Any]]) -> None:
    """Print a structured audit report table to the console.

    Args:
        results: List of audit outcome dictionaries.
    """
    if not results:
        return

    print("\n" + "=" * 78)
    print(f"{'MD&A EXTRACTION AUDIT REPORT':^78}")
    print("=" * 78)

    header = (
        f"{'#':<3} | {'File Name':<34} | {'Words':>7} | "
        f"{'Start':^7} | {'End':^7} | {'Result':^7}"
    )
    print(header)
    print("-" * 78)

    passed_count = 0
    for idx, r in enumerate(results, start=1):
        start_str = "PASS" if r["starts_valid"] else "FAIL"
        end_str = "PASS" if r["ends_valid"] else "FAIL"
        res_str = "PASS" if r["passed"] else "FAIL"
        if r["passed"]:
            passed_count += 1

        print(
            f"{idx:<3} | {r['file_name']:<34} | {r['word_count']:>7,} | "
            f"{start_str:^7} | {end_str:^7} | {res_str:^7}"
        )

    print("-" * 78)
    print(
        f"Summary: {passed_count}/{len(results)} MD&A file(s) passed "
        f"boundary verification."
    )
    print("=" * 78 + "\n")

    # Display sample snippets for transparency
    print("Detailed Boundary Snippets:")
    for idx, r in enumerate(results, start=1):
        print(f"\n[{idx}] {r['file_name']}:")
        print(f"  [Start]: \"{r['start_snippet']}...\"")
        print(f"  [End]:   \"...{r['end_snippet']}\"")


if __name__ == "__main__":
    audit_results = audit_all_mda()
    print_audit_report(audit_results)
