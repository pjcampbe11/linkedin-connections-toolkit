# linkedin_connections.py
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import textwrap
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd


EXPECTED_HEADER_TOKENS = {
    "first name",
    "last name",
    "email address",
    "company",
    "position",
    "connected on",
}


def _eprint(msg: str) -> None:
    print(msg, file=sys.stderr)


def parse_csv_row(line: str) -> list[str]:
    """Parse a single CSV line into cells (best-effort)."""
    try:
        return next(csv.reader([line]))
    except Exception:
        return []


def detect_header_row(
    csv_path: Path,
    *,
    encoding: str,
    max_scan_lines: int,
    min_matches: int = 3,
) -> Optional[int]:
    """
    Try to find the first line that looks like the real header row.

    We consider a line a header row if it contains >= min_matches of
    EXPECTED_HEADER_TOKENS when split as CSV and normalized.
    """
    try:
        with csv_path.open("r", encoding=encoding, errors="replace", newline="") as f:
            for idx, line in enumerate(f):
                if idx >= max_scan_lines:
                    break
                cells = parse_csv_row(line)
                normalized = [
                    (c or "").strip().lower().lstrip("\ufeff") for c in cells
                ]
                hits = sum(1 for t in EXPECTED_HEADER_TOKENS if t in normalized)
                if hits >= min_matches:
                    return idx
    except FileNotFoundError:
        raise
    except Exception as ex:
        _eprint(f"[warn] Header detection failed with error: {ex}")
    return None


def find_connections_csv_in_dir(archive_dir: Path) -> Path:
    """Locate Connections.csv inside an extracted LinkedIn archive directory."""
    candidates = list(archive_dir.rglob("Connections.csv"))
    if not candidates:
        # Some archives may vary in case; be permissive:
        candidates = [p for p in archive_dir.rglob("*.csv") if p.name.lower() == "connections.csv"]

    if not candidates:
        raise FileNotFoundError(
            f"Could not find Connections.csv under: {archive_dir}"
        )

    # Prefer the shortest path (usually the top-level file)
    candidates.sort(key=lambda p: (len(p.parts), str(p)))
    return candidates[0]


def extract_connections_csv_from_zip(archive_zip: Path, extract_to: Path) -> Path:
    """Extract Connections.csv from a LinkedIn ZIP archive into extract_to/ and return the extracted path."""
    extract_to.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(archive_zip, "r") as zf:
        members = zf.namelist()
        matches = [m for m in members if m.lower().endswith("connections.csv")]

        if not matches:
            raise FileNotFoundError(
                f"No Connections.csv found in zip: {archive_zip}"
            )

        # Prefer the shortest member name (closest to root)
        matches.sort(key=lambda m: (len(Path(m).parts), m))
        member = matches[0]

        out_path = extract_to / Path(member).name
        with zf.open(member) as src, out_path.open("wb") as dst:
            dst.write(src.read())
        return out_path


def normalize_whitespace(series: pd.Series) -> pd.Series:
    """Collapse internal whitespace and trim ends."""
    return (
        series.astype("string")
        .fillna("")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def build_keyword_regex(keywords: Iterable[str]) -> Optional[str]:
    cleaned = []
    for k in keywords:
        k = (k or "").strip()
        if k:
            cleaned.append(re.escape(k))
    if not cleaned:
        return None
    return r"(" + "|".join(cleaned) + r")"


def safe_write_csv(df: pd.DataFrame, path: Path, overwrite: bool) -> None:
    if path.exists() and not overwrite:
        raise FileExistsError(f"Refusing to overwrite existing file: {path} (use --overwrite)")
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)


def safe_write_excel(df: pd.DataFrame, path: Path, overwrite: bool) -> None:
    if path.exists() and not overwrite:
        raise FileExistsError(f"Refusing to overwrite existing file: {path} (use --overwrite)")
    path.parent.mkdir(parents=True, exist_ok=True)
    # This will use pandas' default .xlsx writer if available (engine auto-selected by pandas)
    df.to_excel(path, index=False)


def connections_summary(df: pd.DataFrame, input_path: Path) -> dict:
    def col_exists(name: str) -> bool:
        return name in df.columns

    summary = {
        "input_file": str(input_path),
        "rows": int(len(df)),
        "columns": list(df.columns),
        "generated_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }

    if col_exists("Email Address"):
        summary["missing_email_count"] = int(df["Email Address"].isna().sum())
        summary["nonempty_email_count"] = int((df["Email Address"].fillna("").astype(str).str.strip() != "").sum())

    if col_exists("Company"):
        summary["unique_companies"] = int(df["Company"].fillna("").nunique())

    if col_exists("Position"):
        summary["unique_positions"] = int(df["Position"].fillna("").nunique())

    return summary


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="linkedin_connections",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Clean and segment LinkedIn Connections.csv for job-search/outreach workflows.",
        epilog=textwrap.dedent(
            """
            Examples:
              python linkedin_connections.py --input "C:\\path\\to\\Connections.csv" --outdir out --export recruiters --export crm
              python linkedin_connections.py --archive-zip "C:\\path\\LinkedInDataExport.zip" --outdir out --export clean --export targets --target-companies "Google,Microsoft,OpenAI"
            """
        ),
    )

    src = parser.add_mutually_exclusive_group(required=False)
    src.add_argument("--input", type=str, help="Path to Connections.csv")
    src.add_argument("--archive-dir", type=str, help="Path to extracted LinkedIn archive directory (script will search for Connections.csv)")
    src.add_argument("--archive-zip", type=str, help="Path to LinkedIn ZIP archive (script will extract Connections.csv)")

    parser.add_argument("--outdir", type=str, default="out", help="Output directory (default: ./out)")
    parser.add_argument("--basename", type=str, default=None, help="Base filename prefix for outputs (default: derived from input name)")

    parser.add_argument("--skip-rows", type=int, default=None,
                        help="Number of rows to skip at top of CSV. If omitted, script auto-detects header row; if detection fails, falls back to 3.")
    parser.add_argument("--max-scan-lines", type=int, default=25, help="Max lines to scan when auto-detecting the header row (default: 25)")

    parser.add_argument("--encoding", type=str, default="utf-8-sig", help="CSV encoding (default: utf-8-sig)")
    parser.add_argument("--on-bad-lines", type=str, default="warn", choices=["error", "warn", "skip"],
                        help="How pandas should handle malformed CSV rows (default: warn)")

    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing output files (default: off)")
    parser.add_argument("--dry-run", action="store_true", help="Do not write outputs; print summary only")

    parser.add_argument("--dedupe", action="store_true", help="Drop duplicates (heuristic subset)")
    parser.add_argument("--redact-emails", action="store_true", help="Blank out Email Address column in outputs (useful for sharing files)")

    parser.add_argument("--add-full-name", action="store_true", help="Add 'Full Name' column if First/Last Name exist")
    parser.add_argument("--add-email-domain", action="store_true", help="Add 'Email Domain' column if Email Address exists")
    parser.add_argument("--add-connected-iso", action="store_true", help="Add 'Connected On ISO' column if Connected On exists")

    parser.add_argument("--message-template", type=str, default=None,
                        help="Add 'Message' column using a template with {first_name}, {company}, {position}. Example: \"Hi {first_name}, ...\"")

    parser.add_argument("--export", action="append", default=None,
                        choices=["clean", "recruiters", "targets", "crm", "excel"],
                        help="Which outputs to write. Repeatable. Default: clean.")

    parser.add_argument("--recruiter-keywords", type=str, default="recruiter,talent,hr,hiring,people",
                        help="Comma-separated keywords for recruiter/HR filtering (default: recruiter,talent,hr,hiring,people)")
    parser.add_argument("--target-companies", type=str, default="",
                        help="Comma-separated target companies (used with --export targets)")

    parser.add_argument("--summary-json", type=str, default=None,
                        help="If set, write a summary JSON to this path (relative to outdir if not absolute)")

    args = parser.parse_args(argv)

    outdir = Path(args.outdir).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    # Determine input CSV
    extracted_path: Optional[Path] = None

    if args.input:
        input_path = Path(args.input).expanduser().resolve()
    elif args.archive_dir:
        archive_dir = Path(args.archive_dir).expanduser().resolve()
        input_path = find_connections_csv_in_dir(archive_dir)
    elif args.archive_zip:
        archive_zip = Path(args.archive_zip).expanduser().resolve()
        extract_dir = outdir / "_extracted"
        extracted_path = extract_connections_csv_from_zip(archive_zip, extract_dir)
        input_path = extracted_path
    else:
        # Default: look in current working directory
        input_path = Path.cwd() / "Connections.csv"
        input_path = input_path.resolve()

    if not input_path.exists():
        _eprint(f"[error] Input file not found: {input_path}")
        return 2

    basename = args.basename or input_path.stem

    # Determine skiprows
    if args.skip_rows is not None:
        skiprows = args.skip_rows
    else:
        detected = detect_header_row(
            input_path,
            encoding=args.encoding,
            max_scan_lines=args.max_scan_lines,
            min_matches=3,
        )
        skiprows = detected if detected is not None else 3

    # Load
    df = pd.read_csv(
        input_path,
        skiprows=skiprows,
        encoding=args.encoding,
        on_bad_lines=args.on_bad_lines,
    )

    # Basic cleanup
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")

    # Normalize common text fields if present
    for col in ["First Name", "Last Name", "Company", "Position", "Email Address"]:
        if col in df.columns:
            df[col] = normalize_whitespace(df[col])

    # Enhancements
    if args.add_full_name and "First Name" in df.columns and "Last Name" in df.columns:
        df["Full Name"] = (df["First Name"].fillna("") + " " + df["Last Name"].fillna("")).str.strip()

    if args.add_email_domain and "Email Address" in df.columns:
        emails = df["Email Address"].fillna("").astype(str).str.strip()
        df["Email Domain"] = emails.where(emails.str.contains("@"), "").str.split("@").str[-1].str.lower()

    if args.add_connected_iso and "Connected On" in df.columns:
        dt = pd.to_datetime(
    df["Connected On"],
    format="%m/%d/%Y",
    errors="coerce"
)
        df["Connected On ISO"] = dt.dt.date.astype("string")

    if args.message_template:
        # Template variables: first_name, company, position
        def render(row: pd.Series) -> str:
            mapping = {
                "first_name": (row.get("First Name") or "").strip(),
                "company": (row.get("Company") or "").strip(),
                "position": (row.get("Position") or "").strip(),
            }
            try:
                return args.message_template.format(**mapping)
            except Exception:
                # If template is malformed, avoid crashing the whole run:
                return ""

        df["Message"] = df.apply(render, axis=1)

    if args.dedupe:
        # Prefer email if present and nonempty; otherwise fall back to name+company+position.
        if "Email Address" in df.columns and (df["Email Address"].fillna("").astype(str).str.strip() != "").any():
            subset = ["Email Address"]
        else:
            subset = [c for c in ["First Name", "Last Name", "Company", "Position"] if c in df.columns]
            if not subset:
                subset = None  # fall back to full-row duplicates
        df = df.drop_duplicates(subset=subset, keep="first")

    if args.redact_emails and "Email Address" in df.columns:
        df["Email Address"] = ""

    exports = args.export or ["clean"]

    # Always print a quick summary
    summary = connections_summary(df, input_path)
    print(json.dumps(summary, indent=2))

    # Optionally write summary json
    if args.summary_json:
        summary_path = Path(args.summary_json)
        if not summary_path.is_absolute():
            summary_path = outdir / summary_path
        if not args.dry_run:
            summary_path.parent.mkdir(parents=True, exist_ok=True)
            summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    if args.dry_run:
        return 0

    # Export: clean
    if "clean" in exports:
        safe_write_csv(df, outdir / f"{basename}_cleaned.csv", overwrite=args.overwrite)

    # Export: recruiters
    if "recruiters" in exports:
        if "Position" not in df.columns:
            _eprint("[warn] Cannot export recruiters: missing 'Position' column.")
        else:
            keywords = [k.strip() for k in args.recruiter_keywords.split(",") if k.strip()]
            pattern = build_keyword_regex(keywords)
            if not pattern:
                _eprint("[warn] Recruiter keyword list is empty; skipping recruiters export.")
            else:
                mask = df["Position"].fillna("").astype(str).str.contains(pattern, case=False, na=False, regex=True)
                recruiters = df.loc[mask].copy()
                safe_write_csv(recruiters, outdir / f"{basename}_recruiters.csv", overwrite=args.overwrite)

    # Export: targets
    if "targets" in exports:
        if "Company" not in df.columns:
            _eprint("[warn] Cannot export targets: missing 'Company' column.")
        else:
            companies = [c.strip() for c in args.target_companies.split(",") if c.strip()]
            pattern = build_keyword_regex(companies)
            if not pattern:
                _eprint("[warn] No --target-companies provided; skipping targets export.")
            else:
                mask = df["Company"].fillna("").astype(str).str.contains(pattern, case=False, na=False, regex=True)
                targets = df.loc[mask].copy()
                safe_write_csv(targets, outdir / f"{basename}_target_companies.csv", overwrite=args.overwrite)

    # Export: CRM tracker
    if "crm" in exports:
        crm = df.copy()
        for col in ["Status", "Last Contacted", "Next Follow Up", "Notes"]:
            if col not in crm.columns:
                crm[col] = ""
        safe_write_csv(crm, outdir / f"{basename}_crm.csv", overwrite=args.overwrite)

    # Export: Excel (single workbook; simple mode)
    if "excel" in exports:
        safe_write_excel(df, outdir / f"{basename}_cleaned.xlsx", overwrite=args.overwrite)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
