#!/usr/bin/env python3
"""
Teachable School Classifier
Classifies Non-Seller schools as Bypasser or Program Distributor
based on presence of external checkout mechanisms on their Primary Domain.

Usage:
    python classify_schools.py input_file.xlsx
    python classify_schools.py input_file.csv --output results.xlsx
    python classify_schools.py input_file.xlsx --limit 50  # for testing
"""

import sys
import os
import time
import json
import logging
import argparse
import re
import random
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ── Logging ──────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    handlers=[
        logging.FileHandler("classifier.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

# ── Constants ─────────────────────────────────────────────────────────────────

SESSION_TIMEOUT = 15        # seconds per request
MAX_RETRIES = 3             # retries on transient errors
RATE_LIMIT_MIN = 1.0        # minimum sleep between requests (seconds)
RATE_LIMIT_MAX = 2.5        # maximum sleep between requests (seconds)
CHECKPOINT_INTERVAL = 100   # save progress every N new schools

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,pt-BR;q=0.8,pt;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

# ── Checkout detection rules ──────────────────────────────────────────────────
# Each key is the platform label; value is list of regex patterns to search
# in the full HTML source (case-insensitive).

CHECKOUT_PLATFORMS = {
    "Shopify": [
        r"cdn\.shopify\.com",
        r"myshopify\.com",
        r"shopify\.com/s/files",
        r'"Shopify"',
        r"Shopify\.theme",
    ],
    "Stripe": [
        r"js\.stripe\.com",
        r"checkout\.stripe\.com",
        r'href=["\']https://buy\.stripe\.com/[^"\']+["\']',
        r"stripe\.com/pay",
        r"StripeCheckout",
        r'data-stripe["\s]',
    ],
    "WooCommerce": [
        r"woocommerce",
        r"wc-cart",
        r"wc_add_to_cart_params",
        r"woocommerce-cart",
        r"/wc-api/",
        r"WooCommerce",
    ],
    "Hotmart": [
        r"hotmart\.com/product",
        r"pay\.hotmart\.com",
        r"checkout\.hotmart\.com",
        r"hotmart\.product",
        r"go\.hotmart\.com",
        r"hotmart\.com/t/",
    ],
    "Gumroad": [
        r"gumroad\.com",
        r"gum\.co/",
        r"assets\.gumroad\.com",
        r"gumroad-button",
    ],
    "Kajabi": [
        r"kajabi\.com",
        r"app\.kajabi\.com",
        r"kajabi-assets",
        r'"kajabi"',
    ],
    "PayPal": [
        r"paypal\.com/sdk/js",
        r"paypalobjects\.com",
        r"paypal\.com/cgi-bin/webscr",
        r'href=["\']https://www\.paypal\.me/[^"\']+["\']',
        r"paypal\.com/donate",
        r"paypal\.com/ncp/payment",
        r"PayPalButton",
        r"paypal\.Buttons",
    ],
    "Kiwify": [
        r"kiwify\.com\.br",
        r"pay\.kiwify",
        r"go\.kiwify",
    ],
    "Eduzz": [
        r"eduzz\.com",
        r"checkout\.eduzz",
        r"sun\.eduzz",
    ],
    "Monetizze": [
        r"monetizze\.com\.br",
    ],
    "PerfectPay": [
        r"perfectpay\.com\.br",
    ],
    "Lastlink": [
        r"lastlink\.app",
        r"lastlink\.com",
    ],
    "Braip": [
        r"braip\.com",
    ],
    "Ticto": [
        r"ticto\.com\.br",
    ],
    "Thinkific": [
        r"thinkific\.com",
    ],
    "Podia": [
        r"podia\.com",
    ],
    "SamCart": [
        r"samcart\.com",
        r"samcartlive\.com",
    ],
    "ThriveCart": [
        r"thrivecart\.com",
    ],
    "ClickFunnels": [
        r"clickfunnels\.com",
        r"cf-js\.com",
    ],
    "Systeme.io": [
        r"systeme\.io",
        r"affiliatly\.com",
    ],
    "Memberful": [
        r"memberful\.com",
    ],
    "Pagar.me": [
        r"pagar\.me",
        r"pagarme\.com",
    ],
    "MercadoPago": [
        r"mercadopago\.com",
        r"mercadolivre\.com\.br/checkout",
        r"sdk\.mercadopago\.com",
    ],
    "Loja Integrada": [
        r"lojaintegrada\.com\.br",
    ],
    "VTEX": [
        r"vtex\.com",
        r"vtexcommercestable\.com\.br",
    ],
    "Magento": [
        r"Magento_Ui",
        r'"Mage"',
        r"mage/cookies",
    ],
    "LearnWorlds": [
        r"learnworlds\.com",
    ],
    "Udemy": [
        r"udemy\.com",
    ],
    "Coursera": [
        r"coursera\.org",
    ],
    "Hotmart (Club)": [
        r"hotmart\.com/club",
    ],
}

# Extra link-level patterns: if a link href contains these, it's a checkout link
LINK_CHECKOUT_KEYWORDS = [
    "checkout", "cart", "buy", "purchase", "payment", "pay/", "/pay?",
    "enroll", "subscribe", "order", "acquire",
]

# Domains that are part of Teachable itself — not external checkouts
TEACHABLE_DOMAINS = {
    "teachable.com", "teachablecdn.com", "teachableassets.com",
}


# ── URL helpers ───────────────────────────────────────────────────────────────

def normalize_url(domain: str) -> str | None:
    """Return a proper https:// URL for a domain string, or None if empty."""
    if not domain or (isinstance(domain, float)):
        return None
    domain = str(domain).strip().rstrip("/")
    if not domain or domain.lower() in ("nan", "none", "n/a", ""):
        return None
    if not domain.startswith(("http://", "https://")):
        domain = "https://" + domain
    return domain


def same_domain(url1: str, url2: str) -> bool:
    """Return True if both URLs share the same registered domain."""
    try:
        h1 = urlparse(url1).netloc.lower().lstrip("www.")
        h2 = urlparse(url2).netloc.lower().lstrip("www.")
        return h1 == h2
    except Exception:
        return False


def is_teachable_domain(url: str) -> bool:
    host = urlparse(url).netloc.lower().lstrip("www.")
    return any(host.endswith(td) for td in TEACHABLE_DOMAINS)


# ── HTTP fetch with retry ─────────────────────────────────────────────────────

def fetch_page(
    session: requests.Session, url: str, retry: int = 0
) -> tuple[str | None, str, str | None]:
    """
    Fetch *url* and return (html, final_url, error_note).
    Retries up to MAX_RETRIES on transient errors.
    Falls back to http:// on SSL errors.
    """
    try:
        resp = session.get(url, timeout=SESSION_TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        # Try to detect encoding
        resp.encoding = resp.apparent_encoding or "utf-8"
        return resp.text, resp.url, None

    except requests.exceptions.SSLError:
        if url.startswith("https://") and retry == 0:
            return fetch_page(session, url.replace("https://", "http://", 1), retry=1)
        return None, url, "SSL Error"

    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
        if retry < MAX_RETRIES:
            wait = 2 ** retry
            logger.debug(f"  Retrying {url} in {wait}s ({exc.__class__.__name__})")
            time.sleep(wait)
            return fetch_page(session, url, retry=retry + 1)
        return None, url, f"{exc.__class__.__name__}"

    except requests.exceptions.TooManyRedirects:
        return None, url, "Too Many Redirects"

    except requests.exceptions.HTTPError as exc:
        code = exc.response.status_code if exc.response is not None else "?"
        return None, url, f"HTTP {code}"

    except Exception as exc:
        return None, url, f"Error: {str(exc)[:80]}"


# ── Checkout detection ────────────────────────────────────────────────────────

def _extract_link_for_pattern(soup: BeautifulSoup, pattern: str, base_url: str) -> str:
    """
    Try to return a concrete URL that matches *pattern* from
    script src, link href, iframe src, form action, or anchor href.
    Falls back to a note about the pattern.
    """
    selectors = [
        ("script", "src"),
        ("link", "href"),
        ("iframe", "src"),
        ("form", "action"),
        ("a", "href"),
    ]
    for tag, attr in selectors:
        for el in soup.find_all(tag, **{attr: True}):
            val = el.get(attr, "")
            if re.search(pattern, val, re.IGNORECASE):
                return urljoin(base_url, val)[:300]
    return f"[pattern: {pattern}]"


def detect_checkout(
    html: str, base_url: str
) -> tuple[str | None, str | None]:
    """
    Scan *html* for known external checkout signals.
    Returns (platform_name, checkout_url) or (None, None).
    """
    # 1. Fast regex scan of raw HTML for each platform
    for platform, patterns in CHECKOUT_PLATFORMS.items():
        for pat in patterns:
            if re.search(pat, html, re.IGNORECASE):
                try:
                    soup = BeautifulSoup(html, "lxml")
                except Exception:
                    soup = BeautifulSoup(html, "html.parser")
                link = _extract_link_for_pattern(soup, pat, base_url)
                return platform, link

    # 2. DOM-level scan: external links with checkout keywords
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        soup = BeautifulSoup(html, "html.parser")

    current_host = urlparse(base_url).netloc.lower().lstrip("www.")

    for a in soup.find_all("a", href=True):
        href = str(a["href"])
        if not href.startswith(("http://", "https://")):
            continue
        link_host = urlparse(href).netloc.lower().lstrip("www.")
        if link_host == current_host or is_teachable_domain(href):
            continue
        href_lower = href.lower()
        if any(kw in href_lower for kw in LINK_CHECKOUT_KEYWORDS):
            return "External Checkout Link", href[:300]

    # 3. Forms with external / payment-related actions
    for form in soup.find_all("form", action=True):
        action = str(form["action"])
        action_abs = urljoin(base_url, action)
        action_host = urlparse(action_abs).netloc.lower().lstrip("www.")
        if action_host and action_host != current_host and not is_teachable_domain(action_abs):
            return "External Form", action_abs[:300]

    return None, None


# ── Single school classifier ──────────────────────────────────────────────────

def classify_school(
    session: requests.Session,
    school_id: str,
    school_name: str,
    domain: str,
) -> dict:
    """Return a result dict for one school."""
    base = {
        "School ID": school_id,
        "School Name": school_name,
        "Primary Domain": domain,
        "Classification": "Program Distributor",
        "Checkout Link": "",
        "Notes": "",
    }

    url = normalize_url(domain)
    if not url:
        base["Notes"] = "No domain provided"
        return base

    html, final_url, error = fetch_page(session, url)

    if error or not html:
        base["Notes"] = f"Domain unreachable: {error}"
        return base

    platform, checkout_link = detect_checkout(html, final_url)

    if platform:
        base["Classification"] = "Bypasser"
        base["Checkout Link"] = checkout_link or ""
        base["Notes"] = f"External checkout detected: {platform}"
    else:
        base["Notes"] = "No external checkout detected"

    return base


# ── Checkpoint helpers ────────────────────────────────────────────────────────

def load_checkpoint(path: str) -> dict:
    if Path(path).exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            logger.info(f"Resumed from checkpoint: {len(data)} schools already processed.")
            return data
        except Exception as exc:
            logger.warning(f"Could not load checkpoint ({exc}); starting fresh.")
    return {}


def save_checkpoint(path: str, data: dict) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ── Excel export with formatting ──────────────────────────────────────────────

def export_excel(df: pd.DataFrame, output_path: str) -> None:
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    bypasser_fill = PatternFill("solid", fgColor="FFC7CE")      # red
    distributor_fill = PatternFill("solid", fgColor="C6EFCE")   # green
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    col_widths = {"A": 14, "B": 32, "C": 38, "D": 22, "E": 55, "F": 55}

    # Summary data
    bypassers = df[df["Classification"] == "Bypasser"]
    distributors = df[df["Classification"] == "Program Distributor"]

    platform_counts = (
        bypassers["Notes"]
        .str.extract(r"External checkout detected: (.+)")
        [0]
        .value_counts()
    )

    summary_rows = [
        ["Metric", "Value"],
        ["Total Schools", len(df)],
        ["Bypassers", len(bypassers)],
        ["Program Distributors", len(distributors)],
        ["Bypasser Rate", f"{len(bypassers)/max(len(df),1)*100:.1f}%"],
        [],
        ["Platform Breakdown", "Count"],
    ]
    for plat, cnt in platform_counts.items():
        summary_rows.append([plat, cnt])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Classifications")

        # Blank df for summary (we'll write manually via openpyxl)
        pd.DataFrame().to_excel(writer, index=False, sheet_name="Summary")

        wb = writer.book

        # --- Format Classifications sheet ---
        ws = writer.sheets["Classifications"]

        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width
        ws.row_dimensions[1].height = 28

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = border

        for row_idx in range(2, ws.max_row + 1):
            classification = ws.cell(row=row_idx, column=4).value
            fill = bypasser_fill if classification == "Bypasser" else distributor_fill
            for col_idx in range(1, 7):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = fill
                cell.alignment = Alignment(vertical="center", wrap_text=(col_idx >= 5))
                cell.border = border
            ws.row_dimensions[row_idx].height = 18

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # --- Format Summary sheet ---
        ws2 = writer.sheets["Summary"]
        ws2.column_dimensions["A"].width = 30
        ws2.column_dimensions["B"].width = 18

        for r_idx, row_data in enumerate(summary_rows, start=1):
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws2.cell(row=r_idx, column=c_idx, value=val)
                if r_idx == 1 or (r_idx == 7 and val):
                    cell.fill = header_fill
                    cell.font = header_font
                cell.border = border

    logger.info(f"Saved output: {output_path}")


# ── Column auto-detection ─────────────────────────────────────────────────────

def find_column(candidates: list[str], columns: list[str]) -> str | None:
    """Return the first column name that contains any candidate substring (case-insensitive)."""
    for candidate in candidates:
        for col in columns:
            if candidate.lower() in col.lower():
                return col
    return None


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Classify Teachable non-seller schools as Bypasser or Program Distributor"
    )
    parser.add_argument("input_file", help="Input CSV or Excel file")
    parser.add_argument(
        "--output", "-o",
        help="Output Excel file path (default: classified_<timestamp>.xlsx)",
    )
    parser.add_argument(
        "--checkpoint", "-c",
        default="checkpoint.json",
        help="Checkpoint file for resuming (default: checkpoint.json)",
    )
    parser.add_argument(
        "--school-id-col",
        help="Exact column name for School ID (auto-detected if omitted)",
    )
    parser.add_argument(
        "--school-name-col",
        help="Exact column name for School Name (auto-detected if omitted)",
    )
    parser.add_argument(
        "--domain-col",
        help="Exact column name for Primary Domain (auto-detected if omitted)",
    )
    parser.add_argument(
        "--limit", type=int, default=None,
        help="Process only the first N schools (useful for testing)",
    )
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = args.output or f"classified_schools_{timestamp}.xlsx"

    # ── Read input ────────────────────────────────────────────────────────────
    logger.info(f"Reading: {args.input_file}")
    try:
        if args.input_file.lower().endswith(".csv"):
            df_input = pd.read_csv(args.input_file, dtype=str)
        else:
            df_input = pd.read_excel(args.input_file, dtype=str)
    except Exception as exc:
        logger.error(f"Cannot read input file: {exc}")
        sys.exit(1)

    logger.info(f"Loaded {len(df_input):,} rows | columns: {list(df_input.columns)}")

    # ── Detect columns ────────────────────────────────────────────────────────
    cols = list(df_input.columns)
    id_col = args.school_id_col or find_column(
        ["school id", "school_id", "schoolid", "id"], cols
    )
    name_col = args.school_name_col or find_column(
        ["school name", "school_name", "schoolname", "name"], cols
    )
    domain_col = args.domain_col or find_column(
        ["primary domain", "primary_domain", "domain", "url", "website", "site"], cols
    )

    missing = [
        label for label, val in [
            ("School ID", id_col), ("School Name", name_col), ("Primary Domain", domain_col)
        ] if not val
    ]
    if missing:
        logger.error(
            f"Could not auto-detect columns: {missing}\n"
            f"Available columns: {cols}\n"
            f"Use --school-id-col / --school-name-col / --domain-col to specify them."
        )
        sys.exit(1)

    logger.info(f"Columns → ID: '{id_col}' | Name: '{name_col}' | Domain: '{domain_col}'")

    if args.limit:
        df_input = df_input.head(args.limit)
        logger.info(f"Limited to {args.limit} schools (--limit flag)")

    # ── Load checkpoint ───────────────────────────────────────────────────────
    checkpoint = load_checkpoint(args.checkpoint)

    # ── Process ───────────────────────────────────────────────────────────────
    session = requests.Session()
    session.headers.update(HEADERS)

    results: list[dict] = []
    total = len(df_input)
    new_since_checkpoint = 0

    for idx, row in df_input.iterrows():
        school_id = str(row.get(id_col, "")).strip()
        school_name = str(row.get(name_col, "")).strip()
        domain = str(row.get(domain_col, "")).strip()

        # Already in checkpoint?
        if school_id in checkpoint:
            results.append(checkpoint[school_id])
            continue

        processed_count = len(results) + 1
        logger.info(f"[{processed_count}/{total}] {school_name!r}  →  {domain}")

        result = classify_school(session, school_id, school_name, domain)
        results.append(result)
        checkpoint[school_id] = result
        new_since_checkpoint += 1

        cls = result["Classification"]
        note = result["Notes"]
        logger.info(f"        ✓ {cls} | {note}")

        # Periodic checkpoint save
        if new_since_checkpoint % CHECKPOINT_INTERVAL == 0:
            save_checkpoint(args.checkpoint, checkpoint)
            byp = sum(1 for r in results if r["Classification"] == "Bypasser")
            logger.info(
                f"  ── Checkpoint saved ({len(results)}/{total} processed, "
                f"{byp} Bypassers so far) ──"
            )

        # Rate limiting
        time.sleep(random.uniform(RATE_LIMIT_MIN, RATE_LIMIT_MAX))

    # Final checkpoint save
    save_checkpoint(args.checkpoint, checkpoint)

    # ── Build output DataFrame ────────────────────────────────────────────────
    output_df = pd.DataFrame(results, columns=[
        "School ID", "School Name", "Primary Domain",
        "Classification", "Checkout Link", "Notes",
    ])

    # ── Summary ───────────────────────────────────────────────────────────────
    bypassers_df = output_df[output_df["Classification"] == "Bypasser"]
    distributors_df = output_df[output_df["Classification"] == "Program Distributor"]
    n = len(output_df)

    logger.info("")
    logger.info("=" * 60)
    logger.info("FINAL SUMMARY")
    logger.info("=" * 60)
    logger.info(f"Total schools:       {n:,}")
    logger.info(f"Bypassers:           {len(bypassers_df):,}  ({len(bypassers_df)/max(n,1)*100:.1f}%)")
    logger.info(f"Program Distributors:{len(distributors_df):,}  ({len(distributors_df)/max(n,1)*100:.1f}%)")

    platform_counts = (
        bypassers_df["Notes"]
        .str.extract(r"External checkout detected: (.+)")[0]
        .value_counts()
    )
    if not platform_counts.empty:
        logger.info("")
        logger.info("Bypassers by platform:")
        for plat, cnt in platform_counts.items():
            logger.info(f"  {plat:<28} {cnt:>5}")
    logger.info("=" * 60)

    # ── Export Excel ──────────────────────────────────────────────────────────
    export_excel(output_df, output_file)
    logger.info(f"Done! Output file: {output_file}")


if __name__ == "__main__":
    main()
