#!/usr/bin/env python3
"""
url_to_pdf.py — Convert a list of URLs to PDFs using headless Chromium (Playwright).

Usage:
  python3 url_to_pdf.py --urls "https://example.com" "https://mnsu.edu" --output ./pdfs
  python3 url_to_pdf.py --file urls.txt --output ./pdfs
"""

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


def slugify(url: str) -> str:
    """Turn a URL into a safe filename (no extension)."""
    parsed = urlparse(url)
    name = (parsed.netloc + parsed.path).strip("/").replace("/", "_")
    name = re.sub(r"[^\w\-]", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:100] or "page"


def url_to_pdf(page, url: str, output_path: str, timeout: int = 30000) -> dict:
    """Navigate to URL and save as PDF. Returns a result dict."""
    try:
        # Switch to print media before loading — triggers the site's print CSS
        # (same as what the browser does when you right-click → Print → Save as PDF)
        page.emulate_media(media="print")
        page.goto(url, wait_until="networkidle", timeout=timeout)
        page.pdf(
            path=output_path,
            format="A4",
            print_background=False,
            margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"},
        )
        size = Path(output_path).stat().st_size
        return {"url": url, "status": "ok", "file": output_path, "size_kb": round(size / 1024, 1)}
    except PlaywrightTimeout:
        return {"url": url, "status": "error", "reason": "Timeout — page took too long to load"}
    except Exception as e:
        return {"url": url, "status": "error", "reason": str(e)}


def main():
    parser = argparse.ArgumentParser(description="Save URLs as PDFs using headless Chromium.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--urls", nargs="+", metavar="URL", help="One or more URLs to convert")
    group.add_argument("--file", metavar="FILE", help="Text file with one URL per line")
    parser.add_argument("--output", default="./pdfs", metavar="DIR", help="Output directory (default: ./pdfs)")
    parser.add_argument("--timeout", type=int, default=30, help="Page load timeout in seconds (default: 30)")
    parser.add_argument("--json", action="store_true", help="Output results as JSON")
    args = parser.parse_args()

    # Collect URLs
    if args.file:
        with open(args.file) as f:
            urls = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    else:
        urls = args.urls

    if not urls:
        print("No URLs provided.", file=sys.stderr)
        sys.exit(1)

    # Prepare output directory
    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    timeout_ms = args.timeout * 1000
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
        )
        page = context.new_page()

        for i, url in enumerate(urls, 1):
            slug = slugify(url)
            # Add index prefix so ordering is preserved if names clash
            filename = f"{i:02d}_{slug}.pdf"
            output_path = str(out_dir / filename)

            if not args.json:
                print(f"[{i}/{len(urls)}] {url}", end=" ... ", flush=True)

            result = url_to_pdf(page, url, output_path, timeout=timeout_ms)
            results.append(result)

            if not args.json:
                if result["status"] == "ok":
                    print(f"✓  saved → {filename}  ({result['size_kb']} KB)")
                else:
                    print(f"✗  FAILED — {result['reason']}")

            # Small delay between requests to be polite
            if i < len(urls):
                time.sleep(0.5)

        browser.close()

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        ok = sum(1 for r in results if r["status"] == "ok")
        fail = len(results) - ok
        print(f"\nDone: {ok} saved, {fail} failed → {out_dir.resolve()}")

    # Exit with error code if any failed
    if any(r["status"] != "ok" for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
