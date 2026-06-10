#!/usr/bin/env python3
"""
scrape_links.py — Extract all hyperlinks from a webpage and optionally save as URLs file.

Usage:
  python3 scrape_links.py --url "https://mankato.mnsu.edu/..." --output urls.txt
"""

import argparse
import json
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright


def scrape_links(start_url: str, same_domain_only: bool = True, subpath_only: str = None) -> list:
    """Extract all hyperlinks from a webpage.

    Args:
        start_url: URL to scrape
        same_domain_only: Only include links from the same domain
        subpath_only: Only include links that have this subpath (e.g., "/employment/")
    """
    links = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
        )
        page = context.new_page()

        try:
            page.goto(start_url, wait_until="networkidle", timeout=30000)

            # Extract all href attributes
            hrefs = page.query_selector_all("a[href]")

            start_domain = urlparse(start_url).netloc

            for anchor in hrefs:
                href = anchor.get_attribute("href")
                if not href:
                    continue

                # Convert relative URLs to absolute
                absolute_url = urljoin(start_url, href)

                # Remove fragments and query params if needed
                parsed = urlparse(absolute_url)
                clean_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

                # Filter by domain if requested
                if same_domain_only:
                    if urlparse(clean_url).netloc != start_domain:
                        continue

                # Filter by subpath if requested
                if subpath_only:
                    if subpath_only not in clean_url:
                        continue

                links.add(clean_url)

        finally:
            browser.close()

    return sorted(list(links))


def main():
    parser = argparse.ArgumentParser(description="Scrape hyperlinks from a webpage.")
    parser.add_argument("--url", required=True, help="Starting URL to scrape")
    parser.add_argument("--output", help="Save links to file (one per line)")
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    parser.add_argument("--all-domains", action="store_true", help="Include links from other domains")
    parser.add_argument("--subpath", help="Only include links containing this subpath (e.g., '/employment/')")
    args = parser.parse_args()

    print(f"Scraping links from: {args.url}")
    if args.subpath:
        print(f"Filtering for subpath: {args.subpath}")

    links = scrape_links(args.url, same_domain_only=not args.all_domains, subpath_only=args.subpath)

    print(f"Found {len(links)} links")

    if args.json:
        print(json.dumps(links, indent=2))
    else:
        for link in links:
            print(link)

    if args.output:
        with open(args.output, "w") as f:
            for link in links:
                f.write(f"{link}\n")
        print(f"\nSaved to: {args.output}")


if __name__ == "__main__":
    main()
