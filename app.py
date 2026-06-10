#!/usr/bin/env python3
"""
app.py — Local web app: paste a URL, see its sublinks, download all as PDFs.
Run:  python app.py
Open: http://localhost:5000
"""

import json, os, queue, re, tempfile, threading, time, uuid, zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

from flask import Flask, Response, jsonify, render_template, request, send_file
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

app = Flask(__name__)
jobs: dict = {}          # job_id -> {"type": "scrape"|"download", "queue": Queue, "links": list|None, "zip_path": str|None}

OUTPUT_DIR = Path("mankato_pdfs")
MANIFEST_PATH = OUTPUT_DIR / "manifest.json"


# ── helpers ──────────────────────────────────────────────────────────────────

def url_to_filepath(url: str) -> Path:
    """
    Map a URL to a relative file path mirroring the URL structure.
    e.g. /pharmacy/prescriptions/refill/ -> pharmacy/prescriptions/refill/refill.pdf
    The last path segment becomes both the folder name and the filename.
    Index pages (ending in /) are named after their own folder segment.
    """
    parsed = urlparse(url)
    # Strip leading slash, split into segments, drop empty strings
    segments = [s for s in parsed.path.strip("/").split("/") if s]
    if not segments:
        return Path("index.pdf")
    # All segments form the folder path; last segment is also the filename
    folder = Path(*segments)
    filename = segments[-1] + ".pdf"
    return folder / filename


def load_manifest() -> dict:
    """Return manifest as a dict keyed by URL."""
    if MANIFEST_PATH.exists():
        with open(MANIFEST_PATH) as f:
            entries = json.load(f)
        return {e["url"]: e for e in entries}
    return {}


def save_manifest(manifest: dict) -> None:
    """Write manifest dict back to disk."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(MANIFEST_PATH, "w") as f:
        json.dump(list(manifest.values()), f, indent=2)


def _extract_sublinks(page, start_url: str, base: str, domain: str) -> set[str]:
    """Extract all sublinks from the current page that are within base path."""
    found = set()
    for a in page.query_selector_all("a[href]"):
        href = a.get_attribute("href") or ""
        abs_url = urljoin(start_url, href)
        p2 = urlparse(abs_url)
        clean = f"{p2.scheme}://{p2.netloc}{p2.path}"
        clean_path = p2.path.rstrip("/")
        if p2.netloc == domain and (
            clean_path == base or clean_path.startswith(base + "/")
        ):
            found.add(clean)
    return found


def _scrape_worker(job_id: str, start_url: str, q: queue.Queue, max_pages: int = 100) -> None:
    """Background thread: crawl all sublinks under start_url."""
    try:
        p0 = urlparse(start_url)
        base = p0.path.rstrip("/")
        domain = p0.netloc

        visited: set[str] = set()
        to_visit: list[str] = [start_url.rstrip("/") + "/"]
        all_found: set[str] = set()

        with sync_playwright() as pw:
            browser = pw.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
            page = browser.new_context(
                user_agent="Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36"
            ).new_page()
            try:
                while to_visit and len(visited) < max_pages:
                    url = to_visit.pop(0)
                    if url in visited:
                        continue
                    visited.add(url)

                    q.put({"type": "crawling", "url": url, "count": len(visited)})

                    try:
                        page.goto(url, wait_until="networkidle", timeout=15_000)
                    except Exception:
                        continue

                    new_links = _extract_sublinks(page, url, base, domain)
                    all_found.update(new_links)

                    for link in new_links:
                        if link not in visited and link not in to_visit:
                            to_visit.append(link)
            finally:
                browser.close()

        manifest = load_manifest()
        links_with_status = [
            {
                "url": l,
                "done": l in manifest,
                "file": manifest[l]["file"] if l in manifest else None,
                "scraped_at": manifest[l]["scraped_at"] if l in manifest else None,
            }
            for l in sorted(all_found)
        ]
        q.put({"type": "complete", "links": links_with_status})
    except Exception as e:
        q.put({"type": "fatal", "reason": str(e)})


def _pdf_worker(job_id: str, links: list[str], q: queue.Queue) -> None:
    """Background thread: render each URL to PDF, save with mirrored structure, zip, signal done."""
    manifest = load_manifest()
    pdf_paths: list[Path] = []

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
            page = browser.new_context(
                user_agent="Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36"
            ).new_page()

            for i, url in enumerate(links, 1):
                # Skip URLs already in the manifest
                if url in manifest:
                    q.put({"type": "skipped", "current": i, "total": len(links), "url": url,
                           "file": manifest[url]["file"]})
                    continue

                q.put({"type": "progress", "current": i, "total": len(links), "url": url})

                rel_path = url_to_filepath(url)          # e.g. pharmacy/prescriptions/refill/refill.pdf
                out = OUTPUT_DIR / rel_path
                out.parent.mkdir(parents=True, exist_ok=True)

                try:
                    # Use screen media so the page renders like a real browser
                    page.emulate_media(media="screen")
                    page.goto(url, wait_until="networkidle", timeout=30_000)
                    page.wait_for_timeout(500)

                    # Surgically remove only elements known to break multi-page PDFs.
                    # DO NOT use * { position: static } — it collapses the entire layout.
                    page.add_style_tag(content="""
                        /* Hide chrome elements that don't belong in a PDF */
                        header, nav, .nav, .navbar, .site-header,
                        footer, .footer, .site-footer,
                        .cookie-banner, .chat-widget, [class*="overlay"] {
                            display: none !important;
                        }

                        /* Hide all images, icons, and decorative visuals */
                        img, svg, video, audio, canvas,
                        [class*="icon"], [class*="logo"], [class*="banner"],
                        [class*="hero"], [class*="thumbnail"], [class*="carousel"],
                        [class*="slider"], [class*="gallery"], [class*="image"],
                        picture, figure, iframe {
                            display: none !important;
                        }

                        /* Strip background images but keep background colors for structure */
                        * {
                            background-image: none !important;
                        }

                        /* Unpin only fixed/sticky elements so they don't
                           repeat or cover content on pages 2+ */
                        [style*="position: fixed"], [style*="position:fixed"],
                        [style*="position: sticky"], [style*="position:sticky"] {
                            position: static !important;
                        }

                        /* Prevent content being clipped mid-element at page breaks */
                        p, h1, h2, h3, h4, h5, h6, table, li, blockquote {
                            page-break-inside: avoid;
                            break-inside: avoid;
                        }
                    """)

                    page.pdf(
                        path=str(out), format="A4", print_background=True,
                        scale=0.9,
                        margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"},
                    )
                    pdf_paths.append(out)
                    size_kb = round(out.stat().st_size / 1024, 1)

                    # Update manifest immediately after each successful save
                    manifest[url] = {
                        "url": url,
                        "file": str(rel_path),
                        "scraped_at": datetime.now().isoformat(timespec="seconds"),
                        "status": "ok",
                        "size_kb": size_kb,
                    }
                    save_manifest(manifest)

                    q.put({"type": "done_one", "url": url,
                           "file": str(rel_path), "size_kb": size_kb})
                except Exception as e:
                    q.put({"type": "error_one", "url": url, "reason": str(e)})

                if i < len(links):
                    time.sleep(0.4)

            browser.close()

        # Zip preserving the folder structure
        q.put({"type": "status", "message": "Creating ZIP file..."})
        zip_path = OUTPUT_DIR / "pages.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, p in enumerate(pdf_paths, 1):
                zf.write(p, p.relative_to(OUTPUT_DIR))
                if i % 20 == 0:  # Status update every 20 files
                    q.put({"type": "status", "message": f"Compressing ZIP… {i}/{len(pdf_paths)}"})
        jobs[job_id]["zip_path"] = str(zip_path)
        q.put({"type": "complete"})

    except Exception as e:
        q.put({"type": "fatal", "reason": str(e)})


# ── routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/scrape", methods=["POST"])
def scrape():
    url = (request.json or {}).get("url", "").strip()
    if not url:
        return jsonify({"error": "No URL provided"}), 400

    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    jobs[job_id] = {"type": "scrape", "queue": q, "links": None, "zip_path": None}
    threading.Thread(target=_scrape_worker, args=(job_id, url, q), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/start-download", methods=["POST"])
def start_download():
    links = (request.json or {}).get("links", [])
    if not links:
        return jsonify({"error": "No links provided"}), 400
    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    jobs[job_id] = {"queue": q, "zip_path": None}
    threading.Thread(target=_pdf_worker, args=(job_id, links, q), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/progress/<job_id>")
def progress(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    def stream():
        q = job["queue"]
        while True:
            try:
                evt = q.get(timeout=90)
                # Store links if this is a scrape job
                if evt.get("type") == "complete" and job.get("type") == "scrape":
                    job["links"] = evt.get("links", [])
                yield f"data: {json.dumps(evt)}\n\n"
                if evt.get("type") in ("complete", "fatal"):
                    break
            except queue.Empty:
                yield 'data: {"type":"heartbeat"}\n\n'

    return Response(stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/get-zip/<job_id>")
def get_zip(job_id: str):
    job = jobs.get(job_id)
    if not job or not job.get("zip_path"):
        return "Not ready", 404
    return send_file(job["zip_path"], as_attachment=True, download_name="pages.zip")


@app.route("/clear-manifest", methods=["POST"])
def clear_manifest():
    """Clear the manifest to start fresh."""
    try:
        if MANIFEST_PATH.exists():
            MANIFEST_PATH.unlink()
        return jsonify({"status": "cleared"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000, threaded=True)
