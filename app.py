#!/usr/bin/env python3
"""
app.py — Local web app: paste a URL, see its sublinks, download all as PDFs.
Run:  python app.py
Open: http://localhost:5000
"""

import asyncio, hashlib, html, json, os, queue, re, tempfile, threading, time, uuid, zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

import anthropic
import docx
from docx.shared import Pt, Inches
import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, Response, jsonify, render_template, request, send_file
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright
from classifier import normalize_platform, clamp_platform

load_dotenv()
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# ── Content Management Guide → cached Claude system prompt ─────────────────────
# The guide markdown is the single source of truth.  We extract its
# platform-definitions section and inject it verbatim into the system prompt so
# that editing the guide changes the AI's behavior with no code changes.
GUIDE_PATH = Path("University_Comprehensive_Content_Management_Guide.md")
GUIDE_SECTION_START = "## Overview of MSU's Main Content Platforms"
GUIDE_SECTION_END = "## Content Audit"


def load_platform_guide() -> str:
    """Extract the platform-definitions section from the content management guide."""
    try:
        text = GUIDE_PATH.read_text(encoding="utf-8")
    except Exception:
        return ""
    start = text.find(GUIDE_SECTION_START)
    if start == -1:
        return ""
    end = text.find(GUIDE_SECTION_END, start + len(GUIDE_SECTION_START))
    return (text[start:end] if end != -1 else text[start:]).strip()


def guide_version() -> str:
    """Short hash of the guide file, used to invalidate cached suggestions on edit."""
    try:
        return hashlib.md5(GUIDE_PATH.read_bytes()).hexdigest()[:8]
    except Exception:
        return "noguide"


# Fixed instructions; the large, stable guide block is appended and prompt-cached.
TIER2_INSTRUCTIONS = """You are a content strategist at Minnesota State University, Mankato (MNSU).
Your job: read a university web page's content and decide which single MNSU content platform it should live on, based STRICTLY on the official Content Management Guide provided below.

You will receive a page's title, URL, and a content excerpt. Judge by the actual content and its primary audience and purpose — NOT by the URL structure. A page sitting under a public-facing section can still belong on another platform.

VALID PLATFORMS (respond with one of these EXACT names):
- Website
- Maverick OneStop
- The Fountain
- MavLife / Student Hub
- Teams / SharePoint
- No Clear Fit

DECISION GUIDANCE (always defer to the guide below):
- Policies, conduct codes, required procedures, forms, FAQs, and step-by-step task instructions → Maverick OneStop, even under a public URL like /housing/policies/.
- Purely informational or marketing pages for prospective students, families, or the general public → Website.
- Content exclusively for employees (not students) → The Fountain.
- Student involvement, clubs, activities, recreation, engagement → MavLife / Student Hub.
- Internal department/committee collaboration material → Teams / SharePoint.
- If nothing genuinely fits, use "No Clear Fit".

CONFIDENCE CALIBRATION — be honest, never manufacture confidence:
- High: The content unambiguously fits ONE platform; an expert would agree instantly with no reasonable alternative.
- Medium: One platform is the best fit, but a defensible argument exists for one alternative.
- Low: Multiple platforms are genuinely defensible, or the excerpt is too thin to tell. Use this freely.

Respond with a single raw JSON object — no markdown, no code fences, no extra text:
{"platform": "<exact platform name>", "confidence": "High|Medium|Low", "reason": "<one concise sentence grounded in the guide>"}

--- OFFICIAL CONTENT MANAGEMENT GUIDE (SOURCE OF TRUTH) ---
"""

PLATFORM_GUIDE = load_platform_guide()
GUIDE_VERSION = guide_version()

# Anthropic prompt-caching: the large, constant guide block is marked ephemeral
# so it is cached across the whole batch, cutting input cost/latency dramatically.
TIER2_SYSTEM_BLOCKS = [
    {"type": "text", "text": TIER2_INSTRUCTIONS},
    {"type": "text", "text": PLATFORM_GUIDE, "cache_control": {"type": "ephemeral"}},
]

# Concurrency + model settings for the AI analysis pipeline.
# ANALYSIS_CONCURRENCY caps how many pages are fetched + classified at once.
# Each in-flight page is a live Chromium tab, so this is the main CPU/RAM dial.
# Default 3 keeps a typical laptop responsive; raise on a beefy machine, lower
# (e.g. 2) on memory-constrained hosts like the Render free tier (512 MB RAM).
TIER2_MODEL = "claude-haiku-4-5"
AI_CREDIT_ERROR_MARKERS = (
    "credit balance",
    "insufficient credit",
    "insufficient_quota",
    "billing",
    "payment required",
    "quota exceeded",
)


class AIQuotaError(Exception):
    pass


def env_int(name: str, default: int, minimum: int | None = None) -> int:
    try:
        value = int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        value = default
    if minimum is not None:
        value = max(value, minimum)
    return value


ANALYSIS_CONCURRENCY = env_int("ANALYSIS_CONCURRENCY", 2, 1)

# Restart the shared Chromium browser every N page loads.  A single long-lived
# context accumulates memory across thousands of navigations, which is what
# freezes machines on large (~2000 link) runs.  Recycling caps peak RAM.
ANALYSIS_BROWSER_RECYCLE = env_int("ANALYSIS_BROWSER_RECYCLE", 100, 1)

# Per-page navigation timeout (ms) for the analysis fetch.
ANALYSIS_NAV_TIMEOUT = env_int("ANALYSIS_NAV_TIMEOUT", 15000, 1000)
CRAWL_NAV_TIMEOUT = env_int("CRAWL_NAV_TIMEOUT", 10000, 1000)
PDF_NAV_TIMEOUT = env_int("PDF_NAV_TIMEOUT", 20000, 1000)
PDF_SETTLE_MS = env_int("PDF_SETTLE_MS", 750, 0)
PDF_PAGE_RECYCLE = env_int("PDF_PAGE_RECYCLE", 75, 1)
SSE_HEARTBEAT_SECONDS = env_int("SSE_HEARTBEAT_SECONDS", 20, 5)
MAX_ACTIVE_JOBS = env_int("MAX_ACTIVE_JOBS", 2, 1)
JOB_RETENTION_SECONDS = env_int("JOB_RETENTION_SECONDS", 3600, 60)

DATA_DIR = Path(os.getenv("DATA_DIR", tempfile.gettempdir() if os.getenv("RAILWAY_ENVIRONMENT") else "."))
SUGGEST_CACHE_DIR = DATA_DIR / "suggestion_cache"

# Safety ceiling on how many pages a single crawl will visit.  The BFS normally
# stops on its own once it exhausts the in-scope queue; this only caps runaway
# crawls.  Default is well above large sections so full coverage isn't truncated.
CRAWL_MAX_PAGES = env_int("CRAWL_MAX_PAGES", 1000, 1)
SAFE_PATH_SEGMENT_MAX = env_int("SAFE_PATH_SEGMENT_MAX", 48, 16)
SAFE_REL_PATH_MAX = env_int("SAFE_REL_PATH_MAX", 150, 80)

app = Flask(__name__)

# ── CORS ─────────────────────────────────────────────────────────────────────
# Allow any origin so a separately-deployed frontend (e.g. on Vercel) can call
# this backend and open SSE streams.  No credentials are used, so * is safe.
ALLOWED_ORIGIN = os.getenv("ALLOWED_ORIGIN", "*")

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"]  = ALLOWED_ORIGIN
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response

@app.route("/", defaults={"path": ""}, methods=["OPTIONS"])
@app.route("/<path:path>", methods=["OPTIONS"])
def handle_preflight(path):
    """Return 200 for every CORS preflight so fetch() with JSON bodies works cross-origin."""
    resp = app.make_default_options_response()
    resp.headers["Access-Control-Allow-Origin"]  = ALLOWED_ORIGIN
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp
jobs: dict = {}          # job_id -> {"type": "scrape"|"download", "queue": Queue, "links": list|None, "zip_path": str|None}
jobs_lock = threading.Lock()

OUTPUT_DIR = DATA_DIR / "mankato_pdfs"
MANIFEST_PATH = OUTPUT_DIR / "manifest.json"
CRAWL_CACHE_DIR = DATA_DIR / "crawl_cache"  # Cache crawl results by starting URL


def cleanup_jobs() -> None:
    cutoff = time.time() - JOB_RETENTION_SECONDS
    with jobs_lock:
        for job_id, job in list(jobs.items()):
            if job.get("finished_at", 0) and job["finished_at"] < cutoff:
                jobs.pop(job_id, None)


def register_job(job_id: str, job: dict) -> bool:
    cleanup_jobs()
    with jobs_lock:
        active = sum(1 for item in jobs.values() if not item.get("finished_at"))
        if active >= MAX_ACTIVE_JOBS:
            return False
        job["created_at"] = time.time()
        job["cancelled"] = False
        jobs[job_id] = job
        return True


def should_stop(job_id: str) -> bool:
    with jobs_lock:
        job = jobs.get(job_id)
        return bool(job and job.get("cancelled"))


def cancel_job(job_id: str) -> bool:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.get("finished_at"):
            return False
        job["cancelled"] = True
        job.setdefault("status", {}).update(stage="cancelling", updated_at=time.time())
        q = job.get("queue")
    if q:
        q.put({"type": "cancelling"})
    return True


def timestamp_slug() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def update_job_status(job_id: str, **status) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if job:
            current = job.setdefault("status", {})
            current.update(status, updated_at=time.time())


def log_job(job_id: str, message: str, **details) -> None:
    detail_text = " ".join(f"{key}={value}" for key, value in details.items())
    print(f"[job:{job_id}] {message} {detail_text}".rstrip(), flush=True)


def finish_job(job_id: str) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if job and not job.get("finished_at"):
            job["finished_at"] = time.time()


def emit_terminal(q: queue.Queue, job_id: str, event: dict) -> None:
    update_job_status(job_id, stage=event.get("type"), terminal=True)
    finish_job(job_id)
    q.put(event)


HEAVY_RESOURCE_TYPES = {"image", "media", "font"}
BLOCKED_URL_PARTS = ("googletagmanager", "google-analytics", "doubleclick", "hotjar", "fullstory")


def should_block_resource(request, include_stylesheets: bool = False) -> bool:
    if request.resource_type in HEAVY_RESOURCE_TYPES:
        return True
    if include_stylesheets and request.resource_type == "stylesheet":
        return True
    return any(part in request.url.lower() for part in BLOCKED_URL_PARTS)


def block_resource_route(route, include_stylesheets: bool = False) -> None:
    if should_block_resource(route.request, include_stylesheets):
        route.abort()
    else:
        route.continue_()


async def block_resource_route_async(route, include_stylesheets: bool = False) -> None:
    if should_block_resource(route.request, include_stylesheets):
        await route.abort()
    else:
        await route.continue_()


# ── helpers ──────────────────────────────────────────────────────────────────

def safe_path_segment(value: str, fallback: str = "page") -> str:
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip(".-_")
    slug = re.sub(r"-+", "-", slug) or fallback
    return slug[:SAFE_PATH_SEGMENT_MAX].strip(".-_") or fallback


def path_text_length(path: Path) -> int:
    return len(str(path).replace("\\", "/"))


def url_to_filepath(url: str) -> Path:
    parsed = urlparse(url)
    url_hash = hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]
    segments = [safe_path_segment(s) for s in parsed.path.strip("/").split("/") if s]
    if not segments:
        return Path(f"index-{url_hash}.pdf")

    filename_base = safe_path_segment(segments[-1], "page")
    filename = f"{filename_base}-{url_hash}.pdf"
    dirs = segments[:-1]
    rel_path = Path(*dirs, filename) if dirs else Path(filename)

    while dirs and path_text_length(rel_path) > SAFE_REL_PATH_MAX:
        dirs.pop(0)
        rel_path = Path(*dirs, filename) if dirs else Path(filename)

    if path_text_length(rel_path) > SAFE_REL_PATH_MAX:
        filename_base = filename_base[:max(16, SAFE_REL_PATH_MAX - len(url_hash) - 6)].strip(".-_") or "page"
        rel_path = Path(f"{filename_base}-{url_hash}.pdf")

    return rel_path


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


def get_crawl_cache_path(start_url: str) -> Path:
    """Get the cache file path for a given starting URL (keyed by URL hash)."""
    import hashlib
    url_hash = hashlib.md5(start_url.encode()).hexdigest()[:8]
    CRAWL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CRAWL_CACHE_DIR / f"crawl_{url_hash}.json"


def load_crawl_cache(start_url: str) -> list | None:
    """Load cached crawl results for this URL, if they exist."""
    cache_path = get_crawl_cache_path(start_url)
    if cache_path.exists():
        try:
            with open(cache_path) as f:
                return json.load(f)
        except Exception:
            return None
    return None


def save_crawl_cache(start_url: str, links: list) -> None:
    """Save crawl results to cache."""
    cache_path = get_crawl_cache_path(start_url)
    with open(cache_path, "w") as f:
        json.dump(links, f, indent=2)


def _fetch_html(url: str) -> str:
    req = Request(url, headers={
        "User-Agent": "Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })
    with urlopen(req, timeout=CRAWL_NAV_TIMEOUT / 1000) as response:
        content_type = response.headers.get("content-type", "")
        if "html" not in content_type.lower():
            return ""
        raw = response.read(2_000_000)
        charset = response.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="ignore")


def _extract_sublinks_from_html(markup: str, start_url: str, base: str, domain: str) -> set[str]:
    found = set()
    for match in re.finditer(r'''href\s*=\s*["']([^"'#]+)["']''', markup, flags=re.I):
        href = html.unescape(match.group(1)).strip()
        if not href or href.startswith(("mailto:", "tel:", "javascript:")):
            continue
        abs_url = urljoin(start_url, href)
        p2 = urlparse(abs_url)
        clean = f"{p2.scheme}://{p2.netloc}{p2.path}"
        clean_path = p2.path.rstrip("/")
        if "~/link/" in clean_path:
            continue
        if p2.netloc == domain and (clean_path == base or clean_path.startswith(base + "/")):
            found.add(clean)
    return found


def _scrape_worker(job_id: str, start_url: str, q: queue.Queue,
                   max_pages: int = CRAWL_MAX_PAGES, refresh: bool = False) -> None:
    """Background thread: crawl all sublinks under start_url (with caching)."""
    try:
        # Check if we have cached crawl results for this URL (unless a fresh
        # re-crawl was explicitly requested via refresh).
        cached_links = None if refresh else load_crawl_cache(start_url)
        if cached_links:
            q.put({"type": "using_cache", "count": len(cached_links)})
            manifest = load_manifest()
            links_with_status = [
                {
                    "url": l,
                    "done": l in manifest,
                    "file": manifest[l]["file"] if l in manifest else None,
                    "scraped_at": manifest[l]["scraped_at"] if l in manifest else None,
                }
                for l in cached_links
            ]
            emit_terminal(q, job_id, {"type": "complete", "links": links_with_status})
            return

        # No cache — crawl normally
        p0 = urlparse(start_url)
        base = p0.path.rstrip("/")
        domain = p0.netloc

        visited: set[str] = set()
        to_visit: list[str] = [start_url.rstrip("/") + "/"]
        all_found: set[str] = set()
        skipped = 0
        started_at = time.time()
        log_job(job_id, "crawl started", start_url=start_url, max_pages=max_pages, timeout_ms=CRAWL_NAV_TIMEOUT)
        update_job_status(job_id, stage="starting", current_url=start_url, visited=0, found=0, queued=1, skipped=0)

        while to_visit and len(visited) < max_pages:
            if should_stop(job_id):
                log_job(job_id, "crawl cancelled", visited=len(visited), found=len(all_found), skipped=skipped)
                emit_terminal(q, job_id, {"type": "cancelled", "reason": "Crawl cancelled."})
                return
            url = to_visit.pop(0)
            if url in visited:
                continue
            visited.add(url)

            elapsed = round(time.time() - started_at, 1)
            status = {"type": "crawling", "url": url, "count": len(visited),
                      "found": len(all_found), "queued": len(to_visit),
                      "skipped": skipped, "elapsed": elapsed, "stage": "loading"}
            update_job_status(job_id, stage="loading", current_url=url, visited=len(visited),
                              found=len(all_found), queued=len(to_visit), skipped=skipped, elapsed=elapsed)
            log_job(job_id, "loading page", visited=len(visited), found=len(all_found), queued=len(to_visit), url=url)
            q.put(status)

            try:
                markup = _fetch_html(url)
            except Exception as e:
                skipped += 1
                reason = str(e).splitlines()[0][:180]
                update_job_status(job_id, stage="skipped", current_url=url, visited=len(visited),
                                  found=len(all_found), queued=len(to_visit), skipped=skipped,
                                  last_error=reason, elapsed=round(time.time() - started_at, 1))
                log_job(job_id, "skipped page", url=url, reason=reason)
                q.put({"type": "skipped", "url": url, "count": len(visited),
                       "found": len(all_found), "queued": len(to_visit),
                       "skipped": skipped, "reason": reason})
                continue

            update_job_status(job_id, stage="extracting", current_url=url, visited=len(visited),
                              found=len(all_found), queued=len(to_visit), skipped=skipped,
                              elapsed=round(time.time() - started_at, 1))
            new_links = _extract_sublinks_from_html(markup, url, base, domain)
            all_found.update(new_links)
            log_job(job_id, "extracted links", url=url, new_links=len(new_links), total_found=len(all_found))

            for link in new_links:
                if link not in visited and link not in to_visit:
                    to_visit.append(link)
            update_job_status(job_id, stage="queued", current_url=url, visited=len(visited),
                              found=len(all_found), queued=len(to_visit), skipped=skipped,
                              elapsed=round(time.time() - started_at, 1))

        # Save crawl results to cache
        all_found_sorted = sorted(all_found)
        save_crawl_cache(start_url, all_found_sorted)

        manifest = load_manifest()
        links_with_status = [
            {
                "url": l,
                "done": l in manifest,
                "file": manifest[l]["file"] if l in manifest else None,
                "scraped_at": manifest[l]["scraped_at"] if l in manifest else None,
            }
            for l in all_found_sorted
        ]
        elapsed = round(time.time() - started_at, 1)
        log_job(job_id, "crawl complete", visited=len(visited), found=len(all_found_sorted), skipped=skipped, elapsed=elapsed)
        emit_terminal(q, job_id, {"type": "complete", "links": links_with_status,
                                  "visited": len(visited), "skipped": skipped, "elapsed": elapsed})
    except Exception as e:
        log_job(job_id, "crawl fatal", reason=str(e).splitlines()[0][:180])
        emit_terminal(q, job_id, {"type": "fatal", "reason": str(e)})


def _pdf_worker(job_id: str, links: list[str], q: queue.Queue) -> None:
    """Background thread: render each URL to PDF, save with mirrored structure, zip, signal done."""
    manifest = load_manifest()
    pdf_paths: list[Path] = []

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(args=["--no-sandbox", "--disable-setuid-sandbox"])
            context = None
            page = None
            render_count = 0

            def open_page():
                ctx = browser.new_context(
                    user_agent="Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36"
                )
                ctx.route("**/*", lambda route: block_resource_route(route))
                return ctx, ctx.new_page()

            def close_page():
                if page:
                    page.close()
                if context:
                    context.close()

            context, page = open_page()
            try:
                for i, url in enumerate(links, 1):
                    if should_stop(job_id):
                        emit_terminal(q, job_id, {"type": "cancelled", "reason": "PDF generation cancelled."})
                        return
                    if url in manifest:
                        q.put({"type": "skipped", "current": i, "total": len(links), "url": url,
                               "file": manifest[url]["file"]})
                        continue

                    if render_count and render_count % PDF_PAGE_RECYCLE == 0:
                        close_page()
                        context, page = open_page()
                        q.put({"type": "status", "message": f"Recycled browser after {render_count} PDFs…"})

                    update_job_status(job_id, stage="pdf_rendering", current=i, total=len(links), current_url=url,
                                      rendered=render_count, errors=jobs[job_id].get("status", {}).get("errors", 0))
                    q.put({"type": "progress", "current": i, "total": len(links), "url": url})

                    rel_path = url_to_filepath(url)
                    out = OUTPUT_DIR / rel_path
                    out.parent.mkdir(parents=True, exist_ok=True)

                    try:
                        page.emulate_media(media="screen")
                        page.goto(url, wait_until="domcontentloaded", timeout=PDF_NAV_TIMEOUT)
                        page.wait_for_timeout(PDF_SETTLE_MS)
                        page.add_style_tag(content="""
                            header, nav, .nav, .navbar, .site-header,
                            footer, .footer, .site-footer,
                            .cookie-banner, .chat-widget, [class*="overlay"] {
                                display: none !important;
                            }
                            img, svg, video, audio, canvas,
                            [class*="icon"], [class*="logo"], [class*="banner"],
                            [class*="hero"], [class*="thumbnail"], [class*="carousel"],
                            [class*="slider"], [class*="gallery"], [class*="image"],
                            picture, figure, iframe {
                                display: none !important;
                            }
                            * {
                                background-image: none !important;
                            }
                            [style*="position: fixed"], [style*="position:fixed"],
                            [style*="position: sticky"], [style*="position:sticky"] {
                                position: static !important;
                            }
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
                        render_count += 1
                        size_kb = round(out.stat().st_size / 1024, 1)
                        manifest[url] = {
                            "url": url,
                            "file": str(rel_path),
                            "scraped_at": datetime.now().isoformat(timespec="seconds"),
                            "status": "ok",
                            "size_kb": size_kb,
                        }
                        save_manifest(manifest)

                        update_job_status(job_id, stage="pdf_done", current=i, total=len(links), current_url=url,
                                          rendered=render_count, last_file=str(rel_path))
                        q.put({"type": "done_one", "url": url,
                               "file": str(rel_path), "size_kb": size_kb})
                    except Exception as e:
                        current_status = jobs[job_id].get("status", {})
                        error_count = current_status.get("errors", 0) + 1
                        update_job_status(job_id, stage="pdf_error", current=i, total=len(links), current_url=url,
                                          rendered=render_count, errors=error_count, last_error=str(e)[:200])
                        q.put({"type": "error_one", "url": url, "reason": str(e)})

                    if i < len(links):
                        time.sleep(0.4)
            finally:
                close_page()
                browser.close()

        if should_stop(job_id):
            emit_terminal(q, job_id, {"type": "cancelled", "reason": "PDF generation cancelled."})
            return

        update_job_status(job_id, stage="zipping", rendered=render_count, total=len(links))
        q.put({"type": "status", "message": "Creating ZIP file..."})
        zip_path = OUTPUT_DIR / f"pages-{job_id}-{timestamp_slug()}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, p in enumerate(pdf_paths, 1):
                zf.write(p, p.relative_to(OUTPUT_DIR))
                if i % 20 == 0:  # Status update every 20 files
                    q.put({"type": "status", "message": f"Compressing ZIP… {i}/{len(pdf_paths)}"})
        jobs[job_id]["zip_path"] = str(zip_path)
        emit_terminal(q, job_id, {"type": "complete"})

    except Exception as e:
        emit_terminal(q, job_id, {"type": "fatal", "reason": str(e)})



def _docx_worker(job_id: str, links: list[str], q: queue.Queue) -> None:
    manifest = load_manifest()
    docx_paths = []
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(args=['--no-sandbox', '--disable-setuid-sandbox'])
            context = None; page = None; render_count = 0
            def open_page():
                ctx = browser.new_context(user_agent='Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36')
                ctx.route('**/*', lambda route: block_resource_route(route))
                return ctx, ctx.new_page()
            def close_page():
                if page: page.close()
                if context: context.close()
            context, page = open_page()
            try:
                for i, url in enumerate(links, 1):
                    if should_stop(job_id):
                        emit_terminal(q, job_id, {'type': 'cancelled', 'reason': 'DOCX generation cancelled.'})
                        return
                    if render_count and render_count % PDF_PAGE_RECYCLE == 0:
                        close_page(); context, page = open_page()
                        q.put({'type': 'status', 'message': f'Recycled browser after {render_count} DOCXs'})
                    update_job_status(job_id, stage='docx_rendering', current=i, total=len(links), current_url=url, rendered=render_count)
                    q.put({'type': 'progress', 'current': i, 'total': len(links), 'url': url})
                    rel_path = url_to_filepath(url).with_suffix('.docx')
                    out = OUTPUT_DIR / rel_path
                    out.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        page.emulate_media(media='screen')
                        page.goto(url, wait_until='domcontentloaded', timeout=PDF_NAV_TIMEOUT)
                        page.wait_for_timeout(PDF_SETTLE_MS)
                        title = page.title().strip() or 'Untitled'
                        body_text = page.inner_text('body').strip()
                        doc = docx.Document()
                        section = doc.sections[0]
                        section.top_margin = Inches(1); section.bottom_margin = Inches(1)
                        section.left_margin = Inches(1); section.right_margin = Inches(1)
                        heading = doc.add_heading(title, level=1)
                        heading.runs[0].font.size = Pt(16)
                        for para_text in body_text.split('\n'):
                            line = para_text.strip()
                            if line:
                                p = doc.add_paragraph(line)
                                p.style.font.size = Pt(11)
                        doc.save(str(out))
                        docx_paths.append(out); render_count += 1
                        size_kb = round(out.stat().st_size / 1024, 1)
                        update_job_status(job_id, stage='docx_done', current=i, total=len(links), current_url=url, rendered=render_count)
                        q.put({'type': 'done_one', 'url': url, 'file': str(rel_path), 'size_kb': size_kb})
                    except Exception as e:
                        q.put({'type': 'error_one', 'url': url, 'reason': str(e)})
                    if i < len(links):
                        time.sleep(0.4)
            finally:
                close_page(); browser.close()
        if should_stop(job_id):
            emit_terminal(q, job_id, {'type': 'cancelled', 'reason': 'DOCX generation cancelled.'})
            return
        update_job_status(job_id, stage='zipping', rendered=render_count, total=len(links))
        q.put({'type': 'status', 'message': 'Creating DOCX ZIP file...'})
        zip_path = OUTPUT_DIR / f'docx-{job_id}-{timestamp_slug()}.zip'
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i2, p in enumerate(docx_paths, 1):
                zf.write(p, p.relative_to(OUTPUT_DIR))
                if i2 % 20 == 0:
                    q.put({'type': 'status', 'message': f'Compressing ZIP {i2}/{len(docx_paths)}'})
        jobs[job_id]['docx_zip_path'] = str(zip_path)
        emit_terminal(q, job_id, {'type': 'complete'})
    except Exception as e:
        emit_terminal(q, job_id, {'type': 'fatal', 'reason': str(e)})

def export_excel(links: list[dict], start_url: str) -> Path:
    """
    Build an Excel workbook from the scraped link list.
    Columns: #, Page Title (last path segment), Full URL, Folder Path, PDF File, Downloaded, Date, Size KB
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scraped Pages"

    # ── Styles ────────────────────────────────────────────────────────────────
    header_fill   = PatternFill("solid", fgColor="4F46E5")
    header_font   = Font(color="FFFFFF", bold=True, size=11)
    done_fill     = PatternFill("solid", fgColor="DCFCE7")
    pending_fill  = PatternFill("solid", fgColor="FEF9C3")
    center        = Alignment(horizontal="center", vertical="center")
    wrap          = Alignment(wrap_text=True, vertical="top")
    thin          = Side(style="thin", color="E5E7EB")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Header row ────────────────────────────────────────────────────────────
    headers = ["#", "Page Name", "Full URL", "Folder Path", "PDF File", "Downloaded", "Date Scraped", "Size (KB)"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    ws.row_dimensions[1].height = 22

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, link in enumerate(links, 1):
        url      = link.get("url", "")
        done     = link.get("done", False)
        file     = link.get("file") or ""
        scraped  = link.get("scraped_at") or ""
        size_kb  = link.get("size_kb") or ""

        # Derive page name and folder path from the URL
        segments = [s for s in urlparse(url).path.strip("/").split("/") if s]
        page_name   = segments[-1].replace("-", " ").title() if segments else "Home"
        folder_path = "/".join(segments[:-1]) if len(segments) > 1 else "/"

        row_fill = done_fill if done else pending_fill
        row = [i, page_name, url, folder_path, file, "✓ Yes" if done else "✗ No", scraped, size_kb]

        for col, val in enumerate(row, 1):
            cell = ws.cell(row=i + 1, column=col, value=val)
            cell.fill = row_fill
            cell.border = border
            cell.alignment = wrap if col in (3, 4, 5) else center

    # ── Column widths ─────────────────────────────────────────────────────────
    widths = [5, 28, 60, 45, 50, 14, 20, 12]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A2"  # Keep header visible when scrolling

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    total       = len(links)
    downloaded  = sum(1 for l in links if l.get("done"))
    pending     = total - downloaded

    ws2["A1"], ws2["B1"] = "Starting URL", start_url
    ws2["A2"], ws2["B2"] = "Total Pages Found", total
    ws2["A3"], ws2["B3"] = "Downloaded", downloaded
    ws2["A4"], ws2["B4"] = "Pending", pending
    ws2["A5"], ws2["B5"] = "Exported At", datetime.now().isoformat(timespec="seconds")

    for row in ws2.iter_rows(min_row=1, max_row=5, min_col=1, max_col=2):
        for cell in row:
            cell.border = border
            if cell.column == 1:
                cell.font = Font(bold=True)

    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 70

    out_path = OUTPUT_DIR / f"scraped_pages_{uuid.uuid4().hex}.xlsx"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


# ── Suggestion result cache (keyed by URL + guide version) ─────────────────────

def _suggest_cache_path(url: str) -> Path:
    """Cache file for an AI suggestion, invalidated automatically when the guide changes."""
    key = hashlib.md5(f"{url}|{GUIDE_VERSION}".encode()).hexdigest()[:12]
    SUGGEST_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return SUGGEST_CACHE_DIR / f"{key}.json"


def _load_suggest_cache(url: str) -> dict | None:
    path = _suggest_cache_path(url)
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _save_suggest_cache(url: str, data: dict) -> None:
    try:
        _suggest_cache_path(url).write_text(json.dumps(data), encoding="utf-8")
    except Exception:
        pass


def is_credit_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return any(marker in text for marker in AI_CREDIT_ERROR_MARKERS)


def parse_ai_json(raw: str) -> dict:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw.strip())
    match = re.search(r"\{.*\}", raw, flags=re.S)
    if not match:
        raise ValueError("No JSON object found")
    return json.loads(match.group(0))


def fallback_platform_suggestion(page_name: str, url: str, excerpt: str) -> dict:
    text = f"{page_name} {url} {excerpt}".lower()
    if any(word in text for word in ("policy", "policies", "procedure", "form", "report", "complaint", "tobacco", "smoking", "violation")):
        return {
            "platform": "Maverick OneStop",
            "confidence": "Medium",
            "reason": "Fallback rule: policy, reporting, form, or procedure content usually belongs in Maverick OneStop.",
        }
    if any(word in text for word in ("student organization", "club", "homecoming", "activities", "mavlife")):
        return {
            "platform": "MavLife / Student Hub",
            "confidence": "Low",
            "reason": "Fallback rule: student activity content may belong in MavLife / Student Hub.",
        }
    return {
        "platform": "Website",
        "confidence": "Low",
        "reason": "Fallback rule used because AI returned no valid classification; public-facing informational content defaults to Website.",
    }


async def _call_claude(client, page_name: str, url: str, excerpt: str) -> dict | None:
    """Classify one page with Claude. Retries transient API failures with backoff."""
    user_msg = (
        f"Page title: {page_name}\n"
        f"URL: {url}\n\n"
        f"Content excerpt:\n{excerpt}\n\n"
        "Classify this page into the single best MNSU platform per the guide. "
        "If the content is ambiguous or too thin to tell, use Low confidence."
    )
    last_error = "AI classification failed"
    for attempt in range(3):
        try:
            resp = await client.messages.create(
                model=TIER2_MODEL,
                max_tokens=256,
                system=TIER2_SYSTEM_BLOCKS,
                messages=[{"role": "user", "content": user_msg}],
            )
        except Exception as e:
            if is_credit_error(e):
                raise AIQuotaError("Anthropic API credits appear to be exhausted or billing is unavailable. Add credits, then restart the analysis.") from e
            last_error = str(e)
            await asyncio.sleep(1.5 * (attempt + 1))
            continue
        try:
            raw = resp.content[0].text.strip()
            ai = parse_ai_json(raw)
            platform = clamp_platform(normalize_platform(ai.get("platform", "")))
            confidence = ai.get("confidence", "Low")
            if confidence not in ("High", "Medium", "Low"):
                confidence = "Low"
            reason = (ai.get("reason") or "").strip() or "No reason provided."
            return {"platform": platform, "confidence": confidence, "reason": reason}
        except Exception as e:
            last_error = str(e)
            await asyncio.sleep(0.5 * (attempt + 1))
    fallback = fallback_platform_suggestion(page_name, url, excerpt)
    fallback["reason"] = f"{fallback['reason']} AI fallback reason: {last_error[:120]}."
    return fallback


async def _run_ai_analysis(job_id: str, results: list[dict], by_index: dict[int, dict], q: queue.Queue) -> None:
    """
    Fetch + classify every page, streaming upgrades as each finishes.

    Memory safety for large (~2000 link) runs:
      * Cached URLs never launch Chromium — they're emitted straight from disk.
      * Uncached URLs are processed in bounded batches; the browser is fully
        relaunched between batches so Chromium memory can't grow unbounded.
      * Concurrency within a batch is capped by ANALYSIS_CONCURRENCY.
    """
    total = len(results)
    client = anthropic.AsyncAnthropic(api_key=ANTHROPIC_API_KEY)
    done = 0

    def _emit(res: dict) -> None:
        """Stream one finished result (success or error) to the SSE queue."""
        idx = res["index"]
        if res.get("error"):
            q.put({"type": "analysis_error", "index": idx, "reason": res["error"],
                   "done": done, "total": total})
            return
        stored = by_index[idx]
        stored.update(page_name=res["page_name"], platform=res["platform"],
                      confidence=res["confidence"], reason=res["reason"], tier=2)
        q.put({"type": "analysis_upgrade", "index": idx, "page_name": res["page_name"],
               "platform": res["platform"], "confidence": res["confidence"],
               "reason": res["reason"], "cached": res.get("cached", False),
               "done": done, "total": total})

    # ── 1. Serve cached results instantly (no browser involved) ───────────────
    uncached: list[dict] = []
    for r in results:
        if should_stop(job_id):
            return
        cached = _load_suggest_cache(r["url"])
        if cached:
            done += 1
            _emit({"index": r["index"], **cached, "cached": True})
        else:
            uncached.append(r)

    if not uncached:
        return

    # ── 2. Process uncached URLs in browser-recycling batches ─────────────────
    sem = asyncio.Semaphore(ANALYSIS_CONCURRENCY)
    batch_size = max(ANALYSIS_BROWSER_RECYCLE, ANALYSIS_CONCURRENCY)

    async with async_playwright() as pw:
        for start in range(0, len(uncached), batch_size):
            if should_stop(job_id):
                return
            batch = uncached[start:start + batch_size]

            browser = await pw.chromium.launch(
                args=["--no-sandbox", "--disable-setuid-sandbox"])
            context = await browser.new_context(
                user_agent="Mozilla/5.0 AppleWebKit/537.36 Chrome/120 Safari/537.36"
            )
            await context.route("**/*", lambda route: block_resource_route_async(route, True))

            async def analyze_one(result: dict) -> dict:
                idx, url = result["index"], result["url"]
                if should_stop(job_id):
                    return {"index": idx, "error": "cancelled"}
                async with sem:
                    if should_stop(job_id):
                        return {"index": idx, "error": "cancelled"}
                    page = await context.new_page()
                    try:
                        await page.goto(url, wait_until="domcontentloaded",
                                        timeout=ANALYSIS_NAV_TIMEOUT)
                        title = (await page.title()).strip()
                        raw_text = await page.inner_text("body")
                    except Exception as e:
                        return {"index": idx, "error": str(e)}
                    finally:
                        await page.close()

                    excerpt = " ".join(raw_text.split()[:600])
                    page_name = title or result["page_name"]
                    ai = await _call_claude(client, page_name, url, excerpt)

                if ai is None:
                    ai = fallback_platform_suggestion(page_name, url, excerpt)
                out = {"page_name": page_name, **ai}
                _save_suggest_cache(url, out)
                return {"index": idx, **out}

            try:
                tasks = [asyncio.create_task(analyze_one(r)) for r in batch]
                for fut in asyncio.as_completed(tasks):
                    if should_stop(job_id):
                        for task in tasks:
                            task.cancel()
                        await asyncio.gather(*tasks, return_exceptions=True)
                        return
                    try:
                        res = await fut
                    except AIQuotaError:
                        for task in tasks:
                            task.cancel()
                        await asyncio.gather(*tasks, return_exceptions=True)
                        raise
                    if res.get("error") == "cancelled":
                        return
                    done += 1
                    _emit(res)
            finally:
                await context.close()
                await browser.close()


def _analysis_worker(job_id: str, links: list[dict], q: queue.Queue) -> None:
    """
    Background thread: AI-only platform classification.
      1. Emit instant placeholder rows so the table fills immediately.
      2. Concurrently fetch each page and let Claude classify it using the guide.
    """
    total = len(links)
    results: list[dict] = []
    by_index: dict[int, dict] = {}

    try:
        # ── Instant placeholder rows ──────────────────────────────────────────
        for i, link in enumerate(links, 1):
            if should_stop(job_id):
                emit_terminal(q, job_id, {"type": "cancelled", "reason": "Analysis cancelled."})
                return
            url = link.get("url", "")
            segments = [s for s in urlparse(url).path.strip("/").split("/") if s]
            page_name = segments[-1].replace("-", " ").title() if segments else "Home"
            result = {
                "index":      i,
                "total":      total,
                "url":        url,
                "page_name":  page_name,
                "platform":   "Analyzing…",
                "confidence": "—",
                "reason":     "Reading page content…",
                "tier":       1,
            }
            results.append(result)
            by_index[i] = result
            q.put({"type": "analysis_placeholder", **result})

        if not ANTHROPIC_API_KEY:
            emit_terminal(q, job_id, {"type": "fatal",
                                      "reason": "No ANTHROPIC_API_KEY set — add it to .env to enable AI analysis."})
            return

        # ── AI analysis (async, concurrent) ───────────────────────────────────
        q.put({"type": "analysis_started", "total": total})
        asyncio.run(_run_ai_analysis(job_id, results, by_index, q))

        if should_stop(job_id):
            emit_terminal(q, job_id, {"type": "cancelled", "reason": "Analysis cancelled."})
            return

        jobs[job_id]["analysis_results"] = results
        emit_terminal(q, job_id, {"type": "analysis_complete", "total": total})
    except AIQuotaError as e:
        emit_terminal(q, job_id, {"type": "fatal", "reason": str(e)})
    except Exception as e:
        emit_terminal(q, job_id, {"type": "fatal", "reason": str(e)})


def export_suggestions(results: list[dict], start_url: str) -> Path:
    """Build an Excel workbook from platform suggestion results."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Platform Suggestions"

    # ── Styles ────────────────────────────────────────────────────────────────
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    center      = Alignment(horizontal="center", vertical="center")
    wrap        = Alignment(wrap_text=True, vertical="top")
    thin        = Side(style="thin", color="E5E7EB")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    PLATFORM_COLORS = {
        "Website":             "DBEAFE",   # blue-100
        "Maverick OneStop":    "FEF3C7",   # amber-100
        "The Fountain":        "EDE9FE",   # violet-100
        "MavLife / Student Hub": "D1FAE5", # green-100
        "Teams / SharePoint":  "CCFBF1",   # teal-100
        "No Clear Fit":        "F3F4F6",   # gray-100
    }
    CONFIDENCE_COLORS = {"High": "16A34A", "Medium": "D97706", "Low": "DC2626"}

    headers = ["#", "Page Name", "Full URL", "Suggested Platform", "Confidence", "Reason", "Reviewer Override"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[1].height = 22

    for r in results:
        row_num = r["index"] + 1
        row_fill = PatternFill("solid", fgColor=PLATFORM_COLORS.get(r["platform"], "F3F4F6"))
        conf_font = Font(color=CONFIDENCE_COLORS.get(r["confidence"], "374151"), bold=True)
        values = [r["index"], r["page_name"], r["url"], r["platform"], r["confidence"], r["reason"], ""]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.fill = row_fill
            cell.border = border
            cell.alignment = wrap if col in (3, 6) else center
            if col == 5:
                cell.font = conf_font

    widths = [5, 28, 55, 22, 14, 55, 28]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    from collections import Counter
    platform_counts = Counter(r["platform"] for r in results)
    conf_counts     = Counter(r["confidence"] for r in results)
    ws2["A1"], ws2["B1"] = "Starting URL", start_url
    ws2["A2"], ws2["B2"] = "Total Pages Analyzed", len(results)
    ws2["A3"], ws2["B3"] = "Exported At", datetime.now().isoformat(timespec="seconds")
    row = 5
    ws2.cell(row=row, column=1, value="Platform").font = Font(bold=True)
    ws2.cell(row=row, column=2, value="Count").font   = Font(bold=True)
    for platform, count in platform_counts.most_common():
        row += 1
        ws2.cell(row=row, column=1, value=platform)
        ws2.cell(row=row, column=2, value=count)
    row += 2
    ws2.cell(row=row, column=1, value="Confidence").font = Font(bold=True)
    ws2.cell(row=row, column=2, value="Count").font      = Font(bold=True)
    for conf, count in conf_counts.most_common():
        row += 1
        ws2.cell(row=row, column=1, value=conf)
        ws2.cell(row=row, column=2, value=count)
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 15

    out_path = OUTPUT_DIR / f"platform_suggestions_{uuid.uuid4().hex}.xlsx"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


# ── routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/healthz")
def healthz():
    return jsonify({"status": "ok"})


@app.route("/favicon.ico")
def favicon():
    return Response(status=204)


@app.route("/job-status/<job_id>")
def job_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id": job_id,
        "type": job.get("type"),
        "created_at": job.get("created_at"),
        "finished_at": job.get("finished_at"),
        "cancelled": job.get("cancelled", False),
        "status": job.get("status", {}),
    })


@app.route("/cancel/<job_id>", methods=["POST"])
def cancel(job_id: str):
    if not cancel_job(job_id):
        return jsonify({"error": "Job not found or already finished"}), 404
    return jsonify({"status": "cancelling"})


@app.route("/scrape", methods=["POST"])
def scrape():
    body = request.json or {}
    url = body.get("url", "").strip()
    if not url:
        return jsonify({"error": "No URL provided"}), 400
    refresh = bool(body.get("refresh", False))

    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    if not register_job(job_id, {"type": "scrape", "queue": q, "links": None, "zip_path": None}):
        return jsonify({"error": "Too many active jobs. Wait for the current job to finish, then try again."}), 429
    threading.Thread(target=_scrape_worker, args=(job_id, url, q),
                     kwargs={"refresh": refresh}, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/start-download", methods=["POST"])
def start_download():
    links = (request.json or {}).get("links", [])
    if not links:
        return jsonify({"error": "No links provided"}), 400
    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    if not register_job(job_id, {"type": "download", "queue": q, "zip_path": None}):
        return jsonify({"error": "Too many active jobs. Wait for the current job to finish, then try again."}), 429
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
                evt = q.get(timeout=SSE_HEARTBEAT_SECONDS)
                if evt.get("type") == "complete" and job.get("type") == "scrape":
                    job["links"] = evt.get("links", [])
                yield f"data: {json.dumps(evt)}\n\n"
                if evt.get("type") in ("complete", "fatal", "cancelled"):
                    job["finished_at"] = time.time()
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
    return send_file(job["zip_path"], as_attachment=True, download_name=f"pages_{timestamp_slug()}.zip")


@app.route("/start-download-docx", methods=["POST"])
def start_download_docx():
    links = (request.json or {}).get("links", [])
    if not links:
        return jsonify({"error": "No links provided"}), 400
    job_id = str(uuid.uuid4())
    q = queue.Queue()
    if not register_job(job_id, {"type": "download-docx", "queue": q, "docx_zip_path": None}):
        return jsonify({"error": "Too many active jobs. Wait for the current job to finish, then try again."}), 429
    threading.Thread(target=_docx_worker, args=(job_id, links, q), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/get-docx-zip/<job_id>")
def get_docx_zip(job_id: str):
    job = jobs.get(job_id)
    if not job or not job.get("docx_zip_path"):
        return "Not ready", 404
    return send_file(job["docx_zip_path"], as_attachment=True, download_name=f"pages_{timestamp_slug()}.docx.zip")


@app.route("/export-excel", methods=["POST"])

def export_excel_route():
    """Generate and return an Excel file from the current link list + manifest."""
    data       = request.json or {}
    links      = data.get("links", [])
    start_url  = data.get("start_url", "")
    if not links:
        return jsonify({"error": "No links provided"}), 400
    try:
        # Merge manifest status into links in case it's fresher than what the UI has
        manifest = load_manifest()
        for link in links:
            url = link.get("url", "")
            if url in manifest:
                link["done"]       = True
                link["file"]       = manifest[url]["file"]
                link["scraped_at"] = manifest[url]["scraped_at"]
                link["size_kb"]    = manifest[url].get("size_kb", "")
        path = export_excel(links, start_url)
        return send_file(str(path), as_attachment=True, download_name=f"scraped_pages_{timestamp_slug()}.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze-platforms", methods=["POST"])
def analyze_platforms():
    """Start a background platform-suggestion job for the given link list."""
    data  = request.json or {}
    links = data.get("links", [])
    if not links:
        return jsonify({"error": "No links provided"}), 400
    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    if not register_job(job_id, {"type": "analysis", "queue": q, "analysis_results": None}):
        return jsonify({"error": "Too many active jobs. Wait for the current job to finish, then try again."}), 429
    threading.Thread(target=_analysis_worker, args=(job_id, links, q), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/export-suggestions", methods=["POST"])
def export_suggestions_route():
    """Generate and return an Excel file from the platform suggestion results."""
    data      = request.json or {}
    results   = data.get("results", [])
    start_url = data.get("start_url", "")
    if not results:
        return jsonify({"error": "No results provided"}), 400
    try:
        path = export_suggestions(results, start_url)
        return send_file(str(path), as_attachment=True, download_name=f"platform_suggestions_{timestamp_slug()}.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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
    app.run(
        debug=os.getenv("FLASK_DEBUG") == "1",
        host="0.0.0.0",
        port=env_int("PORT", 5000, 1),
        threaded=True,
    )
