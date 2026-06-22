# PLAN V02 — Platform Suggestion Feature + Deployment

> **Status:** DRAFT  
> **Project:** MNSU Web Crawler + PDF Downloader  
> **Goal:** Extend the existing Flask web app with an AI-powered Platform Suggestion engine, then deploy it for broader team access.

---

## Overview

Take the list of scraped university URLs and automatically suggest which MSU content platform each page should migrate to (Website, The Fountain, Maverick OneStop, MavLife, Teams, or No Fit). Results stream in real-time and export to an enhanced Excel report.

---

## Phase 1 — Classification Engine (Local, No API)

**Goal:** Instant, zero-cost platform suggestions using URL path + page title pattern matching.

### 1.1 Define Classification Rules (`classifier.py`)
Create a standalone module with a `classify_url(url, page_name)` function that returns:
```python
{ "platform": "Website", "confidence": "High", "reason": "Path contains /admissions/" }
```

Rules per platform (derived from the Content Management Guide):

| Platform | URL Path Signals | Page Name Keywords |
|---|---|---|
| **Website** | `/admissions/`, `/academics/`, `/housing/`, `/dining/`, `/about/`, `/giving/`, `/alumni/`, `/news/`, `/events/` | prospective, tuition, apply, scholarship, program, degree, campus |
| **Maverick OneStop** | `/services/`, `/forms/`, `/it/`, `/registrar/`, `/policies/` | how-to, FAQ, form, request, process, policy, register, submit, troubleshoot |
| **The Fountain** | `/faculty-staff/`, `/hr/`, `/human-resources/`, `/employee/` | faculty, staff, employee, HR, benefits, committee, internal |
| **MavLife / Student Hub** | `/student-life/`, `/student-organizations/`, `/clubs/`, `/involvement/` | clubs, organizations, activities, involvement, engagement, student life |
| **Teams / SharePoint** | `/departments/`, `/committees/`, `/working-groups/` | meeting notes, collaboration, committee, department resources |
| **No Fit / Archive** | `~/link/`, `.aspx` (non-university), old compliance-only pages | outdated, archived, redirect |

### 1.2 Scoring Logic
- Each URL is scored against all platforms independently (0–100 score)
- Top score wins → becomes the suggestion
- Score gap < 15 between top two → confidence = "Low" → flagged for review
- Score gap ≥ 30 → confidence = "High"

---

## Phase 2 — Tier 2: AI Classification via Anthropic API

**Goal:** For low-confidence URLs (~20–30%), extract page text and send to Claude for a content-aware classification.

### 2.1 Text Extraction
- Re-visit the URL using the already-running Playwright instance
- Extract `page.inner_text("body")` — clean, no PDF parsing needed
- Truncate to first 600 words to keep token cost low

### 2.2 Claude API Call
- Model: `claude-3-5-haiku` (fast + cheap, ~$0.001 per page)
- System prompt: inject the full Content Management Guide as context
- User prompt: `"Page title: {name}\nURL: {url}\nContent excerpt:\n{text}\n\nWhich platform?"`
- Parse structured response: `{ platform, confidence, reason }`
- API key stored in `.env` file, never hardcoded

### 2.3 Cost Estimate
- 600 URLs × Tier 1 only = $0.00
- 600 URLs × 30% low-confidence × Tier 2 = ~180 API calls ≈ $0.18–$0.50 total

---

## Phase 3 — Web App Integration

**Goal:** Surface the classification engine inside the existing Flask UI as a new step after the current workflow.

### 3.1 New Backend Routes (`app.py`)
- `POST /analyze-platforms` — start analysis job, returns `job_id`
- `GET /stream-analysis/<job_id>` — SSE stream for real-time progress (one event per URL)
- `GET /export-suggestions` — download enhanced Excel with suggestion columns

### 3.2 New UI Section (`index.html`)
- Add a **Step 4: Platform Suggestions** panel below the existing workflow
- "🔍 Analyze Platforms" button — triggers analysis of current scraped link list
- Live progress table: URL | Page Name | Suggested Platform | Confidence | Reason
- Rows stream in one by one as they complete (SSE, same pattern as PDF generation)
- Color-coded rows: green (High), yellow (Medium), red (Low/Needs Review)
- "📊 Export Suggestions" button → downloads Excel

### 3.3 Enhanced Excel Output
Add 4 new columns to the existing Excel export:
| Suggested Platform | Confidence | Reason | Reviewer Override |
|---|---|---|---|
| Website | High | Path: /housing/ | _(blank for human input)_ |
| Maverick OneStop | Low | Keyword: "how to" | _(team fills this in)_ |

---

## Phase 4 — Deployment

**Goal:** Host the app so the content team can use it without needing Python installed.

### 4.1 Pre-Deployment Changes Required

#### A. Environment Variables
Move all secrets and config to a `.env` file:
- `ANTHROPIC_API_KEY`
- `SECRET_KEY` (Flask session security)
- `OUTPUT_DIR` (default: `mankato_pdfs/`)
- `MAX_PAGES` (crawl limit)

#### B. Persistent File Storage
Local file system won't persist on cloud platforms. Two options:
- **Option A (Simple):** Supabase Storage — upload PDFs and manifest to a bucket; use existing Supabase project
- **Option B (Simpler for now):** Disable PDF download in deployed version; suggestion engine only reads URLs, no local files needed

#### C. Concurrency
Current app uses Python threads + in-memory `jobs` dict — works for 1 user. For multi-user:
- Add a `MAX_CONCURRENT_JOBS = 1` guard (simplest fix)
- Future: Redis + Celery if scaling is needed

#### D. Playwright in Production
Playwright needs Chromium (~300MB). Requires Docker for reliable deployment.

### 4.2 Containerization (Docker)
Create `Dockerfile`:
1. Base image: `mcr.microsoft.com/playwright/python:v1.44.0-jammy`
2. Copy app files
3. Install Python dependencies (`requirements.txt`)
4. Install Playwright browsers (`playwright install chromium`)
5. Expose port 8080
6. Run with `gunicorn` (not Flask dev server)

### 4.3 Deployment Platform Options

| Platform | Cost | Playwright Support | Difficulty | Recommended For |
|---|---|---|---|---|
| **Railway** | Free tier / ~$5/mo | ✅ via Docker | Low | Best first choice |
| **Render** | Free tier (sleeps) / $7/mo | ✅ via Docker | Low | Good alternative |
| **Fly.io** | Free tier | ✅ via Docker | Medium | More control |
| **AWS App Runner** | ~$10/mo | ✅ via Docker | High | Enterprise later |

**Recommendation: Railway** — Docker-native, no sleep on free hobby plan, easy env variable management, GitHub integration for auto-deploy on push.

### 4.4 Access Control (Simple Auth)
Since this is an internal tool for the content team:
- Add HTTP Basic Auth via Flask (username + password in env vars)
- Or use a shared secret URL token (e.g., `/?token=abc123`)
- Full Supabase Auth if you want per-user login later

---

## Phase 5 — CI/CD and Maintenance

- Connect Railway to the GitHub repo (auto-deploy on push to `main`)
- Keep `excel-feature` branch for ongoing local testing
- Merge to `main` only when a feature is stable
- Add `requirements.txt` (currently missing — needed for deployment)

---

## Implementation Order

| # | Task | Branch | Effort |
|---|---|---|---|
| 1 | Create `classifier.py` with rule engine | `platform-suggestion` | 2–3 hrs |
| 2 | Add `/analyze-platforms` + SSE route | `platform-suggestion` | 2 hrs |
| 3 | Add Step 4 UI panel + streaming table | `platform-suggestion` | 2–3 hrs |
| 4 | Integrate Anthropic API for Tier 2 | `platform-suggestion` | 2 hrs |
| 5 | Add `.env` support + `requirements.txt` | `platform-suggestion` | 30 min |
| 6 | Write `Dockerfile` | `deploy` | 1 hr |
| 7 | Deploy to Railway + test | `deploy` | 1–2 hrs |
| 8 | Merge to `main`, set up auto-deploy | `main` | 30 min |

**Total estimated effort: ~12–14 hours across all phases**

---

## Decisions Locked In

| # | Question | Decision |
|---|---|---|
| 1 | Anthropic API key | ✅ Confirmed active — `claude-3-5-haiku` for Tier 2 |
| 2 | PDF storage on deployment | 🖥️ **Local only** — PDF generation stays in local version; deployed app focuses on crawling + suggestions + Excel export only |
| 3 | Who accesses deployed app | 👥 Content team — **Clerk auth added later** as a separate phase; deploy without auth first |
| 4 | Domain | 🚂 `.railway.app` subdomain is fine |

### What This Means for Architecture

**Local app** (`python app.py`) — Full feature set:
- Crawl → Review → Download PDFs → Export Excel → Platform Suggestions

**Deployed app** (Railway) — Suggestion-focused:
- Crawl → Review → Platform Suggestions → Export Excel
- PDF generation **disabled** in deployed build (Railway file system is ephemeral; files don't persist between restarts)
- A single `DEPLOY_MODE=true` environment variable controls this — hides the Download step in the UI

This keeps the deployed version fast, lightweight, and purposeful for the content team.
