# Scrape → PDF Planning Doc
Target site: `mankato.mnsu.edu` | Test section: `/university-life/health-and-safety/student-health-services/pharmacy/`

---

## Problem 1 — PDF Naming (too long, gets cut off)

**Current behavior**
The full URL is flattened into one long string:
`01_mankato_mnsu_edu_university-life_health-and-safety_..._online-prescription-refill-form.pdf`

**Fix: Mirror the URL path as folders, use the last segment as the filename**

```
mankato_pdfs/
  university-life/
    health-and-safety/
      student-health-services/
        pharmacy/
          pharmacy.pdf                        ← the /pharmacy/ index page
          pharmacy-faqs.pdf
          pharmacy-hours.pdf
          pharmacy-location-and-parking.pdf
          over-the-counter-medications-and-supplies.pdf
          pharmacy-staff/
            pharmacy-staff.pdf                ← the /pharmacy-staff/ index page
            jessica-peterson.pdf
            melanie-moore.pdf
            souk-phaengkhouane.pdf
          prescriptions/
            prescriptions.pdf
            fill-a-prescription.pdf
            transfer-a-prescription.pdf
            refill-a-prescription/
              refill-a-prescription.pdf
              online-prescription-refill-form.pdf
```

**Rule:** Pages whose URL ends in `/` are named after the last folder segment
(e.g. `/pharmacy/` → `pharmacy.pdf`, `/pharmacy-staff/` → `pharmacy-staff.pdf`)

---

## Problem 2 — Duplicate Tracking (across sessions)

**Current behavior**
The `visited` set only lives in memory. Close the app and restart — it forgets everything. Running the tool twice on overlapping sections re-downloads the same pages.

**Fix: A persistent `manifest.json` on disk**

```json
[
  {
    "url": "https://mankato.mnsu.edu/.../pharmacy/",
    "file": "university-life/.../pharmacy/pharmacy.pdf",
    "scraped_at": "2026-06-09T14:00:00",
    "status": "ok",
    "size_kb": 145.2
  }
]
```

- Before downloading any page → check if its URL is already in the manifest → skip if so
- After a successful download → append an entry to the manifest
- The manifest lives at the root of the output folder: `mankato_pdfs/manifest.json`

---

## Problem 3 — Workflow (crawl and download are tangled together)

**Current behavior**
Crawling and PDF generation happen in one shot. No chance to review before committing.

**Fix: Two explicit phases**

**Phase 1 — Crawl** (fast, no PDFs yet)
1. User pastes a starting URL in the web app
2. App spiders every page under that path (BFS, already implemented)
3. App shows the full discovered URL list
4. URLs already in the manifest are flagged as "already downloaded" — user can see what's new vs. done
5. User reviews the list, then clicks Download

**Phase 2 — Download** (slow, runs in background with live progress)
1. App skips any URL already in the manifest
2. Remaining URLs are downloaded as PDFs, saved in the mirrored folder structure
3. Manifest is updated after each successful save
4. Progress shown live in the UI, ZIP offered at the end

---

## Problem 4 — Scale (the full site is huge)

**Strategy: Run section by section, not from the homepage**

| Run | Starting URL | Why |
|-----|-------------|-----|
| 1 | `/university-life/health-and-safety/student-health-services/pharmacy/` | Current test |
| 2 | `/university-life/housing/residential-life/about-residential-life/employment/` | Already tested |
| 3 | `/academics/` | Next section |
| … | … | … |

The shared `manifest.json` across all runs means no page ever gets downloaded twice, even if two sections link to the same page.

---

## Changes needed in code (when ready)

| What | Where | Change |
|------|-------|--------|
| URL → folder/file path mapping | `app.py` | New `url_to_filepath()` replacing `slugify()` |
| Persistent deduplication | `app.py` | `manifest.json` read on start, written after each PDF |
| Show already-done links in UI | `index.html` | Flag links present in manifest before download |
| Separate crawl/download phases | `index.html` | Already split; just add manifest status to the link list |
| Output folder structure | `app.py` | `os.makedirs()` based on URL path segments |
