"""
test_suggestions.py — unit tests for the platform-suggestion engine.

Covers:
  - classifier.normalize_platform : folds Claude's phrasing variations to canonical names
  - classifier.clamp_platform     : validates a name against KNOWN_PLATFORMS
  - app guide loader              : extracts the guide section + builds cached prompt blocks

Run:  python -m unittest test_suggestions      (or)  python test_suggestions.py
No third-party dependencies — uses the standard-library unittest only.
"""

import unittest

from classifier import KNOWN_PLATFORMS, normalize_platform, clamp_platform


class TestNormalizePlatform(unittest.TestCase):
    def test_aliases_fold_to_canonical(self):
        cases = {
            "onestop":            "Maverick OneStop",
            "one stop":           "Maverick OneStop",
            "maverick one stop":  "Maverick OneStop",
            "msu website":        "Website",
            "website":            "Website",
            "fountain":           "The Fountain",
            "mavlife":            "MavLife / Student Hub",
            "student hub":        "MavLife / Student Hub",
            "sharepoint":         "Teams / SharePoint",
            "no fit":             "No Clear Fit",
        }
        for raw, expected in cases.items():
            self.assertEqual(normalize_platform(raw), expected, msg=raw)

    def test_case_and_whitespace_insensitive(self):
        self.assertEqual(normalize_platform("  OneStop  "), "Maverick OneStop")
        self.assertEqual(normalize_platform("MAVLIFE"), "MavLife / Student Hub")

    def test_canonical_names_pass_through(self):
        for name in KNOWN_PLATFORMS:
            self.assertEqual(normalize_platform(name), name)

    def test_unknown_passes_through_unchanged(self):
        # Unknown names are returned as-is so clamp_platform can reject them.
        self.assertEqual(normalize_platform("Totally Unknown"), "Totally Unknown")


class TestClampPlatform(unittest.TestCase):
    def test_known_platforms_clamp_to_themselves(self):
        for name in KNOWN_PLATFORMS:
            self.assertEqual(clamp_platform(name), name)

    def test_unknown_clamps_to_no_clear_fit(self):
        self.assertEqual(clamp_platform("Totally Unknown"), "No Clear Fit")
        self.assertEqual(clamp_platform(""), "No Clear Fit")

    def test_normalize_then_clamp_pipeline(self):
        # This is exactly how app.py validates Claude's raw output.
        self.assertEqual(clamp_platform(normalize_platform("one stop")), "Maverick OneStop")
        self.assertEqual(clamp_platform(normalize_platform("MSU Website")), "Website")
        self.assertEqual(clamp_platform(normalize_platform("garbage value")), "No Clear Fit")


class TestAiFallbacks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app
        cls.app = app

    def test_parse_ai_json_accepts_wrapped_json(self):
        raw = 'Here is the result:\n```json\n{"platform":"Website","confidence":"Low","reason":"Public info"}\n```'
        parsed = self.app.parse_ai_json(raw)
        self.assertEqual(parsed["platform"], "Website")

    def test_tobacco_policy_fallback_goes_to_onestop(self):
        result = self.app.fallback_platform_suggestion(
            "Report Smoking On Campus",
            "https://mankato.mnsu.edu/university-life/health-and-safety/campus-wellness/tobacco-free-campus/report-smoking-on-campus/",
            "Report a tobacco policy violation or smoking on campus.",
        )
        self.assertEqual(result["platform"], "Maverick OneStop")
        self.assertEqual(result["confidence"], "Medium")

    def test_credit_error_detection(self):
        self.assertTrue(self.app.is_credit_error(Exception("credit balance is too low")))
        self.assertFalse(self.app.is_credit_error(Exception("temporary network error")))


class TestProgressReconnect(unittest.TestCase):
    """A dropped SSE stream must not leave the UI stuck on "reconnecting"."""

    @classmethod
    def setUpClass(cls):
        import app
        cls.app = app

    def setUp(self):
        import queue
        self.queue = queue
        self.client = self.app.app.test_client()
        self._original_heartbeat = self.app.SSE_HEARTBEAT_SECONDS
        self.app.SSE_HEARTBEAT_SECONDS = 5

    def tearDown(self):
        self.app.SSE_HEARTBEAT_SECONDS = self._original_heartbeat

    def _first_event(self, job_id):
        import json
        response = self.client.get(f"/progress/{job_id}")
        buffered = ""
        for chunk in response.response:
            buffered += chunk.decode("utf-8", "ignore")
            if "\n\n" in buffered:
                raw = buffered.split("\n\n", 1)[0].strip()
                if raw.startswith("data: "):
                    return json.loads(raw[6:])
        return None

    def _finished_job(self, job_id, job, terminal):
        q = self.queue.Queue()
        job["queue"] = q
        self.assertTrue(self.app.register_job(job_id, job))
        self.app.emit_terminal(q, job_id, terminal)
        while not q.empty():      # the dropped connection consumed the event
            q.get()
        return q

    def test_reconnect_replays_completion_instead_of_heartbeats(self):
        job_id = "test-reconnect-download"
        self._finished_job(job_id, {"type": "download", "zip_path": "x.zip"}, {"type": "complete"})
        event = self._first_event(job_id)
        self.assertIsNotNone(event)
        self.assertEqual(event["type"], "complete")

    def test_reconnect_replays_cancellation(self):
        job_id = "test-reconnect-cancelled"
        self._finished_job(job_id, {"type": "download", "zip_path": None},
                           {"type": "cancelled", "reason": "Download cancelled."})
        self.assertEqual(self._first_event(job_id)["type"], "cancelled")

    def test_reconnect_replay_is_repeatable(self):
        job_id = "test-reconnect-twice"
        self._finished_job(job_id, {"type": "download", "zip_path": "x.zip"}, {"type": "complete"})
        self.assertEqual(self._first_event(job_id)["type"], "complete")
        self.assertEqual(self._first_event(job_id)["type"], "complete")

    def test_reconnect_restores_scrape_links(self):
        job_id = "test-reconnect-scrape"
        links = [{"url": "https://example.edu/a/", "done": False}]
        self._finished_job(job_id, {"type": "scrape", "links": None, "zip_path": None},
                           {"type": "complete", "links": links, "visited": 1, "skipped": 0})
        event = self._first_event(job_id)
        self.assertEqual(event["type"], "complete")
        self.assertEqual(event["links"], links)
        self.assertEqual(self.app.jobs[job_id]["links"], links)

    def test_job_status_reports_zip_readiness(self):
        ready_id, pending_id = "test-zip-ready", "test-zip-pending"
        self._finished_job(ready_id, {"type": "download", "zip_path": "done.zip"}, {"type": "complete"})
        self._finished_job(pending_id, {"type": "download", "zip_path": None},
                           {"type": "cancelled", "reason": "Download cancelled."})
        self.assertTrue(self.client.get(f"/job-status/{ready_id}").get_json()["zip_ready"])
        self.assertFalse(self.client.get(f"/job-status/{pending_id}").get_json()["zip_ready"])


class TestDocxHyperlinks(unittest.TestCase):
    NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    @classmethod
    def setUpClass(cls):
        import app
        cls.app = app

    def _build(self, structured, title="Test Page"):
        import tempfile, docx
        from pathlib import Path
        tmp = Path(tempfile.mkdtemp()) / "out.docx"
        self.app._build_docx_from_content(title, structured, tmp)
        return docx.Document(str(tmp))

    def _hyperlink_texts(self, doc):
        out = []
        for para in doc.paragraphs:
            for h in para._p.findall(self.NS + "hyperlink"):
                out.append("".join(t.text or "" for t in h.iter(self.NS + "t")).strip())
        return out

    def test_only_anchor_text_becomes_hyperlink(self):
        structured = [{"tag": "p", "runs": [
            {"text": "Questions? Contact", "href": ""},
            {"text": "Residential Life", "href": "https://example.edu/reslife/"},
            {"text": "for details.", "href": ""},
        ]}]
        doc = self._build(structured)
        self.assertEqual(self._hyperlink_texts(doc), ["Residential Life"])
        full = " ".join(p.text for p in doc.paragraphs)
        self.assertIn("Questions? Contact", full)
        self.assertIn("for details.", full)

    def test_multiple_links_in_one_block_all_preserved(self):
        structured = [{"tag": "p", "runs": [
            {"text": "Home", "href": "https://example.edu/"},
            {"text": "University Life", "href": "https://example.edu/university-life/"},
            {"text": "Housing", "href": "https://example.edu/housing/"},
        ]}]
        doc = self._build(structured)
        self.assertEqual(
            self._hyperlink_texts(doc),
            ["Home", "University Life", "Housing"],
        )

    def test_javascript_href_is_not_linked(self):
        structured = [{"tag": "p", "runs": [
            {"text": "Open menu", "href": "javascript:void(0)"},
        ]}]
        doc = self._build(structured)
        self.assertEqual(self._hyperlink_texts(doc), [])
        self.assertIn("Open menu", " ".join(p.text for p in doc.paragraphs))

    def test_hyperlink_relationships_are_external(self):
        structured = [{"tag": "li", "runs": [
            {"text": "Email us", "href": "mailto:reslife@mnsu.edu"},
        ]}]
        doc = self._build(structured)
        rels = [r for r in doc.part.rels.values() if "hyperlink" in r.reltype]
        self.assertTrue(rels)
        self.assertTrue(all(r.is_external for r in rels))
        self.assertIn("mailto:reslife@mnsu.edu", [r.target_ref for r in rels])

    def test_title_is_first_heading(self):
        doc = self._build([{"tag": "p", "runs": [{"text": "Body text here", "href": ""}]}], title="My Title")
        self.assertEqual(doc.paragraphs[0].text, "My Title")


class TestUrlToFilepath(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app
        cls.app = app

    def test_long_mnsu_url_stays_under_safe_limit(self):
        url = "https://mankato.mnsu.edu/university-life/activities-and-organizations/Homecoming/student-competitions/schedule-of-events/homecoming-kick-off---live-music-games-in-the-mall-themed-raffle-bag-giveaway/"
        path = self.app.url_to_filepath(url)
        self.assertLessEqual(self.app.path_text_length(path), self.app.SAFE_REL_PATH_MAX)
        self.assertEqual(path.suffix, ".pdf")

    def test_different_long_urls_do_not_collide_after_truncation(self):
        base = "https://mankato.mnsu.edu/university-life/activities-and-organizations/Homecoming/student-competitions/schedule-of-events/"
        first = self.app.url_to_filepath(base + "homecoming-kick-off---live-music-games-in-the-mall-themed-raffle-bag-giveaway/")
        second = self.app.url_to_filepath(base + "homecoming-kick-off---live-music-games-in-the-mall-themed-raffle-bag-giveaway-extra/")
        self.assertNotEqual(first, second)

    def test_root_url_uses_hashed_index_pdf(self):
        path = self.app.url_to_filepath("https://mankato.mnsu.edu/")
        self.assertRegex(path.name, r"^index-[0-9a-f]{10}\.pdf$")


class TestGuideLoader(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app  # imported lazily so a failure here doesn't hide validator tests
        cls.app = app

    def test_guide_section_extracted(self):
        guide = self.app.PLATFORM_GUIDE
        self.assertTrue(guide, "guide section should not be empty")
        self.assertTrue(guide.startswith(self.app.GUIDE_SECTION_START))

    def test_guide_contains_all_platforms(self):
        guide = self.app.PLATFORM_GUIDE
        for marker in ("Website", "The Fountain", "Maverick OneStop", "MavLife"):
            self.assertIn(marker, guide, msg=marker)

    def test_guide_excludes_audit_section(self):
        # The loader must stop before the audit section (the end marker).
        self.assertNotIn(self.app.GUIDE_SECTION_END, self.app.PLATFORM_GUIDE)

    def test_guide_version_is_stable_hash(self):
        v1 = self.app.guide_version()
        v2 = self.app.guide_version()
        self.assertEqual(v1, v2)
        self.assertEqual(len(v1), 8)
        int(v1, 16)  # raises if not valid hex

    def test_system_blocks_use_prompt_caching(self):
        blocks = self.app.TIER2_SYSTEM_BLOCKS
        self.assertEqual(len(blocks), 2)
        # The large, stable guide block must carry the ephemeral cache marker.
        self.assertEqual(blocks[-1].get("cache_control"), {"type": "ephemeral"})
        self.assertEqual(blocks[-1]["text"], self.app.PLATFORM_GUIDE)


if __name__ == "__main__":
    unittest.main(verbosity=2)
