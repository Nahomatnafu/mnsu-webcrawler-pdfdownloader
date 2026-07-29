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
