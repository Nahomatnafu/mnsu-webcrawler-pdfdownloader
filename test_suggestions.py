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
