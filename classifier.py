"""
classifier.py — MNSU content-platform registry and response validators.

Defines the canonical set of content platforms and the helpers used to
normalize and validate platform names returned by the AI classifier.  All
classification decisions are made by the AI tier in app.py using the official
Content Management Guide as the source of truth — this module no longer holds
any URL-based scoring rules.
"""

# ── Canonical platforms ───────────────────────────────────────────────────────
# The single source of truth for valid platform names.  Every AI response is
# clamped to one of these so downstream Excel styling and UI chips always see a
# known value.
KNOWN_PLATFORMS: list[str] = [
    "Website",
    "Maverick OneStop",
    "The Fountain",
    "MavLife / Student Hub",
    "Teams / SharePoint",
    "No Clear Fit",
]

# ── Platform name normalizer ───────────────────────────────────────────────────
# Claude may return slight variations (e.g. "OneStop", "MSU Website").
# Normalize before storing so Excel color-coding and downstream logic always
# see the canonical name from KNOWN_PLATFORMS.
PLATFORM_ALIASES: dict[str, str] = {
    "website":               "Website",
    "msu website":           "Website",
    "maverick onestop":      "Maverick OneStop",
    "onestop":               "Maverick OneStop",
    "one stop":              "Maverick OneStop",
    "maverick one stop":     "Maverick OneStop",
    "the fountain":          "The Fountain",
    "fountain":              "The Fountain",
    "mavlife":               "MavLife / Student Hub",
    "mavlife / student hub": "MavLife / Student Hub",
    "student hub":           "MavLife / Student Hub",
    "mav life":              "MavLife / Student Hub",
    "teams":                 "Teams / SharePoint",
    "sharepoint":            "Teams / SharePoint",
    "teams / sharepoint":    "Teams / SharePoint",
    "no clear fit":          "No Clear Fit",
    "no fit":                "No Clear Fit",
    "unclear":               "No Clear Fit",
}


def normalize_platform(name: str) -> str:
    """Return the canonical platform name, tolerating Claude's phrasing variations."""
    return PLATFORM_ALIASES.get(name.lower().strip(), name)


# ── Public API ────────────────────────────────────────────────────────────────

def clamp_platform(name: str) -> str:
    """
    Validate an AI-returned platform name against KNOWN_PLATFORMS.

    Returns the name unchanged if it is canonical, otherwise "No Clear Fit".
    Call normalize_platform() first to fold phrasing variations into canonical
    names before clamping.
    """
    return name if name in KNOWN_PLATFORMS else "No Clear Fit"
