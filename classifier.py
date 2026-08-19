"""
classifier.py — MNSU content-platform registry and response validators.

Defines the canonical set of content platforms, the vocabularies from the AI
Content Analysis & Platform Recommendation Framework, and the helpers used to
normalize and validate the AI classifier's responses.  All classification
decisions are made by the AI tier in app.py using the official Content
Management Guide and the analysis framework as the sources of truth — this
module no longer holds any URL-based scoring rules.
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


# ── Analysis framework vocabularies ───────────────────────────────────────────
# Value lists from AI_Content_Analysis_Direction_Support.md.  The AI is asked to
# answer using exactly these strings; clamp_analysis_fields() enforces it so the
# UI and Excel export never see an unexpected value.

CONFIDENCE_LEVELS: list[str] = ["High", "Medium", "Low"]

PRIMARY_USER_NEEDS: list[str] = ["Explore", "Understand", "Act", "Connect"]
SECONDARY_USER_NEEDS: list[str] = PRIMARY_USER_NEEDS + ["Not applicable"]

AUDIENCES: list[str] = [
    "Prospective students",
    "Current students",
    "Families/supporters",
    "Faculty",
    "Staff",
    "Alumni",
    "Community/public",
    "Multiple audiences",
    "Unable to determine",
]
SECONDARY_AUDIENCES: list[str] = AUDIENCES + ["Not applicable"]

CONTENT_PURPOSES: list[str] = [
    "Recruitment",
    "Awareness",
    "General information",
    "Service information",
    "Process guidance",
    "Transaction support",
    "Policy or requirement",
    "Academic information",
    "Event or engagement",
    "News or announcement",
    "Contact/support",
    "Resource collection",
    "Reference information",
    "Other",
]

# Shared by all six health dimensions and by ai_search_readiness.
HEALTH_RATINGS: list[str] = [
    "Strong",
    "Acceptable",
    "Needs Improvement",
    "Significant Concern",
    "Unable to Determine",
]

DUPLICATION_STATUSES: list[str] = [
    "No significant duplication found",
    "Possible overlap",
    "Significant overlap",
    "Likely duplicate",
    "Conflicting content detected",
]

RECOMMENDED_TREATMENTS: list[str] = [
    "KEEP",
    "KEEP + IMPROVE",
    "CONSOLIDATE",
    "CONNECT",
    "CONSIDER MOVING",
    "REPLACE WITH ACTION",
    "ARCHIVE / REMOVE REVIEW",
    "EXPERT REVIEW NEEDED",
]

IMPROVEMENT_OPPORTUNITIES: list[str] = [
    "Update outdated information",
    "Verify accuracy",
    "Clarify primary audience",
    "Clarify page purpose",
    "Rewrite for clarity",
    "Reduce jargon",
    "Improve headings",
    "Improve page title",
    "Improve scanability",
    "Add clear next step",
    "Improve call to action",
    "Clarify eligibility",
    "Clarify deadlines",
    "Improve contact/support information",
    "Consolidate duplicate content",
    "Establish authoritative source",
    "Remove redundant information",
    "Link rather than duplicate",
    "Improve platform fit",
    "Improve AI/search readability",
    "Separate multiple audiences",
    "Improve user journey",
    "Review for removal",
    "No significant improvement identified",
]

PRIORITIES: list[str] = [
    "Priority 1 - Act Now",
    "Priority 2 - Important Improvement",
    "Priority 3 - Strategic Opportunity",
    "Priority 4 - Maintain/Monitor",
]

HEALTH_DIMENSIONS: list[str] = [
    "health_accuracy_currency",
    "health_clarity",
    "health_scanability_structure",
    "health_actionability",
    "health_user_focus",
    "health_sustainability",
]

# Single-choice fields → their allowed values.
FIELD_ALLOWED_VALUES: dict[str, list[str]] = {
    "confidence": CONFIDENCE_LEVELS,
    "primary_user_need": PRIMARY_USER_NEEDS,
    "secondary_user_need": SECONDARY_USER_NEEDS,
    "primary_audience": AUDIENCES,
    "secondary_audience": SECONDARY_AUDIENCES,
    "content_purpose": CONTENT_PURPOSES,
    "duplication_status": DUPLICATION_STATUSES,
    "ai_search_readiness": HEALTH_RATINGS,
    "recommended_treatment": RECOMMENDED_TREATMENTS,
    "priority": PRIORITIES,
    **{dim: HEALTH_RATINGS for dim in HEALTH_DIMENSIONS},
}

# Where a field falls back to when the AI returns something unrecognized.
FIELD_DEFAULT_VALUE: dict[str, str] = {
    "confidence": "Low",
    "primary_user_need": "Understand",
    "secondary_user_need": "Not applicable",
    "primary_audience": "Unable to determine",
    "secondary_audience": "Not applicable",
    "content_purpose": "Other",
    "duplication_status": "No significant duplication found",
    "ai_search_readiness": "Unable to Determine",
    "recommended_treatment": "EXPERT REVIEW NEEDED",
    "priority": "Priority 4 - Maintain/Monitor",
    **{dim: "Unable to Determine" for dim in HEALTH_DIMENSIONS},
}

# Multi-choice fields → (allowed values or None for free text, max items).
FIELD_MULTI_LIMITS: dict[str, tuple[list[str] | None, int]] = {
    "improvement_opportunities": (IMPROVEMENT_OPPORTUNITIES, 5),
    "related_urls": (None, 5),
    "content_owner_questions": (None, 3),
}

# Free-text fields carried through verbatim (trimmed).  Empty string is valid.
FREE_TEXT_FIELDS: list[str] = [
    "reason",
    "page_summary",
    "purpose_explanation",
    "health_notes",
    "confidence_gap",
    "major_issues",
    "key_recommendation",
]


def clamp_choice(field: str, value) -> str:
    """
    Validate a single-choice field against its allowed values.

    Matching is case-insensitive so the AI's capitalization drift doesn't
    silently degrade a good answer into the default.
    """
    allowed = FIELD_ALLOWED_VALUES.get(field, [])
    default = FIELD_DEFAULT_VALUE.get(field, "Unable to Determine")
    if not isinstance(value, str):
        return default
    cleaned = value.strip()
    for option in allowed:
        if cleaned.lower() == option.lower():
            return option
    return default


def clamp_multi_choice(field: str, values) -> list[str]:
    """Validate a list field: drop invalid entries, dedupe, cap length."""
    allowed, max_items = FIELD_MULTI_LIMITS.get(field, (None, 5))
    if not isinstance(values, list):
        return []
    lookup = {option.lower(): option for option in allowed} if allowed else None
    result: list[str] = []
    for raw in values:
        if not isinstance(raw, str):
            continue
        cleaned = raw.strip()
        if not cleaned:
            continue
        if lookup is not None:
            canonical = lookup.get(cleaned.lower())
            if canonical is None:
                continue
            cleaned = canonical
        if cleaned not in result:
            result.append(cleaned)
        if len(result) >= max_items:
            break
    return result


def default_analysis_fields() -> dict:
    """
    Every analysis field at its safe default.

    Used as the base for both the fallback suggestion and error rows so a result
    dict always has the same shape no matter how it was produced.
    """
    fields: dict = {field: FIELD_DEFAULT_VALUE[field] for field in FIELD_ALLOWED_VALUES}
    fields["platform"] = "No Clear Fit"
    for field in FREE_TEXT_FIELDS:
        fields[field] = ""
    for field in FIELD_MULTI_LIMITS:
        fields[field] = []
    return fields


def clamp_analysis_fields(ai: dict) -> dict:
    """Validate a full AI response into the canonical analysis field set."""
    fields = default_analysis_fields()
    fields["platform"] = clamp_platform(normalize_platform(str(ai.get("platform", ""))))
    for field in FIELD_ALLOWED_VALUES:
        fields[field] = clamp_choice(field, ai.get(field))
    for field in FIELD_MULTI_LIMITS:
        fields[field] = clamp_multi_choice(field, ai.get(field))
    for field in FREE_TEXT_FIELDS:
        value = ai.get(field)
        fields[field] = value.strip() if isinstance(value, str) else ""
    if not fields["reason"]:
        fields["reason"] = "No reason provided."
    return fields
