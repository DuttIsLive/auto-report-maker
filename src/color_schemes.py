"""
Color scheme definitions for the report generator.

Each topic is hashed to a deterministic palette so that different topics
always get different visual identities, while the same topic always
reproduces the same look.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass


@dataclass(frozen=True)
class ColorScheme:
    """RGB hex strings (without '#') for all roles in the design system."""

    name: str
    primary: str       # main brand colour (headings, cover background band)
    secondary: str     # accent on highlight boxes, table headers
    accent: str        # callout borders, decorative lines
    neutral_light: str # alternating table rows, highlight box background
    neutral_dark: str  # body text, footer text
    cover_text: str    # text placed on the coloured cover band


# ---------------------------------------------------------------------------
# Predefined palettes
# ---------------------------------------------------------------------------

SCHEMES: list[ColorScheme] = [
    ColorScheme(
        name="Academic Blue",
        primary="1B3A6B",
        secondary="2E75B6",
        accent="F0C040",
        neutral_light="EBF3FB",
        neutral_dark="1F1F1F",
        cover_text="FFFFFF",
    ),
    ColorScheme(
        name="Corporate Forest",
        primary="1E4D2B",
        secondary="2E8B57",
        accent="D4A017",
        neutral_light="EAF5ED",
        neutral_dark="1A1A1A",
        cover_text="FFFFFF",
    ),
    ColorScheme(
        name="Executive Burgundy",
        primary="6B1A1A",
        secondary="C0392B",
        accent="D4AC0D",
        neutral_light="FDECEA",
        neutral_dark="1C1C1C",
        cover_text="FFFFFF",
    ),
    ColorScheme(
        name="Modern Slate",
        primary="2C3E50",
        secondary="5D6D7E",
        accent="1ABC9C",
        neutral_light="EAF0F6",
        neutral_dark="1A1A2E",
        cover_text="FFFFFF",
    ),
    ColorScheme(
        name="Ocean Teal",
        primary="005F73",
        secondary="0A9396",
        accent="E9C46A",
        neutral_light="E0F4F5",
        neutral_dark="1A1A1A",
        cover_text="FFFFFF",
    ),
    ColorScheme(
        name="Royal Purple",
        primary="4A235A",
        secondary="7D3C98",
        accent="F39C12",
        neutral_light="F5EEF8",
        neutral_dark="1A1A1A",
        cover_text="FFFFFF",
    ),
]


def pick_scheme(topic: str) -> ColorScheme:
    """Return a deterministic color scheme based on the topic string."""
    digest = int(hashlib.md5(topic.lower().encode()).hexdigest(), 16)
    return SCHEMES[digest % len(SCHEMES)]
