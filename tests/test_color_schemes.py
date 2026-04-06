"""
tests/test_color_schemes.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~
Unit tests for the color scheme selection module.
"""

import pytest

from src.color_schemes import pick_scheme, SCHEMES, ColorScheme


def test_pick_scheme_returns_color_scheme():
    scheme = pick_scheme("Machine Learning")
    assert isinstance(scheme, ColorScheme)


def test_pick_scheme_deterministic():
    """Same topic always produces the same scheme."""
    s1 = pick_scheme("Climate Change")
    s2 = pick_scheme("Climate Change")
    assert s1.name == s2.name


def test_pick_scheme_case_insensitive():
    """Topic comparison is case-insensitive."""
    s1 = pick_scheme("Artificial Intelligence")
    s2 = pick_scheme("artificial intelligence")
    assert s1.name == s2.name


def test_pick_scheme_different_topics_may_differ():
    """Different topics are not required to share a scheme, though collisions
    are possible by design — just verify the function returns valid schemes."""
    schemes = {pick_scheme(t).name for t in [
        "Machine Learning", "Climate Change", "Supply Chain",
        "Healthcare", "Finance", "Cybersecurity",
    ]}
    # At least two distinct schemes should appear across 6 varied topics
    assert len(schemes) >= 2


def test_schemes_list_not_empty():
    assert len(SCHEMES) >= 4


def test_color_scheme_hex_format():
    """All colour fields must be 6-character hex strings (no '#' prefix)."""
    for scheme in SCHEMES:
        for attr in ("primary", "secondary", "accent", "neutral_light",
                     "neutral_dark", "cover_text"):
            value = getattr(scheme, attr)
            assert len(value) == 6, f"{scheme.name}.{attr}={value!r} is not 6 chars"
            int(value, 16)  # raises ValueError if not valid hex
