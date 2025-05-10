"""Tests for the `pptrans` package initializer."""

import pytest

# Since `src/pptrans/__init__.py` only contains a docstring,
# a simple import test is sufficient to "cover" it.
# No executable code paths to test.


def test_import_pptrans_init() -> None:
    """Test that the pptrans package initializer can be imported."""
    try:
        import pptrans  # noqa: F401
    except ImportError as e:
        pytest.fail(f"Failed to import pptrans package: {e}")
