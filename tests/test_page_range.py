"""Unit tests for the `pptrans.page_range.parse_page_range` function."""

import click
import pytest

from pptrans.page_range import parse_page_range


@pytest.mark.kwparametrize(
    dict(range_str="3", total_slides=5, expected_pages_0_indexed={2}),
    dict(range_str="1,5", total_slides=5, expected_pages_0_indexed={0, 4}),
    dict(range_str="2-4", total_slides=5, expected_pages_0_indexed={1, 2, 3}),
    dict(range_str="3-", total_slides=5, expected_pages_0_indexed={2, 3, 4}),
    dict(range_str="-3", total_slides=5, expected_pages_0_indexed={0, 1, 2}),
    dict(range_str="-", total_slides=5, expected_pages_0_indexed={0, 1, 2, 3, 4}),
    dict(range_str="1-5", total_slides=5, expected_pages_0_indexed={0, 1, 2, 3, 4}),
    dict(
        range_str="1,3-4,5-",
        total_slides=7,
        expected_pages_0_indexed={0, 2, 3, 4, 5, 6},
    ),
    dict(
        range_str=" 1 , 3 - 4 , 5 - ",
        total_slides=7,
        expected_pages_0_indexed={0, 2, 3, 4, 5, 6},
    ),
    dict(range_str="1,,3", total_slides=5, expected_pages_0_indexed={0, 2}),
    dict(range_str=",1,3,", total_slides=5, expected_pages_0_indexed={0, 2}),
    dict(range_str="1,", total_slides=5, expected_pages_0_indexed={0}),
    dict(range_str=",3", total_slides=5, expected_pages_0_indexed={2}),
    dict(range_str="", total_slides=0, expected_pages_0_indexed=set()),
    dict(range_str="7-8", total_slides=5, expected_pages_0_indexed=set()),
    dict(range_str="1,7-8,3", total_slides=5, expected_pages_0_indexed={0, 2}),
    dict(range_str="3-7", total_slides=5, expected_pages_0_indexed={2, 3, 4}),
    dict(
        range_str="1,3-7,2",
        total_slides=5,
        expected_pages_0_indexed={0, 1, 2, 3, 4},
    ),
    dict(range_str="1", total_slides=1, expected_pages_0_indexed={0}),
    dict(range_str="1-1", total_slides=1, expected_pages_0_indexed={0}),
    dict(range_str="-", total_slides=1, expected_pages_0_indexed={0}),
    dict(range_str="-7", total_slides=5, expected_pages_0_indexed={0, 1, 2, 3, 4}),
    ids=[
        "basic_single_page",
        "basic_two_pages",
        "basic_simple_range",
        "open_range_to_end",
        "open_range_from_start",
        "open_range_full_dash",
        "closed_range_full",
        "mixed_complex_range",
        "mixed_with_spaces",
        "empty_part_middle_comma",
        "leading_trailing_commas",
        "trailing_comma_only",
        "leading_comma_only",
        "empty_range_str_no_slides",
        "warn_start_gt_total_no_warn_test",  # Warning tested elsewhere
        "warn_part_start_gt_total_no_warn_test",
        "warn_end_gt_total_capped_no_warn_test",
        "warn_part_end_gt_total_capped_no_warn_test",
        "single_page_total_one",
        "range_one_total_one",
        "open_range_total_one",
        "open_start_end_gt_total_capped_no_warn_test",
    ],
)
def test_parse_page_range_valid_inputs(
    range_str: str, total_slides: int, expected_pages_0_indexed: set[int]
) -> None:
    """Test parse_page_range with valid inputs."""
    assert parse_page_range(range_str, total_slides) == expected_pages_0_indexed


@pytest.mark.kwparametrize(
    dict(
        range_str="1",
        total_slides=0,
        expected_message_part=(
            "Cannot specify page range '1' for a presentation with no slides."
        ),
    ),
    dict(
        range_str="1-2",
        total_slides=0,
        expected_message_part=(
            "Cannot specify page range '1-2' for a presentation with no slides."
        ),
    ),
    dict(
        range_str="",
        total_slides=5,
        expected_message_part=(
            "Page range string cannot be empty if --pages option is used"
        ),
    ),
    dict(
        range_str="   ",
        total_slides=5,
        expected_message_part=(
            "Page range string cannot be empty if --pages option is used"
        ),
    ),
    dict(
        range_str="a-5",
        total_slides=10,
        expected_message_part="Invalid start page number 'a' in 'a-5'.",
    ),
    dict(
        range_str="1-b",
        total_slides=10,
        expected_message_part="Invalid end page number 'b' in '1-b'.",
    ),
    dict(
        range_str="a",
        total_slides=10,
        expected_message_part="Invalid page number: 'a'. Must be an integer.",
    ),
    dict(
        range_str="1-2-3",
        total_slides=10,
        expected_message_part="Invalid end page number '2-3' in '1-2-3'.",
    ),
    dict(
        range_str="0-5",
        total_slides=10,
        expected_message_part=(
            "Page numbers must be positive. Found start '0' in '0-5'."
        ),
    ),
    dict(
        range_str="5-0",
        total_slides=10,
        expected_message_part="Page numbers must be positive. Found end '0' in '5-0'.",
    ),
    dict(
        range_str="-0",
        total_slides=10,
        expected_message_part="Page numbers must be positive. Found end '0' in '-0'.",
    ),
    dict(
        range_str="5-3",
        total_slides=10,
        expected_message_part=(
            "Start page 5 cannot be greater than end page 3 in range '5-3'."
        ),
    ),
    dict(
        range_str="7-",  # Becomes 7-5
        total_slides=5,
        expected_message_part=(
            "Start page 7 cannot be greater than end page 5 in range '7-'."
        ),
    ),
    dict(
        range_str="10-",  # Becomes 10-5
        total_slides=5,
        expected_message_part=(
            "Start page 10 cannot be greater than end page 5 in range '10-'."
        ),
    ),
    dict(
        range_str="0",
        total_slides=10,
        expected_message_part="Page number 0 in '0' is out of valid range [1, 10].",
    ),
    dict(
        range_str="11",
        total_slides=10,
        expected_message_part="Page number 11 in '11' is out of valid range [1, 10].",
    ),
    ids=[
        "err_range_for_empty_deck_single",
        "err_range_for_empty_deck_range",
        "err_empty_string_range",
        "err_whitespace_string_range",
        "err_non_int_start",
        "err_non_int_end",
        "err_non_int_single",
        "err_too_many_hyphens",
        "err_zero_start_page",
        "err_zero_end_page",
        "err_zero_end_page_open_start",
        "err_start_gt_end",
        "err_start_gt_total_open_end",
        "err_start_gt_total_high_open_end",
        "err_single_page_zero",
        "err_single_page_gt_total",
    ],
)
def test_parse_page_range_invalid_inputs(
    range_str: str, total_slides: int, expected_message_part: str
) -> None:
    """Test parse_page_range with invalid inputs raising BadParameter."""
    with pytest.raises(click.BadParameter) as excinfo:
        parse_page_range(range_str, total_slides)
    assert expected_message_part in str(excinfo.value)


@pytest.mark.kwparametrize(
    dict(
        range_str="7-8",
        total_slides=5,
        expected_pages=set(),
        expected_warning_part=(
            "Warning: Start page 7 in range '7-8' is beyond the total of 5 slides."
        ),
    ),
    dict(
        range_str="3-7",
        total_slides=5,
        expected_pages={2, 3, 4},
        expected_warning_part=(
            "Warning: End page 7 in range '3-7' is beyond the total of 5 slides."
        ),
    ),
    dict(
        range_str="-7",
        total_slides=5,
        expected_pages={0, 1, 2, 3, 4},
        expected_warning_part=(
            "Warning: End page 7 in range '-7' is beyond the total of 5 slides."
        ),
    ),
    ids=[
        "warn_start_gt_total",
        "warn_end_gt_total_explicit",
        "warn_end_gt_total_implicit_start",
    ],
)
def test_parse_page_range_warnings(
    range_str: str,
    total_slides: int,
    expected_pages: set[int],
    expected_warning_part: str,
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Test parse_page_range for cases that produce warnings on stderr."""
    assert parse_page_range(range_str, total_slides) == expected_pages
    captured = capsys.readouterr()
    assert expected_warning_part in captured.err
