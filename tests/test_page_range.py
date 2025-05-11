"""Unit tests for the `pptrans.page_range.parse_page_range` function."""

import click
import pytest

from pptrans.page_range import parse_page_range


@pytest.mark.kwparametrize(
    dict(range_str="", total_slides=5, expected_pages_0_indexed=set()),
    dict(range_str="   ", total_slides=5, expected_pages_0_indexed=set()),
    dict(range_str="-0", total_slides=10, expected_pages_0_indexed=set()),
    dict(range_str="3", total_slides=5, expected_pages_0_indexed={2}),
    dict(range_str="1,5", total_slides=5, expected_pages_0_indexed={0, 4}),
    dict(range_str="5-3", total_slides=10, expected_pages_0_indexed=set()),
    dict(range_str="7-", total_slides=5, expected_pages_0_indexed=set()),
    dict(range_str="10-", total_slides=5, expected_pages_0_indexed=set()),
    dict(range_str="0", total_slides=10, expected_pages_0_indexed=set()),
    dict(range_str="11", total_slides=10, expected_pages_0_indexed=set()),
    dict(range_str="0-5", total_slides=10, expected_pages_0_indexed={0, 1, 2, 3, 4}),
    dict(range_str="5-0", total_slides=10, expected_pages_0_indexed=set()),
    dict(range_str="2-4", total_slides=5, expected_pages_0_indexed={1, 2, 3}),
    dict(range_str="1-2", total_slides=0, expected_pages_0_indexed=set()),
    dict(range_str="3-", total_slides=5, expected_pages_0_indexed={2, 3, 4}),
    dict(range_str="-3", total_slides=5, expected_pages_0_indexed={0, 1, 2}),
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
    dict(range_str="1", total_slides=0, expected_pages_0_indexed=set()),
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
    dict(range_str="-7", total_slides=5, expected_pages_0_indexed={0, 1, 2, 3, 4}),
)
def test_parse_page_range_valid_inputs(
    range_str: str, total_slides: int, expected_pages_0_indexed: set[int]
) -> None:
    """Test parse_page_range with valid inputs."""
    assert parse_page_range(range_str, total_slides) == expected_pages_0_indexed


@pytest.mark.kwparametrize(
    dict(
        range_str="-",
        total_slides=5,
        expected_message_part="Invalid page range part format: '-'",
    ),
    dict(
        range_str="a-5",
        total_slides=10,
        expected_message_part="Invalid page range part format: 'a-5'",
    ),
    dict(
        range_str="1-b",
        total_slides=10,
        expected_message_part="Invalid page range part format: '1-b'",
    ),
    dict(
        range_str="a",
        total_slides=10,
        expected_message_part="Invalid page range part format: 'a'",
    ),
    dict(
        range_str="1-2-3",
        total_slides=10,
        expected_message_part="Invalid page range part format: '1-2-3'",
    ),
)
def test_parse_page_range_invalid_inputs(
    range_str: str, total_slides: int, expected_message_part: str
) -> None:
    """Test parse_page_range with invalid inputs raising BadParameter."""
    with pytest.raises(click.BadParameter) as excinfo:
        parse_page_range(range_str, total_slides)
    assert expected_message_part in str(excinfo.value)
