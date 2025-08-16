"""Page range parsing for presentation slides."""

import re

import click

# Pre-compile regex for efficiency if this function is called multiple times.
PAGE_PART_REGEX = re.compile(
    # Matches: N-M (groups 1,2), N (group 3), or -M (group 4)
    # \s* allows whitespace around numbers and hyphens.
    # ^ and $ ensure the whole part matches the pattern.
    r"^\s*(?:(\d+)\s*-\s*(\d*)|(\d+)|-\s*(\d+))\s*$"
)


def parse_page_range(range_str: str, total_slides: int) -> set[int]:
    """Parse a page range string (e.g., "1,3-5,8-") to a set of 0-indexed page numbers.

    Relies on standard Python exceptions for invalid formats or values.
    1-indexed page numbers are used in the input string.
    """
    selected_pages_0_indexed: set[int] = set()

    if not range_str.strip():  # Handles empty or whitespace-only input.
        return selected_pages_0_indexed

    for part_raw in range_str.split(","):
        part = part_raw.strip()
        if not part:  # Skips empty parts like in "1,,2".
            continue

        match = PAGE_PART_REGEX.match(part)
        if not match:
            msg = f"Invalid page range part format: '{part}'"
            raise click.BadParameter(msg)

        g = match.groups()
        start_1_idx: int
        end_1_idx: int

        if g[0] is not None:  # Matched "N-M" or "N-"
            # g[0] is N (start of range)
            # g[1] is M (end of range, empty for "N-")
            start_1_idx = int(g[0])
            end_1_idx = int(g[1]) if g[1] else total_slides
        elif g[2] is not None:  # Matched "N" (single page)
            # g[2] is N
            start_1_idx = int(g[2])
            end_1_idx = start_1_idx
        elif g[3] is not None:  # Matched "-M"
            # g[3] is M
            start_1_idx = 1
            end_1_idx = int(g[3])
        else:
            # This case should not happen due to the regex structure.
            # If it does, it indicates an unexpected format.
            msg = f"Unexpected match failure for part: '{part}'"
            raise click.BadParameter(msg)

        # Determine the actual loop bounds, clamped to valid 1-indexed pages.
        # Pages < 1 are effectively treated as 1 for the start.
        # Pages > total_slides are capped at total_slides for the end.
        # If total_slides is 0, actual_loop_end will be <= 0.
        loop_start_clamped = max(1, start_1_idx)
        loop_end_clamped = min(total_slides, end_1_idx)

        # Add 0-indexed pages to the set.
        # The loop `range(A, B+1)` is empty if A > B.
        for i in range(loop_start_clamped, loop_end_clamped + 1):
            selected_pages_0_indexed.add(i - 1)

    return selected_pages_0_indexed
