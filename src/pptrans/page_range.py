"""Page range parsing for presentation slides."""

import click


def parse_page_range(range_str, total_slides):
    """Parse a page range string (e.g., "1,3-5,8-") to a set of 0-indexed page numbers.

    Validates against total_slides. Raises click.BadParameter on error.
    1-indexed page numbers are used in the input string.
    """
    selected_pages_0_indexed = set()

    if total_slides == 0:
        # If a range_str was provided for a presentation with no slides, it's an issue.
        # If range_str is None or empty, this function might not even be called,
        # or it could return an empty set, which is fine.
        if range_str:  # Check if user actually specified a range for an empty deck
            msg = f"Cannot specify page range '{range_str}' for a presentation with no slides."
            raise click.BadParameter(msg)
        return selected_pages_0_indexed  # Correctly returns empty set if no slides

    # If range_str is None (meaning --pages not used), this function shouldn't be called by main.
    # The caller (main) should handle that by selecting all pages.
    # If range_str is an empty string (e.g., --pages=""), it's an invalid input if --pages was explicitly used.
    if (
        range_str is not None and not range_str.strip()
    ):  # Check if --pages was used and an empty string was passed
        msg = "Page range string cannot be empty if --pages option is used with an empty value."
        raise click.BadParameter(msg)

    parts = range_str.split(",")
    for part_specifier in parts:
        part = part_specifier.strip()
        if not part:  # Handles cases like "1,,2" or leading/trailing commas
            continue

        if "-" in part:
            # Range specifier
            elements = part.split("-", 1)
            start_str, end_str = elements[0].strip(), elements[1].strip()

            start_page_1_indexed = None
            end_page_1_indexed = None

            if start_str:  # "N-M" or "N-"
                try:
                    start_page_1_indexed = int(start_str)
                    if start_page_1_indexed < 1:
                        msg = f"Page numbers must be positive. Found start '{start_str}' in '{part_specifier}'."
                        raise click.BadParameter(msg)
                except ValueError:
                    msg = f"Invalid start page number '{start_str}' in '{part_specifier}'. Must be an integer."
                    raise click.BadParameter(msg)
            else:  # "-M"
                start_page_1_indexed = 1  # Default start for ranges like "-5"

            if end_str:  # "N-M" or "-M"
                try:
                    end_page_1_indexed = int(end_str)
                    if end_page_1_indexed < 1:
                        msg = f"Page numbers must be positive. Found end '{end_str}' in '{part_specifier}'."
                        raise click.BadParameter(msg)
                except ValueError:
                    msg = f"Invalid end page number '{end_str}' in '{part_specifier}'. Must be an integer."
                    raise click.BadParameter(msg)
            else:  # "N-"
                end_page_1_indexed = total_slides  # Default end for ranges like "5-"

            if start_page_1_indexed > end_page_1_indexed:
                msg = f"Start page {start_page_1_indexed} cannot be greater than end page {end_page_1_indexed} in range '{part_specifier}'."
                raise click.BadParameter(msg)

            # Validate individual bounds against total_slides
            if start_page_1_indexed > total_slides:
                # This also covers cases where start_page_1_indexed was 1 (for "-M") but total_slides is 0 (already handled)
                # or if user specifies "10-" for a 5-slide deck.
                # No need to add to set if start is already beyond total slides.
                click.echo(
                    f"Warning: Start page {start_page_1_indexed} in range '{part_specifier}' is beyond the total of {total_slides} slides. This part of the range will select no pages.",
                    err=True,
                )
                continue  # Skip this part of the range

            # For "N-M" or "-M", if end_page_1_indexed was explicitly given and exceeds total_slides
            if end_str and end_page_1_indexed > total_slides:
                click.echo(
                    f"Warning: End page {end_page_1_indexed} in range '{part_specifier}' is beyond the total of {total_slides} slides. Range will be capped at {total_slides}.",
                    err=True,
                )
                # end_page_1_indexed will be capped by min() below.

            # Add pages (0-indexed) to the set
            # Cap end_page_1_indexed at total_slides for ranges like "N-" or valid "N-M"
            # This loop will correctly handle start_page_1_indexed up to min(end_page_1_indexed, total_slides)
            for i in range(
                start_page_1_indexed, min(end_page_1_indexed, total_slides) + 1
            ):
                selected_pages_0_indexed.add(i - 1)
        else:
            # Single page specifier
            try:
                page_num_1_indexed = int(part)
                if not (1 <= page_num_1_indexed <= total_slides):
                    msg = f"Page number {page_num_1_indexed} in '{part_specifier}' is out of valid range [1, {total_slides}]."
                    raise click.BadParameter(msg)
                selected_pages_0_indexed.add(page_num_1_indexed - 1)
            except ValueError:
                msg = f"Invalid page number: '{part_specifier}'. Must be an integer."
                raise click.BadParameter(msg)

    return selected_pages_0_indexed
