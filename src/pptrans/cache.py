"""Cache management for translations."""

import hashlib
import json
from pathlib import Path
from typing import Any  # Added for type hints

import click


def load_cache(cache_file_path: str) -> dict:
    """Load the translation cache from a JSON file."""
    cache_file = Path(cache_file_path)
    if cache_file.exists():
        try:
            with cache_file.open(encoding="utf-8") as f:
                # Log successful cache loading
                click.echo(
                    f"Successfully loaded translation cache from: {cache_file_path}"
                )
                return json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            click.echo(
                f"Warning: Could not load cache file {cache_file_path}. Error: {e}. "
                "Starting with an empty cache.",
                err=True,
            )
    else:
        click.echo(
            f"Cache file not found at {cache_file_path}. Starting with an empty cache."
        )
    return {}


def save_cache(cache_data: dict, cache_file_path: str) -> None:
    """Save the translation cache to a JSON file."""
    try:
        with Path(cache_file_path).open("w", encoding="utf-8") as f:
            json.dump(cache_data, f, indent=4, ensure_ascii=False)
        click.echo(f"Translation cache saved to: {cache_file_path}")
    except OSError as e:
        click.echo(
            f"Warning: Could not save cache file {cache_file_path}. Error: {e}",
            err=True,
        )


def generate_page_hash(texts_on_page: list[str]) -> str:
    """Generate a SHA256 hash for a list of text strings from a page."""
    concatenated_texts = "|".join(texts_on_page)  # Use a delimiter
    return hashlib.sha256(concatenated_texts.encode("utf-8")).hexdigest()


def prepare_slide_for_translation(
    slide_run_info: list[dict[str, Any]],
    page_hash: str,
    translation_cache: dict[str, list[dict[str, str]]],
    text_id_counter: int,
    eol_marker: str,
    page_number_1_indexed: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int, bool]:
    """Prepare slide text runs for translation, checking against the cache.

    Args:
        slide_run_info: List of {"original_text": ..., "run_object": ...}.
        page_hash: Hash of the current slide's content.
        translation_cache: The current translation cache.
        text_id_counter: Current unique ID counter for LLM texts.
        eol_marker: End-of-line marker string.
        page_number_1_indexed: 1-indexed page number of the current slide.

    Returns:
        A tuple containing:
            - texts_for_llm: List of items to send to LLM.
            - processed_runs_for_slide: Details for all runs on the slide.
            - updated_text_id_counter: The new text_id_counter.
            - page_requires_llm_processing: Boolean.

    """
    texts_for_llm: list[dict[str, Any]] = []
    processed_runs_for_slide: list[dict[str, Any]] = []
    page_requires_llm_processing = False

    click.echo(
        f"  Slide Hash '{page_hash[:8]}...': "
        f"Processing {len(slide_run_info)} text runs."
    )

    if page_hash in translation_cache:
        click.echo(f"    Page cache hit for hash {page_hash[:8]}...")
        cached_translations_for_page = translation_cache[page_hash]

        for run_detail in slide_run_info:
            original_text = run_detail["original_text"]
            found_in_page_cache = False
            for cached_item in cached_translations_for_page:
                if cached_item["original_text"] == original_text:
                    processed_runs_for_slide.append(
                        {
                            "run_object": run_detail["run_object"],
                            "final_translation": cached_item["translation"],
                            "from_cache": True,
                            "original_text": original_text,
                            "llm_id": None,  # Not sent to LLM
                        }
                    )
                    click.echo(f"      Text cache hit for: '{original_text[:30]}...'")
                    found_in_page_cache = True
                    break

            if not found_in_page_cache:
                click.echo(
                    f"      Partial page cache hit. Text '{original_text[:30]}...' "
                    "not in page's cached list. Sending to LLM."
                )
                text_id = f"pg{page_number_1_indexed}_txt{text_id_counter}"
                text_id_counter += 1
                page_requires_llm_processing = True
                texts_for_llm.append(
                    {
                        "id": text_id,
                        "original_text_for_cache": original_text,
                        "text_to_send": original_text + eol_marker,
                        "run_object": run_detail["run_object"],
                        "page_hash": page_hash,
                    }
                )
                processed_runs_for_slide.append(
                    {
                        "run_object": run_detail["run_object"],
                        "final_translation": None,  # Will be filled by LLM
                        "from_cache": False,
                        "original_text": original_text,
                        "llm_id": text_id,
                    }
                )
    else:  # Page cache miss
        click.echo(
            f"    Page cache miss for hash {page_hash[:8]}. "
            f"Will send {len(slide_run_info)} runs to LLM."
        )
        if slide_run_info:  # Only set to True if there are actual runs to process
            page_requires_llm_processing = True
        for run_detail in slide_run_info:
            original_text = run_detail["original_text"]
            text_id = f"pg{page_number_1_indexed}_txt{text_id_counter}"
            text_id_counter += 1
            texts_for_llm.append(
                {
                    "id": text_id,
                    "original_text_for_cache": original_text,
                    "text_to_send": original_text + eol_marker,
                    "run_object": run_detail["run_object"],
                    "page_hash": page_hash,
                }
            )
            processed_runs_for_slide.append(
                {
                    "run_object": run_detail["run_object"],
                    "final_translation": None,  # Will be filled by LLM
                    "from_cache": False,
                    "original_text": original_text,
                    "llm_id": text_id,
                }
            )
    return (
        texts_for_llm,
        processed_runs_for_slide,
        text_id_counter,
        page_requires_llm_processing,
    )


def update_data_from_llm_response(
    llm_response_lines: list[str],
    global_texts_for_llm_prompt_ref: list[dict[str, Any]],
    all_processed_run_details_ref: list[dict[str, Any]],
    pending_page_cache_updates_ref: dict[str, list[dict[str, str]]],
    eol_marker: str,
) -> None:
    """Update data structures based on the LLM's translation response.

    Args:
        llm_response_lines: Lines from the LLM response.
        global_texts_for_llm_prompt_ref: Reference to list of items sent to LLM.
        all_processed_run_details_ref: Reference to list tracking all text runs.
        pending_page_cache_updates_ref: Reference to dict for staging cache updates.
        eol_marker: End-of-line marker string.

    """
    click.echo("Processing LLM response to update translations and cache...")
    for orig_line in llm_response_lines:
        line = orig_line.strip()
        if not line:
            continue
        try:
            parts = line.split(":", 1)
            if len(parts) == 2:  # noqa: PLR2004
                parsed_text_id = parts[0].strip()
                llm_translation_with_eol = parts[
                    1
                ]  # Keep leading/trailing spaces from LLM

                prompt_item_data = next(
                    (
                        item
                        for item in global_texts_for_llm_prompt_ref
                        if item["id"] == parsed_text_id
                    ),
                    None,
                )

                if prompt_item_data:
                    original_text_for_cache = prompt_item_data[
                        "original_text_for_cache"
                    ]
                    current_page_hash = prompt_item_data["page_hash"]

                    final_llm_translation = llm_translation_with_eol
                    final_llm_translation = final_llm_translation.removesuffix(
                        eol_marker
                    )

                    # Update all_processed_run_details
                    for detail_item in all_processed_run_details_ref:
                        if detail_item.get("llm_id") == parsed_text_id:
                            detail_item["final_translation"] = final_llm_translation
                            break
                            # Found and updated, break from inner loop

                    # Add to pending_page_cache_updates for the specific page
                    if current_page_hash not in pending_page_cache_updates_ref:
                        # This should have been initialized if
                        # page_requires_llm_processing was true
                        click.echo(
                            f"Warning: page_hash {current_page_hash} not "
                            "pre-initialized in pending_page_cache_updates. "
                            "Initializing now.",
                            err=True,
                        )
                        pending_page_cache_updates_ref[current_page_hash] = []

                    # Avoid duplicate entries in the cache list for a page
                    found_in_pending = False
                    for pending_item in pending_page_cache_updates_ref[
                        current_page_hash
                    ]:
                        if pending_item["original_text"] == original_text_for_cache:
                            pending_item["translation"] = (
                                final_llm_translation  # Update if exists
                            )
                            found_in_pending = True
                            break
                    if not found_in_pending:
                        pending_page_cache_updates_ref[current_page_hash].append(
                            {
                                "original_text": original_text_for_cache,
                                "translation": final_llm_translation,
                            }
                        )
                else:
                    click.echo(
                        f"Warning: Could not find original data for ID "
                        f"{parsed_text_id} from LLM response.",
                        err=True,
                    )
            else:
                click.echo(
                    f"Warning: Could not parse translation line: '{line}'",
                    err=True,
                )
        except Exception as e:  # noqa: BLE001
            click.echo(
                f"Warning: Error parsing translation line '{line}': {e}",
                err=True,
            )


def commit_pending_cache_updates(
    translation_cache_ref: dict[str, list[dict[str, str]]],
    pending_page_cache_updates: dict[str, list[dict[str, str]]],
    cache_file_path: str,
) -> None:
    """Merge pending cache updates into the main cache and saves it.

    Args:
        translation_cache_ref: Reference to the main translation cache.
        pending_page_cache_updates: Dict of new translations grouped by page_hash.
        cache_file_path: Path to the cache file.

    """
    click.echo(
        "Updating and preparing to save page-based translation cache to: "
        f"{cache_file_path}"
    )
    for page_hash, translations_list in pending_page_cache_updates.items():
        if translations_list:  # Only update if there are actual new translations
            translation_cache_ref[page_hash] = translations_list
            click.echo(
                f"  Updated cache for page_hash: {page_hash[:8]}... with "
                f"{len(translations_list)} items."
            )
        elif not translations_list:  # translations_list is empty
            # If the page_hash already exists in the cache, and the pending update
            # is an empty list, it means we should clear its translations.
            if page_hash in translation_cache_ref:
                translation_cache_ref[page_hash] = []
                click.echo(
                    f"  Cleared translations for existing page_hash: {page_hash[:8]}..."
                )
            # If page_hash is new and translations_list is empty,
            # we ignore it; it's not added to the cache.

    save_cache(translation_cache_ref, cache_file_path)
