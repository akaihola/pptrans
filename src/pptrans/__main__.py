"""Translate PowerPoint presentations."""

import shutil  # For file copying

import click
import llm  # Simon Willison's LLM library
from pptx import Presentation

from pptrans.cache import generate_page_hash, load_cache, save_cache
from pptrans.page_range import parse_page_range

EOL_MARKER = "<"


def reverse_individual_words(text_string_with_eol):
    """Reverses each word in a space-separated string, preserving an EOL_MARKER if present."""
    text_to_reverse = text_string_with_eol
    had_eol = False
    if text_string_with_eol.endswith(EOL_MARKER):
        text_to_reverse = text_string_with_eol[: -len(EOL_MARKER)]
        had_eol = True

    words = text_to_reverse.split(" ")
    reversed_words = [word[::-1] for word in words]
    result = " ".join(reversed_words)

    if had_eol:
        return result + EOL_MARKER
    return result


@click.command()
@click.option(
    "--mode",
    type=click.Choice(
        ["translate", "reverse-words"],
        case_sensitive=False,  # 'duplicate-only' mode removed
    ),
    default="translate",
    show_default=True,
    help="Operation mode for the script.",
)
@click.option(
    "--pages",
    type=str,
    default=None,
    help="Specify page range to process (e.g., '1,3-5,8-'). 1-indexed. Processes all pages if not specified.",
)
@click.argument("input_path", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_path", type=click.Path(dir_okay=False))
def main(input_path, output_path, mode, pages):
    """Processes a PowerPoint presentation.
    It first copies the input presentation to the output path.
    Then, for 'translate' and 'reverse-words' modes, text on selected slides
    within this copied presentation is modified in place.
    'translate' mode uses an on-disk cache.
    Page range can be specified using the --pages option.
    """
    cache_file_path = (
        "translation_cache.json"  # Cache file in the same directory as script execution
    )

    click.echo(f"Copying '{input_path}' to '{output_path}' to preserve layout...")
    try:
        shutil.copy2(input_path, output_path)
    except Exception as e:
        click.echo(f"Error copying file: {e}", err=True)
        return
    click.echo("File copy complete.")

    click.echo(f"Loading presentation for modification from: {output_path}")
    prs = Presentation(output_path)

    num_original_slides = len(prs.slides)
    if num_original_slides == 0:
        click.echo("Input presentation has no slides. Exiting.")
        return

    selected_pages_0_indexed = set()
    if pages:
        try:
            selected_pages_0_indexed = parse_page_range(pages, num_original_slides)
        except click.BadParameter:
            # parse_page_range raises BadParameter, which click handles by exiting.
            # We can re-raise if we want to be explicit or add more context, but click does it.
            # For now, let click handle the exit.
            # click.echo(f"Error: {e}", err=True) # This would be redundant
            raise  # Re-raise to ensure click handles it as expected
    else:
        # If --pages is not specified, process all slides
        selected_pages_0_indexed = set(range(num_original_slides))

    slides_to_process_objects = []
    if num_original_slides > 0:
        for i, slide_obj in enumerate(prs.slides):
            if i in selected_pages_0_indexed:
                slides_to_process_objects.append(slide_obj)

    if not slides_to_process_objects:
        if pages:  # User specified a range, but it resulted in no slides
            click.echo(
                f"Warning: The specified page range '{pages}' resulted in no slides being selected from {num_original_slides} total slides. No text processing will occur."
            )
        else:  # No pages specified, but presentation was empty (already handled) or became empty (should not happen here)
            click.echo(
                f"No slides selected for processing from {num_original_slides} total slides. No text processing will occur."
            )
        # Save the copied presentation and exit if no slides are to be processed.
        # This is important if only a copy was intended or if the range was valid but empty.
        prs.save(output_path)
        click.echo(f"Presentation saved (no text modifications) to: {output_path}")
        # Save cache even if no text elements were processed, in case it was loaded and needs to be preserved.
        if mode == "translate":
            translation_cache = load_cache(
                cache_file_path
            )  # Ensure cache is loaded if not already
            save_cache(translation_cache, cache_file_path)
        return

    # Renaming slides_for_text_extraction to slides_to_process_objects for clarity
    # The old slides_for_text_extraction is now slides_to_process_objects

    # The original logic for 'slides_for_text_extraction' initialization is now replaced by the above.
    # The following 'if mode == "translate" or mode == "reverse-words":' block's content
    # related to populating slides_for_text_extraction is no longer needed here as slides_to_process_objects is already set.
    # We just need to update the logging.

    if mode == "translate" or mode == "reverse-words":
        # slides_for_text_extraction = list(prs.slides) # This line is now replaced by slides_to_process_objects logic
        if slides_to_process_objects:  # Only print if there are slides to process
            click.echo(
                f"Preparing to process text on {len(slides_to_process_objects)} selected slide(s) (out of {num_original_slides} total) in the copied presentation..."
            )
    # 'duplicate-only' mode and slide duplication logic removed.
    # Text processing will now occur directly on the slides in 'slides_for_text_extraction'.

    # Text processing for 'translate' and 'reverse-words' modes
    text_id_counter = 0  # Shared counter for text element IDs

    if mode == "translate":
        click.echo(f"Loading translation cache from: {cache_file_path}")
        translation_cache = load_cache(
            cache_file_path
        )  # This will now be page-hash based

        global_texts_for_llm_prompt = []  # Stores items for LLM: {id, original_text_for_cache, text_to_send, run_object, page_hash}
        all_processed_run_details = []  # Stores details for all runs: {run_object, final_translation, from_cache, original_text, llm_id (if applicable)}
        pending_page_cache_updates = {}  # {page_hash: [{"original_text": ..., "translation": ...}]}

        if slides_to_process_objects:
            click.echo(
                f"Processing {len(slides_to_process_objects)} selected slide(s) for translation..."
            )
            for slide_idx, slide_to_extract in enumerate(slides_to_process_objects):
                current_page_texts_for_hash = []
                current_page_run_info = []  # List of {"original_text": ..., "run_object": ...}

                # First pass: extract all texts from the current slide for hashing and run info
                for shape in slide_to_extract.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text
                                if original_text:
                                    current_page_texts_for_hash.append(original_text)
                                    current_page_run_info.append(
                                        {
                                            "original_text": original_text,
                                            "run_object": run,
                                        }
                                    )
                    if shape.has_table:
                        for row_idx, row in enumerate(shape.table.rows):
                            for col_idx, cell in enumerate(row.cells):
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        original_text = run.text
                                        if original_text:
                                            current_page_texts_for_hash.append(
                                                original_text
                                            )
                                            current_page_run_info.append(
                                                {
                                                    "original_text": original_text,
                                                    "run_object": run,
                                                }
                                            )

                if not current_page_texts_for_hash:
                    click.echo(f"  Slide {slide_idx + 1}: No text found.")
                    continue

                page_hash = generate_page_hash(current_page_texts_for_hash)
                click.echo(
                    f"  Slide {slide_idx + 1}: Hash '{page_hash[:8]}...', {len(current_page_run_info)} text runs."
                )

                if page_hash in translation_cache:
                    click.echo(f"    Page cache hit for hash {page_hash[:8]}...")
                    cached_translations_for_page = translation_cache[page_hash]

                    for run_detail in current_page_run_info:
                        found_in_page_cache = False
                        for cached_item in cached_translations_for_page:
                            if (
                                cached_item["original_text"]
                                == run_detail["original_text"]
                            ):
                                all_processed_run_details.append(
                                    {
                                        "run_object": run_detail["run_object"],
                                        "final_translation": cached_item["translation"],
                                        "from_cache": True,
                                        "original_text": run_detail["original_text"],
                                    }
                                )
                                found_in_page_cache = True
                                break
                        if not found_in_page_cache:
                            # This case implies the page hash matched, but an individual text on that page
                            # wasn't in the cached list for that page. This might happen if the page structure
                            # is identical but some minor text was edited, then re-edited back to make the hash match,
                            # but the cache entry for that page is from a state where that specific text was different.
                            # For simplicity, we'll treat this as needing LLM translation for this specific run.
                            click.echo(
                                f"    Partial page cache hit for {page_hash[:8]}. Text '{run_detail['original_text'][:30]}...' not in page's cached list. Sending to LLM."
                            )
                            text_id = f"text_{text_id_counter}"
                            text_id_counter += 1
                            global_texts_for_llm_prompt.append(
                                {
                                    "id": text_id,
                                    "original_text_for_cache": run_detail[
                                        "original_text"
                                    ],
                                    "text_to_send": run_detail["original_text"]
                                    + EOL_MARKER,
                                    "run_object": run_detail["run_object"],
                                    "page_hash": page_hash,  # Associate with current page
                                }
                            )
                            all_processed_run_details.append(
                                {
                                    "run_object": run_detail["run_object"],
                                    "final_translation": None,  # Will be filled by LLM
                                    "from_cache": False,
                                    "original_text": run_detail["original_text"],
                                    "llm_id": text_id,
                                }
                            )
                            # Ensure this page is marked for potential cache update
                            if page_hash not in pending_page_cache_updates:
                                pending_page_cache_updates[page_hash] = []

                else:  # Page cache miss
                    click.echo(
                        f"    Page cache miss for hash {page_hash[:8]}. Will send {len(current_page_run_info)} runs to LLM."
                    )
                    pending_page_cache_updates[
                        page_hash
                    ] = []  # Prepare to build this page's cache entry
                    for run_detail in current_page_run_info:
                        text_id = f"text_{text_id_counter}"
                        text_id_counter += 1
                        global_texts_for_llm_prompt.append(
                            {
                                "id": text_id,
                                "original_text_for_cache": run_detail["original_text"],
                                "text_to_send": run_detail["original_text"]
                                + EOL_MARKER,
                                "run_object": run_detail["run_object"],
                                "page_hash": page_hash,
                            }
                        )
                        all_processed_run_details.append(
                            {
                                "run_object": run_detail["run_object"],
                                "final_translation": None,  # Will be filled by LLM
                                "from_cache": False,
                                "original_text": run_detail["original_text"],
                                "llm_id": text_id,
                            }
                        )

        if (
            not all_processed_run_details
        ):  # Check if any text runs were collected at all
            click.echo(
                f"No text found to process on selected slides for mode '{mode}'."
            )
            prs.save(output_path)
            click.echo(
                f"Presentation saved without text modification in '{mode}' mode to: {output_path}"
            )
            # Save cache even if empty/unchanged, as it might have been loaded and format changed
            click.echo(
                f"Saving cache (potentially empty or format updated) to: {cache_file_path}"
            )
            save_cache(translation_cache, cache_file_path)
            return

        click.echo(
            f"Processed {len(all_processed_run_details)} total text runs. {len(global_texts_for_llm_prompt)} to translate via LLM."
        )

        if not global_texts_for_llm_prompt:
            click.echo("All text elements found in page caches. Skipping LLM prompt.")
        else:
            click.echo(
                f"Sending {len(global_texts_for_llm_prompt)} text elements to LLM for translation."
            )
            # Sort by ID to ensure consistent order for LLM prompt, if IDs are not strictly sequential
            # global_texts_for_llm_prompt.sort(key=lambda x: int(x['id'].split('_')[1]))
            # Decided against sorting for now, as text_id_counter should ensure order.

            formatted_text_for_llm = "\n".join(
                [
                    f"{item['id']}:{item['text_to_send']}"
                    for item in global_texts_for_llm_prompt
                ]
            )
            prompt_text = (
                "You are an expert Finnish to English translator. "
                "Translate the following text segments accurately from Finnish to English. "
                "Each segment is prefixed with a unique ID (e.g., text_0, text_1). "
                "IMPORTANT: A sequence of text items (e.g., text_0, text_1, text_2) may represent a single continuous sentence that has been split due to formatting. Interpret and translate such sequences as a coherent whole sentence to maintain context and flow. "
                "The text for each ID might end with an EOL marker: '<'. "
                "Your response MUST consist ONLY of the translated segments, each prefixed with its original ID, "
                "and each on a new line. Maintain the exact ID and format. "
                "PRESERVE ALL LEADING AND TRAILING WHITESPACE from the original segment in your translation. "
                "If an EOL marker '<' was present at the end of the input segment, IT MUST be present at the end of your translated segment, including any whitespace before it.\n"
                "For example, if you receive:\n"
                "text_0: Tämä on pitkä \n"
                "text_1:lause, joka on \n"
                "text_2:jaettu.<\n"
                "text_3:    Toinen lause.   <\n"
                "text_4: Yksittäinen.\n"
                "You MUST return:\n"
                "text_0: This is a long \n"
                "text_1:sentence that has been \n"
                "text_2:split.<\n"
                "text_3:    Another sentence.   <\n"
                "text_4: Standalone.\n\n"
                "Do not add any extra explanations, apologies, or introductory/concluding remarks. "
                "Only provide the ID followed by the translated text for each item.\n\n"
                "Texts to translate:\n"
            )
            click.echo("--- PROMPT TO LLM ---")
            click.echo("System/Instruction Prompt:")
            click.echo(prompt_text)
            click.echo("Data Fragments (for LLM):")
            click.echo(formatted_text_for_llm)
            click.echo("--- END OF PROMPT ---")

            model_instance = llm.get_model()
            response = model_instance.prompt(
                prompt_text, fragments=[formatted_text_for_llm]
            )
            translated_text_response = response.text()

            click.echo("--- RESPONSE FROM LLM ---")
            click.echo(translated_text_response)
            click.echo("--- END OF RESPONSE ---")
            click.echo("Received translation from LLM.")

            for line in translated_text_response.splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        parsed_text_id = parts[0].strip()
                        llm_translation_with_eol = parts[1]

                        prompt_item_data = next(
                            (
                                item
                                for item in global_texts_for_llm_prompt
                                if item["id"] == parsed_text_id
                            ),
                            None,
                        )

                        if prompt_item_data:
                            original_text_for_cache_key = prompt_item_data[
                                "original_text_for_cache"
                            ]
                            current_page_hash = prompt_item_data["page_hash"]

                            final_llm_translation = llm_translation_with_eol
                            final_llm_translation = final_llm_translation.removesuffix(
                                EOL_MARKER
                            )

                            # Add to pending_page_cache_updates for the specific page
                            # Ensure the list for the page_hash exists
                            if current_page_hash not in pending_page_cache_updates:
                                pending_page_cache_updates[current_page_hash] = []

                            # Avoid duplicate entries if a text appears multiple times on a page and was sent to LLM
                            # (though current logic sends each run instance, so original_text might not be unique in the list for a page)
                            # For now, we assume each original_text within a page that went to LLM is distinct enough or handled by run_object uniqueness.
                            # The cache structure is a list of {"original_text": ..., "translation": ...} for the page.

                            # Update the all_processed_run_details list
                            for detail_item in all_processed_run_details:
                                if detail_item.get("llm_id") == parsed_text_id:
                                    detail_item["final_translation"] = (
                                        final_llm_translation
                                    )
                                    # We also need to prepare for saving this to the page's cache entry
                                    # Check if this original_text is already slated for this page_hash update
                                    found_in_pending = False
                                    for pending_item in pending_page_cache_updates[
                                        current_page_hash
                                    ]:
                                        if (
                                            pending_item["original_text"]
                                            == original_text_for_cache_key
                                        ):
                                            pending_item["translation"] = (
                                                final_llm_translation  # Update if somehow already there
                                            )
                                            found_in_pending = True
                                            break
                                    if not found_in_pending:
                                        pending_page_cache_updates[
                                            current_page_hash
                                        ].append(
                                            {
                                                "original_text": original_text_for_cache_key,
                                                "translation": final_llm_translation,
                                            }
                                        )
                                    break
                        else:
                            click.echo(
                                f"Warning: Could not find original data for ID {parsed_text_id} from LLM response.",
                                err=True,
                            )
                    else:
                        click.echo(
                            f"Warning: Could not parse translation line: {line}",
                            err=True,
                        )
                except Exception as e:
                    click.echo(
                        f"Warning: Error parsing translation line '{line}': {e}",
                        err=True,
                    )

        click.echo(
            "Replacing text with translations on slides in the copied presentation..."
        )
        for item in all_processed_run_details:
            if (
                item["final_translation"] is not None
            ):  # Check for None, as empty string is a valid translation
                item["run_object"].text = item["final_translation"]
            elif not item["from_cache"]:
                click.echo(
                    f"Warning: No translation found for run with original text '{item['original_text'][:30]}...' (LLM ID: {item.get('llm_id', 'N/A')}). Leaving original.",
                    err=True,
                )

        click.echo(
            f"Updating and saving page-based translation cache to: {cache_file_path}"
        )
        # Merge pending updates into the main translation_cache
        for page_hash, translations_list in pending_page_cache_updates.items():
            # If a page was partially a cache hit but had some new items for LLM,
            # we need to merge new translations with existing ones for that page.
            # The current logic for `pending_page_cache_updates[page_hash]` only adds LLM results.
            # If a page was a full cache hit, it's not in pending_page_cache_updates.
            # If a page was a full cache miss, pending_page_cache_updates[page_hash] contains all its items.
            # If a page was a partial hit (some runs from cache, some to LLM):
            #   - Runs from cache are already applied.
            #   - Runs to LLM have their translations in pending_page_cache_updates[page_hash].
            #   - We need to ensure the final cache entry for this page_hash contains *all* original_text/translation pairs.

            # Simplest approach for now: if a page_hash is in pending_page_cache_updates,
            # it means it had at least one LLM translation. We'll build its cache entry
            # from all_processed_run_details that belong to that page_hash.

            # Rebuild the cache entry for any page that had LLM involvement or was a full miss.
            if (
                page_hash in pending_page_cache_updates
            ):  # Indicates LLM was involved for this page
                rebuilt_page_cache_entry = []
                # Find all runs associated with this page_hash from all_processed_run_details
                # This is inefficient if done here. Better to build pending_page_cache_updates correctly during LLM response processing.
                # For now, let's assume pending_page_cache_updates[page_hash] has the full list for pages that had misses/LLM calls.
                # The current logic for populating pending_page_cache_updates might be okay if it collects all items for a page_hash that had any LLM calls.

                # Let's refine: pending_page_cache_updates should store the *complete* list of translations for a page if it's being updated.
                # The current LLM response loop populates it with LLM results.
                # If a page was a cache hit, it's not in pending_page_cache_updates.
                # If a page was a cache miss, all its items went to LLM, so pending_page_cache_updates[page_hash] will be complete.
                # The tricky case is a "partial page cache hit" where the hash matched, but some texts were new.
                # In this scenario, `translation_cache[page_hash]` (the old entry) needs to be augmented/replaced.
                # The current logic for "partial page cache hit" adds new items to global_texts_for_llm_prompt,
                # and their translations will be added to pending_page_cache_updates[page_hash].
                # We need to merge these with the *original* cached items for that page that were hits.

                # Revised strategy for updating translation_cache:
                # For any page_hash that appears in pending_page_cache_updates:
                #   Construct its new cache entry by taking all items from all_processed_run_details
                #   that correspond to that page_hash (need to add page_hash to all_processed_run_details items).
                # This is complex. Let's simplify: if a page_hash is in pending_page_cache_updates,
                # it means it was either a full miss, or a partial miss where new items were sent to LLM.
                # The `pending_page_cache_updates[page_hash]` should contain the *newly translated* items.
                # If it was a full miss, this is the complete list for the page.
                # If it was a partial miss, we need to merge with existing cached items for that page.

                # Let's stick to the plan: `pending_page_cache_updates[page_hash]` stores the *full list* for pages that had misses.
                # The LLM response loop correctly adds `{"original_text": ..., "translation": ...}` to `pending_page_cache_updates[current_page_hash]`.
                # This list should be complete for pages that had any LLM calls.
                if translations_list:  # Only update if there are actual translations
                    translation_cache[page_hash] = translations_list
                elif page_hash in translation_cache and not translations_list:
                    # This means a page hash was in pending_updates (so it was a miss or partial)
                    # but ended up with no translations (e.g. LLM returned nothing for its items).
                    # We should probably remove it from cache to force re-translation next time,
                    # or ensure pending_page_cache_updates is built correctly with all original texts
                    # even if translations are empty.
                    # For now, if translations_list is empty, we don't add/update it,
                    # which means if it was a full miss, it won't be cached. If it was partial,
                    # the old cache entry might persist if not explicitly overwritten.
                    # This needs to be robust: if a page is processed and had LLM calls, its entry in
                    # translation_cache should reflect the latest state of all its texts.

                    # Corrected logic: if page_hash is in pending_page_cache_updates, it means we intended to update it.
                    # The list `translations_list` should be the definitive new list of original_text/translation pairs for that page.
                    translation_cache[page_hash] = translations_list

        save_cache(translation_cache, cache_file_path)

    elif mode == "reverse-words":
        # Original logic for reverse-words, uses a simple text_elements list
        text_elements_for_reverse = []
        if slides_to_process_objects:
            click.echo(
                f"Extracting text from {len(slides_to_process_objects)} slides in the copied presentation for mode '{mode}'..."
            )
            for slide_to_extract in slides_to_process_objects:
                for shape in slide_to_extract.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text  # No strip
                                if original_text:
                                    text_with_eol = original_text + EOL_MARKER
                                    text_id = (
                                        f"text_{text_id_counter}"  # Uses global counter
                                    )
                                    text_elements_for_reverse.append(
                                        {
                                            "id": text_id,
                                            "text": text_with_eol,  # Store with EOL
                                            "run_object": run,
                                        }
                                    )
                                    text_id_counter += 1
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        original_text = run.text  # No strip
                                        if original_text:
                                            text_with_eol = original_text + EOL_MARKER
                                            text_id = f"text_{text_id_counter}"  # Uses global counter
                                            text_elements_for_reverse.append(
                                                {
                                                    "id": text_id,
                                                    "text": text_with_eol,  # Store with EOL
                                                    "run_object": run,
                                                }
                                            )
                                            text_id_counter += 1

        if not text_elements_for_reverse:
            click.echo(
                f"No text found to process on slides in the copied presentation for mode '{mode}'."
            )
            prs.save(output_path)
            click.echo(
                f"Presentation saved without text modification in '{mode}' mode to: {output_path}"
            )
            return

        click.echo(
            f"Found {len(text_elements_for_reverse)} text elements to process for mode '{mode}'."
        )
        click.echo("Applying word reversal on slides in the copied presentation...")
        for item in text_elements_for_reverse:
            reversed_text_with_eol = reverse_individual_words(item["text"])
            final_reversed_text = reversed_text_with_eol
            final_reversed_text = final_reversed_text.removesuffix(EOL_MARKER)
            item["run_object"].text = final_reversed_text
        click.echo(
            "Text replaced with reversed-word text on slides in the copied presentation."
        )

    prs.save(output_path)
    click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")


if __name__ == "__main__":
    main()
