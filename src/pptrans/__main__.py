"""Translate PowerPoint presentations."""

import shutil  # For file copying
from typing import Any  # For type hints

import click
import llm  # Simon Willison's LLM library
from pptx import Presentation

from pptrans.cache import (
    commit_pending_cache_updates,
    # save_cache is called by commit_pending_cache_updates
    generate_page_hash,
    load_cache,
    prepare_slide_for_translation,
    update_data_from_llm_response,
)
from pptrans.page_range import parse_page_range

EOL_MARKER = "<"


def reverse_individual_words(text_string_with_eol: str) -> str:
    """Reverse words in a space-separated string, preserving an EOL_MARKER."""
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
        case_sensitive=False,
    ),
    default="translate",
    show_default=True,
    help="Operation mode for the script.",
)
@click.option(
    "--pages",
    type=str,
    default=None,
    help=(
        "Specify page range to process (e.g., '1,3-5,8-'). 1-indexed. "
        "Processes all pages if not specified."
    ),
)
@click.argument("input_path", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_path", type=click.Path(dir_okay=False))
def main(input_path: str, output_path: str, mode: str, pages: str) -> None:
    """Process a PowerPoint presentation.

    It first copies the input presentation to the output path.
    Then, for 'translate' and 'reverse-words' modes, text on selected slides
    within this copied presentation is modified in place.
    'translate' mode uses an on-disk cache.
    Page range can be specified using the --pages option.
    """
    cache_file_path = "translation_cache.json"

    click.echo(f"Copying '{input_path}' to '{output_path}' to preserve layout...")
    shutil.copy2(input_path, output_path)
    click.echo("File copy complete.")

    click.echo(f"Loading presentation for modification from: {output_path}")
    prs = Presentation(output_path)

    num_original_slides = len(prs.slides)
    if num_original_slides == 0:
        click.echo("Input presentation has no slides. Exiting.")
        return

    selected_pages_0_indexed = set()
    if pages:
        selected_pages_0_indexed = parse_page_range(pages, num_original_slides)
    else:
        selected_pages_0_indexed = set(range(num_original_slides))

    slides_to_process_objects = []
    original_page_indices_for_processing = []
    if num_original_slides > 0:
        for i, slide_obj in enumerate(prs.slides):  # i is original 0-indexed page
            if i in selected_pages_0_indexed:
                slides_to_process_objects.append(slide_obj)
                original_page_indices_for_processing.append(i)

    if not slides_to_process_objects:
        if pages:
            click.echo(
                f"Warning: The specified page range '{pages}' resulted in "
                f"no slides being selected from {num_original_slides} total slides. "
                "No text processing will occur."
            )
        else:
            click.echo(
                f"No slides selected for processing from {num_original_slides} "
                "total slides. No text processing will occur."
            )
        prs.save(output_path)
        click.echo(f"Presentation saved (no text modifications) to: {output_path}")
        if mode == "translate":
            # Ensure cache is loaded and saved even if no processing,
            # in case format changed or it needs to be created.
            translation_cache = load_cache(cache_file_path)
            # Pass empty pending updates if no processing occurred
            commit_pending_cache_updates(translation_cache, {}, cache_file_path)
        return

    if mode in {"translate", "reverse-words"} and slides_to_process_objects:
        click.echo(
            f"Preparing to process text on {len(slides_to_process_objects)} selected "
            f"slides (out of {num_original_slides} total) in the copied presentation..."
        )

    text_id_counter = 0

    if mode == "translate":
        translation_cache = load_cache(cache_file_path)  # Logging is in load_cache

        global_texts_for_llm_prompt: list[dict[str, Any]] = []
        all_processed_run_details: list[dict[str, Any]] = []
        pending_page_cache_updates: dict[str, list[dict[str, str]]] = {}

        if slides_to_process_objects:
            click.echo(
                f"Processing {len(slides_to_process_objects)} "
                "selected slide(s) for translation..."
            )
            for slide_idx, slide_to_extract in enumerate(slides_to_process_objects):
                original_page_0_indexed = original_page_indices_for_processing[
                    slide_idx
                ]
                current_page_num_1_indexed = original_page_0_indexed + 1
                current_page_texts_for_hash: list[str] = []
                current_page_run_info: list[dict[str, Any]] = []

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
                        for _row_idx, row in enumerate(shape.table.rows):
                            for _col_idx, cell in enumerate(row.cells):
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

                (
                    texts_for_llm_slide,
                    processed_runs_slide,
                    text_id_counter,  # Updated counter
                    page_requires_llm,
                ) = prepare_slide_for_translation(
                    current_page_run_info,
                    page_hash,
                    translation_cache,
                    text_id_counter,
                    EOL_MARKER,
                    current_page_num_1_indexed,
                )
                global_texts_for_llm_prompt.extend(texts_for_llm_slide)
                all_processed_run_details.extend(processed_runs_slide)

                if page_requires_llm and page_hash not in pending_page_cache_updates:
                    pending_page_cache_updates[page_hash] = []
                    # The actual list for pending_page_cache_updates[page_hash]
                    # will be populated by update_data_from_llm_response

        if not all_processed_run_details:
            click.echo(
                f"No text found to process on selected slides for mode '{mode}'."
            )
            prs.save(output_path)
            click.echo(
                f"Presentation saved without text modification in '{mode}' mode to: "
                f"{output_path}"
            )
            commit_pending_cache_updates(
                translation_cache, pending_page_cache_updates, cache_file_path
            )
            return

        click.echo(
            f"Processed {len(all_processed_run_details)} total text runs. "
            f"{len(global_texts_for_llm_prompt)} to translate via LLM."
        )

        if not global_texts_for_llm_prompt:
            click.echo("All text elements found in page caches. Skipping LLM prompt.")
        else:
            click.echo(
                f"Sending {len(global_texts_for_llm_prompt)} text elements to LLM "
                "for translation."
            )
            formatted_text_for_llm = "\n".join(
                [
                    f"{item['id']}:{item['text_to_send']}"
                    for item in global_texts_for_llm_prompt
                ]
            )
            prompt_text = (
                "You are an expert Finnish to English translator. "
                "Translate the following text segments accurately from Finnish to "
                "English. "
                "Each segment is prefixed with a unique ID (e.g., pg1_txt0, pg1_txt1). "
                "IMPORTANT: A sequence of text items (e.g., pg1_txt0, pg1_txt1, "
                "pg1_txt2) may represent a single continuous sentence that has been "
                "split due to formatting. Interpret and translate such sequences as a "
                "coherent whole sentence to maintain context and flow. "
                "The text for each ID might end with an EOL marker: '<'. "
                "Your response MUST consist ONLY of the translated segments, each "
                "prefixed with its original ID, "
                "and each on a new line. Maintain the exact ID and format. "
                "PRESERVE ALL LEADING AND TRAILING WHITESPACE from the original "
                "segment in your translation. "
                "If an EOL marker '<' was present at the end of the input segment, "
                "IT MUST be present at the end of your translated segment, including "
                "any whitespace before it.\n"
                "For example, if you receive:\n"
                "pg1_txt0: Tämä on pitkä \n"
                "pg1_txt1:lause, joka on \n"
                "pg1_txt2:jaettu.<\n"
                "pg1_txt3:    Toinen lause.   <\n"
                "pg2_txt0: Yksittäinen.\n"
                "You MUST return:\n"
                "pg1_txt0: This is a long \n"
                "pg1_txt1:sentence that has been \n"
                "pg1_txt2:split.<\n"
                "pg1_txt3:    Another sentence.   <\n"
                "pg2_txt0: Standalone.\n\n"
                "Do not add any extra explanations, apologies, or "
                "introductory/concluding remarks. "
                "Only provide the ID followed by the translated text for each item.\n\n"
                "Texts to translate:\n"
            )
            click.echo("--- PROMPT TO LLM ---")
            click.echo("System/Instruction Prompt:")
            click.echo(prompt_text)  # Keep this for debugging/visibility
            click.echo("Data Fragments (for LLM):")
            click.echo(formatted_text_for_llm)  # Keep this
            click.echo("--- END OF PROMPT ---")

            model_instance = llm.get_model()
            response = model_instance.prompt(
                prompt_text, fragments=[formatted_text_for_llm]
            )
            translated_text_response_str = response.text()

            click.echo("--- RESPONSE FROM LLM ---")
            click.echo(translated_text_response_str)  # Keep this
            click.echo("--- END OF RESPONSE ---")
            click.echo("Received translation from LLM.")

            update_data_from_llm_response(
                translated_text_response_str.splitlines(),
                global_texts_for_llm_prompt,
                all_processed_run_details,
                pending_page_cache_updates,
                EOL_MARKER,
            )

        click.echo(
            "Replacing text with translations on slides in the copied presentation..."
        )
        for item in all_processed_run_details:
            if item["final_translation"] is not None:
                item["run_object"].text = item["final_translation"]
            elif not item[
                "from_cache"
            ]:  # Only warn if it was supposed to be LLM translated
                click.echo(
                    f"Warning: No translation found for run with original text "
                    f"'{item['original_text'][:30]}...' "
                    f"(LLM ID: {item.get('llm_id', 'N/A')}). Leaving original.",
                    err=True,
                )

        commit_pending_cache_updates(
            translation_cache, pending_page_cache_updates, cache_file_path
        )

    elif mode == "reverse-words":
        text_elements_for_reverse = []
        if slides_to_process_objects:
            click.echo(
                f"Extracting text from {len(slides_to_process_objects)} slides in the "
                f"copied presentation for mode '{mode}'..."
            )
            for slide_idx_reverse, slide_to_extract in enumerate(
                slides_to_process_objects
            ):
                original_page_0_indexed = original_page_indices_for_processing[
                    slide_idx_reverse
                ]
                current_page_num_1_indexed = original_page_0_indexed + 1
                for shape in slide_to_extract.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text
                                if original_text:
                                    text_with_eol = original_text + EOL_MARKER
                                    text_id = (
                                        f"pg{current_page_num_1_indexed}_"
                                        f"txt{text_id_counter}"
                                    )
                                    text_elements_for_reverse.append(
                                        {
                                            "id": text_id,
                                            "text": text_with_eol,
                                            "run_object": run,
                                        }
                                    )
                                    text_id_counter += 1
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        original_text = run.text
                                        if original_text:
                                            text_with_eol = original_text + EOL_MARKER
                                            text_id = (
                                                f"pg{current_page_num_1_indexed}_"
                                                f"txt{text_id_counter}"
                                            )
                                            text_elements_for_reverse.append(
                                                {
                                                    "id": text_id,
                                                    "text": text_with_eol,
                                                    "run_object": run,
                                                }
                                            )
                                            text_id_counter += 1

        if not text_elements_for_reverse:
            click.echo(
                "No text found to process on slides in the copied presentation for "
                f"mode '{mode}'."
            )
            # No cache involvement for reverse-words, so just save and return
        else:
            click.echo(
                f"Found {len(text_elements_for_reverse)} text elements to process for "
                f"mode '{mode}'."
            )
            click.echo("Applying word reversal on slides in the copied presentation...")
            for item in text_elements_for_reverse:
                reversed_text_with_eol = reverse_individual_words(item["text"])
                final_reversed_text = reversed_text_with_eol.removesuffix(EOL_MARKER)
                item["run_object"].text = final_reversed_text
            click.echo(
                "Text replaced with reversed-word text on slides in the copied "
                "presentation."
            )

    prs.save(output_path)
    click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")


if __name__ == "__main__":
    main()
