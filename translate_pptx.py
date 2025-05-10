import copy  # For deepcopying elements
import json
import os
import shutil  # For file copying

import click
import llm  # Simon Willison's LLM library
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

# import xml.etree.ElementTree as ET # Not strictly needed if using copy.deepcopy for lxml


def load_cache(cache_file_path):
    """Loads the translation cache from a JSON file."""
    if os.path.exists(cache_file_path):
        try:
            with open(cache_file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            click.echo(
                f"Warning: Could not load cache file {cache_file_path}. Error: {e}. Starting with an empty cache.",
                err=True,
            )
    return {}


def save_cache(cache_data, cache_file_path):
    """Saves the translation cache to a JSON file."""
    try:
        with open(cache_file_path, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, indent=4, ensure_ascii=False)
    except IOError as e:
        click.echo(
            f"Warning: Could not save cache file {cache_file_path}. Error: {e}",
            err=True,
        )


def _duplicate_slide_to_end(pres, source_slide):
    """
    Duplicates the source_slide and appends the copy to the end of the presentation.
    Returns the newly created (duplicated) slide.
    """
    target_layout = source_slide.slide_layout
    new_slide = pres.slides.add_slide(target_layout)

    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    if source_slide.background.fill.type:
        new_slide.background.fill.solid()
        try:
            src_rgb = source_slide.background.fill.fore_color.rgb
            new_slide.background.fill.fore_color.rgb = src_rgb
        except (TypeError, AttributeError):
            pass
    return new_slide


def reverse_individual_words(text_string):
    """Reverses each word in a space-separated string."""
    words = text_string.split(" ")
    reversed_words = [word[::-1] for word in words]
    return " ".join(reversed_words)


@click.command()
@click.option(
    "--mode",
    type=click.Choice(
        ["translate", "duplicate-only", "reverse-words"], case_sensitive=False
    ),
    default="translate",
    show_default=True,
    help="Operation mode for the script.",
)
@click.argument("input_path", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_path", type=click.Path(dir_okay=False))
def main(input_path, output_path, mode):
    """
    Processes a PowerPoint presentation.
    In 'translate', 'reverse-words', and 'duplicate-only' modes, it first copies the input
    presentation to the output path. Then, it duplicates each original slide and appends
    the copy to the end of the presentation. For 'translate' and 'reverse-words' modes,
    text on these duplicated slides is then modified. 'translate' mode uses an on-disk cache.

    Output slide order: Original_Slide_1, ..., Original_Slide_N, Modified_Slide_1, ..., Modified_Slide_N
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

    slides_for_text_extraction = []

    click.echo(
        f"Duplicating {num_original_slides} original slide(s) to the end of the presentation..."
    )
    for i in range(num_original_slides):
        original_slide = prs.slides[i]
        click.echo(
            f"  Duplicating original slide {i + 1} ('{original_slide.slide_layout.name}')..."
        )
        duplicated_slide = _duplicate_slide_to_end(prs, original_slide)
        if mode == "translate" or mode == "reverse-words":
            slides_for_text_extraction.append(duplicated_slide)
        elif mode == "duplicate-only":
            slides_for_text_extraction.append(
                duplicated_slide
            )  # Still add for consistent reporting
        click.echo(f"    Slide {i + 1} duplicated. Total slides now: {len(prs.slides)}")

    if mode == "duplicate-only":
        click.echo(
            "Mode 'duplicate-only': Text processing skipped. All slides (originals and their duplicates) are saved."
        )
        prs.save(output_path)
        click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")
        return

    # Text processing for 'translate' and 'reverse-words' modes
    text_id_counter = 0  # Shared counter for text element IDs

    if mode == "translate":
        click.echo(f"Loading translation cache from: {cache_file_path}")
        translation_cache = load_cache(cache_file_path)

        texts_for_llm_prompt = []
        all_text_elements_with_status = []

        if slides_for_text_extraction:
            click.echo(
                f"Extracting text from {len(slides_for_text_extraction)} duplicated slides for mode '{mode}' (with cache checking)..."
            )
            for slide_to_extract in slides_for_text_extraction:
                for shape in slide_to_extract.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text.strip()
                                if original_text:
                                    text_id = f"text_{text_id_counter}"
                                    text_id_counter += 1

                                    if original_text in translation_cache:
                                        click.echo(
                                            f"  Cache hit for ID {text_id}: '{original_text[:30].replace(chr(10), ' ').replace(chr(13), ' ')}...'"
                                        )
                                        all_text_elements_with_status.append(
                                            {
                                                "id": text_id,
                                                "text": original_text,
                                                "run_object": run,
                                                "translation": translation_cache[
                                                    original_text
                                                ],
                                                "from_cache": True,
                                            }
                                        )
                                    else:
                                        click.echo(
                                            f"  Cache miss for ID {text_id}: '{original_text[:30].replace(chr(10), ' ').replace(chr(13), ' ')}...' (will send to LLM)"
                                        )
                                        texts_for_llm_prompt.append(
                                            {"id": text_id, "text": original_text}
                                        )
                                        all_text_elements_with_status.append(
                                            {
                                                "id": text_id,
                                                "text": original_text,
                                                "run_object": run,
                                                "translation": None,
                                                "from_cache": False,
                                            }
                                        )
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        original_text = run.text.strip()
                                        if original_text:
                                            text_id = f"text_{text_id_counter}"
                                            text_id_counter += 1
                                            if original_text in translation_cache:
                                                click.echo(
                                                    f"  Cache hit for ID {text_id} (table): '{original_text[:30].replace(chr(10), ' ').replace(chr(13), ' ')}...'"
                                                )
                                                all_text_elements_with_status.append(
                                                    {
                                                        "id": text_id,
                                                        "text": original_text,
                                                        "run_object": run,
                                                        "translation": translation_cache[
                                                            original_text
                                                        ],
                                                        "from_cache": True,
                                                    }
                                                )
                                            else:
                                                click.echo(
                                                    f"  Cache miss for ID {text_id} (table): '{original_text[:30].replace(chr(10), ' ').replace(chr(13), ' ')}...' (will send to LLM)"
                                                )
                                                texts_for_llm_prompt.append(
                                                    {
                                                        "id": text_id,
                                                        "text": original_text,
                                                    }
                                                )
                                                all_text_elements_with_status.append(
                                                    {
                                                        "id": text_id,
                                                        "text": original_text,
                                                        "run_object": run,
                                                        "translation": None,
                                                        "from_cache": False,
                                                    }
                                                )

        if not all_text_elements_with_status:
            click.echo(
                f"No text found to process on duplicated slides for mode '{mode}'."
            )
            prs.save(output_path)
            click.echo(
                f"Presentation saved without text modification in '{mode}' mode to: {output_path}"
            )
            click.echo(f"Saving cache (even if empty/unchanged) to: {cache_file_path}")
            save_cache(
                translation_cache, cache_file_path
            )  # Save cache even if no text elements
            return

        click.echo(
            f"Found {len(all_text_elements_with_status)} total text elements. {len(texts_for_llm_prompt)} to translate via LLM."
        )

        if not texts_for_llm_prompt:
            click.echo("All text elements found in cache. Skipping LLM prompt.")
        else:
            click.echo(
                f"Sending {len(texts_for_llm_prompt)} text elements to LLM for translation."
            )
            formatted_text_for_llm = "\n".join(
                [f"{item['id']}: {item['text']}" for item in texts_for_llm_prompt]
            )
            prompt_text = (
                "You are an expert Finnish to English translator. "
                "Translate the following text segments accurately from Finnish to English. "
                "Each segment is prefixed with a unique ID (e.g., text_0, text_1). "
                "Your response MUST consist ONLY of the translated segments, each prefixed with its original ID, "
                "and each on a new line. Maintain the exact ID and format.\n"
                "For example, if you receive:\n"
                "text_0: Hei maailma\n"
                "text_1: Kiitos paljon\n"
                "You MUST return:\n"
                "text_0: Hello world\n"
                "text_1: Thank you very much\n\n"
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
                        parsed_text_id, llm_translation = (
                            parts[0].strip(),
                            parts[1].strip(),
                        )
                        original_text_for_cache_key = next(
                            (
                                item["text"]
                                for item in texts_for_llm_prompt
                                if item["id"] == parsed_text_id
                            ),
                            None,
                        )
                        if original_text_for_cache_key:
                            translation_cache[original_text_for_cache_key] = (
                                llm_translation
                            )
                            for elem in all_text_elements_with_status:
                                if elem["id"] == parsed_text_id:
                                    elem["translation"] = llm_translation
                                    elem["from_cache"] = (
                                        False  # Mark as freshly translated
                                    )
                                    break
                        else:
                            click.echo(
                                f"Warning: Could not find original text for ID {parsed_text_id} from LLM response to update cache/status list.",
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

        click.echo("Replacing text with translations on duplicated slides...")
        for item in all_text_elements_with_status:
            if item["translation"]:
                item["run_object"].text = item["translation"]

        click.echo(f"Saving updated translation cache to: {cache_file_path}")
        save_cache(translation_cache, cache_file_path)

    elif mode == "reverse-words":
        # Original logic for reverse-words, uses a simple text_elements list
        text_elements_for_reverse = []
        if slides_for_text_extraction:
            click.echo(
                f"Extracting text from {len(slides_for_text_extraction)} duplicated slides for mode '{mode}'..."
            )
            for slide_to_extract in slides_for_text_extraction:
                for shape in slide_to_extract.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text.strip()
                                if original_text:
                                    text_id = (
                                        f"text_{text_id_counter}"  # Uses global counter
                                    )
                                    text_elements_for_reverse.append(
                                        {
                                            "id": text_id,
                                            "text": original_text,
                                            "run_object": run,
                                        }
                                    )
                                    text_id_counter += 1
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        original_text = run.text.strip()
                                        if original_text:
                                            text_id = f"text_{text_id_counter}"  # Uses global counter
                                            text_elements_for_reverse.append(
                                                {
                                                    "id": text_id,
                                                    "text": original_text,
                                                    "run_object": run,
                                                }
                                            )
                                            text_id_counter += 1

        if not text_elements_for_reverse:
            click.echo(
                f"No text found to process on duplicated slides for mode '{mode}'."
            )
            prs.save(output_path)
            click.echo(
                f"Presentation saved without text modification in '{mode}' mode to: {output_path}"
            )
            return

        click.echo(
            f"Found {len(text_elements_for_reverse)} text elements to process for mode '{mode}'."
        )
        click.echo("Applying word reversal on duplicated slides...")
        for item in text_elements_for_reverse:
            reversed_text = reverse_individual_words(item["text"])
            item["run_object"].text = reversed_text
        click.echo("Text replaced with reversed-word text on duplicated slides.")

    prs.save(output_path)
    click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")


if __name__ == "__main__":
    main()
