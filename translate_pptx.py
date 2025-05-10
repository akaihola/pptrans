import copy  # For deepcopying elements
import shutil  # For file copying

import click
import llm  # Simon Willison's LLM library
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

# import xml.etree.ElementTree as ET # Not strictly needed if using copy.deepcopy for lxml


def _duplicate_slide_to_end(pres, source_slide):
    """
    Duplicates the source_slide and appends the copy to the end of the presentation.
    Returns the newly created (duplicated) slide.
    """
    # Use the source slide's layout directly, as we are in the same presentation context
    target_layout = source_slide.slide_layout
    new_slide = pres.slides.add_slide(target_layout)  # Adds to the end

    # Copy shapes
    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
        # Note: This low-level shape copy is effective for many shape types.
        # However, for shapes with complex relationships (e.g., linked charts,
        # embedded media not stored directly in shape XML), their `rId`s
        # (relationship IDs) would need to be managed by copying the related parts
        # from the source presentation package to the target and updating these rIds.
        # Since we are copying the entire file first, these relationships should be intact
        # for the original slides, and deepcopy should handle most shape-internal data.

    # Copy background (simplified)
    # A full background copy would also consider slide master and layout backgrounds.
    # However, since we are duplicating within the same presentation,
    # the new slide should inherit correctly from its layout/master.
    # This explicit copy might still be useful for direct slide-level background overrides.
    if source_slide.background.fill.type:  # Check if there's a fill type set
        new_slide.background.fill.solid()  # Ensure fill is solid before setting color
        try:
            src_rgb = source_slide.background.fill.fore_color.rgb
            new_slide.background.fill.fore_color.rgb = src_rgb
        except (TypeError, AttributeError):
            # This can happen if the fill isn't a simple solid color or color is inherited.
            # Silently pass for now, as layout inheritance should handle most cases.
            # click.echo(f"DEBUG: Could not copy direct background color for duplicated slide from slide {pres.slides.index(source_slide) + 1}")
            pass

    # Notes slide copying (omitted for simplicity as per PLAN.md scope)
    # if source_slide.has_notes_slide:
    #     ts = source_slide.notes_slide
    #     new_notes_slide = new_slide.notes_slide
    #     # This part needs careful implementation if notes are complex
    #     if ts.notes_text_frame:
    #          new_notes_slide.notes_text_frame.text = ts.notes_text_frame.text

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
    text on these duplicated slides is then modified.

    Output slide order: Original_Slide_1, ..., Original_Slide_N, Modified_Slide_1, ..., Modified_Slide_N
    """
    click.echo(f"Copying '{input_path}' to '{output_path}' to preserve layout...")
    try:
        shutil.copy2(input_path, output_path)
    except Exception as e:
        click.echo(f"Error copying file: {e}", err=True)
        return
    click.echo("File copy complete.")

    click.echo(f"Loading presentation for modification from: {output_path}")
    prs = Presentation(output_path)  # Open the copied file

    num_original_slides = len(prs.slides)
    if num_original_slides == 0:
        click.echo("Input presentation has no slides. Exiting.")
        # prs.save(output_path) # Save the (empty) copy
        return

    slides_for_text_extraction = []  # Will hold the duplicated slides for modification

    click.echo(
        f"Duplicating {num_original_slides} original slide(s) to the end of the presentation..."
    )
    for i in range(num_original_slides):
        original_slide = prs.slides[i]  # This is a slide from the *copied* presentation
        click.echo(
            f"  Duplicating original slide {i + 1} ('{original_slide.slide_layout.name}')..."
        )

        # Duplicate the slide and append it to the end
        duplicated_slide = _duplicate_slide_to_end(prs, original_slide)

        if mode == "translate" or mode == "reverse-words":
            slides_for_text_extraction.append(duplicated_slide)
        elif mode == "duplicate-only":
            # For duplicate-only, we still add it to this list,
            # but it won't be processed for text.
            # This keeps the logic consistent for reporting.
            slides_for_text_extraction.append(duplicated_slide)

        click.echo(f"    Slide {i + 1} duplicated. Total slides now: {len(prs.slides)}")

    if mode == "duplicate-only":
        click.echo(
            "Mode 'duplicate-only': Text processing skipped. All slides (originals and their duplicates) are saved."
        )
        prs.save(output_path)
        click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")
        return

    # Text processing for 'translate' and 'reverse-words' modes
    # This operates on the slides in 'slides_for_text_extraction'
    text_elements = []
    text_id_counter = 0

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
                            if original_text:  # Only process non-empty runs
                                text_id = f"text_{text_id_counter}"
                                text_elements.append(
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
                                    if original_text:  # Only process non-empty runs
                                        text_id = f"text_{text_id_counter}"
                                        text_elements.append(
                                            {
                                                "id": text_id,
                                                "text": original_text,
                                                "run_object": run,
                                            }
                                        )
                                        text_id_counter += 1

    if not text_elements:
        click.echo(f"No text found to process on duplicated slides for mode '{mode}'.")
        prs.save(output_path)
        click.echo(
            f"Presentation saved without text modification in '{mode}' mode to: {output_path}"
        )
        return

    click.echo(
        f"Found {len(text_elements)} text elements to process for mode '{mode}'."
    )

    if mode == "translate":
        formatted_text = "\n".join(
            [f"{item['id']}: {item['text']}" for item in text_elements]
        )
        prompt = (
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
        click.echo("Sending text to LLM for translation...")
        model = llm.get_model()  # Assumes llm is configured
        response = model.prompt(prompt, fragments=[formatted_text])
        translated_text_response = response.text()
        click.echo("Received translation from LLM.")

        id_to_modified_text = {}
        for line in translated_text_response.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parts = line.split(":", 1)
                if len(parts) == 2:
                    text_id = parts[0].strip()
                    translation = parts[1].strip()
                    id_to_modified_text[text_id] = translation
                else:
                    click.echo(f"Warning: Could not parse translation line: {line}")
            except Exception as e:
                click.echo(f"Warning: Error parsing translation line '{line}': {e}")

        click.echo("Replacing text with translations on duplicated slides...")
        for item in text_elements:
            modified_text = id_to_modified_text.get(
                item["id"], item["text"]
            )  # Fallback to original
            item["run_object"].text = modified_text

    elif mode == "reverse-words":
        click.echo("Applying word reversal on duplicated slides...")
        id_to_modified_text = {}  # Not strictly needed here but keeps structure similar
        for item in text_elements:
            reversed_text = reverse_individual_words(item["text"])
            item["run_object"].text = reversed_text  # Apply directly

        click.echo("Text replaced with reversed-word text on duplicated slides.")

    prs.save(output_path)
    click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")


if __name__ == "__main__":
    main()
