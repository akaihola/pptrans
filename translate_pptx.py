import copy  # For deepcopying elements

import click
import llm  # Simon Willison's LLM library
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

# import xml.etree.ElementTree as ET # Not strictly needed if using copy.deepcopy for lxml


# --- Slide Copying Logic (Adapted from https://stackoverflow.com/a/73954830/15770) ---
def _get_blank_slide_layout(pres):
    """Return a blank slide layout from pres"""
    layout_items = [layout for layout in pres.slide_layouts if layout.name == "Blank"]
    if not layout_items:
        # Fallback if "Blank" layout is not found by name (e.g. different language versions)
        # In many default templates, layout index 5 or 6 is often Blank.
        # This is a heuristic and might need adjustment for specific templates.
        if len(pres.slide_layouts) > 5:
            return pres.slide_layouts[5]  # Common index for Blank
        else:  # If very few layouts, pick the last one as a desperate fallback
            return pres.slide_layouts[-1]
    return layout_items[0]


def copy_slide(src_presentation, new_presentation, slide_to_copy_index):
    """
    Copies a slide from src_presentation to new_presentation.
    Adapted from https://stackoverflow.com/a/73954830/15770 user "shredactivate"
    This version is simplified as it copies to a new presentation object `new_presentation`
    which might not have all source masters. The SO answer is more robust for
    copying between arbitrary existing files.
    """
    src_slide = src_presentation.slides[slide_to_copy_index]

    # Create a new slide in the target presentation.
    # Attempt to use the same layout as the source slide.
    # If the named layout doesn't exist in new_presentation's master, this will fail or use a default.
    # A truly robust solution involves copying slide masters and layouts first.
    # For this script, we assume new_presentation starts blank and we add to it.
    # python-pptx creates a default master when Presentation() is called.

    # Try to find layout by name (less reliable across different base templates)
    # Fallback to using the source slide's layout object directly if it's from the same master pool,
    # or a default blank layout.
    target_layout = None
    if new_presentation.slide_masters:  # Check if new_prs has any masters
        try:
            # This assumes new_presentation has been prepared with necessary layouts
            # or that default layouts are sufficient.
            target_layout = new_presentation.slide_layouts.get_by_name(
                src_slide.slide_layout.name
            )
        except KeyError:  # get_by_name raises KeyError if not found
            pass  # target_layout remains None

    if target_layout is None:
        # If specific layout not found, or new_presentation has no masters yet (should not happen with Presentation())
        # fall back to a blank layout from new_presentation.
        # If new_presentation is truly empty (e.g. no default master), this would also need care.
        # However, Presentation() initializes with a default master & layouts.
        target_layout = _get_blank_slide_layout(new_presentation)

    new_slide = new_presentation.slides.add_slide(target_layout)

    # Copy shapes from source slide to new slide
    for shape in src_slide.shapes:
        el = shape.element
        # new_el = ET.fromstring(ET.tostring(el)) # Deep copy using standard library XML
        new_el = copy.deepcopy(
            el
        )  # Deep copy using lxml's capabilities via python-pptx's element objects
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
        # Note: This low-level shape copy is effective for many shape types including text boxes.
        # However, for shapes with complex relationships (e.g., linked charts, embedded media not stored directly in shape XML),
        # their `rId`s (relationship IDs) would need to be managed by copying the related parts
        # from the source presentation package to the target and updating these rIds.
        # The referenced StackOverflow answer provides a much more comprehensive solution for this.
        # This script prioritizes text translation and assumes simpler content or that visual fidelity
        # issues from missing linked parts are acceptable for its primary goal.

    # Copy background (simplified)
    # A full background copy would also consider slide master and layout backgrounds.
    print(
        f"DEBUG: Slide {slide_to_copy_index + 1} background fill object: {src_slide.background.fill}"
    )
    print(
        f"DEBUG: Slide {slide_to_copy_index + 1} background fill type: {src_slide.background.fill.type}"
    )
    if src_slide.background.fill.type:  # Check if there's a fill type set
        new_slide.background.fill.solid()  # Ensure fill is solid before setting color
        # This is a very basic color copy. More complex fills (gradient, picture) are not fully handled here.
        try:
            # Attempt to get the source foreground color's RGB value
            # This line might raise TypeError if fore_color is not directly available (e.g., _NoFill, or inherited)
            # or AttributeError if .rgb is not present.
            src_rgb = src_slide.background.fill.fore_color.rgb
            # If successful, apply it to the new slide's foreground color
            new_slide.background.fill.fore_color.rgb = src_rgb
        except TypeError:
            # This typically occurs if src_slide.background.fill.fore_color itself raises
            # the "fill type ... has no foreground color" error.
            print(
                f"DEBUG: Slide {slide_to_copy_index + 1} - TypeError: No direct foreground color to copy or not an RGBColor object."
            )
        except AttributeError:
            # This could occur if fore_color exists but is None, or doesn't have an .rgb attribute.
            print(
                f"DEBUG: Slide {slide_to_copy_index + 1} - AttributeError: Foreground color attribute missing or not an RGBColor object."
            )
            # elif hasattr(src_slide.background.fill.fore_color, 'theme_color'):
            #     new_slide.background.fill.fore_color.theme_color = src_slide.background.fill.fore_color.theme_color

    # Notes slide copying is omitted for simplicity as per PLAN.md scope.
    return new_slide


# --- End Slide Copying Logic ---


@click.command()
@click.argument("input_path", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_path", type=click.Path(dir_okay=False))
def main(input_path, output_path):
    """
    Translates text in a PowerPoint presentation from Finnish to English.
    Duplicates each slide, translates the duplicated slide, and saves to a new file.
    """
    click.echo(f"Loading presentation from: {input_path}")
    prs = Presentation(input_path)
    new_prs = Presentation()  # Create a new presentation for the output

    # Make sure new_prs has the same slide masters and layouts as prs
    # This is a simplified approach; a full copy would involve iterating masters.
    # For now, we rely on the copy_slide function to use existing or default layouts.
    # To be more robust, one would copy slide masters from prs to new_prs first.
    # Example: (Conceptual, python-pptx doesn't directly support master copy)
    # for master in prs.slide_masters:
    #     # logic to copy master to new_prs

    text_elements = []
    text_id_counter = 0

    slides_for_text_extraction = []

    click.echo("Processing slides...")
    for i, slide in enumerate(prs.slides):
        click.echo(f"  Copying original slide {i + 1}...")
        copy_slide(prs, new_prs, i)  # Copy original slide (first copy)

        click.echo(f"  Copying slide {i + 1} for translation (second copy)...")
        translated_slide_candidate = copy_slide(
            prs, new_prs, i
        )  # Copy slide again for translation
        slides_for_text_extraction.append(translated_slide_candidate)

    click.echo("Extracting text from duplicated slides...")
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
        click.echo("No text found to translate.")
        new_prs.save(output_path)
        click.echo(f"Presentation saved without translation to: {output_path}")
        return

    click.echo(f"Found {len(text_elements)} text elements to translate.")

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
    # Using a placeholder for filename, as from_text doesn't strictly need it.
    attachment = llm.Attachment.from_text(
        formatted_text, filename="texts_to_translate.txt"
    )
    model = llm.get_model()  # Get default model
    response = model.prompt(prompt, attachments=[attachment])

    translated_text_response = response.text()
    click.echo("Received translation from LLM.")

    id_to_translation = {}
    for line in translated_text_response.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            parts = line.split(":", 1)
            if len(parts) == 2:
                text_id = parts[0].strip()
                translation = parts[1].strip()
                id_to_translation[text_id] = translation
            else:
                click.echo(f"Warning: Could not parse translation line: {line}")
        except Exception as e:
            click.echo(f"Warning: Error parsing translation line '{line}': {e}")

    click.echo("Replacing text with translations...")
    for item in text_elements:
        translated_text = id_to_translation.get(
            item["id"], item["text"]
        )  # Fallback to original
        item["run_object"].text = translated_text

    new_prs.save(output_path)
    click.echo(f"Translated presentation saved to: {output_path}")


if __name__ == "__main__":
    main()
