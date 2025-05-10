import click
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import llm # Simon Willison's LLM library
import xml.etree.ElementTree as ET # For slide copying

# --- Slide Copying Logic (Adapted from https://stackoverflow.com/a/73954830/15770) ---
def _get_blank_slide_layout(pres):
    """Return a blank slide layout from pres"""
    layout_items = [layout for layout in pres.slide_layouts if layout.name == 'Blank']
    if not layout_items:
        # Fallback if "Blank" layout is not found by name (e.g. different language versions)
        # In many default templates, layout index 5 or 6 is often Blank.
        # This is a heuristic and might need adjustment for specific templates.
        if len(pres.slide_layouts) > 5:
            return pres.slide_layouts[5] # Common index for Blank
        else: # If very few layouts, pick the last one as a desperate fallback
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

    # Create a new slide in the target presentation, using the same layout as the source slide.
    # This requires the layout to exist in the target presentation's slide master.
    # For simplicity here, we'll try to use the source slide's layout.
    # A more robust solution would involve copying slide masters and layouts first if they don't exist.
    try:
        slide_layout = new_presentation.slide_layouts.get_by_name(src_slide.slide_layout.name)
        if slide_layout is None: # If layout not found by name, use a default blank one
            slide_layout = _get_blank_slide_layout(new_presentation)
    except Exception: # Broad exception if get_by_name fails or other issues
        slide_layout = _get_blank_slide_layout(new_presentation)
        
    new_slide = new_presentation.slides.add_slide(slide_layout)

    # Copy shapes from source slide to new slide
    for shape in src_slide.shapes:
        el = shape.element
        new_el = ET.fromstring(ET.tostring(el)) # Deep copy of element
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        # Note: This is a low-level way to add shapes. Relationships (like images, charts)
        # might need explicit handling (relIds need to be copied and potentially updated).
        # The original SO answer has more comprehensive handling for this.
        # For text, this should generally work.

    # Copy background
    if src_slide.has_notes_slide: # Check if notes_slide exists
        # notes_slide background copy is more complex and omitted for this version
        pass

    if src_slide.background.fill.type:
        new_slide.background.fill.type = src_slide.background.fill.type
        if src_slide.background.fill.type == MSO_SHAPE_TYPE.PICTURE:
             # Picture fill copy is complex (requires image part) and omitted
            pass
        elif src_slide.background.fill.fore_color:
            new_slide.background.fill.fore_color.rgb = src_slide.background.fill.fore_color.rgb


    # Copy slide master background (if not overridden by slide background)
    # This is also complex and omitted for brevity in this adaptation.
    # The full SO answer attempts to handle this by looking at `slideLayout.slideMaster.background`.

    return new_slide
# --- End Slide Copying Logic ---


@click.command()
@click.argument('input_path', type=click.Path(exists=True, dir_okay=False))
@click.argument('output_path', type=click.Path(dir_okay=False))
def main(input_path, output_path):
    """
    Translates text in a PowerPoint presentation from Finnish to English.
    Duplicates each slide, translates the duplicated slide, and saves to a new file.
    """
    click.echo(f"Loading presentation from: {input_path}")
    prs = Presentation(input_path)
    new_prs = Presentation() # Create a new presentation for the output

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
        click.echo(f"  Copying original slide {i+1}...")
        copy_slide(prs, new_prs, i) # Copy original slide (first copy)
        
        click.echo(f"  Copying slide {i+1} for translation (second copy)...")
        translated_slide_candidate = copy_slide(prs, new_prs, i) # Copy slide again for translation
        slides_for_text_extraction.append(translated_slide_candidate)

    click.echo("Extracting text from duplicated slides...")
    for slide_to_extract in slides_for_text_extraction:
        for shape in slide_to_extract.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        original_text = run.text.strip()
                        if original_text: # Only process non-empty runs
                            text_id = f"text_{text_id_counter}"
                            text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})
                            text_id_counter += 1
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text.strip()
                                if original_text: # Only process non-empty runs
                                    text_id = f"text_{text_id_counter}"
                                    text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})
                                    text_id_counter += 1
    
    if not text_elements:
        click.echo("No text found to translate.")
        new_prs.save(output_path)
        click.echo(f"Presentation saved without translation to: {output_path}")
        return

    click.echo(f"Found {len(text_elements)} text elements to translate.")

    formatted_text = "\n".join([f"{item['id']}: {item['text']}" for item in text_elements])
    
    prompt = (
        "Translate the following text segments from Finnish to English. "
        "Each segment is prefixed with an ID. "
        "Return the translations in the exact same format, preserving the IDs, with each translation on a new line.\n"
        "For example, if you receive:\n"
        "ID1: Hei maailma\n"
        "ID2: Kiitos\n"
        "You should return:\n"
        "ID1: Hello world\n"
        "ID2: Thank you\n\n"
        "Here are the texts to translate:\n"
    )

    click.echo("Sending text to LLM for translation...")
    # Using a placeholder for filename, as from_text doesn't strictly need it.
    attachment = llm.Attachment.from_text(formatted_text, filename="texts_to_translate.txt")
    model = llm.get_model() # Get default model
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
        translated_text = id_to_translation.get(item['id'], item['text']) # Fallback to original
        item['run_object'].text = translated_text
        
    new_prs.save(output_path)
    click.echo(f"Translated presentation saved to: {output_path}")

if __name__ == '__main__':
    main()