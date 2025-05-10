# Goal

Create a Python script which
- reads a PowerPoint `.pptx` file
- duplicates each page: `[p1, p2, p3]` becomes `[p1, p1copy, p2, p2copy, p3, p3copy]`
- translates all text elements on each `p?copy` page from Finnish to English
- saves the presentation as a new `.pptx` file with a different name

# Details

The script should perform the following steps:

1.  **Parse Command-Line Arguments:**
    *   Use the `click` package to handle command-line arguments.
    *   Define arguments for the input `.pptx` file path (required) and the output `.pptx` file path (required).
2.  **Open Presentation:**
    *   Import the `Presentation` class from `pptx`.
    *   Load the input presentation using `prs = Presentation(input_path)`.
3.  **Iterate and Duplicate Slides:**
    *   Create a new presentation object to build the modified presentation: `new_prs = Presentation()`.
    *   Iterate through the slides of the original presentation: `for slide in prs.slides:`.
    *   For each original slide, copy and append it to the `new_prs` **twice**.
    *   The latter copy is `new_slide`.
    *   For an example solution of how to copy slides, see https://stackoverflow.com/a/73954830/15770
4.  **Extract Text and Map IDs:**
    *   Initialize an empty list to store text elements: `text_elements = []`.
    *   Initialize a counter for unique IDs: `text_id_counter = 0`.
    *   Iterate through the *duplicated* slides in `new_prs` (even pages, i.e. each `new_slide` from step 3).
    *   For each duplicated slide, iterate through its shapes: `for shape in duplicated_slide.shapes:`.
    *   Check if the shape has text: `if shape.has_text_frame:`.
        *   If it has a text frame, iterate through paragraphs: `for paragraph in shape.text_frame.paragraphs:`.
            *   Iterate through runs: `for run in paragraph.runs:`.
                *   Extract text: `text = run.text`.
                *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                *   Store the data and references: `text_elements.append({'id': text_id, 'text': text, 'shape': shape, 'paragraph': paragraph, 'run': run})`.
                *   Increment counter: `text_id_counter += 1`.
    *   Check if the shape is a table: `if shape.has_table:`.
        *   If it's a table, iterate through rows and columns: `for row in shape.table.rows: for cell in row.cells:`.
            *   Iterate through paragraphs in the cell's text frame: `for paragraph in cell.text_frame.paragraphs:`.
                *   Iterate through runs: `for run in paragraph.runs:`.
                    *   Extract text: `text = run.text`.
                    *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                    *   Store the data and references: `text_elements.append({'id': text_id, 'text': text, 'cell': cell, 'paragraph': paragraph, 'run': run})`.
                    *   Increment counter: `text_id_counter += 1`.
    *   Consider other shape types that might contain text (e.g., chart data labels, though this is more complex).
5.  **Prepare Data for Language Model:**
    *   Format the `text_elements` list into a string. A simple format like `ID: Original Text\n` for each element is suitable for an ASCII attachment.
    *   `formatted_text = "\n".join([f"{item['id']}: {item['text']}" for item in text_elements])`.
6.  **Invoke Language Model:**
    *   Import the `llm` library.
    *   Define the prompt. The prompt should clearly instruct the model to translate the text provided in the attachment from Finnish to English and return the translations in the format `ID: Translated Text\n`, preserving the original IDs.
    *   Call the language model using `llm.get_model().prompt(prompt, attachments=[llm.Attachment(...)])`. Capture the output.
7.  **Process Language Model Response:**
    *   Initialize an empty dictionary for translations: `id_to_translation = {}`.
    *   Parse the model's output string line by line.
    *   For each line, split it into ID and translated text based on the expected format (e.g., splitting at the first `: `).
    *   Store the translation in the dictionary: `id_to_translation[id] = translated_text`. Handle potential parsing errors or unexpected output formats.
8.  **Replace Text on Duplicated Slides:**
    *   Iterate through the `text_elements` list created in step 4.
    *   For each element, retrieve the stored references (shape, paragraph, run, or cell).
    *   Look up the translated text using the element's ID in the `id_to_translation` dictionary: `translated_text = id_to_translation.get(element['id'], element['text'])`. Use the original text as a fallback if translation is missing.
    *   Replace the text in the presentation. For runs, set `element['run'].text = translated_text`. For paragraphs, clear existing runs and add a new run with the translated text: `element['paragraph'].clear()`, `new_run = element['paragraph'].add_run()`, `new_run.text = translated_text`. For table cells, clear and add run similarly.
9.  **Save Presentation:**
    *   Save the modified presentation: `new_prs.save(output_path)`.
10. **Error Handling:**
    *   Don't use any error handling. Let errors crash the script with a traceback.

# Libraries to use

- `python-pptx` (available on PyPI) to read and write PowerPoint files
- Simon Willison's `llm` library (available on PyPI) to invoke the language model
