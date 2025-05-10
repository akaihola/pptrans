# Goal

Create a Python script which
- reads a PowerPoint `.pptx` file
- duplicates each page: `[p1, p2, p3]` becomes `[p1, p1copy, p2, p2copy, p3, p3copy]`
- translates all text elements on each `p?copy` page from Finnish to English, preserving run-level formatting (in `translate` mode)
- saves the presentation as a new `.pptx` file with a different name

# Key Design Decisions & Clarifications

Based on review and discussion, the following clarifications and design choices have been confirmed:

1.  **Formatting Preservation:** Text modification (translation or word reversal) will occur at the `run` level. The new text will replace the original text within its existing `run` object (`run.text = new_text`). This approach is intended to preserve the individual formatting (font, size, color, bold, italic, etc.) of each text run.
2.  **Error Handling:** For this version, the script will not implement extensive error handling beyond basic parsing of LLM responses. If other errors occur (e.g., issues with file processing), the script is expected to crash and provide a traceback.
3.  **Scope of Modifiable Text:** The script will focus on text found within standard shapes (text boxes) and tables on the slides. Other text elements (e.g., chart labels, SmartArt, slide notes) are out of scope.
4.  **Language Model (for `translate` mode):** The script will use the default language model accessible via Simon Willison's `llm` library (`llm.get_model()`). No specific model or advanced configuration will be hardcoded for the `translate` mode.
5.  **Segmented Translation (for `translate` mode):** Text is extracted from individual runs. These (potentially short) text segments will be sent to the LLM with unique IDs. This approach is accepted, prioritizing formatting preservation over providing broader paragraph-level context to the LLM for each individual translation request. The LLM will still receive all text segments from the presentation in one batch.
6.  **Slide Copying:** The method for duplicating slides will follow the principles outlined in the Stack Overflow solution (https://stackoverflow.com/a/73954830/15770), which aims to preserve slide layouts and masters when copying to a new presentation object.

## New: Dry-Run Modes

To facilitate debugging and testing without incurring LLM costs or delays, two dry-run modes will be implemented:

1.  **`duplicate-only` Mode:**
    *   The script will perform the initial slide duplication (creating two copies of each original slide).
    *   All subsequent text extraction, LLM interaction, and text modification steps will be skipped.
    *   The output presentation will contain two identical, untranslated copies of each original slide.
2.  **`reverse-words` Mode:**
    *   The script will duplicate slides as usual.
    *   Text will be extracted from the second copy of each slide.
    *   Instead of LLM translation, a local function will reverse each word in the extracted text strings (e.g., "Hei maailma" becomes "ieH amliaam"). Punctuation attached to words will be reversed along with the word (e.g., "Hello, world!" becomes ",olleH !dlrow"). This simple reversal is deemed acceptable for a debugging tool.
    *   The original text in the second-copy slides will be replaced with these word-reversed versions.

These modes will be selectable via a new command-line option (`--mode`).

# Process Flow Diagram

```mermaid
graph TD
    A[Start] --> B(Parse CLI Args: input_file, output_file, mode);
    B --> C{Open input_file with python-pptx};
    C --> D[Create new_prs Presentation object];
    D --> E{Iterate prs.slides};
    E -- For each slide --> F[Copy original slide to new_prs (1st copy)];
    F --> G[Copy original slide again to new_prs (2nd copy - for modification)];
    G --> E;
    E -- All slides processed --> H{What is the --mode?};

    H -- Mode: 'translate' --> I[Extract Text from 2nd Copies];
        I --> J[Format Text & Send to LLM];
        J --> K[Parse LLM's Translated Text];
        K --> L[Replace Text in 2nd Copies with Translations];
        L --> Z[Save new_prs to output_file];

    H -- Mode: 'reverse-words' --> N[Extract Text from 2nd Copies];
        N --> O[Define/Use local reverse_individual_words function];
        O --> P[Apply reverse_individual_words to Each Text Element];
        P --> Q[Replace Text in 2nd Copies with Reversed-Word Text];
        Q --> Z;

    H -- Mode: 'duplicate-only' --> R[Skip Text Processing];
        R --> Z;

    Z --> W[End];

    subgraph TextExtractionAndID [Step 4a: Extract Text and Map IDs (for translate/reverse-words)]
        direction LR
        ExtractRunText["Extract run.text, assign ID, store run object reference"]
    end

    subgraph TextReplacement [Conditional Text Replacement]
        direction LR
        L_Replace["Update run.text = translated_text (translate mode)"]
        Q_Replace["Update run.text = reversed_word_text (reverse-words mode)"]
    end
```

# Details

The script should perform the following steps:

1.  **Parse Command-Line Arguments:**
    *   Use the `click` package to handle command-line arguments.
    *   Define arguments for the input `.pptx` file path (required, `type=click.Path(exists=True, dir_okay=False)`) and the output `.pptx` file path (required, `type=click.Path(dir_okay=False)`).
    *   Add a `--mode` option using `click.option('--mode', type=click.Choice(['translate', 'duplicate-only', 'reverse-words'], case_sensitive=False), default='translate', show_default=True, help='Operation mode for the script.')`.

2.  **Open Presentation:**
    *   Import the `Presentation` class from `pptx`.
    *   Load the input presentation using `prs = Presentation(input_path)`.
3.  **Iterate and Duplicate Slides:**
    *   Create a new presentation object to build the modified presentation: `new_prs = Presentation()`.
    *   Iterate through the slides of the original presentation: `for i, slide in enumerate(prs.slides):`.
    *   For each original slide, copy and append it to the `new_prs` **twice**. The first copy remains as is; the second copy is designated for modification based on the selected mode. Use a robust slide copying function (e.g., adapted from the SO link or the existing `copy_slide` in the script).

4.  **Conditional Processing Based on Mode:**

    *   A list `slides_for_text_extraction` should be populated with references to the second copies of slides if the mode is `translate` or `reverse-words`.

    *   **If `mode == 'duplicate-only'`:**
        *   Proceed directly to Step 9 (Save Presentation). No text extraction or modification is performed on the duplicated slides.

    *   **If `mode == 'translate'` OR `mode == 'reverse-words'`:**
        *   **4a. Extract Text and Map IDs (from `slides_for_text_extraction`):**
            *   Initialize an empty list to store text elements: `text_elements = []`.
            *   Initialize a counter for unique IDs: `text_id_counter = 0`.
            *   Iterate through the slides in `slides_for_text_extraction`.
            *   For each such slide, iterate through its shapes: `for shape in slide_to_extract.shapes:`.
            *   If `shape.has_text_frame`:
                *   Iterate through paragraphs: `for paragraph in shape.text_frame.paragraphs:`.
                    *   Iterate through runs: `for run in paragraph.runs:`.
                        *   Extract text: `original_text = run.text.strip()`.
                        *   If `original_text`:
                            *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                            *   Store the data: `text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})`.
                            *   Increment counter: `text_id_counter += 1`.
            *   If `shape.has_table`:
                *   Iterate through rows and cells: `for row in shape.table.rows: for cell in row.cells:`.
                    *   Iterate through paragraphs in the cell's text frame: `for paragraph in cell.text_frame.paragraphs:`.
                        *   Iterate through runs: `for run in paragraph.runs:`.
                            *   Extract text: `original_text = run.text.strip()`.
                            *   If `original_text`:
                                *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                                *   Store the data: `text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})`.
                                *   Increment counter: `text_id_counter += 1`.
        *   If no text elements are found (`if not text_elements:`), print a message and proceed to Step 9.

    *   **If `mode == 'translate'`:**
        *   **5. Prepare Data for Language Model:**
            *   Format the `text_elements` list into a single string. Each element should be on a new line, formatted as `ID: Original Text`.
            *   Example: `formatted_text = "\n".join([f"{item['id']}: {item['text']}" for item in text_elements])`.
        *   **6. Invoke Language Model:**
            *   Import the `llm` library.
            *   Define the prompt. The prompt should instruct the model to translate the text provided from Finnish to English and return the translations in the format `ID: Translated Text\n`, preserving the original IDs.
            *   Call the language model (e.g., `response = llm.get_model().prompt(prompt, fragments=[formatted_text])`). Capture the response text: `translated_text_response = response.text()`.
        *   **7. Process Language Model Response:**
            *   Initialize an empty dictionary for translations: `id_to_translation = {}`.
            *   Parse the `translated_text_response` line by line.
            *   For each line, split it into ID and translated text (e.g., splitting at the first occurrence of `: `).
            *   Store the translation: `id_to_translation[id.strip()] = translated_text.strip()`. Handle potential parsing errors gracefully (e.g., by printing a warning and skipping the line).
        *   **8. Replace Text with Translations:**
            *   Iterate through the `text_elements` list.
            *   For each element, retrieve the `run_object`.
            *   Look up the translated text using the element's ID: `translated_text = id_to_translation.get(element['id'], element['text'])` (fallback to original text if translation is missing or ID not found).
            *   Update the text of the run object directly: `element['run_object'].text = translated_text`.

    *   **If `mode == 'reverse-words'`:**
        *   **(No LLM data prep needed for step 5b)**
        *   **6b. Apply Word Reversal:**
            *   Define a helper function `def reverse_individual_words(text_string): words = text_string.split(' '); reversed_words = [word[::-1] for word in words]; return ' '.join(reversed_words)`.
            *   Initialize `id_to_modified_text = {}`.
            *   Iterate through `text_elements`:
                *   `modified_text = reverse_individual_words(item['text'])`
                *   `id_to_modified_text[item['id']] = modified_text`
        *   **(No LLM response processing needed for step 7b)**
        *   **8b. Replace Text with Reversed-Word Text:**
            *   Iterate through the `text_elements` list.
            *   For each element, retrieve the `run_object`.
            *   Look up the modified text using the element's ID: `modified_text = id_to_modified_text.get(element['id'], element['text'])` (fallback to original text if ID missing).
            *   Update the text of the run object directly: `element['run_object'].text = modified_text`.

9.  **Save Presentation:**
    *   Save the modified presentation: `new_prs.save(output_path)`.
    *   Print a confirmation message indicating the mode used and the output file path (e.g., `click.echo(f"Presentation saved in '{mode}' mode to: {output_path}")`).
10. **Error Handling:**
    *   As per design decision, no explicit error handling blocks (try-except) will be implemented for issues beyond basic parsing of LLM response. Script errors will result in a crash with a traceback.

# Libraries to use

- `python-pptx` (available on PyPI) to read and write PowerPoint files
- Simon Willison's `llm` library (available on PyPI) to invoke the language model (for `translate` mode)
- `click` (available on PyPI) for command-line argument parsing
- `copy` (standard library) for `deepcopy` used in the slide copying logic.
