# Goal

Create a Python script which
- reads a PowerPoint `.pptx` file
- duplicates each page: `[p1, p2, p3]` becomes `[p1, p1copy, p2, p2copy, p3, p3copy]`
- translates all text elements on each `p?copy` page from Finnish to English, preserving run-level formatting
- saves the presentation as a new `.pptx` file with a different name

# Key Design Decisions & Clarifications

Based on review and discussion, the following clarifications and design choices have been confirmed:

1.  **Formatting Preservation:** Text translation will occur at the `run` level. The translated text will replace the original text within its existing `run` object (`run.text = translated_text`). This approach is intended to preserve the individual formatting (font, size, color, bold, italic, etc.) of each text run.
2.  **Error Handling:** For this version, the script will not implement extensive error handling. If errors occur (e.g., malformed LLM response, issues with file processing), the script is expected to crash and provide a traceback.
3.  **Scope of Translatable Text:** The initial version will focus on translating text found within standard shapes (text boxes) and tables on the slides. Other text elements (e.g., chart labels, SmartArt, slide notes) are out of scope for this version.
4.  **Language Model:** The script will use the default language model accessible via Simon Willison's `llm` library (`llm.get_model()`). No specific model or advanced configuration will be hardcoded.
5.  **Segmented Translation:** Text is extracted from individual runs. These (potentially short) text segments will be sent to the LLM with unique IDs. This approach is accepted, prioritizing formatting preservation over providing broader paragraph-level context to the LLM for each individual translation request. The LLM will still receive all text segments from the presentation in one batch.
6.  **Slide Copying:** The method for duplicating slides will follow the principles outlined in the Stack Overflow solution (https://stackoverflow.com/a/73954830/15770), which aims to preserve slide layouts and masters when copying to a new presentation object.

# Process Flow Diagram

```mermaid
graph TD
    A[Start] --> B(Parse CLI Args: input_file, output_file);
    B --> C{Open input_file with python-pptx};
    C --> D[Create new_prs Presentation object];
    D --> E{Iterate prs.slides};
    E -- For each slide --> F[Copy original slide to new_prs (1st copy)];
    F --> G[Copy original slide again to new_prs (2nd copy - for translation)];
    G --> H{Iterate shapes in 2nd copy (translated slide)};
    H -- For each shape --> I{Has Text Frame or Table?};
    I -- Yes --> J[Iterate paragraphs];
    J -- For each paragraph --> K[Iterate runs];
    K -- For each run --> L[Extract run.text, assign ID, store run object reference];
    L --> M(Collect all text_elements with IDs, original text, and run references);
    H -- No / Next Shape --> H;
    K -- Next Paragraph / Shape --> J;
    J -- Next Shape --> H;
    E -- All slides processed --> M;
    M --> N[Format text_elements for LLM: "ID1: text1\\nID2: text2..."];
    N --> O[Send to LLM for translation];
    O --> P[Receive LLM response];
    P --> Q{Parse LLM response into id_to_translation map};
    Q --> R{Iterate through stored text_elements};
    R -- For each element --> S[Get original run object reference];
    S --> T[Get translated_text for element's ID (fallback to original if missing)];
    T --> U[**Update element['run'].text = translated_text** (Preserves run's formatting)];
    U --> R;
    R -- All elements processed --> V[Save new_prs to output_file];
    V --> W[End];

    subgraph TextExtractionAndID [Step 4: Extract Text and Map IDs]
        direction LR
        L
    end

    subgraph TextReplacement [Step 8: Replace Text on Duplicated Slides]
        direction LR
        S --> T --> U
    end
```

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
    *   For each original slide, copy and append it to the `new_prs` **twice**, following the methodology from https://stackoverflow.com/a/73954830/15770 to ensure layouts and masters are handled appropriately. The first copy remains as is; the second copy is designated for translation.
4.  **Extract Text and Map IDs:**
    *   Initialize an empty list to store text elements: `text_elements = []`.
    *   Initialize a counter for unique IDs: `text_id_counter = 0`.
    *   Iterate through the *duplicated* slides in `new_prs` (i.e., the second copy of each original slide, intended for translation).
    *   For each such slide, iterate through its shapes: `for shape in duplicated_slide.shapes:`.
    *   If `shape.has_text_frame`:
        *   Iterate through paragraphs: `for paragraph in shape.text_frame.paragraphs:`.
            *   Iterate through runs: `for run in paragraph.runs:`.
                *   Extract text: `original_text = run.text`.
                *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                *   Store the data: `text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})`.
                *   Increment counter: `text_id_counter += 1`.
    *   If `shape.has_table`:
        *   Iterate through rows and columns: `for row in shape.table.rows: for cell in row.cells:`.
            *   Iterate through paragraphs in the cell's text frame: `for paragraph in cell.text_frame.paragraphs:`.
                *   Iterate through runs: `for run in paragraph.runs:`.
                    *   Extract text: `original_text = run.text`.
                    *   Assign a unique ID: `text_id = f"text_{text_id_counter}"`.
                    *   Store the data: `text_elements.append({'id': text_id, 'text': original_text, 'run_object': run})`.
                    *   Increment counter: `text_id_counter += 1`.
5.  **Prepare Data for Language Model:**
    *   Format the `text_elements` list into a single string. Each element should be on a new line, formatted as `ID: Original Text`.
    *   Example: `formatted_text = "\n".join([f"{item['id']}: {item['text']}" for item in text_elements])`.
6.  **Invoke Language Model:**
    *   Import the `llm` library.
    *   Define the prompt. The prompt should instruct the model to translate the text provided in the attachment from Finnish to English and return the translations in the format `ID: Translated Text\n`, preserving the original IDs.
    *   Call the language model using `llm.get_model().prompt(prompt, attachments=[llm.Attachment.from_text(formatted_text, filename="texts_to_translate.txt")])`. Capture the output.
7.  **Process Language Model Response:**
    *   Initialize an empty dictionary for translations: `id_to_translation = {}`.
    *   Parse the model's output string line by line.
    *   For each line, split it into ID and translated text (e.g., splitting at the first occurrence of `: `).
    *   Store the translation: `id_to_translation[id.strip()] = translated_text.strip()`.
8.  **Replace Text on Duplicated Slides:**
    *   Iterate through the `text_elements` list created in Step 4.
    *   For each element:
        *   Retrieve the `run_object` reference.
        *   Look up the translated text using the element's ID: `translated_text = id_to_translation.get(element['id'], element['text'])` (fallback to original text if translation is missing or ID not found).
        *   Update the text of the run object directly: `element['run_object'].text = translated_text`. This preserves the original formatting of the run.
9.  **Save Presentation:**
    *   Save the modified presentation: `new_prs.save(output_path)`.
10. **Error Handling:**
    *   As per design decision, no explicit error handling blocks (try-except) will be implemented. Script errors will result in a crash with a traceback.

# Libraries to use

- `python-pptx` (available on PyPI) to read and write PowerPoint files
- Simon Willison's `llm` library (available on PyPI) to invoke the language model
- `click` (available on PyPI) for command-line argument parsing
