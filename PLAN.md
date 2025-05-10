# Plan to Display LLM Prompt and Response in `translate_pptx.py`

**Objective:** Modify the `translate_pptx.py` script to display the full prompt sent to the Language Model (LLM) and the raw response received from it directly in the terminal when the script is run in `translate` mode.

**Background:**

The script `translate_pptx.py` uses Simon Willison's `llm` library. The key parts for LLM interaction are:

1.  **Prompt Construction (Lines 195-213 of `translate_pptx.py`):**
    *   `formatted_text` (lines 195-197): A string containing all unique text elements from the PowerPoint slides, each prefixed with an ID (e.g., `text_0: Hei maailma`).
    *   `prompt` variable (lines 198-213): Contains the main instructions for the LLM on how to perform the translation and format the output.
    *   The LLM is called via: `model.prompt(prompt, fragments=[formatted_text])` (line 216). Here, `prompt` is the system prompt/main instruction, and `formatted_text` is passed as `fragments` (user input data).

2.  **Response Retrieval (Line 217 of `translate_pptx.py`):**
    *   The raw text response from the LLM is obtained using `response.text()` and stored in the `translated_text_response` variable.

**Proposed Modifications:**

The script will be modified to print these values using `click.echo()`.

1.  **Display the Full Prompt Sent to the LLM:**
    *   **Location:** Insert code immediately before the line `click.echo("Sending text to LLM for translation...")` (currently line 214).
    *   **Action:** Add `click.echo()` statements to print:
        *   A header: `--- PROMPT TO LLM ---`
        *   A sub-header: `System/Instruction Prompt:`
        *   The content of the `prompt` variable.
        *   A sub-header: `Data Fragments:`
        *   The content of the `formatted_text` variable.
        *   A footer: `--- END OF PROMPT ---`

2.  **Display the Raw Response from the LLM:**
    *   **Location:** Insert code immediately after the line `translated_text_response = response.text()` (currently line 217) and before `click.echo("Received translation from LLM.")` (currently line 218).
    *   **Action:** Add `click.echo()` statements to print:
        *   A header: `--- RESPONSE FROM LLM ---`
        *   The content of the `translated_text_response` variable.
        *   A footer: `--- END OF RESPONSE ---`

**Visual Representation (Mermaid Diagram):**

```mermaid
graph TD
    A[Original Script Flow] --> B{Mode == 'translate'?};
    B -- Yes --> C[Define `formatted_text` (lines 195-197)];
    C --> D[Define `prompt` (lines 198-213)];
    D --> E[Current: `click.echo("Sending text...")` (line 214)];
    E --> F[Call LLM: `model.prompt(prompt, fragments=[formatted_text])` (line 216)];
    F --> G[Get `translated_text_response` (line 217)];
    G --> H[Current: `click.echo("Received translation...")` (line 218)];
    H --> I[Process response];

    subgraph "Proposed Modifications"
        D --> D_mod1["NEW: click.echo('--- PROMPT TO LLM ---')"];
        D_mod1 --> D_mod2["click.echo('System/Instruction Prompt:')"];
        D_mod2 --> D_mod3["click.echo(prompt)"];
        D_mod3 --> D_mod4["click.echo('Data Fragments:')"];
        D_mod4 --> D_mod5["click.echo(formatted_text)"];
        D_mod5 --> D_mod6["click.echo('--- END OF PROMPT ---')"];
        D_mod6 --> E;

        G --> G_mod1["NEW: click.echo('--- RESPONSE FROM LLM ---')"];
        G_mod1 --> G_mod2["click.echo(translated_text_response)"];
        G_mod2 --> G_mod3["click.echo('--- END OF RESPONSE ---')"];
        G_mod3 --> H;
    end
```

This plan will ensure that the complete prompt (instructions + data) and the raw response are clearly printed to the terminal when the script is run in `translate` mode.
