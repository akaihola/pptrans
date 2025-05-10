# Goal

Create a Python script which
- reads a PowerPoint `.pptx` file
- duplicates each page: `[p1, p2, p3]` becomes `[p1, p1copy, p2, p2copy, p3, p3copy]`
- translates all text elements on each `p?copy` page from Finnish to English
- saves the presentation as a new `.pptx` file with a different name

# Details

- extract all atomic text elements into a list and make sure each item has an ID
- feed the list as an ASCII attachment to a language model
- use a prompt which asks the language model to translate each item while keeping the ID
- extract translated items from the language model response
- replace matching elements on `p?copy` pages with the translations

# Libraries to use

- Simon Willison's `llm` library (available on PyPI) to invoke the language model
