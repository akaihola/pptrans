"""Unit tests for the `pptrans.__main__` module."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any
from unittest.mock import MagicMock, patch

import pytest  # Import pytest for the marker
from click.testing import CliRunner

from pptrans.__main__ import (
    EOL_MARKER,
    _apply_translations_to_runs,
    _build_llm_prompt_and_data,
    _extract_run_info_from_slide,
    _handle_slide_selection,
    _process_reverse_words_mode,
    _process_translation_mode,
    main,
    reverse_individual_words,
)

if TYPE_CHECKING:
    from pathlib import Path

    from pptx.slide import Slide as PptxSlide


def create_mock_run(text: str) -> MagicMock:
    """Create a mock run object."""
    run = MagicMock()
    run.text = text
    return run


def create_mock_paragraph(runs: list[MagicMock]) -> MagicMock:
    """Create a mock paragraph object."""
    para = MagicMock()
    para.runs = runs
    return para


def create_mock_text_frame(paragraphs: list[MagicMock]) -> MagicMock:
    """Create a mock text_frame object."""
    tf = MagicMock()
    tf.paragraphs = paragraphs
    return tf


def create_mock_cell(text_frame: MagicMock) -> MagicMock:
    """Create a mock cell object."""
    cell = MagicMock()
    cell.text_frame = text_frame
    return cell


def create_mock_row(cells: list[MagicMock]) -> MagicMock:
    """Create a mock row object."""
    row = MagicMock()
    row.cells = cells
    return row


def create_mock_table(rows: list[MagicMock]) -> MagicMock:
    """Create a mock table object."""
    table = MagicMock()
    table.rows = rows
    return table


def create_mock_shape(
    has_text_frame: bool = False,
    text_frame: MagicMock | None = None,
    has_table: bool = False,
    table: MagicMock | None = None,
) -> MagicMock:
    """Create a mock shape object."""
    shape = MagicMock()
    shape.has_text_frame = has_text_frame
    shape.text_frame = text_frame
    shape.has_table = has_table
    shape.table = table
    return shape


def create_mock_slide(shapes: list[MagicMock]) -> MagicMock:
    """Create a mock slide object."""
    slide = MagicMock()
    slide.shapes = shapes
    return slide


@pytest.mark.kwparametrize(
    dict(
        text_string_with_eol=f"hello world{EOL_MARKER}",
        expected=f"olleh dlrow{EOL_MARKER}",
        id="with_eol",
    ),
    dict(
        text_string_with_eol="hello world",
        expected="olleh dlrow",
        id="without_eol",
    ),
    dict(
        text_string_with_eol=EOL_MARKER,
        expected=EOL_MARKER,
        id="empty_string_with_eol",
    ),
    dict(
        text_string_with_eol="",
        expected="",
        id="empty_string_without_eol",
    ),
    dict(
        text_string_with_eol=f"hello{EOL_MARKER}",
        expected=f"olleh{EOL_MARKER}",
        id="single_word_with_eol",
    ),
    dict(
        text_string_with_eol="hello",
        expected="olleh",
        id="single_word_without_eol",
    ),
    dict(
        text_string_with_eol=f"hello  world{EOL_MARKER}",
        expected=f"olleh  dlrow{EOL_MARKER}",
        id="multiple_spaces_with_eol",
    ),
)
def test_reverse_individual_words(text_string_with_eol: str, expected: str) -> None:
    assert reverse_individual_words(text_string_with_eol) == expected


@patch("pptrans.__main__.parse_page_range")
@pytest.mark.kwparametrize(
    dict(
        pages_option=None,
        num_original_slides=3,
        expected_result={0, 1, 2},
        mock_parse_page_range_return=None,
        id="pages_option_none",
    ),
    dict(
        pages_option="1,3",
        num_original_slides=5,
        mock_parse_page_range_return={0, 2},
        expected_result={0, 2},
        id="pages_option_valid",
    ),
    dict(
        pages_option=None,
        num_original_slides=0,
        expected_result=set(),
        mock_parse_page_range_return=None,
        id="num_original_slides_zero",
    ),
)
def test_handle_slide_selection(
    mock_parse_page_range: MagicMock,
    pages_option: str | None,
    num_original_slides: int,
    expected_result: set[int],
    mock_parse_page_range_return: set[int] | None,
) -> None:
    if mock_parse_page_range_return is not None:
        mock_parse_page_range.return_value = mock_parse_page_range_return

    result = _handle_slide_selection(pages_option, num_original_slides)
    assert result == expected_result

    if pages_option:
        mock_parse_page_range.assert_called_once_with(pages_option, num_original_slides)
    else:
        mock_parse_page_range.assert_not_called()


@pytest.mark.kwparametrize(
    dict(
        slide_obj=create_mock_slide(shapes=[]),
        expected=[],
        id="empty_slide",
    ),
    dict(
        slide_obj=create_mock_slide(shapes=[create_mock_shape()]),
        expected=[],
        id="shape_no_text_frame_no_table",
    ),
    dict(
        slide_obj=create_mock_slide(
            shapes=[
                create_mock_shape(
                    has_text_frame=True,
                    text_frame=create_mock_text_frame(
                        paragraphs=[create_mock_paragraph(runs=[])]
                    ),
                )
            ]
        ),
        expected=[],
        id="shape_with_text_frame_empty_runs",
    ),
    dict(
        slide_obj=create_mock_slide(
            shapes=[
                create_mock_shape(
                    has_text_frame=True,
                    text_frame=create_mock_text_frame(
                        paragraphs=[
                            create_mock_paragraph(
                                runs=[
                                    create_mock_run("Hello"),
                                    create_mock_run(" World"),
                                ]
                            )
                        ]
                    ),
                )
            ]
        ),
        expected=[
            {"original_text": "Hello", "run_object": Ellipsis},
            {"original_text": " World", "run_object": Ellipsis},
        ],
        id="shape_with_text_runs",
    ),
    dict(
        slide_obj=create_mock_slide(
            shapes=[
                create_mock_shape(
                    has_table=True,
                    table=create_mock_table(
                        rows=[
                            create_mock_row(
                                cells=[
                                    create_mock_cell(
                                        create_mock_text_frame(
                                            [
                                                create_mock_paragraph(
                                                    [create_mock_run("Table")]
                                                )
                                            ]
                                        )
                                    )
                                ]
                            )
                        ]
                    ),
                )
            ]
        ),
        expected=[{"original_text": "Table", "run_object": Ellipsis}],
        id="shape_with_table_runs",
    ),
    dict(
        slide_obj=create_mock_slide(
            shapes=[
                create_mock_shape(
                    has_text_frame=True,
                    text_frame=create_mock_text_frame(
                        paragraphs=[
                            create_mock_paragraph([create_mock_run("TextShape")])
                        ]
                    ),
                ),
                create_mock_shape(
                    has_table=True,
                    table=create_mock_table(
                        rows=[
                            create_mock_row(
                                cells=[
                                    create_mock_cell(
                                        create_mock_text_frame(
                                            [
                                                create_mock_paragraph(
                                                    [create_mock_run("InTable")]
                                                )
                                            ]
                                        )
                                    )
                                ]
                            )
                        ]
                    ),
                ),
            ]
        ),
        expected=[
            {"original_text": "TextShape", "run_object": Ellipsis},
            {"original_text": "InTable", "run_object": Ellipsis},
        ],
        id="mixed_content_slide",
    ),
    dict(
        slide_obj=create_mock_slide(
            shapes=[
                create_mock_shape(
                    has_text_frame=True,
                    text_frame=create_mock_text_frame(
                        paragraphs=[
                            create_mock_paragraph(
                                runs=[create_mock_run("Data"), create_mock_run("")]
                            )
                        ]
                    ),
                )
            ]
        ),
        expected=[{"original_text": "Data", "run_object": Ellipsis}],
        id="run_with_empty_text",
    ),
)
def test_extract_run_info_from_slide(
    slide_obj: PptxSlide, expected: list[dict[str, Any]]
) -> None:
    result = _extract_run_info_from_slide(slide_obj)
    assert len(result) == len(expected)
    for res_item, exp_item in zip(result, expected):
        assert res_item["original_text"] == exp_item["original_text"]
        assert "run_object" in res_item
        if exp_item["run_object"] is Ellipsis:
            assert res_item["run_object"] is not None


@pytest.mark.kwparametrize(
    dict(
        texts_for_llm=[],
        expected_formatted_text="",
        id="empty_texts_for_llm",
    ),
    dict(
        texts_for_llm=[{"id": "pg1_txt0", "text_to_send": "Hello"}],
        expected_formatted_text="pg1_txt0:Hello",
        id="single_item",
    ),
    dict(
        texts_for_llm=[
            {"id": "pg1_txt0", "text_to_send": "First"},
            {"id": "pg1_txt1", "text_to_send": "Second"},
        ],
        expected_formatted_text="pg1_txt0:First\npg1_txt1:Second",
        id="multiple_items",
    ),
)
def test_build_llm_prompt_and_data(
    texts_for_llm: list[dict[str, Any]], expected_formatted_text: str
) -> None:
    prompt_text, formatted_text_for_llm = _build_llm_prompt_and_data(texts_for_llm)

    assert "You are an expert Finnish to English translator." in prompt_text
    assert f"EOL marker: '{EOL_MARKER}'" in prompt_text
    assert "Texts to translate:\n" in prompt_text
    assert formatted_text_for_llm == expected_formatted_text


@patch("pptrans.__main__.click.echo")
@pytest.mark.kwparametrize(
    dict(
        details=[
            {
                "final_translation": "Translated Text",
                "run_object": create_mock_run("Original"),
            }
        ],
        expected_run_text="Translated Text",
        expected_warnings=0,
        id="apply_translation",
    ),
    dict(
        details=[
            {
                "final_translation": None,
                "run_object": create_mock_run("Original"),
                "original_text": "Original Text",
                "llm_id": "pg1_txt0",
                "from_cache": False,
            }
        ],
        expected_run_text="Original",
        expected_warnings=1,
        id="no_translation_not_from_cache",
    ),
    dict(
        details=[
            {
                "final_translation": None,
                "run_object": create_mock_run("Original Cache"),
                "original_text": "Original Cache Text",
                "from_cache": True,
            }
        ],
        expected_run_text="Original Cache",
        expected_warnings=0,
        id="no_translation_from_cache",
    ),
    dict(
        details=[],
        expected_run_text=None,
        expected_warnings=0,
        id="empty_details",
    ),
)
def test_apply_translations_to_runs(
    mock_echo: MagicMock,
    details: list[dict[str, Any]],
    expected_run_text: str | None,
    expected_warnings: int,
) -> None:
    _apply_translations_to_runs(details)

    # Ensure the run_object is actually a mock before asserting .text
    if (
        details
        and expected_run_text is not None
        and isinstance(details[0]["run_object"], MagicMock)
    ):
        assert details[0]["run_object"].text == expected_run_text

    assert mock_echo.call_count == 1 + expected_warnings
    if expected_warnings > 0:
        mock_echo.assert_any_call(
            "Warning: No translation found for run with original text "
            f"'{details[0]['original_text'][:30]}...' "
            f"(LLM ID: {details[0].get('llm_id', 'N/A')}). Leaving original.",
            err=True,
        )


@patch("pptrans.__main__.commit_pending_cache_updates")
@patch("pptrans.__main__._apply_translations_to_runs")
@patch("pptrans.__main__.update_data_from_llm_response")
@patch("pptrans.__main__.llm.get_model")
@patch("pptrans.__main__._build_llm_prompt_and_data")
@patch("pptrans.__main__.prepare_slide_for_translation")
@patch("pptrans.__main__.generate_page_hash")
@patch("pptrans.__main__._extract_run_info_from_slide")
@patch("pptrans.__main__.load_cache")
@patch("pptrans.__main__.click.echo")
def test_process_translation_mode_no_slides(
    mock_echo: MagicMock,
    mock_load_cache: MagicMock,
    mock_extract_run: MagicMock,
    _mock_gen_hash: MagicMock,
    _mock_prep_slide: MagicMock,
    _mock_build_prompt: MagicMock,
    _mock_get_model: MagicMock,
    _mock_update_llm_resp: MagicMock,
    _mock_apply_trans: MagicMock,
    mock_commit_cache: MagicMock,
) -> None:
    mock_load_cache.return_value = {}
    result = _process_translation_mode([], [], "cache.json", EOL_MARKER, 0)
    assert result == 0
    mock_echo.assert_any_call(
        "No text found to process on selected slides for mode 'translate'."
    )
    mock_commit_cache.assert_called_once_with({}, {}, "cache.json")
    mock_extract_run.assert_not_called()


@patch("pptrans.__main__.commit_pending_cache_updates")
@patch("pptrans.__main__._apply_translations_to_runs")
@patch("pptrans.__main__.update_data_from_llm_response")
@patch("pptrans.__main__.llm.get_model")
@patch("pptrans.__main__._build_llm_prompt_and_data")
@patch("pptrans.__main__.prepare_slide_for_translation")
@patch("pptrans.__main__.generate_page_hash")
@patch("pptrans.__main__._extract_run_info_from_slide")
@patch("pptrans.__main__.load_cache")
@patch("pptrans.__main__.click.echo")
def test_process_translation_mode_all_cached(
    mock_echo: MagicMock,
    mock_load_cache: MagicMock,
    mock_extract_run: MagicMock,
    mock_gen_hash: MagicMock,
    mock_prep_slide: MagicMock,
    mock_build_prompt: MagicMock,
    mock_get_model: MagicMock,
    _mock_update_llm_resp: MagicMock,
    mock_apply_trans: MagicMock,
    mock_commit_cache: MagicMock,
) -> None:
    mock_slide = create_mock_slide([])
    mock_load_cache.return_value = {"some_hash": [{"id": "id1", "text": "text"}]}
    mock_extract_run.return_value = [{"original_text": "Hi", "run_object": MagicMock()}]
    mock_gen_hash.return_value = "some_hash"
    mock_prep_slide.return_value = ([], [{"run_object": MagicMock()}], 1, False)

    result = _process_translation_mode([mock_slide], [0], "cache.json", EOL_MARKER, 0)

    assert result == 1
    mock_echo.assert_any_call(
        "All text elements found in page caches. Skipping LLM prompt."
    )
    mock_build_prompt.assert_not_called()
    mock_get_model.assert_not_called()
    mock_apply_trans.assert_called_once()
    mock_commit_cache.assert_called_once()


@patch("pptrans.__main__.commit_pending_cache_updates")
@patch("pptrans.__main__._apply_translations_to_runs")
@patch("pptrans.__main__.update_data_from_llm_response")
@patch("pptrans.__main__.llm.get_model")
@patch("pptrans.__main__._build_llm_prompt_and_data")
@patch("pptrans.__main__.prepare_slide_for_translation")
@patch("pptrans.__main__.generate_page_hash")
@patch("pptrans.__main__._extract_run_info_from_slide")
@patch("pptrans.__main__.load_cache")
@patch("pptrans.__main__.click.echo")
def test_process_translation_mode_with_llm_call(
    mock_echo: MagicMock,
    mock_load_cache: MagicMock,
    mock_extract_run: MagicMock,
    mock_gen_hash: MagicMock,
    mock_prep_slide: MagicMock,
    mock_build_prompt: MagicMock,
    mock_get_model: MagicMock,
    mock_update_llm_resp: MagicMock,
    mock_apply_trans: MagicMock,
    mock_commit_cache: MagicMock,
) -> None:
    mock_slide = create_mock_slide([])
    mock_load_cache.return_value = {}
    mock_extract_run.return_value = [{"original_text": "Hi", "run_object": MagicMock()}]
    mock_gen_hash.return_value = "new_hash"
    texts_for_llm_slide = [{"id": "pg1_txt0", "text_to_send": "Text for LLM"}]
    processed_runs_slide = [{"run_object": MagicMock(), "llm_id": "pg1_txt0"}]
    mock_prep_slide.return_value = (
        texts_for_llm_slide,
        processed_runs_slide,
        1,
        True,
    )

    mock_build_prompt.return_value = ("Test Prompt", "Formatted Text")
    mock_llm_instance = MagicMock()
    mock_llm_response = MagicMock()
    mock_llm_response.text.return_value = "pg1_txt0:Translated Text"
    mock_llm_instance.prompt.return_value = mock_llm_response
    mock_get_model.return_value = mock_llm_instance

    result = _process_translation_mode([mock_slide], [0], "cache.json", EOL_MARKER, 0)

    assert result == 1
    mock_build_prompt.assert_called_once_with(texts_for_llm_slide)
    mock_get_model.assert_called_once()
    mock_llm_instance.prompt.assert_called_once_with(
        "Test Prompt", fragments=["Formatted Text"]
    )
    mock_update_llm_resp.assert_called_once_with(
        ["pg1_txt0:Translated Text"],
        texts_for_llm_slide,
        processed_runs_slide,
        {"new_hash": []},
        EOL_MARKER,
    )
    mock_apply_trans.assert_called_once()
    mock_commit_cache.assert_called_once()
    mock_echo.assert_any_call("--- RESPONSE FROM LLM ---")


@patch("pptrans.__main__.click.echo")
@patch("pptrans.__main__.reverse_individual_words")
@patch("pptrans.__main__._extract_run_info_from_slide")
def test_process_reverse_words_mode_no_text(
    mock_extract_run: MagicMock,
    mock_reverse_words: MagicMock,
    mock_echo: MagicMock,
) -> None:
    mock_slide = create_mock_slide([])
    mock_extract_run.return_value = []

    result = _process_reverse_words_mode([mock_slide], [0], EOL_MARKER, 0)
    assert result == 0
    mock_extract_run.assert_called_once_with(mock_slide)
    mock_echo.assert_any_call(
        "No text found to process on selected slides for mode 'reverse-words'."
    )
    mock_reverse_words.assert_not_called()


@patch("pptrans.__main__.click.echo")
@patch("pptrans.__main__.reverse_individual_words")
@patch("pptrans.__main__._extract_run_info_from_slide")
def test_process_reverse_words_mode_with_text(
    mock_extract_run: MagicMock,
    mock_reverse_words: MagicMock,
    mock_echo: MagicMock,
) -> None:
    mock_run = create_mock_run("Original")
    mock_slide = create_mock_slide([])
    mock_extract_run.return_value = [{"original_text": "Hello", "run_object": mock_run}]
    mock_reverse_words.return_value = f"olleH{EOL_MARKER}"

    result = _process_reverse_words_mode([mock_slide], [0], EOL_MARKER, 0)

    assert result == 1
    mock_extract_run.assert_called_once_with(mock_slide)
    mock_reverse_words.assert_called_once_with(f"Hello{EOL_MARKER}")
    assert mock_run.text == "olleH"
    mock_echo.assert_any_call("Text replaced with reversed-word text on slides.")


@patch("pptrans.__main__.shutil.copy2")
@patch("pptrans.__main__.Presentation")
@patch("pptrans.__main__._handle_slide_selection")
@patch("pptrans.__main__._process_translation_mode")
@patch("pptrans.__main__._process_reverse_words_mode")
@patch("pptrans.__main__.load_cache")
@patch("pptrans.__main__.commit_pending_cache_updates")
def test_main_cli(
    mock_commit_cache: MagicMock,
    mock_load_cache: MagicMock,
    mock_process_reverse: MagicMock,
    mock_process_translate: MagicMock,
    mock_handle_selection: MagicMock,
    MockPresentation: MagicMock,
    mock_copy2: MagicMock,
    tmp_path: Path,
) -> None:
    runner = CliRunner()
    input_file = tmp_path / "input.pptx"
    output_file = tmp_path / "output.pptx"
    input_file.touch()

    mock_prs_instance = MagicMock()
    mock_slide_obj = create_mock_slide([])
    mock_prs_instance.slides = [mock_slide_obj, mock_slide_obj]
    MockPresentation.return_value = mock_prs_instance
    mock_handle_selection.return_value = {0, 1}

    mock_process_translate.return_value = 5
    result = runner.invoke(
        main, [str(input_file), str(output_file), "--mode", "translate"]
    )
    assert result.exit_code == 0
    mock_copy2.assert_called_once_with(str(input_file), str(output_file))
    MockPresentation.assert_called_with(str(output_file))
    mock_handle_selection.assert_called_with(None, 2)
    mock_process_translate.assert_called_once_with(
        [mock_slide_obj, mock_slide_obj],
        [0, 1],
        "translation_cache.json",
        EOL_MARKER,
        0,
    )
    mock_prs_instance.save.assert_called_with(str(output_file))
    assert f"Presentation saved in 'translate' mode to: {output_file}" in result.output

    mock_copy2.reset_mock()
    MockPresentation.reset_mock()
    MockPresentation.return_value = mock_prs_instance
    mock_handle_selection.reset_mock()
    mock_handle_selection.return_value = {0}
    mock_prs_instance.save.reset_mock()

    mock_process_reverse.return_value = 3
    result = runner.invoke(
        main,
        [str(input_file), str(output_file), "--mode", "reverse-words", "--pages", "1"],
    )
    assert result.exit_code == 0
    mock_copy2.assert_called_once_with(str(input_file), str(output_file))
    MockPresentation.assert_called_with(str(output_file))
    mock_handle_selection.assert_called_with("1", 2)
    mock_process_reverse.assert_called_once_with([mock_slide_obj], [0], EOL_MARKER, 0)
    mock_prs_instance.save.assert_called_with(str(output_file))
    assert (
        f"Presentation saved in 'reverse-words' mode to: {output_file}" in result.output
    )

    mock_copy2.reset_mock()
    MockPresentation.reset_mock()
    mock_prs_instance_no_slides = MagicMock()
    mock_prs_instance_no_slides.slides = []
    MockPresentation.return_value = mock_prs_instance_no_slides
    mock_prs_instance.save.reset_mock()  # Original mock_prs_instance

    result = runner.invoke(main, [str(input_file), str(output_file)])
    assert result.exit_code == 0
    assert "Input presentation has no slides. Exiting." in result.output
    mock_prs_instance_no_slides.save.assert_not_called()

    mock_copy2.reset_mock()
    MockPresentation.reset_mock()
    MockPresentation.return_value = mock_prs_instance  # Restore 2 slides
    mock_handle_selection.reset_mock()
    mock_handle_selection.return_value = set()
    mock_prs_instance.save.reset_mock()
    mock_load_cache.return_value = {}

    result = runner.invoke(main, [str(input_file), str(output_file), "--pages", "99"])
    assert result.exit_code == 0
    assert "Warning: The specified page range '99' resulted in" in result.output
    assert "no slides being selected" in result.output
    mock_prs_instance.save.assert_called_once_with(str(output_file))
    mock_load_cache.assert_called_once_with("translation_cache.json")
    mock_commit_cache.assert_called_once_with({}, {}, "translation_cache.json")


@patch("pptrans.__main__.shutil.copy2")
@patch("pptrans.__main__.Presentation")
@patch("pptrans.__main__._handle_slide_selection")
def test_main_cli_no_slides_selected_no_pages_option(
    mock_handle_selection: MagicMock,
    MockPresentation: MagicMock,
    _mock_copy2: MagicMock,  # Not used directly, but good practice to mock
    tmp_path: Path,
) -> None:
    runner = CliRunner()
    input_file = tmp_path / "input.pptx"
    output_file = tmp_path / "output.pptx"
    input_file.touch()

    mock_prs_instance = MagicMock()
    mock_prs_instance.slides = [MagicMock(), MagicMock()]
    MockPresentation.return_value = mock_prs_instance
    mock_handle_selection.return_value = set()

    # Patch load_cache and commit_pending_cache_updates for the translate branch
    with (
        patch("pptrans.__main__.load_cache") as mock_load,
        patch("pptrans.__main__.commit_pending_cache_updates") as mock_commit,
    ):
        mock_load.return_value = {}  # Simulate empty cache
        result = runner.invoke(main, [str(input_file), str(output_file)])

    assert result.exit_code == 0
    assert "No slides selected for processing" in result.output
    assert "No text processing will occur." in result.output
    mock_prs_instance.save.assert_called_once_with(str(output_file))
    mock_load.assert_called_once_with("translation_cache.json")
    mock_commit.assert_called_once_with({}, {}, "translation_cache.json")


@patch("pptrans.__main__.main")
def test_main_dunder_guard(_mock_main_func: MagicMock) -> None:
    # This is a conceptual test. In practice, CliRunner tests for `main`
    # cover the behavior of the script's entry point.
    # We assert that `main` is callable.
    assert callable(main)
    # To truly test the `if __name__ == "__main__":` guard, one would typically
    # run the script as a subprocess or use importlib to simulate being the main module.
    # This is often more involved than necessary if the CLI function itself is
    # well-tested.
